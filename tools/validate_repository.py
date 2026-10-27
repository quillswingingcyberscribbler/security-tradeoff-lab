#!/usr/bin/env python3
"""Validate a workbook repository without modifying or recalculating its files.

Python 3 standard library only. Checks OOXML packages and relationships, saved
formula results, local Markdown links, and SHA-256 manifest entries. It does
not execute Excel, macros, external links, or workbook formulas.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import unquote, urlsplit
from zipfile import BadZipFile, ZipFile

SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
SKIP_DIRS = {".git", ".venv", "venv", "node_modules", "__pycache__"}
REQUIRED_FILES = ("README.md", "CITATION.cff", "assets/logo.png", "validation/manifest.json")


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def add_issue(issues: list, severity: str, code: str, path: str, message: str, **details) -> None:
    issue = {"severity": severity, "code": code, "path": path, "message": message}
    if details:
        issue["details"] = details
    issues.append(issue)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def safe_local_path(root: Path, relative: str) -> Path | None:
    if not relative or "\\" in relative or Path(relative).is_absolute():
        return None
    candidate = (root / relative).resolve()
    try:
        candidate.relative_to(root.resolve())
    except ValueError:
        return None
    return candidate


def inspect_ooxml(path: Path, root: Path, issues: list, check_relationships: bool = True) -> dict:
    label = path.relative_to(root).as_posix()
    stats = {"path": label, "xml_parts": 0, "formula_cells": 0,
             "missing_or_empty_numeric_caches": 0, "cached_blank_strings": 0,
             "cached_error_cells": 0, "broken_formula_references": 0,
             "external_hyperlinks": 0, "macros": False, "external_workbook_links": False}
    try:
        with ZipFile(path) as archive:
            names = archive.namelist()
            name_set = set(names)
            if len(names) != len(name_set):
                add_issue(issues, "error", "zip_duplicate_members", label,
                          "The ZIP package contains duplicate member names.")
            unsafe = [name for name in names if name.startswith("/") or "\\" in name
                      or posixpath.normpath(name).startswith("../")]
            if unsafe:
                add_issue(issues, "error", "zip_unsafe_member", label,
                          "The package contains non-canonical or escaping member paths.", examples=unsafe[:5])
            broken_crc = archive.testzip()
            if broken_crc:
                add_issue(issues, "error", "zip_crc_failure", label,
                          "A ZIP member failed its CRC check.", member=broken_crc)
            main_part = "xl/workbook.xml" if path.suffix.lower() == ".xlsx" else "word/document.xml"
            for required in ("[Content_Types].xml", "_rels/.rels", main_part):
                if required not in name_set:
                    add_issue(issues, "error", "ooxml_required_part", label,
                              "A required OOXML part is missing.", member=required)
            stats["macros"] = any("vbaproject" in name.lower() for name in names)
            if stats["macros"]:
                add_issue(issues, "error", "unexpected_macro_part", label,
                          "A VBA project is present in a file labelled as macro-free .xlsx or .docx.")
            stats["external_workbook_links"] = any(name.startswith("xl/externalLinks/") for name in names)
            if stats["external_workbook_links"]:
                add_issue(issues, "warning", "external_workbook_dependency", label,
                          "The workbook contains external-workbook link parts; linked data was not fetched.")
            if "xl/connections.xml" in name_set:
                add_issue(issues, "warning", "data_connections", label,
                          "The workbook contains data connections; refresh and availability were not tested.")
            cache_errors, bad_refs, missing_targets = [], [], []
            for name in names:
                if not (name.endswith(".xml") or name.endswith(".rels")):
                    continue
                try:
                    tree = ET.fromstring(archive.read(name))
                except (ET.ParseError, ValueError) as exc:
                    add_issue(issues, "error", "xml_parse_failure", label,
                              "An XML part is not well formed.", member=name, reason=str(exc))
                    continue
                stats["xml_parts"] += 1
                if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                    for cell in tree.iter():
                        if local_name(cell.tag) != "c":
                            continue
                        children = {local_name(child.tag): child for child in cell}
                        formula, value = children.get("f"), children.get("v")
                        if cell.attrib.get("t") == "e":
                            cache_errors.append({"part": name, "cell": cell.attrib.get("r", "?"),
                                                 "error": value.text if value is not None else "unknown"})
                        if formula is not None:
                            stats["formula_cells"] += 1
                            if value is None or value.text is None:
                                if value is not None and cell.attrib.get("t") == "str":
                                    stats["cached_blank_strings"] += 1
                                else:
                                    stats["missing_or_empty_numeric_caches"] += 1
                if name.startswith("xl/"):
                    for element in tree.iter():
                        if local_name(element.tag) not in {"f", "definedName", "calculatedColumnFormula", "totalsRowFormula"}:
                            continue
                        formula_text = element.text or ""
                        # '#REF!' inside a quoted explanation is ordinary text.
                        unquoted = re.sub(r'"(?:[^"]|"")*"', '""', formula_text)
                        if "#REF!" in unquoted.upper():
                            bad_refs.append({"part": name, "name": element.attrib.get("name", ""),
                                             "formula_excerpt": formula_text[:180]})
                    if name == "xl/workbook.xml":
                        for element in tree.iter():
                            if local_name(element.tag) == "calcPr" and element.attrib.get("calcMode") == "manual":
                                add_issue(issues, "warning", "manual_calculation", label,
                                          "Workbook calculation mode is manual; saved results may become stale after input changes.")
                if name.endswith(".rels"):
                    ids = []
                    for relationship in tree:
                        if local_name(relationship.tag) != "Relationship":
                            continue
                        ids.append(relationship.attrib.get("Id"))
                        target = relationship.attrib.get("Target")
                        relation_type = relationship.attrib.get("Type", "")
                        if not target:
                            missing_targets.append({"part": name, "target": "<missing>"})
                            continue
                        if relationship.attrib.get("TargetMode") == "External":
                            if relation_type.endswith("/hyperlink"):
                                stats["external_hyperlinks"] += 1
                            continue
                        if not check_relationships:
                            continue
                        split = urlsplit(target)
                        if split.scheme or split.netloc:
                            missing_targets.append({"part": name, "target": target,
                                                    "reason": "External URI lacks TargetMode=External"})
                            continue
                        target_path = unquote(split.path)
                        if not target_path:
                            continue  # Fragment-only relationship.
                        if name == "_rels/.rels":
                            base = ""
                        else:
                            base = posixpath.dirname(posixpath.dirname(name))
                        resolved = posixpath.normpath(target_path.lstrip("/") if target_path.startswith("/")
                                                     else posixpath.join(base, target_path))
                        if resolved.startswith("../") or resolved not in name_set:
                            missing_targets.append({"part": name, "target": target, "resolved": resolved})
                    if len(ids) != len(set(ids)) or None in ids:
                        add_issue(issues, "error", "relationship_ids", label,
                                  "A relationships part has duplicate or missing relationship IDs.", member=name)
            stats["cached_error_cells"] = len(cache_errors)
            stats["broken_formula_references"] = len(bad_refs)
            if cache_errors:
                add_issue(issues, "error", "cached_excel_errors", label,
                          f"Found {len(cache_errors)} saved Excel error cell(s).", examples=cache_errors[:8])
            if bad_refs:
                add_issue(issues, "error", "broken_formula_references", label,
                          f"Found {len(bad_refs)} formula/name definition(s) containing an unquoted #REF!.", examples=bad_refs[:8])
            if missing_targets:
                add_issue(issues, "error", "broken_internal_relationships", label,
                          f"Found {len(missing_targets)} missing or invalid internal relationship target(s).", examples=missing_targets[:8])
            if stats["missing_or_empty_numeric_caches"]:
                add_issue(issues, "warning", "formula_caches_unavailable", label,
                          f"{stats['missing_or_empty_numeric_caches']} formula cell(s) lack usable saved numeric/text results. "
                          "This is not proof of a calculation failure; open and recalculate in the intended spreadsheet application.")
    except (BadZipFile, OSError, RuntimeError, ValueError, NotImplementedError) as exc:
        add_issue(issues, "error", "ooxml_open_failure", label, "The OOXML package could not be read.", reason=str(exc))
    return stats


def markdown_targets(text: str):
    """Yield common inline/reference link targets, excluding code examples.

    This deliberately does not claim a complete CommonMark implementation.
    Supports angle-bracket targets, escaped punctuation, and balanced brackets
    in filenames such as workbook(2).xlsx. Fragment existence is not checked.
    """
    lines = []
    fence = None
    for line in text.splitlines():
        match = re.match(r"^\s{0,3}(`{3,}|~{3,})", line)
        if match:
            marker = match.group(1)
            if fence is None:
                fence = marker
            elif marker[0] == fence[0] and len(marker) >= len(fence):
                fence = None
            lines.append("")
        elif fence is None:
            lines.append(re.sub(r"(`+).*?\1", "", line))
        else:
            lines.append("")
    clean = "\n".join(lines)
    for match in re.finditer(r"^\s{0,3}\[(?!\^)[^\]\n]+\]:\s*(?:<([^>\n]+)>|(\S+))", clean, re.MULTILINE):
        yield match.group(1) or match.group(2)
    for match in re.finditer(r"\]\(", clean):
        index = match.end()
        while index < len(clean) and clean[index].isspace():
            index += 1
        if index >= len(clean):
            continue
        if clean[index] == "<":
            end = clean.find(">", index + 1)
            if end >= 0:
                yield clean[index + 1:end]
            continue
        target, depth = [], 0
        while index < len(clean):
            char = clean[index]
            if char == "\\" and index + 1 < len(clean):
                target.append(clean[index + 1]); index += 2; continue
            if char == "(":
                depth += 1
            elif char == ")":
                if depth == 0:
                    break
                depth -= 1
            if char.isspace() and depth == 0:
                break
            target.append(char); index += 1
        if target:
            yield "".join(target)
    for match in re.finditer(r"<(?:img|a)\b[^>]*\b(?:src|href)\s*=\s*['\"]([^'\"]+)['\"]", clean, re.IGNORECASE):
        yield match.group(1)


def inspect_markdown(path: Path, root: Path, issues: list) -> int:
    label = path.relative_to(root).as_posix()
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as exc:
        add_issue(issues, "error", "markdown_read_failure", label, "Markdown file is not readable UTF-8.", reason=str(exc))
        return 0
    checked = 0
    for target in set(markdown_targets(text)):
        split = urlsplit(target)
        if split.scheme or split.netloc or not split.path:
            continue
        decoded = unquote(split.path)
        candidate = (root / decoded.lstrip("/") if decoded.startswith("/") else path.parent / decoded).resolve()
        checked += 1
        try:
            candidate.relative_to(root)
        except ValueError:
            add_issue(issues, "error", "markdown_link_escapes_repo", label,
                      "A relative Markdown link points outside this standalone repository.", target=target)
            continue
        if not candidate.exists():
            add_issue(issues, "error", "markdown_local_link_missing", label,
                      "A local Markdown/image link has no matching file or directory.", target=target)
    return checked


def inspect_manifest(root: Path, issues: list, artifacts: list[Path]) -> int:
    manifest_path = root / "validation/manifest.json"
    if not manifest_path.is_file():
        return 0  # Required-file check already reports this.
    try:
        data = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        add_issue(issues, "error", "manifest_parse_failure", "validation/manifest.json",
                  "Manifest is not readable JSON.", reason=str(exc))
        return 0
    if not isinstance(data, dict) or not isinstance(data.get("files"), list):
        add_issue(issues, "error", "manifest_shape", "validation/manifest.json",
                  "Manifest must be an object containing a files array of path/sha256 objects.")
        return 0
    checked, seen = 0, set()
    for index, entry in enumerate(data["files"]):
        if not isinstance(entry, dict) or not isinstance(entry.get("path"), str) or not isinstance(entry.get("sha256"), str):
            add_issue(issues, "error", "manifest_entry", "validation/manifest.json",
                      "Manifest entry requires string path and sha256 fields.", index=index)
            continue
        name, expected = entry["path"], entry["sha256"].lower()
        if name in seen:
            add_issue(issues, "error", "manifest_duplicate_path", name, "Manifest lists this path more than once.")
            continue
        seen.add(name)
        target = safe_local_path(root, name)
        if target is None:
            add_issue(issues, "error", "manifest_unsafe_path", name, "Manifest path must stay inside the repository.")
            continue
        if not re.fullmatch(r"[0-9a-f]{64}", expected):
            add_issue(issues, "error", "manifest_invalid_hash", name, "Expected SHA-256 must contain 64 hexadecimal characters.")
            continue
        if not target.is_file():
            add_issue(issues, "error", "manifest_file_missing", name, "Manifest file does not exist.")
            continue
        try:
            actual = sha256_file(target)
        except OSError as exc:
            add_issue(issues, "error", "manifest_file_unreadable", name, "Manifest file could not be hashed.", reason=str(exc))
            continue
        checked += 1
        if actual != expected:
            add_issue(issues, "error", "manifest_hash_mismatch", name, "File bytes do not match the manifest SHA-256.", expected=expected, actual=actual)
    for artifact in artifacts:
        name = artifact.relative_to(root).as_posix()
        if name not in seen:
            add_issue(issues, "warning", "artifact_not_in_manifest", name, "This OOXML artifact is not listed in the integrity manifest.")
    return checked


def validate_repository(root: Path, check_relationships: bool = True) -> dict:
    root = root.resolve()
    issues = []
    if not root.is_dir():
        add_issue(issues, "error", "repository_missing", str(root), "Repository root does not exist.")
        return {"status": "failed", "errors": 1, "warnings": 0, "issues": issues, "artifacts": []}
    for name in REQUIRED_FILES:
        if not (root / name).is_file():
            add_issue(issues, "error", "required_file_missing", name, "Required project file is missing.")
    if not (root / "docs").is_dir():
        add_issue(issues, "error", "docs_directory_missing", "docs", "The documentation directory is missing.")
    # A missing LICENSE is deliberately not treated as a structural failure.
    files = sorted(p for p in root.rglob("*") if p.is_file()
                   and not any(part in SKIP_DIRS for part in p.relative_to(root).parts))
    artifacts = [p for p in files if p.suffix.lower() in {".xlsx", ".docx"}]
    if not any(p.suffix.lower() == ".xlsx" and p.parent == root / "workbooks" for p in artifacts):
        add_issue(issues, "error", "workbook_missing", "workbooks", "Expected at least one .xlsx file directly in workbooks/.")
    artifact_stats = [inspect_ooxml(p, root, issues, check_relationships) for p in artifacts]
    markdown = [p for p in files if p.suffix.lower() == ".md"]
    local_links = sum(inspect_markdown(p, root, issues) for p in markdown)
    manifest_files = inspect_manifest(root, issues, artifacts)
    error_count = sum(i["severity"] == "error" for i in issues)
    warning_count = sum(i["severity"] == "warning" for i in issues)
    return {"status": "failed" if error_count else "passed_with_warnings" if warning_count else "passed",
            "scope": "Read-only package, cached-result, link and integrity checks. No native Excel recalculation or behavioural certification.",
            "errors": error_count, "warnings": warning_count,
            "files_checked": {"ooxml": len(artifacts), "markdown": len(markdown), "local_links": local_links, "manifest_hashes": manifest_files},
            "artifacts": artifact_stats, "issues": issues}


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd(), help="Repository root (default: current directory).")
    parser.add_argument("--json", type=Path, metavar="PATH", help="Optionally write the complete machine-readable report.")
    parser.add_argument("--no-relations", action="store_true", help="Skip checking whether internal OOXML relationship targets exist.")
    args = parser.parse_args(argv)
    report = validate_repository(args.root, check_relationships=not args.no_relations)
    if args.json:
        try:
            args.json.parent.mkdir(parents=True, exist_ok=True)
            args.json.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        except OSError as exc:
            print(f"ERROR: could not write JSON report: {exc}", file=sys.stderr)
            return 2
    counts = report.get("files_checked", {})
    print(f"Repository validation: {report['status']} | {report['errors']} error(s), {report['warnings']} warning(s)")
    print(f"Checked {counts.get('ooxml', 0)} OOXML file(s), {counts.get('markdown', 0)} Markdown file(s), "
          f"{counts.get('local_links', 0)} local link(s), {counts.get('manifest_hashes', 0)} manifest hash(es).")
    for issue in report["issues"][:12]:
        print(f"{issue['severity'].upper()} [{issue['code']}] {issue['path']}: {issue['message']}")
    if len(report["issues"]) > 12:
        print(f"{len(report['issues']) - 12} additional issue(s); use --json PATH for the full report.")
    print("These checks do not recalculate or certify workbook logic in Microsoft Excel.")
    return 1 if report["errors"] else 0


if __name__ == "__main__":
    sys.exit(main())
