#!/usr/bin/env python3
"""Explicit stable-release check; package integrity is a separate question."""
import json
from pathlib import Path
import sys

root = Path(__file__).resolve().parents[1]
status = json.loads((root / 'validation/release-status.json').read_text(encoding='utf-8'))
blockers = status.get('blockers', [])


def evidence_file(key):
    value = status.get(key)
    return isinstance(value, str) and bool(value) and (root / value).is_file()


evidence_keys = ('native_excel_evidence', 'licensing_evidence', 'publication_metadata_evidence')
ready = (status.get('stable_release_ready') is True
         and status.get('native_excel_verified') is True
         and status.get('licensing_confirmed') is True
         and status.get('publication_metadata_reviewed') is True
         and not blockers
         and all(evidence_file(key) for key in evidence_keys))
if ready:
    print('Recorded stable-release requirements and evidence files are complete. Review the evidence before publishing.')
else:
    print('STABLE RELEASE NOT READY: this is a reviewed prerelease package.')
    for item in blockers:
        print('- ' + item)
    if not status.get('native_excel_verified'):
        print('- Native Excel edit/recalculate/save/reopen acceptance is not recorded.')
    if not status.get('licensing_confirmed'):
        print('- The overall release license and applicable third-party terms are not confirmed.')
    if not status.get('publication_metadata_reviewed'):
        print('- The publication sensitivity-label and document-metadata review is not recorded.')
    for key in evidence_keys:
        if not evidence_file(key):
            print(f'- Evidence file is missing or not recorded: {key}.')
sys.exit(0 if ready else 1)
