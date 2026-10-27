# Publish Security Tradeoff Lab through the GitHub web interface

This is a browser-only publication guide for the reviewed prerelease package `0.1.0-review.2`. No command line, Git installation, or GitHub CLI is required. Repository-only preparation was completed on 2026-10-26; the supplied source artifacts and functional-review baseline remain unchanged.

Read [known issues](docs/known-issues.md), [LICENSE-STATUS.md](LICENSE-STATUS.md), and the [publication-metadata record](validation/publication-metadata.json) before changing repository visibility.

## 1. Extract and inspect the package

Extract the supplied ZIP on your computer. Open the extracted `security-tradeoff-lab` folder. You will upload the contents of this folder, including `.github`, `.gitignore`, and the other dotfiles. Do not upload the ZIP itself, the outer collection folder, or a second nested `security-tradeoff-lab` folder.

Keep real organizational assessments, credentials, customer data, and local working copies out of the upload. The package contains fewer than 100 files and every file is below GitHub's browser-upload limit of 25 MiB.

## 2. Create an empty repository

1. Sign in to GitHub. In the upper-right corner, select **+**, then **New repository**.
2. Choose the intended owner. Use `security-tradeoff-lab` as the repository name, subject to availability in that account.
3. Use this description: **Compare security-control changes through employee time, specialist capacity, cash and evidence.**
4. Choose **Private** for the initial import. Do not choose a template.
5. Do not add a README, `.gitignore`, or license. Those choices would create an initial commit and could conflict with the supplied files.
6. Click **Create repository**.

## 3. Upload the repository contents

1. On the empty repository's Quick Setup page, choose **uploading an existing file**. If the repository already has an initial commit, use **Add file**, then **Upload files** instead.
2. Open the extracted `security-tradeoff-lab` folder in your file manager. Select everything inside it and drag the selection into the browser upload area. Include `.github` and the other dotfiles. On macOS, **Command+Shift+.** toggles hidden-file visibility in Finder.
3. Before committing, inspect the upload list. Confirm that it includes `README.md`, `.gitignore`, `.github/workflows/validate.yml`, `CITATION.cff`, and the files under `workbooks/`.
4. Use `Add reviewed prerelease package` as the commit message. For this initial import into an empty repository, commit directly to the default branch.

GitHub supports up to 100 files in one browser upload. Its browser uploader does not apply `.gitattributes` processing. The supplied files are already in the exact reviewed byte state required by the checks, including the recorded CSV line endings, and the Office/image files are unchanged binary artifacts. No conversion or text normalization is needed for this import. Keep that limitation in mind for later browser-based bulk changes.

## 4. Check the imported repository

1. Open **Actions**. If GitHub asks for permission to run workflows, enable them for this repository. Open the first **Repository checks** run and inspect every step. The workflow runs the complete v6.2 scenario, arithmetic, and structural checker in addition to package integrity. A green run confirms the recorded scope; it is not native Excel acceptance.
2. Return to the repository main page. Open the **About** gear and add the topics listed in [repo-settings.json](repo-settings.json).
3. Open **Issues** and create the issues you want to track from [docs/issues-to-open.md](docs/issues-to-open.md). Use synthetic examples only.
4. Review **Settings**, then **Code security and analysis**. Enable private vulnerability reporting if it is available for the repository and plan. Otherwise, provide the verified private contact route described in [SECURITY.md](SECURITY.md).
5. After the repository exists, add its actual `repository-code` URL to `CITATION.cff` through GitHub's file editor. Set or update `date-released` only when the corresponding release is actually published.

## 5. Complete the publication decisions

The MIT/CC BY 4.0 license allocation is confirmed in [LICENSE-STATUS.md](LICENSE-STATUS.md). Complete the publication-metadata review. The scoped calculation checks pass; native Excel acceptance and print-layout work remain before a stable release.

## 6. Change visibility when ready

When the license, metadata, and known-issue presentation are acceptable for public viewing, open **Settings**, then **General**. In the danger zone, choose **Change repository visibility**, then **Make public**, and follow GitHub's confirmation prompts. Review the repository main page and every linked download afterwards.

Public visibility and open-source licensing are separate decisions. A public repository without an operative license does not grant general reuse rights.

## 7. Create a reviewed prerelease

1. On the repository main page, open **Releases**, then **Draft a new release**.
2. Under **Choose a tag**, create `v0.1.0-review.2` targeting the default branch.
3. Use **Security Tradeoff Lab 0.1.0-review.2** as the title.
4. Explain that this is a reviewed prerelease and link to `docs/known-issues.md` and `validation/release-status.json`. State the remaining acceptance work plainly.
5. Attach the principal artifacts and source manifest for convenient downloading:
   - `workbooks/Opportunity-Cost-Security-Model-v6.2.xlsx`
   - `validation/manifest.json`
6. Select **This is a pre-release**. Save the draft first, inspect the assets and notes, and then choose **Publish release**.

GitHub also creates source ZIP and tar.gz archives automatically. The attached files above give users direct access to the principal artifacts without searching the repository tree.

## 8. Before any stable release

Resolve the applicable functional blockers, complete [native Excel acceptance](validation/native-excel-acceptance.md), and finish the publication-metadata review. Keep the confirmed allocation in [LICENSE-STATUS.md](LICENSE-STATUS.md) consistent with the released artifacts. Update `validation/release-status.json`, its evidence paths, the source manifest, citation/version metadata, and change log to describe the same artifact set.

You can run the explicit release declaration check without leaving the browser. Open **Actions**, then **Repository checks**, then **Run workflow**, and enable `require_stable_release`. It intentionally fails while blockers or required acceptance evidence remain. A manually requested green gate confirms the recorded declarations and evidence-file presence; the underlying evidence still needs human review.

## Official GitHub guidance

- [Creating a new repository](https://docs.github.com/en/repositories/creating-and-managing-repositories/creating-a-new-repository)
- [Adding a file to a repository through the browser](https://docs.github.com/en/repositories/working-with-files/managing-files/adding-a-file-to-a-repository)
- [Managing releases in a repository](https://docs.github.com/en/repositories/releasing-projects-on-github/managing-releases-in-a-repository)
- [CITATION.cff support](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/about-citation-files)
- [Repository licensing](https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/customizing-your-repository/licensing-a-repository)
