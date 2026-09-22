# Benjamin Rhoads research website

A static website for publications and computer vision projects. Open `index.html` locally to preview it. The full paper page is `gnn-paper.html`; its figures are in `assets/gnn/`.

## Publish with GitHub Desktop

The downloaded `bmrhoads51.github.io-main` folder is a GitHub ZIP download. It has no `.git` directory, so it is not connected to GitHub and cannot be pushed as it stands. **Clone** means making a local copy that stays connected to the repository; **push** means uploading your committed changes.

1. Install [GitHub Desktop](https://desktop.github.com/) and sign in to the GitHub account that owns the website.
2. In GitHub Desktop, choose **File → Clone repository**, select `bmrhoads51/bmrhoads51.github.io` (or paste the actual repository URL from GitHub), and choose a local folder. If that repository name is different, use the one that contains your current website.
3. In File Explorer, copy the contents of this updated `bmrhoads51.github.io-main` folder **into the cloned folder**. Copy `index.html`, `gnn-paper.html`, `style.css`, `assets/`, `tools/`, and the existing project pages and image folders. Do not copy the outer `bmrhoads51.github.io-main` folder itself; `index.html` needs to sit at the top level of the repository.
4. Return to GitHub Desktop. Review the changed files, enter a summary such as `Create research website and GNN paper page`, click **Commit to main**, then **Push origin**.
5. On GitHub, open the repository’s **Settings → Pages**. Under **Build and deployment**, choose **Deploy from a branch**, branch `main`, folder `/ (root)`, then save. If Pages is already configured this way, the push will update the site automatically.
6. For a repository named exactly `bmrhoads51.github.io`, the site address is `https://bmrhoads51.github.io/`. GitHub Pages may take a few minutes to update after a push.

For later edits, work in the **cloned folder**. Open it in your editor, make changes, preview `index.html`, then use GitHub Desktop to commit and push. You only clone once per computer.

## Paper source

The paper page was generated from `GNN/GNNPaper_.docx`, the newest supplied Word file, dated April 1, 2026. The displayed publication title, authors, and DOI come from the published article. The Word manuscript is a later draft, so the page labels the journal article as the version of record. The original Word file is outside this website folder and is not published with the site.

To regenerate the paper page after editing the Word manuscript, keep a sibling `GNN/GNNPaper_.docx` next to this website folder and run `tools/build_paper.py` with Python and `python-docx` installed. The generated HTML and images are committed to the website; visitors do not need Python.
