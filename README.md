# Benjamin Rhoads research website

A static website for publications and computer vision projects. Open `index.html` locally to preview it. The full paper page is `gnn-paper.html`; its figures are in `assets/gnn/`.

## Edit and publish with GitHub Desktop

This folder is already the clone of [bmrhoads51/bmrhoads51.github.io](https://github.com/bmrhoads51/bmrhoads51.github.io). It is connected to GitHub, so there is no need to clone it again or copy files from the old ZIP download. **Commit** saves a named version locally; **push** uploads your commits to GitHub.

1. In GitHub Desktop, choose **File → Add local repository** and select this `bmrhoads51.github.io` folder if it is not already listed. Sign in to the GitHub account that owns the repository if prompted.
2. Edit files in this folder. The homepage is `index.html`, the publication and project indexes are `publications.html` and `projects.html`, the paper is `gnn-paper.html`, and the shared design is in `style.css`. Open `index.html` in a browser to preview local edits.
3. Return to GitHub Desktop. Review the changed files, write a short summary, and click **Commit to main**.
4. Click **Push origin**. GitHub Pages publishes the `main` branch from the repository root at [https://bmrhoads51.github.io/](https://bmrhoads51.github.io/). Publication can take a few minutes.
5. Check that address in a private browser window if your usual browser still shows an older cached page. A hard refresh also usually clears a stale copy (Ctrl+Shift+R on Windows).

For later edits, repeat steps 2–5. Clone only when setting up a different computer.

## Choose homepage images

Edit the entries in `featured-images.js` to choose which figures rotate on the homepage. Each entry gives an `image` path, a short `alt` description, a `caption`, and a `paper` link. The first image is also in `index.html` as a fallback for visitors whose browsers do not run JavaScript. When you change the first entry, update that fallback image and link too. `carousel.js` handles the automatic rotation and Previous/Next buttons; you normally do not need to edit it.

The current three entries are Figures 2.1, 2.2, and 2.3 from the GNN paper. Their links lead directly to those figures on the paper page. For another publication, put its image files in `assets/`, add a publication page, and add entries pointing to that page. Keep the `paper` value as a relative path such as `new-paper.html#figure-1`.

## Paper source

The paper page was generated from `GNN/GNNPaper_.docx`, the newest supplied Word file, dated April 1, 2026. The displayed publication title, authors, and DOI come from the published article. The Word manuscript is a later draft, so the page labels the journal article as the version of record. The original Word file is outside this website folder and is not published with the site.

To regenerate the paper page after editing the Word manuscript, keep a sibling `GNN/GNNPaper_.docx` next to this website folder and run `tools/build_paper.py` with Python and `python-docx` installed. The generated HTML and images are committed to the website; visitors do not need Python.
