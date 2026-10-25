# Investor Survey Monitor

Field monitor for the Investor Survey. Sign in with username `rapid`.

## Files
- `index.html` - the dashboard page
- `data/data.json` - the data the page reads (rebuilt by the Action)
- `scripts/build.py` - turns the Kobo XLSX export into `data/data.json` and runs the data checks
- `scripts/fetch_kobo.py` - downloads the export from Kobo
- `scripts/bd_div.json` - division boundaries (geoBoundaries, CC BY 4.0)
- `.github/workflows/update.yml` - runs every 30 minutes

## Setup
1. Upload everything in this folder to a new GitHub repository (keep the folder structure, including `.github`).
2. In Kobo, open the form > Data > Downloads. Choose XLS, "Labels", "Include all versions", then "Save selection as..." to make a saved export.
   Copy its synchronous export link (looks like `https://kf.kobotoolbox.org/api/v2/assets/<form id>/export-settings/<setting id>/data.xlsx`).
3. Get your Kobo API token: Account settings > Security > API key.
4. In GitHub: Settings > Secrets and variables > Actions > New repository secret. Add `KOBO_EXPORT_URL` and `KOBO_TOKEN`.
5. Settings > Pages > Deploy from branch > `main`, folder `/ (root)`.
6. Actions > Update dashboard data > Run workflow, to fill the data the first time.

## Note
`build.py` reads export columns by position. If questions are added to or removed from the form, the column numbers in `build.py` must be updated.
The password only hides the page. Anyone with the address of `data/data.json` can open it, so use a private repository if the firm names must stay private (GitHub Pages from a private repo needs a paid plan).
