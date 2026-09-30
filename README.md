# TSET-AHIF Project Plan Progress Dashboard — v1.2

A public, always-on web dashboard that tracks each project-plan milestone
from **Not Started** (grey) → **In Progress** (yellow) → **Complete** (green),
organized in one card per project year and split into **Construction** and
**Administrative** groups.

It's a static site: `index.html` reads `data/dashboard.json`, and
`build_data.py` creates that JSON from the project-plan spreadsheet.
It runs free on GitHub Pages — nothing to install on a server.

---

## What's in this folder

| File / folder | What it is |
|---|---|
| `index.html` | The dashboard page. |
| `data/dashboard.json` | The data the page displays. **Regenerate this whenever the spreadsheet changes.** |
| `assets/dashboard-graphic.png` | The logo banner at the top of the page (web-sized copy of `Graphics/Dashboard Graphic.png`). |
| `build_data.py` | Reads the spreadsheet and writes `data/dashboard.json`. |
| `requirements.txt` | The one Python package the build script needs (`openpyxl`). |
| `.nojekyll` | Tells GitHub Pages to serve the files as-is. Keep it. |
| `README.md` | This file. |

The spreadsheet itself does **not** go on GitHub — only the generated JSON does.

---

## Where the data comes from

- Workbook: `TSET_Project_Plan_Tracking_Data.xlsx`
- Tab: **Dashboard_Data** (headers on row 1)
- Columns used: Year, Project_Type, Project Milestones, Task Assigned to,
  Target Start Date, Target End Date, Actual Start Date, Actual End Date,
  Task Status, Notes

Notes on how the columns are read:

- **Year** — grouped by the "Year N" part (e.g. `2026 - Year 1` → Year 1).
  The calendar year shown on the dashboard is the earliest one typed for that
  year number.
- **Project_Type** — "Construction" or "Administrative"; drives the category pill.
- **Task Status** — must be `Not Started`, `In Progress`, or `Complete`
  (the drop-down values). Anything else is treated as Not Started and listed in
  the dashboard's Data-quality notes.
- **Dates** — real Excel dates are best. Text dates like `3/8/2027` also work.
  `n/a` or blank means "no date yet." An impossible date (e.g. `9/31/2027`)
  is shown as typed and flagged in the Data-quality notes.
- **Notes** — shown when someone clicks a milestone row.

> **Important:** the Dashboard_Data tab is filled by formulas that pull from the
> "Project Plan" tab. The build script reads the values Excel saved with the
> file, so **always open the workbook in Excel and click Save** after editing,
> before running the script. (If a whole year's first milestone goes missing,
> that's the sign the file wasn't saved from Excel.)

---

## Updating the dashboard (each time the spreadsheet changes)

1. **Edit and save the spreadsheet in Excel.**
   It lives in `TSET_Project-Plan_Dashboard\Data\`.

2. **Run the build script.** Open a Command Prompt (press the Windows key,
   type `cmd`, press Enter). First move into the live dashboard folder —
   **the script won't run correctly without this step**:

   ```
   cd C:\Users\marlinc\Desktop\TSET_Project-Plan_Dashboard\TSET_progress_dashboard_Live
   ```

   Then run:

   ```
   python build_data.py
   ```

   The script automatically finds the workbook in
   `TSET_Project-Plan_Dashboard\Data\`. To point it at a different file:

   ```
   python build_data.py "C:\path\to\TSET_Project_Plan_Tracking_Data.xlsx"
   ```

   (If you ever move or rename the dashboard folder, change the `cd` path to
   match. If the folder is on a different drive, use `cd /d` instead of `cd`.)

   It prints a summary (milestone counts by status) and any data-quality notes.

3. **Upload the new `data/dashboard.json` to GitHub** (see "Uploading an
   updated file" below). The live site refreshes within a minute or two.

Uploading only the spreadsheet does nothing — the site only reads `dashboard.json`.

**First-time setup only:** install the one package the script needs
(from the same folder):

```
cd C:\Users\marlinc\Desktop\TSET_Project-Plan_Dashboard\TSET_progress_dashboard_Live
pip install -r requirements.txt
```

---

## Previewing on your own computer

Double-clicking `index.html` shows a "Couldn't load data" message, because
browsers block pages opened from disk from reading other files. To preview,
open a Command Prompt and run:

```
cd C:\Users\marlinc\Desktop\TSET_Project-Plan_Dashboard\TSET_progress_dashboard_Live
python -m http.server
```

then open <http://localhost:8000> in your browser. Press Ctrl+C in the
Command Prompt to stop it.

---

## Publishing on GitHub Pages (first time)

1. Sign in at <https://github.com>.
2. Click **+** (top right) → **New repository**.
   - Name it, e.g. `TSET-Dashboard`.
   - Set it to **Public** (required for free GitHub Pages).
   - Leave "Add a README" unchecked. Click **Create repository**.
3. On the new repo's page, click **uploading an existing file**.
4. Open your dashboard folder in File Explorer, select **everything inside it**
   (`index.html`, `README.md`, `build_data.py`, `requirements.txt`, `.nojekyll`,
   and the `data` and `assets` folders) and drag it into the browser window.
   - `.nojekyll` is a hidden-style file; if you can't see it, turn on
     View → Show → Hidden items in File Explorer.
   - Make sure the folders keep their names: the site must have
     `data/dashboard.json` and `assets/dashboard-graphic.png`.
5. Scroll down and click **Commit changes**.
6. Go to **Settings** (tab at the top of the repo) → **Pages** (left sidebar).
   - Under **Build and deployment → Source**, choose **Deploy from a branch**.
   - Branch: **main**, folder: **/ (root)**. Click **Save**.
7. Wait 1–2 minutes and refresh the Pages settings screen. It will show your
   live link, which looks like
   `https://<your-username>.github.io/TSET-Dashboard/`.

## Uploading an updated file

1. Open the repo on GitHub and click into the `data` folder.
2. Click **Add file → Upload files** and drag in the new `dashboard.json`.
   GitHub replaces the old one with the same name.
3. Click **Commit changes**. The live site updates in a minute or two
   (press Ctrl+F5 to force-refresh if you still see the old numbers).

The same steps work for any other file (e.g. a new `index.html` from a new
version) — just upload it into the matching folder.

---

## How the dashboard reads

- **KPI tiles** — milestones complete (with overall %), counts per status, and
  **Past target end**: milestones not marked Complete whose target end date is
  before today (calculated live in the viewer's browser, so it stays current
  even between data updates).
- **Filters** — Year, Category and Status pills. Click to show/hide; every card
  updates. "Show all" resets.
- **Progress by project year** — one bar per year plus an all-years bar,
  showing the share of milestones in each status. Hover for counts;
  "View as table" for the numbers.
- **Year cards** — one per project year, split into Construction and
  Administrative. Each milestone row shows:
  - three progress blocks that fill left to right
    (grey = Not Started, + yellow = In Progress, + green = Complete),
    with the status written beside them;
  - target start – end dates and length in days, plus a red
    "Past target end" tag when applicable;
  - actual start – end dates (— until entered).
  Click a row to see who it's assigned to and its notes.
  "View as table" shows the whole year as a plain table.
- **Data-quality notes** — how to read the page, plus anything odd the build
  script found in the spreadsheet.

---

## Customizing

- **Title:** `DASHBOARD_TITLE` near the top of `build_data.py`.
- **Banner graphic:** replace `assets/dashboard-graphic.png` (keep the name).
- **Colors:** the `:root { ... }` block at the top of `index.html`
  (`--st-ns`, `--st-ip`, `--st-c` are the status colors;
  `--cat-con`, `--cat-adm` are the category colors).
