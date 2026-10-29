# APAAR Creation status :: SSM, Purba Bardhaman

A web portal for monitoring APAAR ID creation by block, school, management and category, from the
**APAAR Entry Status - School Wise (Only Operational)** report of the UDISE+ portal.
It is a static site: no server or database is needed, and it runs on **GitHub Pages**.
No Tampermonkey or other browser script is needed: the portal's own Excel download is uploaded as it is.

## Features

| Page | What it shows |
|---|---|
| Dashboard | District students / APAAR generated / not generated / generated %, with change since the previous day; requested, failed and not-yet-requested counts; 100% / 1-99% / 0% schools; key insights; charts; block-wise generated %; daily trend; top/bottom 5 blocks; Management and Category summaries |
| Block-wise Status | Rank, schools, students, generated, not generated, %, daily change, requested, failed, 100% / 1-99% / 0% school counts, Leading / Moderate / Lagging status. Tap a block to open its schools |
| Leading & Lagging | Top 10 and bottom 10 blocks, top gainers, least progress, most students without APAAR, most 0% schools, most failed requests, priority schools, large 0% schools, schools with most failures, best improving schools |
| Block × Management | Every block with Dept. of Education, SSK & MSK, Govt. Sponsored, Other Managements and All Managements side by side: schools, enrolment, APAAR created, pending and %. District tiles per management; % shown green / red against the district % of the same management. Downloads as Excel, PDF (2 A4 pages) or **PowerPoint (2 slides)** |
| Management / Category | APAAR generation by school Management, Category or School Type (optionally for one block), with daily change, a chart, and a block × group matrix |
| School lists | All / 100% / 1-99% / 0% generated, schools with failed requests, schools with requests in process; filtered by block, Management, Category, School Type and search |
| Day-wise Comparison | Any two dates side by side: block-wise change, plus schools that improved, became 100%, moved out of 0%, stayed at 0% or decreased |
| Daily Trend | District summary for every date, plus a block × date matrix (%, generated, not generated, failed, 0% schools, 100% schools) |
| Upload / Manage Data | Upload daily files or the school master, delete a day, back up and restore |

Measures:
- **Generated %** = Total Generated / Total Student.
- **Not Generated** = Total Student − Total Generated (includes Requested and Failed).
- **Not yet requested** = Not Generated − Requested − Failed.
- A school is *100%* when every student has an APAAR ID, *0%* when none has, and *No students* when Total Student is 0.

**Leading / Lagging rule:** blocks are ranked by generated %. The top quarter are *Leading*, the bottom quarter are *Lagging*, and the rest are *Moderate*.

Every table has **Excel** and **PDF** download buttons. Each file carries the header
*"APAAR Creation status :: SSM, Purba Bardhaman"*, the report title, the data date and the generation time.
The layout works on phones, tablets and desktops.

## Management / Category / School Type

- **Management** and **Category** come from the `School Management` and `School Category` codes in the APAAR file
  (UDISE+ codes, named in `js/app.js`: `MG_CODES`, `CT_CODES`). A code not listed there takes its name from the school master.
- **School Type** (Co Education / Girls / Boys) comes from the **School Search** export (`UDISE Code`, `Management`, `Category`, `Type`), matched by UDISE code.
  - For everyone: save it as `data/school_master.csv` in the repository (the name is fixed, it is not listed in the manifest).
  - For one browser: upload the CSV on **Upload / Manage Data**. It overrides the repo copy in that browser.

## Daily use

1. On the UDISE+ portal, download **APAAR Entry Status - School Wise (Only Operational)** for the district as Excel.
2. Open the portal and go to **Upload / Manage Data**, then upload the file as it is.
   The date is read from the *Data as on* line at the top of the file. You can override it in the *Data date* box.
3. Uploading again for the same date replaces that day's data.

**Upload / Manage Data is password protected** (same password as the Prashast monitor). After unlocking, it stays open until the browser tab is closed or *Lock* is pressed.
Only a SHA-256 hash of the password is in `js/app.js` (`ADMIN_HASH`). To change the password, put the hash of the new one there.
This is a static site, so the lock keeps casual visitors out of the upload page; it is not server-side security.

Uploaded data is stored **in that browser only** (IndexedDB). Use *Download backup* to move it to another device.

## Publish on GitHub Pages

1. The files are in the repository `apaar`: `index.html`, `css/`, `js/`, `data/`, `.nojekyll`.
2. Go to **Settings → Pages → Build and deployment**. Set *Source* to **Deploy from a branch**, choose branch `main` and folder `/ (root)`, then save.
3. After about a minute the portal is live at `https://<your-username>.github.io/apaar/`.

### Let everyone see the same data

So that every officer who opens the link sees the full history without uploading anything, commit the daily files to the repo:

1. Put each day's Excel file in `data/`, renamed with its data date, for example `APAAR_Entry_Status_School_Wise_2026-10-28.xlsx`.
2. List the files in `data/manifest.json`:
   ```json
   [
     "APAAR_Entry_Status_School_Wise_2026-10-28.xlsx",
     "APAAR_Entry_Status_School_Wise_2026-10-29.xlsx"
   ]
   ```
   An entry can also be `{ "file": "...", "date": "2026-10-29" }`. Without a date in the name or entry, the *Data as on* date in the file is used.
3. Commit. These dates appear for everyone, marked "from repo". If someone uploads a file for the same date in their own browser, their upload is used instead.

> If the repository is public, anyone with the link can download files in `data/`. Make the repo private (GitHub Pages on private repos needs a paid plan) or keep data upload-only if that matters.

## Accepted file format

The first sheet that has a **School Name** column is used. The columns are found by name:
`Block Name`, `School Name`, `UDISE Code`, `School Management`, `School Category`, `Total Student`, `Total Generated`, `Total Requested`, `Total Failed`.
The title rows above the header and the `(1) (2) (3)…` column-number row are skipped.

## Local testing

Open the folder through any web server, for example XAMPP: `http://localhost/apaar/`.
Opening `index.html` directly from disk also works, but the `data/manifest.json` loading needs a web server.
