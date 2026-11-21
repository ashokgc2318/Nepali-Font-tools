# Case Expiry Monitor (BS)

Tracks the one-year expiry of registered cases from their **Darta Miti** in Bikram Sambat.

* Upload the case register (.xlsx) → dashboard, Red / Yellow / Green / Overdue lists.
* Add, edit or delete cases, then **Download Excel** to get the updated master file.
* Everything runs in the browser. The Excel file is never uploaded anywhere and nothing is stored.

## Publish on GitHub Pages
Put these files in the repository (root, or a folder such as `/cases/`), then
Settings → Pages → Deploy from branch. Do **not** commit any case Excel files to the repository.

Libraries bundled in `vendor/`: Chart.js (MIT), SheetJS Community Edition (Apache-2.0), ExcelJS (MIT).
Fonts in `fonts/`: Mukta and Poppins (SIL Open Font License).
BS month lengths: 2000–2099 BS; 2084 onward are projected until the official Panchanga is published.
