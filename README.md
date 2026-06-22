# PayrollTool Web

A browser-only web version of the existing PayrollTool. It runs the same Python scripts locally inside the user’s browser through Pyodide, so HR Excel files do not leave the device.

## What It Does

- Runs the four-step payroll flow: detailed calendar, attendance/rules fill, final calendar, final summary.
- Accepts `.xls` and `.xlsx` HR reports with the same exact basenames used by the Python tool.
- Requires the two styled private templates on each real run so employee rosters, formulas, formats, merged cells, borders, and column widths are preserved.
- Supports prepared permission files and raw permission request files.
- Includes a standalone “Prepare Permissions Only” tool for converting `Nagwa_Permission_Request_Report.xls[x]` into `Nagwa_Permission_Request_permission_details.xls[x]`.
- Exposes business-rule settings for Ramadan, schedules, special-rule pairs, hour reductions, lunch windows, abbreviations, request cutoff defaults, and debug mode.

## Privacy

No backend server, API key, or upload service is required for payroll processing. All files are read in the browser and downloaded back to the same computer.

The first run downloads the Pyodide runtime and Python packages from the CDN; after that the browser cache normally reuses them. Do not commit real HR raw reports, generated payroll outputs, or unsanitized employee templates.

## HR Usage

1. Open the deployed GitHub Pages URL.
2. Drop the raw reports or select a folder containing the expected file names.
3. Upload `Nagwa Technologies.xlsx` and `Final Nagwa Technologies.xlsx` in the Private Templates section.
4. Check the detected payroll period.
5. Review Settings if the month has special rules.
6. Click `Run Payroll`.
7. Download `Nagwa Technologies.xlsx` and `Final Nagwa Technologies.xlsx`.

For permission preparation only, use the separate panel, upload `Nagwa_Permission_Request_Report.xls[x]`, choose month/year/cutoff options, and download the prepared report.

## Local Development

```bash
npm install
npm run dev
```

## Verification

```bash
npm run test
npm run build
```

Smoke tests live under `src/**/*.test.ts`. The acceptance test is a real browser run with private reports/templates, then comparing downloaded workbooks with the desktop-generated `output/` workbooks.

## Deployment

The workflow in `.github/workflows/deploy.yml` builds this folder and deploys the static Vite output to GitHub Pages on pushes to `main`. Runtime Python packages are loaded by the browser from the Pyodide CDN; no server process is deployed.

Target public repository:

```bash
git init
git remote add origin git@github.com:farouknagwa/PayrollTool.git
git add .
git commit -m "Add PayrollTool web app"
git push -u origin main
```

Run tests, build, and a privacy review before committing or pushing.
