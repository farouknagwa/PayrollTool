# PayrollTool Web

A browser-only web version of the existing Python PayrollTool. It runs payroll processing locally in the user’s browser, so HR Excel files do not leave the device.

## What It Does

- Runs the four-step payroll flow: detailed calendar, attendance/rules fill, final calendar, final summary.
- Accepts `.xls` and `.xlsx` HR reports with the same exact basenames used by the Python tool.
- Supports prepared permission files and raw permission request files.
- Includes a standalone “Prepare Permissions Only” tool for converting `Nagwa_Permission_Request_Report.xls[x]` into `Nagwa_Permission_Request_permission_details.xls[x]`.
- Exposes business-rule settings for Ramadan, schedules, special-rule pairs, hour reductions, lunch windows, abbreviations, request cutoff defaults, and debug mode.

## Privacy

No backend server, API key, or external service is required for payroll processing. All files are read in the browser and downloaded back to the same computer.

Do not commit real HR raw reports, generated payroll outputs, or unsanitized employee templates. Use the optional template upload controls for private/current rosters.

## HR Usage

1. Open the deployed GitHub Pages URL.
2. Drop the raw reports or select a folder containing the expected file names.
3. Check the file checklist and detected payroll period.
4. Review Settings if the month has special rules.
5. Click `Run Payroll`.
6. Download `Nagwa Technologies.xlsx` and `Final Nagwa Technologies.xlsx`.

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

Synthetic tests live under `src/**/*.test.ts`. Private parity tests can compare web outputs with Python-generated workbooks locally, but private fixtures must not be committed.

## Deployment

The workflow in `.github/workflows/deploy.yml` builds this folder and deploys the static Vite output to GitHub Pages on pushes to `main`.

Target public repository:

```bash
git init
git remote add origin git@github.com:farouknagwa/PayrollTool.git
git add .
git commit -m "Add PayrollTool web app"
git push -u origin main
```

Run tests, build, and a privacy review before committing or pushing.
