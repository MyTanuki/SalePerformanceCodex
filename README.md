# Sale Performance Dashboard 2026

Static web app for tracking sales target, won amount, revenue schedule, pipeline, deal detail, and new lead activity from CSV files.

## Project Structure

```text
sale-performance-dashboard/
  index.html                  Main dashboard layout
  app.js                      CSV parsing, mapping, filters, charts, tables, and UI logic
  styles.css                  Light/dark theme and responsive layout
  data/
    dashboard-data.js         Empty skeleton data loaded by the page
    Sale Target.csv           Optional default target file auto-loaded at startup
    Stage Mapping.csv         Optional default mapping file auto-loaded at startup
    DEAL_*.csv                Sample/source deal CSV, not auto-loaded
  scripts/
    preview-server.js         Local static file server
    build-dashboard-data.js   Optional offline data builder
```

## Run Offline Preview

From this folder:

```powershell
node .\scripts\preview-server.js
```

Open:

```text
http://127.0.0.1:4174/
```

Use a local server instead of opening `index.html` directly so the browser can fetch files from `./data`.

## Data Loading

The app starts with empty skeleton data from:

```text
data/dashboard-data.js
```

On page load, the app auto-loads these setup files when present:

- `data/Stage Mapping.csv`
- `data/Sale Target.csv`

Deal data is not auto-loaded. Use the gear button in the header to upload a Deal CSV, then press `Refresh Dashboard`.

## Upload Files In The App

Open `Setting` from the gear icon and upload:

- `Sale Target CSV File`
- `Mapping CSV File` (`Stage Mapping.csv` or compatible mapping CSV)
- `Deal CSV File`

The dashboard stores uploaded data in browser local storage so refreshes can keep the latest workspace data. Use `Reset Data` to clear uploaded data and return to the skeleton plus auto-loaded setup files.

## Main Tabs

- `Overview`: summary charts, status mix, and top open opportunities.
- `Sales Performance`: performance charts and transaction detail. Default chart status filter shows only `Won`.
- `Revenue Analysis`: monthly revenue schedule from `Won` / `Pre-WON` deals for finance planning and accounting review. One-time/OTC deals are recognized in the contract start month, while recurring deals are spread across contract months. The monthly revenue chart compares actual recurring revenue with the 2026 target estimated from 2025 recurring revenue and shows achievement %. Revenue Detail is sorted chronologically and supports AND / NOT search.
- `All Deals`: all deal detail in the selected period, including included/excluded/unmapped pipeline counting status.
- `Pipeline Risk`: open deals with risk signals such as overdue, stale, missing expected date, and not contacted.
- `Data Quality`: mapping assumptions, normalization totals, and unmapped responsible names.
- `New Deals`: lead creation analysis by ISO Week using `Created Date`.

## Filters

Global filters appear above the tabs:

- `Filter มุมมอง`: view mode, sale group, sale, global search, and New/Renew type buttons.
- `Filter ช่วงเวลา`: Year, Quarter, Month, or Range. Week filtering was removed from the global period filter.

`New Deals` uses the same view filter card on the left, and uses a dedicated `New Deals by ISO Week` filter card on the right:

- `ตั้งแต่ Week`
- `ถึง Week`
- `ค้นหา Lead`

Dropdown controls are rendered as custom lists so the current selected value is centered when the list opens.

ISO Week labels use this format:

```text
2026-W18 (2026-04-27)
```

When opening the ISO Week dropdown, the current selected week is centered in the list.

## Mapping And Counting Rules

Mapping is read from `Stage Mapping.csv` or an uploaded compatible mapping file. Supported mapping areas include:

- Stage source to display stage
- Deal Type source to display deal type
- Sales to Sale Group
- Pipeline Matching by Sale Group and New/Renew type

Pipeline Matching can also be edited in the Setting modal. When matching rules exist, only deals whose pipeline is included for the sale group and sale type are counted in sales performance.

## Core Assumptions

- `Pre-WON` and `Won` are treated as won.
- `Pre-LOST` and `Lost` are treated as lost.
- Sales performance date filtering uses Expected Close Date for pipeline tracking.
- Revenue Analysis filters deals by view/search/type, then filters the output schedule by the selected Year, Quarter, Month, or Range.
- Revenue Analysis recognizes `One-time` / `OTC` deals in the contract start month. Recurring deals spread contract value straight-line by month using `Contract Start Date` and `Contract End Date`; if those are incomplete it falls back to `Contract Period (จำนวนเดือน)`, then MRR inference, then a one-month temporary schedule with a data check flag.
- The 2026 recurring target table uses only `Billing Type = Recurring`, takes monthly revenue from 2025, and shifts each month to the same month in 2026 as a baseline target estimate.
- Monthly Revenue Schedule compares `Actual Recurring Revenue` against the calculated recurring target. Total Revenue is still shown separately so one-time revenue does not inflate recurring achievement.
- The 2026 recurring target table can be exported to CSV from the `Revenue Analysis` tab.
- New Deals uses Created Date and ISO Week.
- `New` and `Renew` can be filtered independently.

## Optional Offline Build

The app can also rebuild `data/dashboard-data.js` from CSV files:

```powershell
node .\scripts\build-dashboard-data.js
```

Optional source paths:

```powershell
node .\scripts\build-dashboard-data.js "C:\Temp\DEAL_20260506_2e04b49b_69fabfc019c48.csv" "C:\path\to\Sale Target.csv" "C:\path\to\Stage Mapping.csv"
```

The current app is designed to start empty, so use this builder only when you intentionally want to pre-generate dashboard data.
