# Sale Performance Dashboard 2026

Static dashboard for combining deal pipeline data with Sale Target 2026.

## Open

The dashboard is available locally at:

```text
http://127.0.0.1:4173/
```

## Rebuild Data

Run this when either CSV changes:

```powershell
node .\scripts\build-dashboard-data.js
```

Optional custom source paths:

```powershell
node .\scripts\build-dashboard-data.js "C:\Temp\DEAL_20260506_2e04b49b_69fabfc019c48.csv" "C:\Users\noppadol.s\OneDrive - 1-TO-ALL Co., Ltd\Project\1toAll\Sales Dashboard\Sales Amount\Sale Target 2026_r1.5.csv"
```

## Stage Assumptions

- `Pre-WON` is counted as `Won`.
- `Pre-LOST` is counted as `Lost`.
- Won/Lost month uses `Stage change date`, then falls back to `Expected close date`, then `Created`.
- Renew category uses Pipeline containing `Renew` or Deal Type starting with `Re-New`; otherwise the deal is treated as New.
