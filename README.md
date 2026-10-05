# Metropolis electricity billing

Web app for The Metropolis monthly CLP electricity: parse the 15 bills, pass the gate, roll Cost Allocation and Cost Sheet forward, and chart consumption from the history workbook.

Repository: https://github.com/chunghonkit/fortune-metropolis-billing

Live site: https://mpelect.citybaseiot.duckdns.org

OpenResty proxies that host to uvicorn on the Omarchy machine (`0.0.0.0:8000`). `elect.citybaseiot.duckdns.org` is retired.

Pull request [#3](https://github.com/chunghonkit/fortune-metropolis-billing/pull/3) holds the current dashboard and month-end work until it is merged. `main` may still be the older intake-only tree.

## What a month-end does

1. Set the billing month (`YYYY-MM`).
2. Upload the 15 CLP PDFs in the browser. The server parses each file in a temp path and deletes it. PDFs are not kept. The dashboard does not scan a bills folder.
3. Enter the check-meter `6681757` (散熱水泵電) previous and present readings. One decimal place is allowed. Present must be at least previous.
4. The gate passes only when all 15 accounts are present, for that month, without duplicates, and the meter log is complete.
5. Process month-end. The app copies the previous month’s Cost Allocation and Cost Sheet, then updates the copies. Search order for the master is the previous month’s `out/` first, then that month’s folder, then that month’s `masters/`. The month being written is never its own master.

Cost Allocation (Elect Charge):

- Column BC moves into BB, and the new month is written into BC.
- Check meter `6681757` previous/present are written to BR31/BS31. The delta formula in BQ31 is left alone.

Cost Sheet:

- Year-grid sheet `01-12 2025` gets that month’s centre column (including June).
- Sheet `Comparison ` (the name has a trailing space) and all 15 account sheets are updated from the bills and the live allocation.
- Older year grids and the budget sheet are not written. Staff names and titles are copied, not invented.

Reprocessing a month overwrites that month’s `out/` and replaces only that month’s rows in the history workbook.

## Where files live

`METROPOLIS_ROOT` defaults to `~/Metropolis`. On Omarchy it is `/home/kit/Metropolis`.

```
METROPOLIS_ROOT/
  2025-05/
    out/
      Cost Allocation - Electricity 2007-2.xlsx
      Cost Sheet-2025-05.xlsx
      dashboard.json
  masters/
    electricity_history.xlsx
```

`electricity_history.xlsx` is not inside a month `out/` folder.

- Sheet `Bills`: one row per account per month (kWh and the charges printed on the bill).
- Sheet `Meters`: one row per meter. `charge` is filled only when the bill prints an amount for that meter (a FiT line). `proportional_charge` is calculated: account charge times (meter kWh / account kWh). It is not a printed meter charge. For FiT accounts `52167-13569-2` and `00776-78552-1` the base is net due plus absolute FiT. If account kWh is zero, `proportional_charge` is blank.

Real Citybase workbooks are `.xls`. The updater uses openpyxl on `.xlsx` copies.

## Dashboard

https://mpelect.citybaseiot.duckdns.org serves this UI.

- English / 繁體中文. The choice is stored in the browser (`localStorage`, key `dashboard-lang`). English is the default.
- Consumption shows one set of buttons at a time: electricity bills, meters, or cost centres. Nothing is charted until a button is clicked.
- A bill charts stored kWh and the bill total. A cost centre charts each bill’s saved kWh times that month’s live percent, and the allocation amount. A meter charts kWh and `proportional_charge`.
- Downloads for the selected month: Cost Allocation (`/api/download/{YYYY-MM}/cost-allocation`) and Cost Sheet (`/api/download/{YYYY-MM}/cost-sheet`). History: `/api/history/download`. Meter log: `/api/download-meter-log`.

Account numbers, meter ids, cost-centre codes, file names, and figures are not translated.

## Allocation rules

- Centre dollars are the account allocation base times that month’s live percent from the master (`AC DEPT`). Percents are not hardcoded.
- FiT accounts `52167-13569-2` and `00776-78552-1`: base is net due plus absolute FiT, then the live percents.
- Office chillers: AO 92% / OC 8%.
- Food Court: 100% FC, outside the operational master chain.
- Retail chillers (`55861-52267-1`): AC/SW comes from check meter `6681757`. `O7 = (dial delta × 160) / kWh` of private meter `9024222` (legacy `9046787`).
- One cent is a match. Elect Charge ledger amounts are bookkeeping and are not Cost Sheet centre dollars.
- Do not invent a meter charge the CLP bill does not print. `proportional_charge` is the calculated share above, stored in its own column.

Month-end uses `app/citybase_model.py`, not `app/allocation_engine.py`.

## Run locally

Python 3.8+, then:

```bash
git clone https://github.com/chunghonkit/fortune-metropolis-billing.git
cd fortune-metropolis-billing
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export METROPOLIS_ROOT="$HOME/Metropolis"
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

Open http://127.0.0.1:8000. Health check:

```bash
curl -s http://127.0.0.1:8000/api/health
```

```json
{"status":"ok","version":"1.0.0-part1"}
```

The version string is still `1.0.0-part1`.

### Omarchy service

systemd unit `metropolis-dashboard.service`:

- `WorkingDirectory` `~/Projects/fortune-metropolis-billing`
- `METROPOLIS_ROOT=/home/kit/Metropolis`
- `.venv` uvicorn: `app.main:app --host 0.0.0.0 --port 8000`

## Layout

```
app/main.py                 HTTP API and dashboard
app/clp_parser.py           CLP PDF parser
app/citybase_model.py       live allocation from the master
app/workbook_updater.py     Elect Charge and Cost Sheet writes
app/dashboard_data.py       month report and workbook downloads
app/consumption_history.py  electricity_history.xlsx
app/master_parser.py        AC DEPT workbook reader
app/excel_exporter.py       meter-log spreadsheet
static/index.html           dashboard
static/i18n.js              English and Traditional Chinese
tests/test_dashboard.py
tests/test_consumption_history.py
tests/test_consumption_page.py
tests/test_language.py
tests/test_elect_charge_ledger.py
tests/test_cost_sheet_months.py
```

`tests/` also has parser, meter-unit, and citybase tests. `app/allocation_engine.py` is not the month-end path.

## API

| Method | Path | Role |
|--------|------|------|
| GET | `/` | Dashboard HTML |
| GET | `/i18n.js` | UI strings |
| GET | `/api/health` | `{"status":"ok","version":"1.0.0-part1"}` |
| GET | `/api/config` | Root path, the 15 accounts, check meter `6681757` |
| GET | `/api/status` | In-memory month, parse count, gate |
| POST | `/api/set-month` | JSON `{"month":"YYYY-MM"}` |
| POST | `/api/upload-bills` | Multipart PDF upload. Temp-parsed, not stored |
| POST | `/api/meter-log` | Form `previous`, `present`, optional `read_date` for meter `6681757` |
| GET | `/api/download-meter-log` | Meter-log spreadsheet |
| POST | `/api/process-month-end` | Copy previous workbooks and write this month |
| GET | `/api/months` | Months with a report or output workbook |
| GET | `/api/dashboard/{YYYY-MM}` | Saved consumption and allocation for that month |
| GET | `/api/charts` | Query `months` (comma-separated), `account`, `centre` |
| GET | `/api/history/items?kind=` | Buttons. `kind` is `bill`, `meter`, or `centre` |
| GET | `/api/history/item?kind=&item=` | kWh and cost for one clicked item |
| GET | `/api/history` | Query `account`, `meter`. Series from the history workbook |
| GET | `/api/history/download` | `electricity_history.xlsx` |
| GET | `/api/download/{YYYY-MM}/cost-allocation` | That month’s Cost Allocation file |
| GET | `/api/download/{YYYY-MM}/cost-sheet` | That month’s Cost Sheet file |
| POST | `/api/reset` | Clear the in-memory session |

`POST /api/scan-folder` is still registered and is not used by the dashboard. Intake is `POST /api/upload-bills`.
