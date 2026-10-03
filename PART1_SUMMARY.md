# Part 1 Implementation Summary

## ✅ Completed (2026-09-28)

### Product Flow Reshaped
- **Old flow (Draft PR #1):** Master-first 4-step (Upload Master → Upload Bills → Allocate → Download Excel)
- **New flow (Part 1):** Bills-first intake (Set Month → Scan/Upload Bills → Fill Meter Log → Gate Check)

### Features Implemented

1. **Billing Month Picker**
   - YYYY-MM format validation (e.g., `2025-04`)
   - Required before scan/upload
   - Session tracking

2. **Two Bill Intake Modes**
   - **Mode A (Preferred for Omarchy):** Designated folder scan
     - Default: `~/Metropolis/{YYYY-MM}/bills/`
     - Configurable via `METROPOLIS_ROOT` env var
     - Scan button parses all PDFs found
   - **Mode B (Fallback):** Browser upload
     - Multi-PDF dropzone
     - Same validation as folder mode

3. **Hard Gate Validation**
   - Exactly 15 expected Metropolis CLP accounts
   - Same billing month (extracted from bill dates)
   - No duplicates
   - No unrecognised accounts
   - Meter log complete (present ≥ previous)
   - **Continue button disabled until gate PASS**

4. **Meter Log for Check-Meter 6681757**
   - Form fields: previous/present readings + optional read date
   - Validation: present must be ≥ previous
   - Drives account 55861-52267-1 AC/SW allocation (Part 2)

5. **Meter Log Excel Generation**
   - Download `meter_log_YYYY-MM.xlsx`
   - CheckMeter sheet: meter 6681757, previous/present/delta, read date
   - Session sheet: billing month, 15 PDFs ↔ accounts, validation status
   - Written to `~/Metropolis/{YYYY-MM}/out/` when folder mode

6. **UI Checklist (15 Accounts)**
   - Live status for each expected account:
     - ✅ Matched (correct month)
     - ❌ Missing
     - ⚠️ Duplicate
     - ⚠️ Wrong month
     - ❓ Unrecognised (not in Metropolis 15)

7. **Session Storage**
   - In-memory session stores parsed bill summaries
   - Ready for Part 2 (no re-upload required)

### Tests Written (21 pass)

**`tests/test_part1.py`:**
- Gate PASS with 15 mocked accounts
- Gate FAIL on missing account
- Gate FAIL on duplicate account
- Gate FAIL on wrong month
- Gate FAIL on invalid meter log
- Meter log completeness validation
- Billing month extraction from bill dates
- Account checklist status display

**Test results:**
```
21 passed in 0.29s
```

### Files Modified/Created

**Modified:**
- `app/main.py` - Reshaped for Part 1 intake & gate
- `static/index.html` - New UI (bills + meter log first)
- `README.md` - Updated for Part 1 + local Omarchy usage

**Created:**
- `tests/test_part1.py` - 21 unit/integration tests

**Preserved (not modified):**
- `app/clp_parser.py` - Reused for bill parsing
- Other modules (master_parser, allocation_engine, excel_exporter) - TODO for Parts 2-3

### API Endpoints (Part 1)

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Web UI (Part 1) |
| `/api/config` | GET | App config (METROPOLIS_ROOT, expected accounts) |
| `/api/status` | GET | Current Part 1 status |
| `/api/set-month` | POST | Set billing month (YYYY-MM) |
| `/api/scan-folder` | POST | Scan designated folder for PDFs |
| `/api/upload-bills` | POST | Upload PDFs (browser fallback) |
| `/api/meter-log` | POST | Set meter log (6681757 previous/present) |
| `/api/download-meter-log` | GET | Download meter_log_YYYY-MM.xlsx |
| `/api/reset` | POST | Reset Part 1 session |

### Local Omarchy Test

**Start app:**
```bash
python3 -m uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

**Access:**
```
http://127.0.0.1:8000
```

**Example workflow:**
1. Create folder: `mkdir -p ~/Metropolis/2025-04/bills`
2. Place 15 CLP bill PDFs in `~/Metropolis/2025-04/bills/`
3. Open app at http://127.0.0.1:8000
4. Set month to `2025-04`
5. Click **Scan Folder**
6. Fill meter log (6681757: previous/present)
7. Check gate → **GATE PASSED**

### Git & PR

**Branch:** `cursor/metropolis-allocation-app`  
**PR:** #1 (updated) - https://github.com/chunghonkit/fortune-metropolis-billing/pull/1  
**Commit:** `c4487e9` - Implement Part 1: Bills + Meter Log Intake & Gate

**PR Status:** Draft (Part 1 complete, Parts 2-3 TODO)

---

## 🚫 NOT in Part 1 (TODO for Parts 2-3)

### Part 2 (TODO)
- Cost Allocation master upload
- Live % recomputation from AC DEPT column C
- Power master (BI) export (bill_monthly, meter_monthly, check_meter_log, etc.)
- Detail cost allocation (audit) export (allocation_by_account_centre, etc.)
- Allocation engine (FiT gross base, office chillers 92/8, FC 100%, AC/SW live)

### Part 3 (TODO)
- In-place update Citybase Excel templates
- Cost Sheet: year-grid month column + per-account forms
- Cost Allocation Elect Charge: check-meter readings + month/date/kWh
- Preserve format, formulas, other months' columns

**Part 1 does NOT run allocation.** It only validates intake and prepares session.

---

## 📋 Known Issues (Part 1)

- FC label: FC = **Food Court** (not "Fitness Centre") - documented in SPEC §2.4
- Month format: Only `YYYY-MM` supported (no `MM-YYYY` or `April 2025`)
- Session storage: In-memory only (resets on server restart)

---

## ✅ Part 1 Acceptance Criteria (All Met)

- [x] User can select month and either scan folder or upload PDFs
- [x] Folder mode works with `~/Metropolis/YYYY-MM/bills/`
- [x] Missing account → Continue disabled
- [x] Duplicate / wrong-month / unrecognised called out
- [x] Meter log produces `meter_log_YYYY-MM.xlsx`
- [x] Session stores parsed bill summaries for Part 2-3
- [x] No allocation math in Part 1
- [x] Works locally with uvicorn on 127.0.0.1:8000

---

## 🎯 Next Steps

1. **Implement Part 2:** Master workbook upload, power master export, allocation detail export
2. **Implement Part 3:** In-place Citybase Excel template updater
3. **Golden pack test:** Validate with real Apr/May 2025 bills on Omarchy

---

**Part 1 is complete, tested, and ready for local Omarchy smoke test.**

**⚡ Metropolis Electricity Cost Allocation v1.0.0-part1**
