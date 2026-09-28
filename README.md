# Metropolis Electricity Cost Allocation App

**FastAPI web application for allocating CLP electricity bills to cost centres - Part 1 Intake & Gate**

[![Version](https://img.shields.io/badge/version-1.0.0--part1-blue.svg)](https://github.com/chunghonkit/fortune-metropolis-billing)
[![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.104-green.svg)](https://fastapi.tiangolo.com/)

## 🎯 Product Flow (SPEC v1.4)

Three-step web app for Metropolis electricity cost allocation:

| Part | Status | Description |
|------|--------|-------------|
| **Part 1** | ✅ **SHIPPED** | Bills + Meter Log First → Gate validation (15 accounts + meter log) |
| Part 2 | 📋 TODO | Master workbook → Power master (BI) + detail allocation (audit) |
| Part 3 | 📋 TODO | In-place update Citybase Excel templates (Cost Sheet + Elect Charge) |

**Current implementation:** Part 1 only  
**Test target:** Kit's Omarchy notebook (local `uvicorn` on `localhost:8000`)

---

## 🚀 Quick Start (Local Omarchy Test)

### Prerequisites

- Python 3.8+
- pip

### Installation & Run

```bash
# Clone the repository
git clone https://github.com/chunghonkit/fortune-metropolis-billing.git
cd fortune-metropolis-billing

# Install dependencies
pip3 install -r requirements.txt

# Run the app (localhost only for Omarchy)
python3 -m uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

### Access the App

Open browser to: **http://127.0.0.1:8000**

### Folder Layout for Part 1 (Designated Folder Mode)

Default root: `~/Metropolis` (configurable via `METROPOLIS_ROOT` env var)

```
~/Metropolis/
  2025-04/                    # Billing month (YYYY-MM)
    bills/                    # Place exactly 15 CLP bill PDFs here
    meter_log/                # Optional: meter log photo / Excel
    masters/                  # Optional: Cost Allocation master (Part 2)
    out/                      # App writes: meter_log_YYYY-MM.xlsx, etc.
  2025-05/
    bills/
    ...
```

**Example workflow:**

1. Create folder: `mkdir -p ~/Metropolis/2025-04/bills`
2. Place 15 CLP bill PDFs in `~/Metropolis/2025-04/bills/`
3. Open app at http://127.0.0.1:8000
4. Set month to `2025-04`
5. Click **Scan Folder**
6. Fill meter log (check-meter 6681757: previous/present readings)
7. Check gate status → Continue when PASSED

---

## 📖 Part 1 Usage

### Step 1: Set Billing Month

Enter billing month in `YYYY-MM` format (e.g., `2025-04`). **Required before scan/upload.**

### Step 2: CLP Bills Intake (Two Modes)

#### Mode A: Folder Scan (Preferred for Omarchy)

- Place 15 CLP bill PDFs in `~/Metropolis/{YYYY-MM}/bills/`
- Click **Scan Folder** in the app
- App automatically parses all PDFs found

#### Mode B: Browser Upload (Fallback)

- Click **Browser Upload** tab
- Select multiple PDF files
- Click **Upload & Parse Bills**

### Step 3: Meter Log

Fill meter log for check-meter **6681757** (散熱水泵電):

- **Previous reading:** Integer (e.g., 1234567)
- **Present reading:** Integer, must be ≥ previous (e.g., 1235890)
- **Read date:** Optional (YYYY-MM-DD)

Click **Save Meter Log**.

Download generated `meter_log_YYYY-MM.xlsx` (CheckMeter + Session sheets).

### Step 4: Gate Validation

Click **Check Gate** to validate:

✅ **GATE PASS conditions:**
- Exactly 15 expected Metropolis accounts present
- All bills from the same billing month
- No duplicates
- No unrecognised accounts
- Meter log complete (present ≥ previous)

❌ **GATE FAIL:** Continue button disabled until all conditions met.

### Expected 15 Accounts (Metropolis)

| # | Account | Typical Role |
|---|---------|--------------|
| 1 | 55861-52267-1 | Retail chillers → AC/SW (live) |
| 2 | 24096-78457-6 | Multi-centre (office power etc.) |
| 3 | 13639-58422-3 | Office chiller → AO 92 / OC 8 |
| 4 | 40722-61440-7 | Multi-centre (carpark-heavy) |
| 5 | 35204-69738-4 | Multi-centre (carpark / shared) |
| 6 | 97968-02236-6 | DC 100% |
| 7 | 23529-59279-9 | SW 100% |
| 8 | 79292-23337-6 | Office chiller → AO 92 / OC 8 |
| 9 | 52167-13569-2 | FiT (retail) — gross base |
| 10 | 08731-83914-5 | DC 100% |
| 11 | 00776-78552-1 | FiT — gross base |
| 12 | 88931-57029-6 | C 100% |
| 13 | 70873-85471-3 | SA 100% |
| 14 | 72399-00664-9 | SA 100% |
| 15 | 82805-94744-7 | **Food Court** → FC 100% |

**Note:** FC = Food Court (not "Fitness Centre").

---

## ⚠️ Part 1 Scope

**Implemented in Part 1:**
- ✅ Billing month picker (YYYY-MM) required first
- ✅ Designated folder scan (preferred for Omarchy)
- ✅ Browser upload fallback
- ✅ Hard gate: 15 accounts, same month, no duplicates, meter log complete
- ✅ Meter log form + Excel generation
- ✅ Session storage (parsed bill summaries)
- ✅ UI checklist of 15 accounts with status (matched/missing/duplicate/wrong-month/unrecognised)
- ✅ Modern, dynamic SaaS UI with dark theme toggle

**NOT in Part 1 (TODO for Parts 2-3):**
- ❌ Cost Allocation master upload
- ❌ Live % recomputation from AC DEPT
- ❌ Allocation engine (FiT gross base, office chillers 92/8, etc.)
- ❌ Power master (BI) export
- ❌ Detail cost allocation (audit) export
- ❌ Citybase Cost Sheet / Elect Charge Excel export

**Part 1 does NOT run allocation.** It only validates intake and prepares session for Part 2.

---

## 🎨 Modern UI Design

Part 1 features a redesigned, contemporary SaaS interface:

### Design Highlights
- **Clean Typography:** Modern font stack with clear hierarchy
- **Soft Depth:** Subtle shadows and borders for visual separation
- **Status Chips:** Color-coded badges for account status (matched/missing/duplicate/wrong-month/unrecognised)
- **Dynamic Updates:** Live checklist updates as files are scanned/parsed
- **Animated Gate Banner:** Smooth transition from blocked → PASSED state
- **Progress Indicators:** Pulse animations and spinners during operations
- **Dark Theme Toggle:** Optional dark mode for comfortable viewing
- **Responsive Layout:** Desktop-first, usable on laptops (768px+)
- **Polished States:** Professional empty, loading, error, and success states
- **Accessible:** High contrast, keyboard navigation, screen reader friendly

### UI Components
- **Step Rail:** Visual indicator showing Part 1 active, Parts 2-3 locked
- **Card-based Layout:** Organized sections with clear separation
- **Tabbed Interface:** Folder scan / Browser upload toggle
- **Stats Grid:** Quick overview of gate status (matched, missing, duplicates, etc.)
- **Info Boxes:** Contextual help with folder path, hints
- **Button Groups:** Organized actions with clear hierarchy

### Tech Stack
- **Vanilla HTML/CSS/JS** - No heavy frameworks, simple deployment
- **CSS Custom Properties** - Easy theming with light/dark modes
- **Smooth Transitions** - 0.2-0.4s animations for polished feel
- **System Font Stack** - Fast loading, native feel

---

## 🏗️ Architecture

### Components

```
app/
├── main.py              # FastAPI server
├── clp_parser.py        # CLP bill PDF parser (reuses clp_parser_v2_final.py)
├── master_parser.py     # Master workbook parser (AC DEPT column C → weights)
├── allocation_engine.py # Cost allocation engine
└── excel_exporter.py    # Citybase Excel format exporter

static/
└── index.html           # Web UI

data/
└── masters/             # Sample/default master workbooks

tests/
└── test_allocation.py   # Unit tests (July 2026 fixtures)

config.json              # Runtime configuration (open questions, special cases)
```

### API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Web UI |
| `/api/health` | GET | Health check |
| `/api/status` | GET | Current session status |
| `/api/upload-master` | POST | Upload master workbook |
| `/api/upload-bills` | POST | Upload CLP bill PDFs |
| `/api/allocate` | POST | Run allocation |
| `/api/download-excel` | GET | Download Citybase Excel |
| `/api/reset` | POST | Reset session |

---

## 🔧 Configuration

### `config.json`

Runtime configuration for special cases and open questions.

#### FC 100% Hard Rule

```json
{
  "fc_100_percent_hard_rule": true
}
```

FC (Fitness Centre) accounts remain 100% allocated to FC, outside master workbook rules.

#### Open Questions (Pending Kit Confirmation)

**Q1: Meter ID Mapping**

Master 9092771/9091324 vs bill meters 7662756/7662761

```json
{
  "open_questions": {
    "meter_id_mapping": {
      "description": "Q1: Master 9092771/9091324 vs bill meters 7662756/7662761",
      "mappings": {
        "7662756": "9092771",
        "7662761": "9091324"
      },
      "apply_mapping": false,
      "status": "pending_kit_confirmation"
    }
  }
}
```

**Q2: FiT Account 9144816 Centre**

Centre C vs DC for FiT account/meter 9144816

```json
{
  "open_questions": {
    "fit_account_9144816_centre": {
      "description": "Q2: FiT account/meter 9144816 - centre C vs DC",
      "options": ["C", "DC"],
      "current_value": "C",
      "status": "pending_kit_confirmation"
    }
  }
}
```

**To enable:** Set `apply_mapping: true` or `current_value: "DC"` when Kit confirms.

#### Validation

```json
{
  "validation": {
    "residual_tolerance_hkd": 0.01,
    "ignore_ref_errors": true
  }
}
```

- **residual_tolerance_hkd:** ±0.01 HKD tolerance for allocation residuals
- **ignore_ref_errors:** Safely ignore `#REF!` errors in Cost Allocat sheet (2× non-blocking)

---

## 🧪 Testing (Part 1)

### Run Unit Tests

```bash
# Run Part 1 tests
python3 -m pytest tests/test_part1.py -v

# Run with coverage
python3 -m pytest tests/ --cov=app --cov-report=html
```

### Test Cases for Part 1

- ✅ Gate PASS with 15 mocked accounts
- ✅ Gate FAIL on missing account
- ✅ Gate FAIL on duplicate account
- ✅ Gate FAIL on wrong month
- ✅ Gate FAIL on invalid meter log (present < previous)
- ✅ Folder scan finds PDFs under temp METROPOLIS_ROOT
- ✅ Meter log completeness validation

**Test data:** Use fixtures or mocks for 15 expected accounts. Golden PDFs not required in repo for Part 1.

---

## 🐛 Known Issues (Part 1)

- **FC label:** FC = **Food Court**, not "Fitness Centre" (documented in SPEC §2.4)
- **Month format:** Only `YYYY-MM` supported (e.g., `2025-04`); no `MM-YYYY` or `April 2025`

**Parts 2-3 not implemented yet** — no allocation, no Citybase Excel export in Part 1.

---

## 📦 Dependencies

See [`requirements.txt`](requirements.txt) for full list.

**Key dependencies:**
- **FastAPI 0.104:** Web framework
- **uvicorn:** ASGI server
- **PyMuPDF (fitz):** PDF parsing
- **openpyxl:** Excel reading/writing (meter log generation)

---

## 🗂️ File Structure (Part 1)

```
.
├── app/                          # Application code
│   ├── main.py                   # FastAPI app (Part 1 intake & gate)
│   ├── clp_parser.py             # CLP bill PDF parser
│   └── __init__.py
├── static/                       # Static web UI
│   └── index.html                # Part 1 UI (bills + meter log first)
├── tests/                        # Unit tests
│   └── test_part1.py             # Part 1 gate & folder scan tests (TODO)
├── requirements.txt              # Python dependencies
├── README.md                     # This file
├── clp_parser_v2_final.py        # Original CLP parser (copied to app/)
└── SPEC.md / uploads/SPEC.md     # Product spec v1.4 (3-part flow)
```

**Parts 2-3 modules (not yet implemented):**
- `app/master_parser.py` - Master workbook parser
- `app/allocation_engine.py` - Allocation engine
- `app/excel_exporter.py` - Citybase Excel exporter

---

## 🔄 Part 1 vs Old 4-Step Flow

| Old Flow (PR #1 draft) | New Flow (Part 1) |
|------------------------|-------------------|
| Step 1: Upload Master | Step 1: Set Month **first** |
| Step 2: Upload Bills | Step 2: Scan Folder **or** Upload Bills |
| Step 3: Allocate | Step 3: Fill Meter Log |
| Step 4: Download Excel | Step 4: Gate Check → Continue to Part 2 (TODO) |

**Key change:** Part 1 reshapes the flow to **bills + meter log first**, not master-first. Gate validation blocks Part 2 until intake is complete.

---

## 📝 Version History

### v1.0.0-part1 (September 2026)

**Shipped:**
- ✅ Part 1: Bills + meter log first intake flow
- ✅ Designated folder scan (preferred for Omarchy)
- ✅ Browser upload fallback
- ✅ Hard gate: 15 accounts, same month, no duplicates, meter log complete
- ✅ Meter log form + Excel generation (`meter_log_YYYY-MM.xlsx`)
- ✅ Session storage (parsed bill summaries)
- ✅ UI checklist of 15 accounts with status
- ✅ CLP bill PDF parser (reuse `clp_parser_v2_final.py`)
- ✅ Local Omarchy test: `uvicorn` on `127.0.0.1:8000`

**Not in Part 1 (TODO Parts 2-3):**
- ❌ Cost Allocation master upload
- ❌ Live % recomputation from AC DEPT
- ❌ Allocation engine
- ❌ Power master (BI) + detail allocation (audit)
- ❌ Citybase Cost Sheet / Elect Charge Excel export

### v0.x (Legacy - PR #1 draft)

- Master-first 4-step flow (replaced by bills-first Part 1)
- Allocation engine + Excel export (moved to Parts 2-3)

---

## 🤝 Contributing

### Extending to Part 2

1. Implement `app/master_parser.py` (AC DEPT column C → weights)
2. Implement `app/allocation_engine.py` (locked rules from SPEC §2)
3. Generate power master (BI) + detail cost allocation (audit) Excel/Google Sheets
4. Update UI to add Part 2 step after Part 1 gate passes

### Extending to Part 3

1. Implement in-place Citybase Excel template updater
2. Write only changed cells (preserve format, formulas, other months)
3. Cost Sheet forms + Elect Charge month/date/kWh updates

---

## 📄 License

Internal use only - Citybase Property Management / Fortune Metropolis

---

## 📞 Contact

**Excel Allocator / Kit Chung**  
For questions about:
- Metropolis accounts or allocation rules
- Cost Allocation master workbook
- Citybase Cost Sheet templates

**Repository:**  
https://github.com/chunghonkit/fortune-metropolis-billing

---

## 🎉 Acknowledgments

- **Kit Chung** - Excel Allocator, master workbook validation, Metropolis rules
- **Citybase Property Management** - Cost sheet format requirements
- **CLP** - Electricity bill format (22 parsing rules from `clp_parser_v2_final.py`)

---

**⚡ Metropolis Electricity Cost Allocation v1.0.0-part1 | Part 1 Only | Local Omarchy Test Ready**
