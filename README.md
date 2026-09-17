# Metropolis Electricity Cost Allocation App

**FastAPI web application for allocating CLP electricity bills to cost centres using live monthly master workbook weights.**

[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/chunghonkit/fortune-metropolis-billing)
[![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.104-green.svg)](https://fastapi.tiangolo.com/)

## 🎯 Product Goal

Upload ~15 CLP electricity bill PDFs for one month → parse → allocate to cost centres using Kit-approved fixed-% rules → download an Excel workbook matching the Citybase cost sheet (year rollup column + per-account ELECTRICITY COST ALLOCATION forms).

**First live target month:** August 2026  
**Dry-run month:** July 2026 (fixtures validated)

---

## ⚠️ CRITICAL POLICY CHANGE

### Allocation % Must Be Recomputed Each Month from Live Master Workbook

**NOT frozen July-2026 constants.**

#### Runtime Source of Truth
- **Live master workbook** (Cost Allocation master.xlsx/xls)
- **AC DEPT sheet, Column C** → `weight / Σ(weights)` per account
- Recompute every month from the current master supplied by Kit

#### Frozen JSON/MD Files
- `metropolis_allocation_rules.json`
- `allocation_config.json`
- `uploads/METROPOLIS_ALLOCATION_RULES.md`

**These are FIXTURES/TESTS ONLY**, not runtime source of truth.

---

## 🚀 Quick Start

### Prerequisites

- Python 3.8+
- pip

### Installation

```bash
# Clone the repository
git clone https://github.com/chunghonkit/fortune-metropolis-billing.git
cd fortune-metropolis-billing

# Install dependencies
pip3 install -r requirements.txt

# Run the app
export PATH="/home/ubuntu/.local/bin:$PATH"  # If needed
python3 -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

### Access the App

Open browser to: **http://localhost:8000**

---

## 📖 Usage

### Step 1: Upload Master Workbook

Upload the **current Cost Allocation master workbook** for the target month.

**Required sheets:**
- `AC DEPT` (runtime source of truth for weights)
- `Allocation` (fallback/supplementary)

**Default master:** A sample master is pre-loaded at `data/masters/cost_allocation_master.xlsx` (generated from July 2026 allocation_config.json for testing).

**⚠️ For production:** Upload the current master each month. Weights are recomputed live from AC DEPT column C.

### Step 2: Upload CLP Bills

Upload ~15 CLP electricity bill PDFs for the target month.

**Format:** Standard CLP bills (PDF)  
**Month:** Enter month label (e.g., "August 2026")

### Step 3: Run Allocation

Click **"Run Allocation"** to:
1. Compute weights from master workbook (live, not frozen)
2. Allocate electricity costs to cost centres
3. Validate residuals within ±0.01 HKD tolerance

### Step 4: Download Excel

Download Citybase-format Excel workbook containing:
- **Year Rollup Sheet:** 27 cost centres in FC→SA row order
- **Per-Account Forms:** ELECTRICITY COST ALLOCATION forms for each account

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

## 🧪 Testing

### Run Unit Tests

```bash
# Run all tests
python3 -m pytest tests/ -v

# Run allocation tests only
python3 -m pytest tests/test_allocation.py -v

# Run with coverage
python3 -m pytest tests/ --cov=app --cov-report=html
```

### Test Data

- **July 2026 fixture bills:** Validated ground truth (1,332,208 kWh total)
- **allocation_config.json:** July 2026 snapshot for test fixtures
- **Sample master:** `data/masters/cost_allocation_master.xlsx` (generated from allocation_config.json)

**⚠️ Tests use frozen July 2026 data.** Production uses live master workbook.

---

## 📋 Master Workbook Format

### AC DEPT Sheet (Runtime Source of Truth)

| Column | Content | Example |
|--------|---------|---------|
| A or B | Account/Meter ID | 55861-52267-1 |
| B or C | Cost Centre | AC |
| **C** | **Weight** (source of truth) | **525.48** |

**Percentage Calculation:**

```
AC%  = 525.48 / (525.48 + 198.99) = 72.54%
SW%  = 198.99 / (525.48 + 198.99) = 27.46%
```

### Allocation Sheet (Fallback)

Pre-computed percentages (used if account not in AC DEPT).

### Cost Allocat Sheet

May contain `#REF!` errors (2×) – these are non-blocking and safely ignored per `config.json`.

---

## 🐛 Known Issues & Open Questions

### Open Questions (Config-Gated)

| # | Question | Status | Config Flag |
|---|----------|--------|-------------|
| Q1 | Meter ID mapping (9092771/9091324 vs 7662756/7662761) | Pending Kit | `meter_id_mapping.apply_mapping` |
| Q2 | FiT account 9144816: centre C vs DC | Pending Kit | `fit_account_9144816_centre.current_value` |

**Default behavior:** Q1 mapping OFF, Q2 centre = C (until Kit confirms).

### Non-Blocking Issues

- **Cost Allocat #REF! (2×):** Ignored per config (`ignore_ref_errors: true`)
- **Missing PDFs:** If < 15 bills, validation warns but continues

---

## 📦 Dependencies

See [`requirements.txt`](requirements.txt) for full list.

**Key dependencies:**
- **FastAPI 0.104:** Web framework
- **uvicorn:** ASGI server
- **PyMuPDF (fitz):** PDF parsing
- **openpyxl:** Excel reading/writing
- **pandas:** Data processing

---

## 🗂️ File Structure

```
.
├── app/                          # Application code
│   ├── main.py                   # FastAPI app
│   ├── clp_parser.py             # CLP bill parser
│   ├── master_parser.py          # Master workbook parser (AC DEPT column C)
│   ├── allocation_engine.py      # Allocation engine (live weights)
│   └── excel_exporter.py         # Citybase Excel exporter
├── static/                       # Static web UI
│   └── index.html                # Web interface
├── data/                         # Data files
│   └── masters/                  # Sample/uploaded master workbooks
├── tests/                        # Unit tests
│   └── test_allocation.py        # Allocation tests (July 2026 fixtures)
├── fixtures/                     # Test fixtures (planned)
├── config.json                   # Runtime configuration
├── requirements.txt              # Python dependencies
├── README.md                     # This file
├── allocation_config.json        # July 2026 fixture (TEST ONLY)
├── clp_parser_v2_final.py        # Original CLP parser (copied to app/)
├── ELECTRICITY_PARSING_RULES.md  # CLP parsing rules documentation
└── index.html / Index.html       # Legacy client-side demo (replaced)
```

---

## 📝 Validation Report

See `uploads/MASTER_VALIDATION.md` for full validation details.

**Summary:**
- ✅ Citybase cache: 0 mismatches
- ✅ Allocation!O spot-checks: 32/32 pass
- ✅ AC DEPT column C weights verified
- ✅ FC 100% hard rule confirmed

---

## 🔄 Monthly Workflow

1. **Receive current Cost Allocation master from Kit** (Excel .xlsx or .xls)
2. **Collect ~15 CLP bill PDFs** for the target month
3. **Upload master + bills** to the app
4. **Run allocation** (weights computed live from master AC DEPT column C)
5. **Validate** residuals within ±0.01 HKD
6. **Download Citybase Excel** workbook
7. **Submit** to Citybase Property Management

**⚠️ Critical:** Always use the **current master** for the month. Do not reuse old masters.

---

## 🤝 Contributing

### Updating Allocation Rules

1. Rules come from the **live master workbook** (AC DEPT column C)
2. Do NOT hardcode percentages in code
3. Update master workbook via Kit
4. Test with July 2026 fixtures before production use

### Adding Features

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Commit changes: `git commit -m "Add feature"`
4. Push: `git push origin feature/my-feature`
5. Open a Pull Request

---

## 📄 License

Internal use only - Citybase Property Management / Fortune Metropolis

---

## 📞 Contact

**Excel Allocator / Kit Chung**  
For questions about:
- Master workbook updates
- Allocation rule changes
- Open questions (Q1, Q2)

**Repository:**  
https://github.com/chunghonkit/fortune-metropolis-billing

---

## 🎉 Acknowledgments

- **Kit Chung** - Excel Allocator, master workbook validation
- **Citybase Property Management** - Cost sheet format requirements
- **CLP** - Electricity bill format (22 parsing rules)

---

## 🔖 Version History

### v1.0.0 (September 2026)
- ✅ Live master workbook integration (AC DEPT column C → weights)
- ✅ CLP bill PDF parser (reuse clp_parser_v2_final.py)
- ✅ Allocation engine with ±0.01 HKD residual tolerance
- ✅ Citybase Excel export (year rollup + per-account forms)
- ✅ FastAPI web UI
- ✅ July 2026 fixture tests
- ✅ Config-gated open questions (Q1, Q2)
- ✅ FC 100% hard rule
- ⚠️ Open: Q1 meter mapping, Q2 FiT centre (pending Kit)

### v0.x (Legacy)
- Client-side demo (index.html + PDF.js)
- Frozen allocation_config.json (replaced by live master)

---

**⚡ Metropolis Electricity Cost Allocation v1.0 | Fortune Metropolis | Citybase Property Management**
