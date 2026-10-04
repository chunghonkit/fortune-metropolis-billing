"""
Metropolis Electricity Cost Allocation App - Main API Server
Part 1: Bills + Meter Log Intake & Gate

FastAPI application for:
Part 1: Bill intake (folder scan or upload) + meter log + gate validation
Part 2: Master workbook + power master + allocation detail (TODO)
Part 3: Citybase Excel export (TODO)
"""

from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Body
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from typing import List, Optional, Dict, Any
import tempfile
import os
import shutil
from pathlib import Path
import logging
from datetime import datetime
import json
from collections import defaultdict

from app.clp_parser import parse_clp_bill

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Metropolis Electricity Cost Allocation - Part 1",
    description="Bills + Meter Log First → Gate → Parts 2-3 TODO",
    version="1.0.0-part1"
)

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Expected 15 CLP accounts for Metropolis (SPEC §3)
EXPECTED_ACCOUNTS = [
    "55861-52267-1",  # Retail chillers → AC/SW (live)
    "24096-78457-6",  # Multi-centre (office power etc.)
    "13639-58422-3",  # Office chiller → AO 92 / OC 8
    "40722-61440-7",  # Multi-centre (carpark-heavy)
    "35204-69738-4",  # Multi-centre (carpark / shared)
    "97968-02236-6",  # DC 100%
    "23529-59279-9",  # SW 100%
    "79292-23337-6",  # Office chiller → AO 92 / OC 8
    "52167-13569-2",  # FiT (retail) — gross base
    "08731-83914-5",  # DC 100%
    "00776-78552-1",  # FiT — gross base
    "88931-57029-6",  # C 100%
    "70873-85471-3",  # SA 100%
    "72399-00664-9",  # SA 100%
    "82805-94744-7",  # Food Court → FC 100%
]

# Get METROPOLIS_ROOT from env, default to ~/Metropolis
def get_metropolis_root():
    root = os.environ.get("METROPOLIS_ROOT", "~/Metropolis")
    return Path(root).expanduser()

# Session storage (in-memory for v1; write to disk under out/)
session_storage = {
    'billing_month': None,  # YYYY-MM
    'folder_mode': False,
    'folder_path': None,
    'parsed_bills': [],
    'gate_status': None,
    'meter_log': {
        'meter_no': '6681757',
        'label': '散熱水泵電',
        'previous': None,
        'present': None,
        'read_date': None,
        'photo_path': None,
    },
    'session_id': None,
    'created_at': None,
}


@app.get("/", response_class=HTMLResponse)
async def root():
    """Serve the web UI"""
    return FileResponse("static/index.html")


@app.get("/api/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "ok", "version": "1.0.0-part1"}


@app.get("/api/config")
async def get_config():
    """Get app configuration"""
    root = get_metropolis_root()
    return {
        "metropolis_root": str(root),
        "expected_accounts": EXPECTED_ACCOUNTS,
        "expected_count": len(EXPECTED_ACCOUNTS),
        "check_meter_no": "6681757",
        "check_meter_label": "散熱水泵電",
    }


@app.get("/api/status")
async def get_status():
    """Get current Part 1 status"""
    gate = validate_gate()
    return {
        "part": 1,
        "billing_month": session_storage['billing_month'],
        "folder_mode": session_storage['folder_mode'],
        "folder_path": str(session_storage['folder_path']) if session_storage['folder_path'] else None,
        "bills_parsed": len(session_storage['parsed_bills']),
        "gate_passed": gate['passed'],
        "gate_status": gate,
        "meter_log_complete": is_meter_log_complete(),
        "session_id": session_storage['session_id'],
    }


@app.post("/api/set-month")
async def set_billing_month(month: str = Body(..., embed=True)):
    """
    Set billing month (YYYY-MM format) - required before scan/upload.
    """
    try:
        # Validate YYYY-MM format
        datetime.strptime(month, "%Y-%m")
        session_storage['billing_month'] = month
        session_storage['session_id'] = f"session_{month}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        session_storage['created_at'] = datetime.now().isoformat()
        logger.info(f"Billing month set to: {month}")
        return {"success": True, "month": month, "session_id": session_storage['session_id']}
    except ValueError:
        raise HTTPException(400, "Invalid month format. Use YYYY-MM (e.g., 2025-04)")


def is_meter_log_complete():
    """Check if meter log is complete"""
    ml = session_storage['meter_log']
    if ml['previous'] is None or ml['present'] is None:
        return False
    if ml['present'] < ml['previous']:
        return False
    return True


def validate_gate():
    """
    Validate Part 1 gate conditions.
    Returns gate status dict with passed flag and details.
    """
    status = {
        'passed': False,
        'checks': [],
        'errors': [],
        'warnings': [],
        'account_checklist': {},
    }
    
    # Check 1: Billing month set
    if not session_storage['billing_month']:
        status['errors'].append("Billing month not set")
        status['checks'].append({'name': 'Billing month set', 'passed': False})
        return status
    else:
        status['checks'].append({'name': 'Billing month set', 'passed': True})
    
    # Check 2: Meter log complete
    meter_complete = is_meter_log_complete()
    status['checks'].append({'name': 'Meter log complete', 'passed': meter_complete})
    if not meter_complete:
        status['errors'].append("Meter log incomplete or invalid (present must be ≥ previous)")
    
    # Check 3: Parse all bills
    parsed_accounts = {}
    for bill in session_storage['parsed_bills']:
        account = bill.get('account')
        if account:
            if account not in parsed_accounts:
                parsed_accounts[account] = []
            parsed_accounts[account].append(bill)
    
    # Build checklist for all 15 expected accounts
    expected_month_ym = session_storage['billing_month']
    missing = []
    duplicates = []
    wrong_month = []
    matched = []
    
    for account in EXPECTED_ACCOUNTS:
        if account not in parsed_accounts:
            status['account_checklist'][account] = {'status': 'missing', 'bills': []}
            missing.append(account)
        elif len(parsed_accounts[account]) > 1:
            status['account_checklist'][account] = {'status': 'duplicate', 'bills': parsed_accounts[account]}
            duplicates.append(account)
        else:
            bill = parsed_accounts[account][0]
            # Check month match - extract YYYY-MM from bill dates
            bill_month = extract_bill_month(bill)
            if bill_month and bill_month != expected_month_ym:
                status['account_checklist'][account] = {'status': 'wrong_month', 'bills': [bill], 'bill_month': bill_month}
                wrong_month.append(account)
            else:
                status['account_checklist'][account] = {'status': 'matched', 'bills': [bill]}
                matched.append(account)
    
    # Check for unrecognised accounts
    unrecognised = []
    for account in parsed_accounts:
        if account not in EXPECTED_ACCOUNTS:
            unrecognised.append(account)
            status['account_checklist'][account] = {'status': 'unrecognised', 'bills': parsed_accounts[account]}
    
    # Summarize checks
    status['checks'].append({'name': f'Exactly 15 expected accounts', 'passed': len(matched) == 15 and not missing and not duplicates and not wrong_month})
    status['checks'].append({'name': 'No missing accounts', 'passed': len(missing) == 0})
    status['checks'].append({'name': 'No duplicates', 'passed': len(duplicates) == 0})
    status['checks'].append({'name': 'No wrong month', 'passed': len(wrong_month) == 0})
    status['checks'].append({'name': 'No unrecognised accounts', 'passed': len(unrecognised) == 0})
    
    # Add to errors
    if missing:
        status['errors'].append(f"Missing {len(missing)} accounts: {', '.join(missing[:5])}{'...' if len(missing) > 5 else ''}")
    if duplicates:
        status['errors'].append(f"Duplicate accounts: {', '.join(duplicates)}")
    if wrong_month:
        status['errors'].append(f"Wrong month: {', '.join(wrong_month)}")
    if unrecognised:
        status['errors'].append(f"Unrecognised accounts (not in Metropolis 15): {', '.join(unrecognised)}")
    
    # Summary stats
    status['summary'] = {
        'expected': len(EXPECTED_ACCOUNTS),
        'matched': len(matched),
        'missing': len(missing),
        'duplicates': len(duplicates),
        'wrong_month': len(wrong_month),
        'unrecognised': len(unrecognised),
    }
    
    # Gate PASS condition (strict: no unrecognised accounts)
    status['passed'] = (
        len(matched) == 15 and
        len(missing) == 0 and
        len(duplicates) == 0 and
        len(wrong_month) == 0 and
        len(unrecognised) == 0 and
        meter_complete
    )
    
    return status


def extract_bill_month(bill):
    """
    Extract YYYY-MM from bill dates.
    Uses to_date (period end) since billing month is the month the period ends in.
    Format: DD-MM-YY
    """
    to_date = bill.get('to_date')  # Use period END date, not start
    if to_date:
        try:
            # Parse DD-MM-YY
            parts = to_date.split('-')
            if len(parts) == 3:
                dd, mm, yy = parts
                # Convert YY to YYYY (assume 20YY for yy < 80, else 19YY)
                yyyy = f"20{yy}" if int(yy) < 80 else f"19{yy}"
                return f"{yyyy}-{mm}"
        except:
            pass
    return None


@app.post("/api/scan-folder")
async def scan_folder():
    """
    Scan designated folder for CLP bill PDFs.
    Folder path: METROPOLIS_ROOT/{YYYY-MM}/bills/
    """
    try:
        if not session_storage['billing_month']:
            raise HTTPException(400, "Please set billing month first")
        
        # Build folder path
        root = get_metropolis_root()
        month_folder = root / session_storage['billing_month'] / "bills"
        
        if not month_folder.exists():
            raise HTTPException(404, f"Folder not found: {month_folder}\nCreate the folder and place CLP bill PDFs inside.")
        
        # Scan for PDFs
        pdf_files = list(month_folder.glob("*.pdf"))
        if not pdf_files:
            raise HTTPException(404, f"No PDF files found in {month_folder}")
        
        logger.info(f"Found {len(pdf_files)} PDFs in {month_folder}")
        
        # Parse each PDF
        parsed_bills = []
        errors = []
        
        for pdf_path in pdf_files:
            try:
                bill = parse_clp_bill(str(pdf_path))
                if bill and bill.get('kwh'):
                    parsed_bills.append(bill)
                    logger.info(f"Parsed {pdf_path.name}: {bill.get('account')} - {bill.get('kwh')} kWh")
                else:
                    errors.append(f"{pdf_path.name}: Failed to parse (no KWH)")
            except Exception as e:
                logger.error(f"Error parsing {pdf_path.name}: {e}")
                errors.append(f"{pdf_path.name}: {str(e)}")
        
        # Update session
        session_storage['parsed_bills'] = parsed_bills
        session_storage['folder_mode'] = True
        session_storage['folder_path'] = month_folder
        
        # Validate gate
        gate = validate_gate()
        
        return {
            "success": True,
            "folder_path": str(month_folder),
            "pdfs_found": len(pdf_files),
            "bills_parsed": len(parsed_bills),
            "errors": errors,
            "gate_status": gate,
        }
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error scanning folder: {e}", exc_info=True)
        raise HTTPException(500, f"Error scanning folder: {str(e)}")


@app.post("/api/upload-bills")
async def upload_bills(files: List[UploadFile] = File(...)):
    """
    Upload CLP electricity bill PDFs (browser upload fallback).
    Billing month must be set first via /api/set-month.
    """
    try:
        if not session_storage['billing_month']:
            raise HTTPException(400, "Please set billing month first")
        
        if not files:
            raise HTTPException(400, "No files uploaded")
        
        parsed_bills = []
        errors = []
        
        # Parse each PDF
        for file in files:
            if not file.filename.endswith('.pdf'):
                errors.append(f"{file.filename}: Not a PDF")
                continue
            
            try:
                # Save to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                    shutil.copyfileobj(file.file, tmp)
                    tmp_path = tmp.name
                
                # Parse the bill
                bill = parse_clp_bill(tmp_path)
                
                # Clean up temp file
                os.remove(tmp_path)
                
                if bill and bill.get('kwh'):
                    parsed_bills.append(bill)
                    logger.info(f"Parsed {file.filename}: {bill.get('account')} - {bill.get('kwh')} kWh")
                else:
                    errors.append(f"{file.filename}: Failed to parse (no KWH)")
            
            except Exception as e:
                logger.error(f"Error parsing {file.filename}: {e}")
                errors.append(f"{file.filename}: {str(e)}")
        
        # Update storage
        session_storage['parsed_bills'] = parsed_bills
        session_storage['folder_mode'] = False
        session_storage['folder_path'] = None
        
        # Validate gate
        gate = validate_gate()
        
        return {
            "success": True,
            "bills_parsed": len(parsed_bills),
            "errors": errors,
            "gate_status": gate,
        }
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error uploading bills: {e}", exc_info=True)
        raise HTTPException(500, f"Error uploading bills: {str(e)}")


@app.post("/api/meter-log")
async def set_meter_log(
    previous: float = Form(...),
    present: float = Form(...),
    read_date: Optional[str] = Form(None),
):
    """
    Set meter log for check-meter 6681757 (散熱水泵電).
    
    Args:
        previous: Previous reading (float, up to 1 decimal place)
        present: Present reading (float, up to 1 decimal place, must be >= previous)
        read_date: Optional read date (YYYY-MM-DD or DD-MM-YY)
    """
    try:
        # Round to 1 decimal place
        previous = round(previous, 1)
        present = round(present, 1)
        
        if present < previous:
            raise HTTPException(400, "Present reading must be >= previous reading")
        
        session_storage['meter_log']['previous'] = previous
        session_storage['meter_log']['present'] = present
        session_storage['meter_log']['read_date'] = read_date
        
        # Validate gate after update
        gate = validate_gate()
        
        # Calculate delta with 1 decimal precision
        delta = round(present - previous, 1)
        
        return {
            "success": True,
            "meter_log": session_storage['meter_log'],
            "delta": delta,
            "gate_status": gate,
        }
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error setting meter log: {e}", exc_info=True)
        raise HTTPException(500, f"Error setting meter log: {str(e)}")


@app.get("/api/download-meter-log")
async def download_meter_log():
    """
    Generate and download meter_log_YYYY-MM.xlsx.
    Contains CheckMeter and Session sheets per SPEC §5.3.
    """
    try:
        if not session_storage['billing_month']:
            raise HTTPException(400, "No billing month set")
        
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment
        
        wb = Workbook()
        
        # Sheet 1: CheckMeter
        ws1 = wb.active
        ws1.title = "CheckMeter"
        
        # Headers
        headers = ["Meter No.", "Label", "Previous", "Present", "Delta", "Read Date", "Notes"]
        ws1.append(headers)
        for cell in ws1[1]:
            cell.font = Font(bold=True)
        
        # Data
        ml = session_storage['meter_log']
        delta = round(ml['present'] - ml['previous'], 1) if ml['present'] and ml['previous'] else None
        ws1.append([
            ml['meter_no'],
            ml['label'],
            ml['previous'],
            ml['present'],
            delta,
            ml['read_date'] or "",
            f"Drives account 55861-52267-1 AC/SW"
        ])
        
        # Sheet 2: Session
        ws2 = wb.create_sheet("Session")
        
        # Session info
        ws2.append(["Billing Month", session_storage['billing_month']])
        ws2.append(["Session ID", session_storage['session_id']])
        ws2.append(["Created At", session_storage['created_at']])
        ws2.append(["Mode", "Folder Scan" if session_storage['folder_mode'] else "Browser Upload"])
        if session_storage['folder_path']:
            ws2.append(["Folder Path", str(session_storage['folder_path'])])
        ws2.append([])
        
        # Bills table
        ws2.append(["Account", "Filename", "KWH", "Total Due", "FiT Amount", "Period From", "Period To"])
        for cell in ws2[ws2.max_row]:
            cell.font = Font(bold=True)
        
        for bill in session_storage['parsed_bills']:
            ws2.append([
                bill.get('account', ''),
                bill.get('file', ''),
                bill.get('kwh', ''),
                bill.get('total_amount', ''),
                bill.get('fit_amount', 0),
                bill.get('from_date', ''),
                bill.get('to_date', ''),
            ])
        
        # Save to output folder or temp
        month = session_storage['billing_month']
        filename = f"meter_log_{month}.xlsx"
        
        # Try to save to METROPOLIS_ROOT/{YYYY-MM}/out/
        if session_storage['folder_mode'] and session_storage['folder_path']:
            out_folder = session_storage['folder_path'].parent / "out"
            out_folder.mkdir(exist_ok=True, parents=True)
            output_path = out_folder / filename
        else:
            # Save to temp
            output_path = Path(tempfile.gettempdir()) / filename
        
        wb.save(output_path)
        logger.info(f"Generated meter log Excel: {output_path}")
        
        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename
        )
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error generating meter log Excel: {e}", exc_info=True)
        raise HTTPException(500, f"Error generating meter log Excel: {str(e)}")


def get_previous_month(billing_month: str) -> str:
    """
    Get previous month from billing month (YYYY-MM).
    
    Args:
        billing_month: Current month in YYYY-MM format
        
    Returns:
        Previous month in YYYY-MM format
    """
    try:
        # Parse YYYY-MM
        current = datetime.strptime(billing_month, '%Y-%m')
        # Subtract one month
        if current.month == 1:
            previous = datetime(current.year - 1, 12, 1)
        else:
            previous = datetime(current.year, current.month - 1, 1)
        return previous.strftime('%Y-%m')
    except:
        return None


@app.post("/api/process-month-end")
async def process_month_end():
    """
    Process Part 2: Copy previous month's workbooks and update with current data.
    
    Requires:
    - Part 1 gate passed (15 bills + meter log)
    - Previous month's Cost Sheet and Cost Allocation workbooks exist
    
    Returns:
        Paths to updated workbooks in current month's out/ directory
    """
    try:
        # Validate gate first
        gate = validate_gate()
        if not gate['passed']:
            raise HTTPException(400, f"Gate must pass before processing month-end. Errors: {gate['errors']}")
        
        billing_month = session_storage['billing_month']
        if not billing_month:
            raise HTTPException(400, "Billing month not set")
        
        parsed_bills = session_storage['parsed_bills']
        if not parsed_bills:
            raise HTTPException(400, "No bills parsed")
        
        meter_log = session_storage['meter_log']
        
        # Get previous month
        previous_month = get_previous_month(billing_month)
        if not previous_month:
            raise HTTPException(400, f"Could not compute previous month from {billing_month}")
        
        # Get metropolis root
        metropolis_root = get_metropolis_root()
        
        # Find the Cost Allocation workbook to use as master (AC DEPT sheet defines allocations)
        # CRITICAL: Never use current month's out/ directory - that's the file we're about to overwrite!
        # Priority:
        # 1. Previous month's main directory (untouched source)
        # 2. Previous month's out/ directory (if that month was processed)
        # 3. Current month's main directory (if not yet moved to out/)
        # 4. masters/ subfolder
        # 5. Test data fallback
        # The chain reads the previous month, never this month's out/.
        # June copies May's outputs; May copies the April pair in the month folder.
        master_candidates = [
            metropolis_root / previous_month / 'out' / 'Cost Allocation - Electricity 2007-2.xls',
            metropolis_root / previous_month / 'out' / 'Cost Allocation - Electricity 2007-2.xlsx',
            metropolis_root / previous_month / 'out' / 'Cost Allocation.xls',
            metropolis_root / previous_month / 'out' / 'Cost Allocation.xlsx',
            metropolis_root / previous_month / 'Cost Allocation - Electricity 2007-2.xls',
            metropolis_root / previous_month / 'Cost Allocation - Electricity 2007-2.xlsx',
            metropolis_root / previous_month / 'Cost Allocation.xls',
            metropolis_root / previous_month / 'Cost Allocation.xlsx',
            metropolis_root / previous_month / 'masters' / 'Cost Allocation - Electricity 2007-2.xlsx',
            metropolis_root / previous_month / 'masters' / 'Cost Allocation.xlsx',
            metropolis_root / previous_month / 'masters' / 'cost_allocation_master.xlsx',
            # A workbook sitting in the month folder (not out/) is only a fallback.
            metropolis_root / billing_month / 'Cost Allocation - Electricity 2007-2.xls',
            metropolis_root / billing_month / 'Cost Allocation - Electricity 2007-2.xlsx',
            metropolis_root / billing_month / 'Cost Allocation.xls',
            metropolis_root / billing_month / 'Cost Allocation.xlsx',
            metropolis_root / billing_month / 'masters' / 'Cost Allocation.xlsx',
            metropolis_root / billing_month / 'masters' / 'cost_allocation_master.xlsx',
            Path('data/masters/cost_allocation_master.xlsx'),
        ]
        
        master_path = None
        for candidate in master_candidates:
            if candidate.exists():
                master_path = candidate
                logger.info(f"Using master workbook: {candidate}")
                break
        
        if not master_path:
            raise HTTPException(400, f"Cost Allocation master not found. Checked previous/current month folders.")
        
        # Parse master workbook to get allocation rules (from AC DEPT sheet)
        from app.master_parser import compute_weights_from_master
        allocation_rules = compute_weights_from_master(str(master_path))
        
        logger.info(f"Loaded allocation rules for {len(allocation_rules)} accounts from master")
        
        if not allocation_rules:
            raise HTTPException(400, "No allocation rules found in master workbook AC DEPT sheet")
        
        # Process month-end: copy and update workbooks
        from app.workbook_updater import process_month_end as process_workbooks
        
        updated_files = process_workbooks(
            metropolis_root,
            billing_month,
            previous_month,
            parsed_bills,
            allocation_rules,
            meter_log,
            master_path,
        )

        from app.dashboard_data import load_month_report
        report = load_month_report(metropolis_root, billing_month)
        
        return {
            "success": True,
            "billing_month": billing_month,
            "previous_month": previous_month,
            "updated_files": {k: str(v) for k, v in updated_files.items()},
            "report": report,
            "message": f"Month-end processed: copied from {previous_month}, updated for {billing_month}"
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error processing month-end: {e}", exc_info=True)
        raise HTTPException(500, f"Error processing month-end: {str(e)}")


@app.get("/api/months")
async def processed_months():
    """Months that have a dashboard report or output workbooks on disk."""
    from app.dashboard_data import list_processed_months
    return {"months": list_processed_months(get_metropolis_root())}


@app.get("/api/dashboard/{month}")
async def dashboard_month(month: str):
    """Consumption and cost allocation saved when that month was processed."""
    from app.dashboard_data import load_month_report, valid_month
    if not valid_month(month):
        raise HTTPException(400, "Month must be YYYY-MM")
    report = load_month_report(get_metropolis_root(), month)
    if report is None:
        raise HTTPException(404, f"No dashboard report for {month}. Process that month first.")
    return report


@app.get("/api/charts")
async def dashboard_charts(
    months: Optional[str] = None,
    account: Optional[str] = None,
    centre: Optional[str] = None,
):
    """
    Chart series from saved bill kWh and allocation lines.

    months is a comma-separated YYYY-MM list. Omit it for every processed month.
    """
    from app.dashboard_data import chart_series, load_all_reports, valid_month
    selected = None
    if months:
        selected = [part.strip() for part in months.split(',') if part.strip()]
        bad = [part for part in selected if not valid_month(part)]
        if bad:
            raise HTTPException(400, f"Month must be YYYY-MM: {', '.join(bad)}")
    return chart_series(
        load_all_reports(get_metropolis_root()),
        months=selected,
        account=account or None,
        centre=centre or None,
    )


@app.get("/api/history/items")
async def history_item_buttons(kind: str = "bill"):
    """Button labels for bills, meters, or cost centres. No kWh or charges."""
    from app.consumption_history import history_items, history_path, load_history
    from app.dashboard_data import load_all_reports
    if kind not in ("bill", "meter", "centre"):
        raise HTTPException(400, "kind must be bill, meter, or centre")
    root = get_metropolis_root()
    return {
        "kind": kind,
        "items": history_items(load_history(root), load_all_reports(root), kind),
        "history": history_path(root).is_file(),
    }


@app.get("/api/history/item")
async def history_item_chart(kind: str, item: str):
    """kWh and cost for one clicked item, read from the history workbook."""
    from app.consumption_history import item_chart_series, load_history
    from app.dashboard_data import load_all_reports
    if kind not in ("bill", "meter", "centre"):
        raise HTTPException(400, "kind must be bill, meter, or centre")
    if not item:
        raise HTTPException(400, "item is required")
    root = get_metropolis_root()
    series = item_chart_series(load_history(root), load_all_reports(root), kind, item)
    if series is None:
        raise HTTPException(404, "No saved history for that item")
    return series


@app.get("/api/history")
async def electricity_history(
    account: Optional[str] = None,
    meter: Optional[str] = None,
):
    """Bill and meter trends from masters/electricity_history.xlsx."""
    from app.consumption_history import history_choices, history_series, load_history
    records = load_history(get_metropolis_root())
    series = history_series(records, account=account or None, meter=meter or None)
    choices = history_choices(records)
    series['accounts'] = choices['accounts']
    series['meter_choices'] = choices['meters']
    return series


@app.get("/api/history/download")
async def download_electricity_history():
    """Download the history workbook. It is not a month's out/ file."""
    from app.consumption_history import HISTORY_NAME, history_path
    path = history_path(get_metropolis_root())
    if not path.is_file():
        raise HTTPException(404, "No electricity history yet. Process a month first.")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=HISTORY_NAME,
    )


@app.get("/api/download/{month}/{kind}")
async def download_output_workbook(month: str, kind: str):
    """
    Download a month-end workbook from METROPOLIS_ROOT/{YYYY-MM}/out/.

    kind is cost-allocation or cost-sheet. The meter-log route is unchanged.
    """
    from app.dashboard_data import find_output_workbook, valid_month
    if not valid_month(month):
        raise HTTPException(400, "Month must be YYYY-MM")
    if kind not in ("cost-allocation", "cost-sheet"):
        raise HTTPException(400, "kind must be cost-allocation or cost-sheet")
    try:
        path = find_output_workbook(get_metropolis_root(), month, kind)
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    if path is None:
        raise HTTPException(404, f"No {kind} workbook for {month}")
    media = "application/vnd.ms-excel"
    if path.suffix.lower() == ".xlsx":
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return FileResponse(path, media_type=media, filename=path.name)


@app.post("/api/reset")
async def reset_session():
    """Reset Part 1 session"""
    session_storage['billing_month'] = None
    session_storage['folder_mode'] = False
    session_storage['folder_path'] = None
    session_storage['parsed_bills'] = []
    session_storage['gate_status'] = None
    session_storage['meter_log'] = {
        'meter_no': '6681757',
        'label': '散熱水泵電',
        'previous': None,
        'present': None,
        'read_date': None,
        'photo_path': None,
    }
    session_storage['session_id'] = None
    session_storage['created_at'] = None
    
    return {"success": True, "message": "Session reset"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
