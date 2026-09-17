"""
Metropolis Electricity Cost Allocation App - Main API Server

FastAPI application for:
1. Uploading CLP electricity bill PDFs
2. Uploading/using master Cost Allocation workbook
3. Computing allocation weights from master (LIVE, not frozen JSON)
4. Allocating costs to cost centres
5. Exporting to Citybase Excel format
"""

from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from typing import List, Optional
import tempfile
import os
import shutil
from pathlib import Path
import logging
from datetime import datetime

from app.clp_parser import parse_clp_bill
from app.master_parser import compute_weights_from_master
from app.allocation_engine import allocate_costs
from app.excel_exporter import export_to_citybase_excel

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Metropolis Electricity Cost Allocation",
    description="Upload CLP bills + master workbook → allocate costs → download Citybase Excel",
    version="1.0.0"
)

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Temp storage for uploaded files (in-memory for demo; use DB for production)
temp_storage = {
    'master_path': 'data/masters/cost_allocation_master.xlsx',  # Default master
    'parsed_bills': [],
    'allocation_results': None,
    'current_month': None
}


@app.get("/", response_class=HTMLResponse)
async def root():
    """Serve the web UI"""
    return FileResponse("static/index.html")


@app.get("/api/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "ok", "version": "1.0.0"}


@app.get("/api/status")
async def get_status():
    """Get current processing status"""
    return {
        "master_loaded": os.path.exists(temp_storage['master_path']),
        "master_path": temp_storage['master_path'],
        "bills_parsed": len(temp_storage['parsed_bills']),
        "allocation_complete": temp_storage['allocation_results'] is not None,
        "current_month": temp_storage['current_month']
    }


@app.post("/api/upload-master")
async def upload_master(file: UploadFile = File(...)):
    """
    Upload master Cost Allocation workbook.
    
    This workbook contains AC DEPT and Allocation sheets with current
    cost-centre weights. Weights are recomputed from this master each month.
    """
    try:
        # Validate file type
        if not file.filename.endswith(('.xlsx', '.xls', '.xlsm')):
            raise HTTPException(400, "Master workbook must be .xlsx, .xls, or .xlsm")
        
        # Save to temp location
        temp_path = f"data/masters/uploaded_master{Path(file.filename).suffix}"
        with open(temp_path, 'wb') as f:
            shutil.copyfileobj(file.file, f)
        
        # Try to parse it to validate
        try:
            weights = compute_weights_from_master(temp_path)
            logger.info(f"Master parsed: {len(weights)} accounts with allocation rules")
        except Exception as e:
            os.remove(temp_path)
            raise HTTPException(400, f"Failed to parse master workbook: {str(e)}")
        
        # Update temp storage
        temp_storage['master_path'] = temp_path
        
        return {
            "success": True,
            "message": f"Master workbook uploaded: {len(weights)} accounts",
            "accounts": list(weights.keys())[:10],  # First 10 accounts
            "total_accounts": len(weights)
        }
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error uploading master: {e}", exc_info=True)
        raise HTTPException(500, f"Error uploading master: {str(e)}")


@app.post("/api/upload-bills")
async def upload_bills(
    files: List[UploadFile] = File(...),
    month: str = Form(...)
):
    """
    Upload CLP electricity bill PDFs for parsing.
    
    Args:
        files: List of PDF files (CLP bills)
        month: Month label (e.g., "August 2026")
    """
    try:
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
        temp_storage['parsed_bills'] = parsed_bills
        temp_storage['current_month'] = month
        temp_storage['allocation_results'] = None  # Reset allocation
        
        # Calculate totals
        total_kwh = sum(b.get('kwh', 0) for b in parsed_bills)
        total_amount = sum(b.get('total_amount', 0) for b in parsed_bills if b.get('total_amount'))
        
        return {
            "success": True,
            "bills_parsed": len(parsed_bills),
            "total_kwh": total_kwh,
            "total_amount": round(total_amount, 2),
            "errors": errors,
            "accounts": [b.get('account') for b in parsed_bills]
        }
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error uploading bills: {e}", exc_info=True)
        raise HTTPException(500, f"Error uploading bills: {str(e)}")


@app.post("/api/allocate")
async def perform_allocation():
    """
    Allocate electricity costs to cost centres.
    
    Uses:
    - Parsed bills from /api/upload-bills
    - Current master workbook weights (computed live, not frozen JSON)
    
    Returns allocation results with validation.
    """
    try:
        # Check prerequisites
        if not temp_storage['parsed_bills']:
            raise HTTPException(400, "No bills parsed. Upload bills first.")
        
        if not os.path.exists(temp_storage['master_path']):
            raise HTTPException(400, "No master workbook available")
        
        # Compute weights from master workbook (LIVE computation)
        logger.info(f"Computing allocation weights from master: {temp_storage['master_path']}")
        allocation_rules = compute_weights_from_master(temp_storage['master_path'])
        logger.info(f"Computed weights for {len(allocation_rules)} accounts")
        
        # Perform allocation
        results = allocate_costs(temp_storage['parsed_bills'], allocation_rules)
        
        # Store results
        temp_storage['allocation_results'] = results
        
        return {
            "success": True,
            "allocations": results['allocations'],
            "summary": results['summary'],
            "validation": results['validation'],
            "note": "Weights computed from live master workbook, not frozen JSON"
        }
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error performing allocation: {e}", exc_info=True)
        raise HTTPException(500, f"Error performing allocation: {str(e)}")


@app.get("/api/download-excel")
async def download_excel():
    """
    Download Citybase-format Excel workbook with allocation results.
    
    Format:
    - Year rollup sheet (27 cost centres, FC → SA order)
    - Per-account ELECTRICITY COST ALLOCATION forms
    """
    try:
        # Check if allocation is done
        if not temp_storage['allocation_results']:
            raise HTTPException(400, "No allocation results. Run allocation first.")
        
        results = temp_storage['allocation_results']
        month = temp_storage['current_month'] or datetime.now().strftime("%B %Y")
        
        # Generate Excel file
        output_path = f"data/Metropolis_Cost_Sheet_{month.replace(' ', '_')}.xlsx"
        export_to_citybase_excel(
            results['allocations'],
            results['summary'],
            month,
            output_path
        )
        
        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(output_path)
        )
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error generating Excel: {e}", exc_info=True)
        raise HTTPException(500, f"Error generating Excel: {str(e)}")


@app.post("/api/reset")
async def reset_session():
    """Reset current session"""
    temp_storage['parsed_bills'] = []
    temp_storage['allocation_results'] = None
    temp_storage['current_month'] = None
    # Keep master_path (don't reset to default)
    
    return {"success": True, "message": "Session reset"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
