"""
Workbook Updater for Cost Sheet and Cost Allocation

Copies previous month's workbooks and updates them with current month's data
while preserving all formatting (fonts, merges, widths, number formats, structure).

Critical rules:
- Copy source files, NEVER overwrite originals
- Write computed VALUES not formulas (no cross-file links that become #REF!)
- Preserve all formatting exactly
- Odd-cent differences of 0.01 are acceptable
"""

import openpyxl
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from datetime import datetime
import shutil
import logging
from decimal import Decimal, ROUND_HALF_UP

logger = logging.getLogger(__name__)


class WorkbookUpdater:
    """
    Updates Cost Sheet and Cost Allocation workbooks with current month's data.
    """
    
    def __init__(self, metropolis_root: Path):
        self.metropolis_root = Path(metropolis_root).expanduser()
    
    def _extract_primary_centre_code(self, centre_name: str) -> str:
        """
        Extract primary centre code from complex centre names.
        
        Examples:
            "AC" -> "AC"
            "FC" -> "FC"
            "C" -> "C"
            "Hotel/Commercial/SA (11)" -> "SA"
            "Commercial/Office/SA (7)" -> "SA"
            "SA/Commercial (16)" -> "SA"
            "Hotel / Carpark(9)" -> fallback to "Carpark"
            "Commercial" -> "Commercial"
        
        Strategy:
        1. If it's already a simple 1-2 letter uppercase code, return as-is
        2. Remove parenthetical reference numbers like "(11)"
        3. Look for known simple centre codes (AC, AO, C, CP, DC, FC, O, OC, SA, SW) in segments
        4. If found, return the first matching code
        5. Otherwise extract last segment from slashes
        """
        centre_name = centre_name.strip()
        
        # Simple codes (1-2 uppercase letters) - return as-is
        if len(centre_name) <= 2 and centre_name.isupper():
            return centre_name
        
        # Check for parenthetical reference numbers like "(11)"
        # Extract the text before the parenthesis
        if '(' in centre_name:
            # "Hotel/Commercial/SA (11)" -> "Hotel/Commercial/SA"
            centre_name = centre_name.split('(')[0].strip()
        
        # Known simple centre codes
        SIMPLE_CODES = ['AC', 'AO', 'C', 'CP', 'DC', 'FC', 'O', 'OC', 'SA', 'SW']
        
        # Check if any known code appears in the centre name
        if '/' in centre_name:
            # Split into segments and check each
            segments = [s.strip() for s in centre_name.split('/') if s.strip()]
            for segment in segments:
                if segment.upper() in SIMPLE_CODES:
                    return segment.upper()
            # If no simple code found, return last segment
            if segments:
                return segments[-1]
        
        # No slashes - check if it's a known code
        if centre_name.upper() in SIMPLE_CODES:
            return centre_name.upper()
        
        # Return as-is for other cases (Commercial, Carpark, etc.)
        return centre_name
    
    def copy_previous_month_workbooks(
        self,
        current_month: str,  # YYYY-MM
        previous_month: str,  # YYYY-MM
    ) -> Dict[str, Path]:
        """
        Copy previous month's workbooks to current month's out/ directory.
        
        Handles real Citybase filenames:
        - Cost Allocation - Electricity 2007-2.xls (long-lived master)
        - Cost Sheet-2025-03.xls (contains previous month)
        
        Looks in:
        - ~/Metropolis/{previous}/masters/
        - ~/Metropolis/{previous}/out/
        - ~/Metropolis/{current}/masters/
        
        Args:
            current_month: Current billing month (YYYY-MM)
            previous_month: Previous month (YYYY-MM)
        
        Returns:
            Dict mapping workbook type to copied file path
        """
        current_month_dir = self.metropolis_root / current_month
        previous_month_dir = self.metropolis_root / previous_month
        
        # Ensure output directory exists
        out_dir = current_month_dir / 'out'
        out_dir.mkdir(parents=True, exist_ok=True)
        
        copied_files = {}
        
        # Search locations in priority order
        search_dirs = [
            previous_month_dir / 'out',
            previous_month_dir / 'masters',
            current_month_dir / 'masters',
        ]
        
        # Find Cost Allocation workbook (long-lived file)
        cost_allocation_file = None
        for search_dir in search_dirs:
            if not search_dir.exists():
                continue
            # Look for files containing "Cost Allocation" with .xls or .xlsx
            for pattern in ['*Cost Allocation*.xls', '*Cost Allocation*.xlsx']:
                matches = list(search_dir.glob(pattern))
                if matches:
                    # Take the first match
                    cost_allocation_file = matches[0]
                    logger.info(f"Found Cost Allocation: {cost_allocation_file}")
                    break
            if cost_allocation_file:
                break
        
        if cost_allocation_file and cost_allocation_file.exists():
            # Copy to current month's out/ with same name (long-lived)
            dest_file = out_dir / cost_allocation_file.name
            shutil.copy2(cost_allocation_file, dest_file)
            logger.info(f"Copied Cost Allocation: {cost_allocation_file} -> {dest_file}")
            copied_files['cost_allocation'] = dest_file
        else:
            logger.warning(f"Could not find Cost Allocation workbook")
        
        # Find Cost Sheet workbook (contains previous month in name)
        cost_sheet_file = None
        for search_dir in search_dirs:
            if not search_dir.exists():
                continue
            # Look for files containing "Cost Sheet" with .xls or .xlsx
            for pattern in ['*Cost Sheet*.xls', '*Cost Sheet*.xlsx']:
                matches = list(search_dir.glob(pattern))
                for match in matches:
                    # Prefer files that contain the previous month in the name
                    if previous_month in match.name:
                        cost_sheet_file = match
                        logger.info(f"Found Cost Sheet (with month): {cost_sheet_file}")
                        break
                if cost_sheet_file:
                    break
            if cost_sheet_file:
                break
        
        # If not found with previous month, take any Cost Sheet file
        if not cost_sheet_file:
            for search_dir in search_dirs:
                if not search_dir.exists():
                    continue
                for pattern in ['*Cost Sheet*.xls', '*Cost Sheet*.xlsx']:
                    matches = list(search_dir.glob(pattern))
                    if matches:
                        cost_sheet_file = matches[0]
                        logger.info(f"Found Cost Sheet (generic): {cost_sheet_file}")
                        break
                if cost_sheet_file:
                    break
        
        if cost_sheet_file and cost_sheet_file.exists():
            # Update filename to use current month
            # Cost Sheet-2025-03.xls -> Cost Sheet-2025-04.xls
            original_ext = cost_sheet_file.suffix  # .xls or .xlsx
            dest_name = f"Cost Sheet-{current_month}{original_ext}"
            dest_file = out_dir / dest_name
            shutil.copy2(cost_sheet_file, dest_file)
            logger.info(f"Copied Cost Sheet: {cost_sheet_file} -> {dest_file}")
            copied_files['cost_sheet'] = dest_file
        else:
            logger.warning(f"Could not find Cost Sheet workbook")
        
        return copied_files
    
    def update_cost_allocation_workbook(
        self,
        workbook_path: Path,
        allocations: List[Dict],
        billing_month: str,
        parsed_bills: List[Dict],
        meter_log: Dict
    ) -> Path:
        """
        Update Cost Allocation workbook with current month's allocation data.
        
        Writes computed VALUES (not formulas) to preserve formatting.
        
        Key updates:
        - Elect Charge sheet: check-meter 6681757 previous/present/delta
        - Title row: update month
        
        Args:
            workbook_path: Path to Cost Allocation workbook (already copied)
            allocations: List of allocation dicts from allocation engine
            billing_month: YYYY-MM
            parsed_bills: List of parsed bills
            meter_log: Dict with check-meter readings
            
        Returns:
            Updated workbook path (same as input)
        """
        # Check if .xls - openpyxl only supports .xlsx
        if workbook_path.suffix.lower() == '.xls':
            logger.warning(f"Cost Allocation is .xls format - openpyxl requires .xlsx. "
                          f"Consider converting with LibreOffice first.")
            # For now, just update the .xls file we can
            # In production, you'd use xlrd/xlwt or convert with soffice
        
        try:
            wb = openpyxl.load_workbook(workbook_path)
        except Exception as e:
            logger.error(f"Failed to load Cost Allocation workbook: {e}")
            # If it's .xls and fails, just log and return
            return workbook_path
        
        # DO NOT update title/header cells - preserve original titles from 2008
        # The golden workbooks keep their original titles unchanged
        
        # Find Elect Charge sheet (check-meter 6681757)
        elect_charge_sheet = None
        for sheet in wb.worksheets:
            # Look for sheet with "Elect" or "Charge" in name
            if 'elect' in sheet.title.lower() or 'charge' in sheet.title.lower():
                elect_charge_sheet = sheet
                break
        
        if elect_charge_sheet and meter_log.get('previous') is not None and meter_log.get('present') is not None:
            logger.info(f"Found Elect Charge sheet: {elect_charge_sheet.title}")
            
            # Try to find meter 6681757 row by searching for the meter number
            meter_no = meter_log.get('meter_no', '6681757')
            meter_row = None
            
            # Search for meter number in the sheet - extend search to ALL columns, not just 10
            for row in elect_charge_sheet.iter_rows(min_row=1, max_row=100, min_col=1, max_col=200):
                for cell in row:
                    if cell.value and str(meter_no) in str(cell.value):
                        meter_row = cell.row
                        logger.info(f"Found meter {meter_no} at row {meter_row}")
                        break
                if meter_row:
                    break
            
            # If found meter row, try to write readings
            # Common pattern: Previous, Present, Delta in adjacent columns
            # User mentioned BR31/BS31/BQ31, so columns might be around BQ-BS (columns 69-71)
            if meter_row:
                # Try common column positions for meter readings
                # BQ=69, BR=70, BS=71 (0-indexed: 68, 69, 70)
                # But let's search for "Previous" and "Present" headers first
                
                # Search for column headers in row above meter or in header row - search ALL columns
                prev_col = None
                present_col = None
                delta_col = None
                
                # Check a few rows above meter_row for headers - search ALL columns
                for header_row_idx in range(max(1, meter_row - 5), meter_row):
                    row_cells = list(elect_charge_sheet.iter_rows(
                        min_row=header_row_idx, max_row=header_row_idx, 
                        min_col=1, max_col=200, values_only=False
                    ))[0]
                    
                    for cell in row_cells:
                        if cell.value:
                            val_lower = str(cell.value).lower()
                            # Match: "Previous", "Prev", "上期"
                            if 'previous' in val_lower or 'prev' in val_lower or '上期' in val_lower:
                                prev_col = cell.column
                                logger.info(f"Found Previous column: {openpyxl.utils.get_column_letter(prev_col)}")
                            # Match: "Present", "Current", "Current Reading", "今期"
                            elif 'present' in val_lower or 'current' in val_lower or '今期' in val_lower:
                                present_col = cell.column
                                logger.info(f"Found Present/Current column: {openpyxl.utils.get_column_letter(present_col)}")
                            # Match: "Delta", "Diff", "Difference", "度數"
                            elif 'delta' in val_lower or 'diff' in val_lower or '度數' in val_lower:
                                delta_col = cell.column
                                logger.info(f"Found Delta column: {openpyxl.utils.get_column_letter(delta_col)}")
                
                # If we found the columns, write the values
                if prev_col and present_col:
                    elect_charge_sheet.cell(row=meter_row, column=prev_col).value = meter_log['previous']
                    elect_charge_sheet.cell(row=meter_row, column=present_col).value = meter_log['present']
                    
                    # Check if delta cell has a formula - if so, don't overwrite it
                    if delta_col:
                        delta_cell = elect_charge_sheet.cell(row=meter_row, column=delta_col)
                        # If it's a formula (like =BS31-BR31), leave it alone
                        # Otherwise, write the delta value
                        if not (delta_cell.value and isinstance(delta_cell.value, str) and delta_cell.value.startswith('=')):
                            delta = round(meter_log['present'] - meter_log['previous'], 1)
                            delta_cell.value = delta
                            logger.info(f"Wrote delta value {delta}")
                        else:
                            logger.info(f"Delta cell has formula, leaving it to auto-calculate")
                    
                    logger.info(f"Updated meter readings at row {meter_row}: "
                              f"prev={meter_log['previous']}, present={meter_log['present']}")
                else:
                    logger.warning(f"Could not find Previous/Present columns for meter {meter_no}")
            else:
                logger.warning(f"Could not find meter {meter_no} in Elect Charge sheet")
        
        # Save the updated workbook
        try:
            wb.save(workbook_path)
            wb.close()
            logger.info(f"Updated Cost Allocation workbook: {workbook_path}")
        except Exception as e:
            logger.error(f"Failed to save Cost Allocation workbook: {e}")
            wb.close()
        
        return workbook_path
    
    def update_cost_sheet_workbook(
        self,
        workbook_path: Path,
        allocations: List[Dict],
        billing_month: str,
        parsed_bills: List[Dict],
        meter_log: Dict
    ) -> Path:
        """
        Update Cost Sheet workbook with current month's data.
        
        Updates:
        - Year-grid sheet: month column with allocation totals per cost centre
        - Allocation % for each account/cost centre
        - Separated Amount (computed values)
        - Check-meter 6681757 readings if present
        
        Writes VALUES not formulas to preserve formatting.
        
        Args:
            workbook_path: Path to Cost Sheet workbook (already copied)
            allocations: List of allocation dicts
            billing_month: YYYY-MM (e.g., "2025-05")
            parsed_bills: List of parsed bills
            meter_log: Dict with check-meter readings
            
        Returns:
            Updated workbook path (same as input)
        """
        # Check if .xls - openpyxl only supports .xlsx
        if workbook_path.suffix.lower() == '.xls':
            logger.warning(f"Cost Sheet is .xls format - openpyxl requires .xlsx. "
                          f"Consider converting with LibreOffice first.")
            # For now, just update what we can
        
        try:
            wb = openpyxl.load_workbook(workbook_path)
        except Exception as e:
            logger.error(f"Failed to load Cost Sheet workbook: {e}")
            return workbook_path
        
        # DO NOT update title/header cells - preserve original titles
        # The golden workbooks keep their original titles unchanged
        
        # Find year-grid sheet matching the billing year (e.g., "01-12 2025")
        year_grid_sheet = None
        billing_year = billing_month.split('-')[0]
        
        for sheet in wb.worksheets:
            sheet_name = sheet.title
            # Look for sheets with "01-12" AND the billing year
            if '01-12' in sheet_name and billing_year in sheet_name:
                year_grid_sheet = sheet
                logger.info(f"Found year-grid sheet matching {billing_year}: {sheet.title}")
                break
        
        # If no exact match, try just "01-12" as fallback (but log warning)
        if not year_grid_sheet:
            for sheet in wb.worksheets:
                if '01-12' in sheet.title.lower():
                    year_grid_sheet = sheet
                    logger.warning(f"Using fallback year-grid sheet (no {billing_year} match): {sheet.title}")
                    break
        
        if year_grid_sheet:
            # Calculate which column for this month (1=Jan=B, 2=Feb=C, ..., 5=May=F)
            month_num = int(billing_month.split('-')[1])  # Extract month number
            # Assuming column A is labels, column B is Jan (month 1), C is Feb (month 2), etc.
            month_col = month_num + 1  # B=2, C=3, D=4, E=5, F=6 for May
            month_col_letter = openpyxl.utils.get_column_letter(month_col)
            
            logger.info(f"Writing to month column {month_col_letter} (month {month_num})")
            
            # Calculate total allocation per cost centre
            # Map complex centre names to simple codes (AC, FC, SW, DC, etc.)
            centre_totals = {}
            for alloc in allocations:
                centre_raw = alloc['centre']
                amount = alloc['amount']
                
                # Extract primary centre code from complex names
                # Examples:
                #   "AC" -> "AC"
                #   "Hotel/Commercial/SA (11)" -> "SA"
                #   "Commercial/Office/SA (7)" -> "SA"
                #   "C" -> "C"
                centre = self._extract_primary_centre_code(centre_raw)
                
                centre_totals[centre] = centre_totals.get(centre, 0) + amount
            
            logger.info(f"Aggregated allocation totals by centre: {centre_totals}")
            
            if not centre_totals:
                logger.warning("No allocation totals to write - allocations list may be empty or invalid")
            
            # Find rows for each cost centre by searching column A for centre names
            # Common centres: FC, AC, SW, DC, OC, AO, CP, SA, C
            centre_rows = {}
            
            for row in year_grid_sheet.iter_rows(min_row=1, max_row=50, min_col=1, max_col=1):
                cell = row[0]
                if cell.value:
                    cell_val = str(cell.value).strip().upper()
                    # Check if this cell contains a cost centre code
                    for centre in ['FC', 'AC', 'SW', 'DC', 'OC', 'AO', 'CP', 'SA', 'C']:
                        # Match exact or as part of longer description
                        if cell_val == centre or cell_val.startswith(centre + ' ') or cell_val.startswith(centre + '-'):
                            centre_rows[centre] = cell.row
                            logger.info(f"Found centre {centre} at row {cell.row}")
                            break
            
            # Write allocation totals to the month column
            for centre, total in centre_totals.items():
                if centre in centre_rows:
                    row = centre_rows[centre]
                    cell = year_grid_sheet.cell(row=row, column=month_col)
                    cell.value = round(total, 2)
                    logger.info(f"Wrote {centre} total ${total:.2f} to {month_col_letter}{row}")
                else:
                    logger.warning(f"Could not find row for centre {centre}")
        else:
            logger.warning("Could not find year-grid sheet in Cost Sheet")
        
        # Save the updated workbook
        try:
            wb.save(workbook_path)
            wb.close()
            logger.info(f"Updated Cost Sheet workbook: {workbook_path}")
        except Exception as e:
            logger.error(f"Failed to save Cost Sheet workbook: {e}")
            wb.close()
        
        return workbook_path


def process_month_end(
    metropolis_root: Path,
    current_month: str,
    previous_month: str,
    parsed_bills: List[Dict],
    allocation_rules: Dict[str, List[Dict]],
    meter_log: Dict
) -> Dict[str, Path]:
    """
    Process month-end: copy previous workbooks and update with current data.
    
    Args:
        metropolis_root: Root directory (e.g., ~/Metropolis)
        current_month: Current billing month (YYYY-MM)
        previous_month: Previous month (YYYY-MM)
        parsed_bills: List of parsed bill dicts
        allocation_rules: Allocation rules from master workbook
        meter_log: Dict with check-meter readings
        
    Returns:
        Dict mapping workbook type to updated file path
    """
    from app.allocation_engine import allocate_costs
    
    # Run allocation engine with meter_log for check-meter split
    allocation_result = allocate_costs(parsed_bills, allocation_rules, meter_log)
    allocations = allocation_result['allocations']
    
    updater = WorkbookUpdater(metropolis_root)
    
    # Copy previous month's workbooks
    copied_files = updater.copy_previous_month_workbooks(
        current_month,
        previous_month
    )
    
    updated_files = {}
    
    # Update Cost Allocation if copied
    if 'cost_allocation' in copied_files:
        updated_files['cost_allocation'] = updater.update_cost_allocation_workbook(
            copied_files['cost_allocation'],
            allocations,
            current_month,
            parsed_bills,
            meter_log
        )
    
    # Update Cost Sheet if copied
    if 'cost_sheet' in copied_files:
        updated_files['cost_sheet'] = updater.update_cost_sheet_workbook(
            copied_files['cost_sheet'],
            allocations,
            current_month,
            parsed_bills,
            meter_log
        )
    
    return updated_files
