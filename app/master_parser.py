"""
Master Workbook Parser

Parses the master AC DEPT / Cost Allocation workbook to compute 
monthly cost-centre weights for each CLP account.

Critical: Weights are recomputed EACH MONTH from the live master workbook,
NOT from frozen JSON constants.

Runtime source of truth: AC DEPT sheet column C → weight/Σ(weights) per account
"""

import openpyxl
import pandas as pd
from typing import Dict, List, Optional, Tuple
from pathlib import Path
from decimal import Decimal
import logging

logger = logging.getLogger(__name__)


class MasterWorkbookParser:
    """
    Parses master AC DEPT / Cost Allocation Excel workbook.
    
    Expected sheets:
    - "Allocation": Contains allocation % by account/meter
    - "AC DEPT": Contains AC department mappings
    
    Output: Dict mapping account/meter -> List[{centre, percentage}]
    """
    
    def __init__(self, workbook_path: str):
        self.workbook_path = Path(workbook_path)
        self.workbook = None
        self.allocations = {}
        
    def parse(self) -> Dict[str, List[Dict]]:
        """
        Parse master workbook and return allocation rules.
        
        Returns:
            Dict mapping account/meter ID to list of {centre, percentage} allocations
        """
        if self.workbook_path.suffix in ['.xlsx', '.xlsm']:
            return self._parse_xlsx()
        elif self.workbook_path.suffix == '.xls':
            return self._parse_xls()
        else:
            raise ValueError(f"Unsupported file type: {self.workbook_path.suffix}")
    
    def _parse_xlsx(self) -> Dict[str, List[Dict]]:
        """
        Parse .xlsx or .xlsm using openpyxl.
        
        Priority order:
        1. AC DEPT sheet (runtime source of truth: column C weights)
        2. Allocation sheet (fallback or supplementary)
        """
        wb = openpyxl.load_workbook(self.workbook_path, data_only=True)
        
        allocations = {}
        
        # PRIMARY SOURCE: AC DEPT sheet (column C weights → compute percentages)
        ac_dept_sheet = None
        for name in wb.sheetnames:
            if 'ac dept' in name.lower() or 'acdept' in name.lower():
                ac_dept_sheet = wb[name]
                break
        
        if ac_dept_sheet:
            logger.info("Parsing AC DEPT sheet (runtime source of truth)")
            ac_allocations = self._parse_ac_dept_column_c_weights(ac_dept_sheet)
            allocations.update(ac_allocations)
        
        # FALLBACK: Allocation sheet (for accounts not in AC DEPT)
        allocation_sheet = None
        for name in wb.sheetnames:
            if 'allocation' in name.lower():
                allocation_sheet = wb[name]
                break
        
        if allocation_sheet:
            logger.info("Parsing Allocation sheet (supplementary)")
            alloc_rules = self._parse_allocation_sheet_openpyxl(allocation_sheet)
            # Only add accounts not already in AC DEPT
            for account, rules in alloc_rules.items():
                if account not in allocations:
                    allocations[account] = rules
                    
        wb.close()
        
        logger.info(f"Loaded allocation rules for {len(allocations)} accounts from master")
        return allocations
    
    def _parse_xls(self) -> Dict[str, List[Dict]]:
        """Parse .xls using pandas/xlrd"""
        try:
            # Read all sheets
            all_sheets = pd.read_excel(self.workbook_path, sheet_name=None, engine='xlrd')
        except:
            # Fallback to openpyxl if xlrd fails
            all_sheets = pd.read_excel(self.workbook_path, sheet_name=None, engine='openpyxl')
        
        allocations = {}
        
        # Find Allocation sheet
        for sheet_name, df in all_sheets.items():
            if 'allocation' in sheet_name.lower():
                allocations.update(self._parse_allocation_sheet_pandas(df))
            elif 'ac dept' in sheet_name.lower() or 'acdept' in sheet_name.lower():
                ac_allocations = self._parse_ac_dept_sheet_pandas(df)
                for account, allocs in ac_allocations.items():
                    if account not in allocations:
                        allocations[account] = allocs
        
        return allocations
    
    def _parse_allocation_sheet_openpyxl(self, sheet) -> Dict[str, List[Dict]]:
        """
        Parse Allocation sheet using openpyxl.
        
        Real Citybase structure (verified from April master):
        - Row 3: Headers
        - Column B: "Elect. Meter No.:" (may have merged cells)
        - Column O: "Percentage:" (already computed)
        - Column Q: "Allocated Cost Centre:" (long names like "Commercial - Chiller Plant")
        
        Multiple rows per meter if split across centres (Column B merged down).
        
        Example:
            Row 4: Meter=9046064, Pct=1.0, Centre="Commercial - Chiller Plant"
            Row 36: Meter=9044168, Pct=0.24, Centre="DC - Lifts /Escalators"
            Row 37: Meter=(merged), Pct=0.09, Centre="DC - Lighting"
        """
        allocations = {}
        
        # Convert to list of rows
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            return allocations
        
        # Find header row - look for "Elect. Meter No.:"
        header_row_idx = None
        for idx, row in enumerate(rows[:10]):
            if any('meter' in str(cell).lower() for cell in row if cell):
                header_row_idx = idx
                logger.info(f"Found Allocation sheet header row at index {idx}")
                break
        
        if header_row_idx is None:
            logger.warning("Could not find header row in Allocation sheet")
            return allocations
        
        # Find column positions from headers
        headers = [str(cell).lower() if cell else '' for cell in rows[header_row_idx]]
        
        meter_col = None
        percentage_col = None
        centre_col = None
        
        for idx, h in enumerate(headers):
            if 'meter' in h:
                meter_col = idx
                logger.info(f"Found meter column: {idx}")
            elif 'percentage' in h:
                percentage_col = idx
                logger.info(f"Found percentage column: {idx}")
            elif 'allocated' in h and 'centre' in h:
                centre_col = idx
                logger.info(f"Found centre column: {idx}")
        
        # Default positions if not found in headers
        if meter_col is None:
            meter_col = 1  # Column B = index 1
            logger.warning("Meter column not found in headers, using default: 1 (Column B)")
        
        # For percentage and centre columns, only use defaults if sheet has enough columns
        max_col = len(headers) if headers else 0
        
        if percentage_col is None:
            if max_col >= 15:  # Real Citybase structure
                percentage_col = 14  # Column O = index 14
                logger.warning("Percentage column not found in headers, using default: 14 (Column O)")
            else:
                # Small test workbook - can't parse
                logger.error(f"Percentage column not found and sheet only has {max_col} columns (need 15+)")
                return allocations
        
        if centre_col is None:
            if max_col >= 17:  # Real Citybase structure
                centre_col = 16  # Column Q = index 16
                logger.warning("Centre column not found in headers, using default: 16 (Column Q)")
            else:
                # Small test workbook - can't parse
                logger.error(f"Centre column not found and sheet only has {max_col} columns (need 17+)")
                return allocations
        
        logger.info(f"Allocation sheet columns: meter={meter_col}, percentage={percentage_col}, centre={centre_col}")
        
        # Parse data rows - handle merged cells for meter column
        current_meter = None
        
        for row_idx, row in enumerate(rows[header_row_idx + 1:], start=header_row_idx + 2):
            if not row or all(cell is None for cell in row):
                continue
            
            # Get meter number (may be None if cell is merged)
            meter_val = row[meter_col] if meter_col < len(row) else None
            
            if meter_val and not str(meter_val).startswith('#'):
                # New meter found - update current_meter
                meter = str(meter_val).strip()
                # Remove ** suffix if present (9046787** -> 9046787)
                meter = meter.rstrip('*').strip()
                if meter and meter.lower() not in ['none', 'total']:
                    current_meter = meter
                    logger.debug(f"Row {row_idx}: New meter {current_meter}")
            
            # Skip if no current meter
            if not current_meter:
                continue
            
            # Get percentage from column O
            pct_val = row[percentage_col] if percentage_col < len(row) else None
            if pct_val is None or str(pct_val).startswith('#'):
                continue
            
            try:
                pct = float(pct_val)
                if pct <= 0:
                    continue
            except (ValueError, TypeError):
                continue
            
            # Get centre name from column Q
            centre_val = row[centre_col] if centre_col < len(row) else None
            if centre_val is None or str(centre_val).startswith('#'):
                continue
            
            centre_full = str(centre_val).strip()
            if not centre_full:
                continue
            
            # Extract short code from long name
            # "Commercial - Chiller Plant" -> "C"
            # "DC - Lighting" -> "DC"
            # "Office Accommodation - Chiller Plant" -> "AO"
            centre_code = self._extract_centre_code_from_long_name(centre_full)
            
            # Skip if no valid centre code extracted
            if not centre_code or centre_code.strip() == '':
                logger.debug(f"Row {row_idx}: No valid centre code from '{centre_full}' - SKIPPING")
                continue
            
            # CRITICAL: Validate centre is not numeric
            try:
                float(centre_code)
                logger.warning(f"Row {row_idx}: NUMERIC centre '{centre_code}' from '{centre_full}' - SKIPPING")
                continue
            except (ValueError, TypeError):
                pass  # Good - not numeric
            
            # Add to meter allocations
            if current_meter not in allocations:
                allocations[current_meter] = []
            
            # Handle shared allocations (centre codes with "/")
            if '/' in centre_code:
                centre_codes = [c.strip() for c in centre_code.split('/')]
                split_pct = pct / len(centre_codes)
                
                for code in centre_codes:
                    allocations[current_meter].append({
                        'centre': code,
                        'percentage': split_pct
                    })
                    logger.debug(f"Row {row_idx}: meter {current_meter}, centre {code} ({centre_full} split), pct {split_pct:.4f}")
            else:
                # Single centre
                allocations[current_meter].append({
                    'centre': centre_code,
                    'percentage': pct
                })
                logger.debug(f"Row {row_idx}: meter {current_meter}, centre {centre_code} ({centre_full}), pct {pct:.4f}")
        
        # CRITICAL: Normalize percentages for each meter to sum to 1.0
        # The Allocation sheet has percentages that can sum to >1.0 due to
        # multiple allocation scenarios or merged cell errors
        for meter, entries in allocations.items():
            total_pct = sum(e['percentage'] for e in entries)
            if total_pct == 0:
                logger.warning(f"Meter {meter} has zero total percentage, skipping")
                continue
            
            if abs(total_pct - 1.0) > 0.01:
                logger.warning(f"Meter {meter} percentages sum to {total_pct:.4f}, normalizing to 1.0")
                # Normalize
                for entry in entries:
                    entry['percentage'] /= total_pct
        
        logger.info(f"Parsed {len(allocations)} meters from Allocation sheet")
        return allocations
    
    def _parse_ac_dept_sheet_openpyxl(self, sheet) -> Dict[str, List[Dict]]:
        """Parse AC DEPT sheet - similar structure to Allocation"""
        return self._parse_allocation_sheet_openpyxl(sheet)
    
    def _parse_ac_dept_column_c_weights(self, sheet) -> Dict[str, List[Dict]]:
        """
        Parse AC DEPT sheet - REAL Citybase structure.
        
        Actual structure (verified from real April workbook):
        - Row 3: Headers
        - Column A: "Elect. Meter No.:" - Account/meter (only on first row of group)
        - Columns B-M: Monthly charges (e.g. Column C "May-08 Charge")  
        - Column N: "Percentage:" - formulas pointing to OLD month (Jun-06), NOT current!
        - Column O: "Allocated Cost Centre:" - Centre names (long form)
        
        Example April values:
            Row 73: A="55861-52267-1 (A)", C=287864.81, O="AC - Commercial Air Conditioning"
            Row 74: A="",                  C=82049.19,  O="SW - Sea Water Pump House"
            → AC: 287864.81 / 369914.0 = 77.82%, SW: 22.18%
        
        Key insights:
        - Account in column A only on first row of group; subsequent rows blank
        - Column C has CHARGE amounts to use as weights (compute percentages from them)
        - Column N percentages point to old month - DO NOT USE
        - Centre in column O as "AC - Commercial Air Conditioning" (extract "AC" code)
        
        Strategy:
        1. Find account and centre columns by header
        2. Find a charge column (columns B-M, look for "Charge" in header)
        3. Track current account (persists across rows until new account found)
        4. Extract short code from centre name ("AC" from "AC - Commercial...")
        5. Compute percentages from charge weights (weight / Σ per account group)
        """
        allocations = {}
        
        # Convert to list of rows
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            return allocations
        
        # Find header row - look for row with "Percentage:" or "Elect. Meter No.:"
        header_row_idx = None
        for idx, row in enumerate(rows[:10]):
            if not row:
                continue
            # Check if this row has distinctive header keywords
            row_str_lower = ' '.join(str(cell or '').lower() for cell in row)
            if 'percentage:' in row_str_lower or 'elect. meter no' in row_str_lower:
                header_row_idx = idx
                logger.info(f"Found AC DEPT header row at index {idx}")
                break
        
        if header_row_idx is None:
            logger.warning("Could not find AC DEPT header row - trying allocation sheet parser as fallback")
            return self._parse_allocation_sheet_openpyxl(sheet)
        
        # Find column positions from headers
        headers = [str(cell) if cell else '' for cell in rows[header_row_idx]]
        
        account_col = None
        charge_col = 2  # HARDCODE: Always use Column C (index 2) for charges
        centre_col = None
        
        for idx, h in enumerate(headers):
            h_lower = h.lower()
            if ('meter' in h_lower or 'account' in h_lower or 'acct' in h_lower) and account_col is None:
                account_col = idx
                logger.info(f"Found account column: {idx}")
            elif 'allocated' in h_lower and 'centre' in h_lower and centre_col is None:
                centre_col = idx
                logger.info(f"Found centre column: {idx}")
        
        # Log which charge column we're using
        if charge_col < len(headers):
            logger.info(f"Using charge column: {charge_col} ({headers[charge_col]})")
        else:
            logger.info(f"Using charge column: {charge_col} (Column C)")
        
        if centre_col is None:
            logger.warning("Centre column not found, cannot parse AC DEPT")
            return allocations
        
        if account_col is None:
            account_col = 0
            logger.warning("Account column not found, using column 0")
        
        logger.info(f"AC DEPT columns: account={account_col}, charge={charge_col}, centre={centre_col}")
        
        # Parse data rows - group by account
        account_groups = {}  # account -> [(charge, centre_full), ...]
        current_account = None
        
        for row_idx, row in enumerate(rows[header_row_idx + 1:], start=header_row_idx + 2):
            if not row or all(cell is None for cell in row):
                continue
            
            # Check for new account in column A
            account_val = row[account_col] if account_col < len(row) else None
            if account_val and str(account_val).strip() and not str(account_val).startswith('#'):
                account_str = str(account_val).strip()
                # Remove trailing letters in parentheses like "(A)" or "(B)"
                import re
                account_match = re.match(r'^([\d\-]+)', account_str)
                if account_match:
                    current_account = account_match.group(1)
                else:
                    current_account = account_str
                
                if current_account not in account_groups:
                    account_groups[current_account] = []
                
                logger.debug(f"Row {row_idx}: New account {current_account}")
            
            # Skip if no current account
            if not current_account:
                continue
            
            # Get charge from charge column
            charge_val = row[charge_col] if charge_col < len(row) else None
            if charge_val is None or str(charge_val).startswith('#'):
                continue
            
            try:
                charge = float(charge_val)
                if charge <= 0:
                    continue
            except (ValueError, TypeError):
                continue
            
            # Get centre name
            centre_val = row[centre_col] if centre_col < len(row) else None
            if centre_val is None or str(centre_val).startswith('#'):
                continue
            
            centre_full = str(centre_val).strip()
            if not centre_full:
                continue
            
            # Add to account group
            account_groups[current_account].append((charge, centre_full))
            logger.debug(f"Row {row_idx}: account {current_account}, charge {charge:.2f}, centre {centre_full}")
        
        # Compute percentages for each account group from charges
        for account, entries in account_groups.items():
            total_charge = sum(charge for charge, _ in entries)
            if total_charge == 0:
                logger.warning(f"Account {account} has zero total charge")
                continue
            
            account_allocations = []
            for charge, centre_full in entries:
                pct = charge / total_charge
                
                # Extract short code
                centre_code = self._extract_centre_code_from_long_name(centre_full)
                
                # Skip if no valid centre code extracted
                if not centre_code or centre_code.strip() == '':
                    logger.debug(f"Account {account}: No valid centre code from '{centre_full}' - SKIPPING")
                    continue
                
                # CRITICAL: Validate centre is not numeric
                try:
                    float(centre_code)
                    logger.warning(f"Account {account}: NUMERIC centre '{centre_code}' from '{centre_full}' - SKIPPING")
                    continue
                except (ValueError, TypeError):
                    pass  # Good - not numeric
                
                # Handle shared allocations (centre codes with "/")
                if '/' in centre_code:
                    centre_codes = [c.strip() for c in centre_code.split('/')]
                    split_pct = pct / len(centre_codes)
                    
                    for code in centre_codes:
                        account_allocations.append({
                            'centre': code,
                            'percentage': split_pct
                        })
                        logger.debug(f"Account {account}: centre {code} ({centre_full} split), pct {split_pct:.4f}, charge {charge * split_pct:.2f}")
                else:
                    # Single centre
                    account_allocations.append({
                        'centre': centre_code,
                        'percentage': pct
                    })
                    logger.debug(f"Account {account}: centre {centre_code} ({centre_full}), pct {pct:.4f}, charge {charge:.2f}")
            
            if account_allocations:
                allocations[account] = account_allocations
                logger.info(f"Saved account {account}: {len(account_allocations)} centres, total charge {total_charge:.2f}")
        
        return allocations
    
    def _extract_centre_code_from_long_name(self, centre_full: str) -> str:
        """
        Extract short centre code from long name.
        
        Examples:
            "AC - Commercial Air Conditioning" -> "AC"
            "SW - Sea Water Pump House" -> "SW"
            "C - Commercial" -> "C"
            "OC -Office Common" -> "OC" (handle missing space)
            "Commercial Common" -> "C"
            "Hotel / Commercial / SA" -> "C/SA" (Hotel not a valid code, extract C and SA)
            "SA / Commercial" -> "SA/C"
            "Hotel / Commercial" -> "C"
        
        Strategy:
        1. If format is "CODE - Description" or "CODE -Description", extract CODE
        2. If format has slashes, extract all known codes from segments
        3. If no code prefix, check for known keywords (Commercial -> C, Office -> O, etc.)
        4. Return extracted code(s) joined with "/" or original if no match
        
        Known codes: AC, AO, C, CP, DC, FC, O, OC, SA, SW
        """
        # Known centre codes
        KNOWN_CODES = {'AC', 'AO', 'C', 'CP', 'DC', 'FC', 'O', 'OC', 'SA', 'SW'}
        
        # Try "CODE - Description" or "CODE -Description" format first (handle missing space)
        if ' -' in centre_full:
            code = centre_full.split(' -')[0].strip()
            if code in KNOWN_CODES or len(code) <= 3:
                return code
        
        # If it has slashes, extract all known codes from segments
        if '/' in centre_full:
            segments = [s.strip() for s in centre_full.split('/')]
            matched_codes = []
            
            for seg in segments:
                # Direct match
                if seg in KNOWN_CODES:
                    matched_codes.append(seg)
                # Check if segment contains a known code word
                elif 'Commercial' in seg and 'C' not in matched_codes:
                    matched_codes.append('C')
                elif 'Serviced Apartment' in seg or seg == 'SA':
                    if 'SA' not in matched_codes:
                        matched_codes.append('SA')
                elif 'Office' in seg and 'O' not in matched_codes:
                    # Could be O or OC or AO - default to O
                    matched_codes.append('O')
                # Skip non-code segments like "Hotel", "L8 Premises", "Carpark"
                # These are facility names, not cost centres
            
            if matched_codes:
                return '/'.join(matched_codes)
            else:
                # No recognized codes found in slash-separated name
                # Return empty to skip this entry
                logger.warning(f"Slash-separated centre '{centre_full}' has no recognized codes, skipping")
                return ''
        
        # Check for known keywords without code prefix
        centre_lower = centre_full.lower()
        if 'commercial' in centre_lower and 'air' not in centre_lower:
            return 'C'
        elif 'office' in centre_lower and 'accommodation' not in centre_lower:
            return 'OC'  # "Office Common" -> OC
        elif 'carpark' in centre_lower or 'car park' in centre_lower:
            return 'CP'
        
        # Return as-is if no extraction worked
        return centre_full
    
    def _parse_allocation_sheet_pandas(self, df: pd.DataFrame) -> Dict[str, List[Dict]]:
        """Parse Allocation sheet using pandas"""
        allocations = {}
        
        # Find account column
        account_col = None
        for col in df.columns:
            if any(keyword in str(col).lower() for keyword in ['account', 'meter', 'acct', 'a/c']):
                account_col = col
                break
        
        if account_col is None:
            account_col = df.columns[0]
        
        # Find cost centre columns
        centre_cols = []
        for col in df.columns:
            if col == account_col:
                continue
            # Include ALL non-account columns as potential cost centres
            col_str = str(col).strip()
            if col_str and col_str.lower() not in ['account', 'meter', 'acct', 'a/c', 'account/meter']:
                centre_cols.append(col)
        
        # Parse rows
        for idx, row in df.iterrows():
            account = str(row[account_col]) if pd.notna(row[account_col]) else None
            if not account or account.lower() in ['none', 'total', 'nan', '']:
                continue
            
            account = account.strip().replace(' ', '')
            
            account_allocations = []
            total_pct = 0.0
            
            for centre_col in centre_cols:
                value = row[centre_col]
                if pd.isna(value):
                    continue
                
                try:
                    if isinstance(value, str):
                        value = value.replace('%', '').strip()
                        pct = float(value)
                        if pct > 1:
                            pct = pct / 100.0
                    else:
                        pct = float(value)
                        if pct > 1:
                            pct = pct / 100.0
                    
                    if pct > 0:
                        account_allocations.append({
                            'centre': str(centre_col).strip(),
                            'percentage': pct
                        })
                        total_pct += pct
                except (ValueError, TypeError):
                    continue
            
            if account_allocations:
                if 0.99 <= total_pct <= 1.01:
                    for alloc in account_allocations:
                        alloc['percentage'] /= total_pct
                
                allocations[account] = account_allocations
        
        return allocations
    
    def _parse_ac_dept_sheet_pandas(self, df: pd.DataFrame) -> Dict[str, List[Dict]]:
        """Parse AC DEPT sheet"""
        return self._parse_allocation_sheet_pandas(df)


def compute_weights_from_master(workbook_path: str) -> Dict[str, List[Dict]]:
    """
    Compute allocation weights from master workbook.
    
    Args:
        workbook_path: Path to master AC DEPT / Cost Allocation workbook
        
    Returns:
        Dict mapping account/meter -> List[{centre, percentage}]
    """
    parser = MasterWorkbookParser(workbook_path)
    return parser.parse()
