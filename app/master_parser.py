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
        
        Expected structure (flexible - will auto-detect):
        - Account/Meter column
        - Cost centre columns with percentages
        - May have header rows
        """
        allocations = {}
        
        # Convert to list of rows
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            return allocations
        
        # Try to find header row
        header_row_idx = 0
        for idx, row in enumerate(rows[:10]):  # Check first 10 rows
            if any(str(cell).lower() in ['account', 'meter', 'acct'] for cell in row if cell):
                header_row_idx = idx
                break
        
        headers = [str(cell).strip() if cell else '' for cell in rows[header_row_idx]]
        
        # Find account/meter column
        account_col_idx = None
        for idx, h in enumerate(headers):
            if any(keyword in h.lower() for keyword in ['account', 'meter', 'acct', 'a/c']):
                account_col_idx = idx
                break
        
        if account_col_idx is None:
            # Default to first column
            account_col_idx = 0
        
        # Find cost centre columns (columns with FC, SW, AC, C, DC, etc.)
        centre_cols = []
        for idx, h in enumerate(headers):
            if idx == account_col_idx:
                continue
            # Include ALL non-account columns as potential cost centres
            # This handles both simple codes (AC, FC, SW, C, DC, OC, AO, CP, SA)
            # and complex descriptions (Hotel/Commercial/SA (11), etc.)
            if h and h.strip():  # Any non-empty header
                centre_cols.append((idx, h))
        
        # Parse data rows
        for row in rows[header_row_idx + 1:]:
            if not row or all(cell is None for cell in row):
                continue
            
            account = str(row[account_col_idx]) if row[account_col_idx] else None
            if not account or account.lower() in ['none', 'total', '']:
                continue
            
            # Clean account number
            account = account.strip().replace(' ', '')
            
            # Extract allocations for this account
            account_allocations = []
            total_pct = 0.0
            
            for col_idx, centre_name in centre_cols:
                value = row[col_idx] if col_idx < len(row) else None
                if value is None:
                    continue
                
                # Try to parse as percentage
                try:
                    if isinstance(value, str):
                        # Remove % sign and convert
                        value = value.replace('%', '').strip()
                        pct = float(value)
                        # If value is > 1, assume it's a percentage (e.g. 50 = 50%)
                        if pct > 1:
                            pct = pct / 100.0
                    else:
                        pct = float(value)
                        if pct > 1:
                            pct = pct / 100.0
                    
                    if pct > 0:
                        account_allocations.append({
                            'centre': centre_name.strip(),
                            'percentage': pct
                        })
                        total_pct += pct
                except (ValueError, TypeError):
                    continue
            
            if account_allocations:
                # Normalize if total is close to but not exactly 1.0
                if 0.99 <= total_pct <= 1.01:
                    # Normalize to exactly 1.0
                    for alloc in account_allocations:
                        alloc['percentage'] /= total_pct
                
                allocations[account] = account_allocations
        
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
        - Columns B-M: Monthly charges (formulas/numbers - NOT weights!)
        - Column N: "Percentage:" - The actual percentages (already computed!)
        - Column O: "Allocated Cost Centre:" - Centre names (long form)
        
        Example:
            Row 73: A="55861-52267-1 (A)", N=0.848169931, O="AC - Commercial Air Conditioning"
            Row 74: A="",                  N=0.151830069, O="SW - Sea Water Pump House"
        
        Key insights:
        - Account in column A only on first row of group; subsequent rows blank
        - Percentages ALREADY COMPUTED in column N (not weights to compute from!)
        - Centre in column O as "AC - Commercial Air Conditioning" (extract "AC" code)
        - Column C is charges (numeric) - must NOT be used as centre!
        
        Strategy:
        1. Find "Percentage:" and "Allocated Cost Centre:" columns by header
        2. Track current account (persists across rows until new account found)
        3. Extract short code from centre name ("AC" from "AC - Commercial...")
        4. Use percentage from column N directly
        """
        allocations = {}
        
        # Convert to list of rows
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            return allocations
        
        # Find header row - look for "Percentage:" and "Allocated Cost Centre:"
        header_row_idx = None
        for idx, row in enumerate(rows[:10]):
            if any('percentage' in str(cell).lower() for cell in row if cell):
                header_row_idx = idx
                logger.info(f"Found header row at index {idx}")
                break
        
        if header_row_idx is None:
            logger.warning("Could not find header row with 'Percentage:' - trying allocation sheet parser as fallback")
            # Fall back to the allocation sheet parser for test fixtures
            return self._parse_allocation_sheet_openpyxl(sheet)
        
        # Find column positions from headers
        headers = [str(cell).lower() if cell else '' for cell in rows[header_row_idx]]
        
        account_col = None
        percentage_col = None
        centre_col = None
        
        for idx, h in enumerate(headers):
            if 'meter' in h or 'account' in h or 'acct' in h:
                account_col = idx
                logger.info(f"Found account column: {idx}")
            elif 'percentage' in h:
                percentage_col = idx
                logger.info(f"Found percentage column: {idx}")
            elif 'allocated' in h and 'centre' in h:
                centre_col = idx
                logger.info(f"Found centre column: {idx}")
        
        if percentage_col is None or centre_col is None:
            logger.warning(f"Missing columns: percentage_col={percentage_col}, centre_col={centre_col}")
            return allocations
        
        if account_col is None:
            account_col = 0
            logger.warning("Account column not found in headers, using column 0")
        
        logger.info(f"AC DEPT columns: account={account_col}, percentage={percentage_col}, centre={centre_col}")
        
        # Parse data rows - track current account across rows
        current_account = None
        account_allocations = []
        
        for row_idx, row in enumerate(rows[header_row_idx + 1:], start=header_row_idx + 2):
            if not row or all(cell is None for cell in row):
                continue
            
            # Check for new account in column A
            account_val = row[account_col] if account_col < len(row) else None
            if account_val and str(account_val).strip() and not str(account_val).startswith('#'):
                # New account found - save previous account if exists
                if current_account and account_allocations:
                    allocations[current_account] = account_allocations
                    logger.info(f"Saved account {current_account}: {len(account_allocations)} centres")
                
                # Start new account
                account_str = str(account_val).strip()
                # Remove trailing letters in parentheses like "(A)" or "(B)"
                import re
                account_match = re.match(r'^([\d\-]+)', account_str)
                if account_match:
                    current_account = account_match.group(1)
                else:
                    current_account = account_str
                
                account_allocations = []
                logger.debug(f"Row {row_idx}: New account {current_account}")
            
            # Skip if no current account
            if not current_account:
                continue
            
            # Get percentage from column N
            pct_val = row[percentage_col] if percentage_col < len(row) else None
            if pct_val is None or str(pct_val).startswith('#'):
                continue
            
            try:
                pct = float(pct_val)
                if pct <= 0:
                    continue
            except (ValueError, TypeError):
                continue
            
            # Get centre name from column O
            centre_val = row[centre_col] if centre_col < len(row) else None
            if centre_val is None or str(centre_val).startswith('#'):
                continue
            
            centre_full = str(centre_val).strip()
            if not centre_full:
                continue
            
            # Extract short code from long name
            # "AC - Commercial Air Conditioning" -> "AC"
            # "Hotel / Commercial / SA" -> ["C", "SA"]
            centre_code = self._extract_centre_code_from_long_name(centre_full)
            
            # CRITICAL: Validate centre is not numeric
            try:
                float(centre_code)
                logger.warning(f"Row {row_idx}: NUMERIC centre '{centre_code}' from '{centre_full}' - SKIPPING")
                continue
            except (ValueError, TypeError):
                pass  # Good - not numeric
            
            # Handle shared allocations (centre codes with "/")
            # Example: "C/SA" means split percentage evenly across C and SA
            if '/' in centre_code:
                centre_codes = [c.strip() for c in centre_code.split('/')]
                split_pct = pct / len(centre_codes)
                
                for code in centre_codes:
                    account_allocations.append({
                        'centre': code,
                        'percentage': split_pct
                    })
                    logger.debug(f"Row {row_idx}: account {current_account}, centre {code} ({centre_full} split), pct {split_pct:.4f}")
            else:
                # Single centre
                account_allocations.append({
                    'centre': centre_code,
                    'percentage': pct
                })
                logger.debug(f"Row {row_idx}: account {current_account}, centre {centre_code} ({centre_full}), pct {pct:.4f}")
        
        # Save last account
        if current_account and account_allocations:
            allocations[current_account] = account_allocations
            logger.info(f"Saved account {current_account}: {len(account_allocations)} centres")
        
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
            
            if matched_codes:
                return '/'.join(matched_codes)
        
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
