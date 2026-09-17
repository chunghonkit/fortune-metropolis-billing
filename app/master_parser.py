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
            # Check if this looks like a cost centre column
            if any(cc in h.upper() for cc in ['FC', 'SW', 'AC', 'DC', 'OC', 'AO', 'CP', 'SA']):
                centre_cols.append((idx, h))
            elif 'centre' in h.lower() or 'center' in h.lower():
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
        Parse AC DEPT sheet using column C weights (runtime source of truth).
        
        Expected structure:
        - Column A or B: Account/Meter ID
        - Column C: Weight value
        - Rows grouped by account (multiple cost centres per account)
        - Compute percentage = weight / Σ(weights) for each account group
        
        Example:
            Account     Centre    Weight
            55861       AC        525.48
            55861       SW        198.99
            → AC: 525.48 / (525.48 + 198.99) = 72.54%
            → SW: 198.99 / (525.48 + 198.99) = 27.46%
        """
        allocations = {}
        
        # Convert to list of rows
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            return allocations
        
        # Find header row
        header_row_idx = 0
        for idx, row in enumerate(rows[:10]):
            if any(str(cell).lower() in ['account', 'meter', 'acct', 'a/c', 'weight'] 
                   for cell in row if cell):
                header_row_idx = idx
                break
        
        # Parse structure: look for account, centre, weight columns
        # Try to auto-detect column positions
        headers = [str(cell).strip().lower() if cell else '' for cell in rows[header_row_idx]]
        
        # Find columns
        account_col = None
        centre_col = None
        weight_col = None
        
        for idx, h in enumerate(headers):
            if 'account' in h or 'meter' in h or 'acct' in h or 'a/c' in h:
                account_col = idx
            elif 'centre' in h or 'center' in h or 'dept' in h or 'cost' in h:
                centre_col = idx
            elif 'weight' in h or h == 'c':  # Column C is often labeled 'weight' or just 'c'
                weight_col = idx
        
        # If columns not found by header, use common positions
        if account_col is None:
            account_col = 0  # Column A
        if centre_col is None:
            centre_col = 1  # Column B
        if weight_col is None:
            weight_col = 2  # Column C
        
        # Group rows by account and collect weights
        account_groups = {}
        
        for row in rows[header_row_idx + 1:]:
            if not row or all(cell is None for cell in row):
                continue
            
            # Skip #REF! and error rows
            account_val = row[account_col] if account_col < len(row) else None
            if account_val is None or str(account_val).startswith('#'):
                continue
            
            account = str(account_val).strip().replace(' ', '')
            if not account or account.lower() in ['none', 'total', '']:
                continue
            
            # Get centre
            centre_val = row[centre_col] if centre_col < len(row) else None
            if centre_val is None or str(centre_val).startswith('#'):
                continue
            centre = str(centre_val).strip()
            
            # Get weight
            weight_val = row[weight_col] if weight_col < len(row) else None
            if weight_val is None or str(weight_val).startswith('#'):
                continue
            
            try:
                weight = float(weight_val)
                if weight <= 0:
                    continue
            except (ValueError, TypeError):
                continue
            
            # Add to account group
            if account not in account_groups:
                account_groups[account] = []
            
            account_groups[account].append({
                'centre': centre,
                'weight': weight
            })
        
        # Compute percentages for each account
        for account, entries in account_groups.items():
            total_weight = sum(e['weight'] for e in entries)
            if total_weight == 0:
                continue
            
            allocations[account] = [
                {
                    'centre': e['centre'],
                    'percentage': e['weight'] / total_weight
                }
                for e in entries
            ]
            
            logger.debug(f"Account {account}: {len(entries)} centres, total weight {total_weight}")
        
        return allocations
    
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
            col_str = str(col).upper()
            if any(cc in col_str for cc in ['FC', 'SW', 'AC', 'DC', 'OC', 'AO', 'CP', 'SA']):
                centre_cols.append(col)
            elif 'centre' in str(col).lower() or 'center' in str(col).lower():
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
