"""
Excel Exporter for Citybase Cost Sheet Format

Generates Excel workbook matching Citybase template structure:
- Year rollup sheet with 27 cost centres (FC → SA row order)
- Per-account ELECTRICITY COST ALLOCATION forms
"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from typing import Dict, List
from datetime import datetime
import logging

logger = logging.getLogger(__name__)


class CitybaseExporter:
    """
    Exports allocation results to Citybase-format Excel workbook.
    """
    
    # Standard cost centre order (FC → SA, 27 centres)
    COST_CENTRE_ORDER = [
        'FC',  # Fitness Centre
        'SW',  # Swimming Pool
        'AC',  # Air Conditioning
        'C',   # Commercial
        'DC',  # Domestic
        'OC',  # Office/Commercial
        'O',   # Office
        'AO',  # Apartment/Office
        'CP',  # Carpark
        'SA',  # Service Apartment
        # Additional centres that may appear in allocations
        'Hotel/Commercial/SA (11)',
        'Hotel/Commercial (10)',
        'SA/Commercial (16)',
        'Hotel/SA (13)',
        'Commercial/Office/SA (7)',
        'Office/Commercial (14)',
        'Office/Commercial/Carpark (15)',
        'Commercial Common (2)',
        'Public/Carpark(19)',
        'Hotel / L8 Premises(12)',
        'SA/Hotel/Carpark(18)',
        'SA/Commercial/Carpark(17)',
        'Hotel / Carpark(9)',
        'Commercial/Office/Hotel(6)',
        'Commercial Common(2)',
        'Commerical/Carpark/Hotel(5)',
        'Commercial/Carpark (4)',
    ]
    
    def __init__(self):
        self.workbook = None
        
    def export(self, 
               allocations: List[Dict],
               summary: Dict[str, Dict],
               month: str,
               output_path: str) -> str:
        """
        Export allocation results to Excel workbook.
        
        Args:
            allocations: List of allocation records
            summary: Summary dict by cost centre
            month: Month label (e.g., "August 2026")
            output_path: Output file path
            
        Returns:
            Output file path
        """
        self.workbook = Workbook()
        
        # Remove default sheet
        if 'Sheet' in self.workbook.sheetnames:
            del self.workbook['Sheet']
        
        # Create year rollup sheet
        self._create_year_rollup_sheet(summary, month)
        
        # Create per-account forms
        self._create_account_forms(allocations, month)
        
        # Save workbook
        self.workbook.save(output_path)
        logger.info(f"Exported to {output_path}")
        
        return output_path
    
    def _create_year_rollup_sheet(self, summary: Dict[str, Dict], month: str):
        """
        Create year rollup sheet with 27 cost centres in FC→SA order.
        """
        ws = self.workbook.create_sheet("Year Rollup")
        
        # Title
        ws['A1'] = 'ELECTRICITY COST ALLOCATION'
        ws['A1'].font = Font(bold=True, size=14)
        
        ws['A2'] = f'Period: {month}'
        ws['A2'].font = Font(italic=True)
        
        # Headers
        row = 4
        ws[f'A{row}'] = 'Cost Centre'
        ws[f'B{row}'] = 'Total Amount (HKD)'
        ws[f'C{row}'] = 'KWH'
        ws[f'D{row}'] = 'Accounts'
        
        for col in ['A', 'B', 'C', 'D']:
            cell = ws[f'{col}{row}']
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
            cell.border = Border(
                bottom=Side(style='thin'),
                top=Side(style='thin'),
                left=Side(style='thin'),
                right=Side(style='thin')
            )
        
        # Data rows in FC→SA order
        row = 5
        total_amount = 0
        total_kwh = 0
        
        for centre in self.COST_CENTRE_ORDER:
            if centre not in summary:
                continue
            
            data = summary[centre]
            ws[f'A{row}'] = centre
            ws[f'B{row}'] = data['total_amount']
            ws[f'B{row}'].number_format = '#,##0.00'
            ws[f'C{row}'] = data['kwh']
            ws[f'C{row}'].number_format = '#,##0'
            ws[f'D{row}'] = ', '.join(data['accounts'][:5])  # First 5 accounts
            
            total_amount += data['total_amount']
            total_kwh += data['kwh']
            row += 1
        
        # Total row
        ws[f'A{row}'] = 'TOTAL'
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'] = total_amount
        ws[f'B{row}'].number_format = '#,##0.00'
        ws[f'B{row}'].font = Font(bold=True)
        ws[f'C{row}'] = total_kwh
        ws[f'C{row}'].number_format = '#,##0'
        ws[f'C{row}'].font = Font(bold=True)
        
        for col in ['A', 'B', 'C', 'D']:
            cell = ws[f'{col}{row}']
            cell.border = Border(top=Side(style='double'))
        
        # Adjust column widths
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 40
    
    def _create_account_forms(self, allocations: List[Dict], month: str):
        """
        Create per-account ELECTRICITY COST ALLOCATION forms.
        """
        # Group allocations by account
        accounts_data = {}
        for alloc in allocations:
            account = alloc['account']
            if account not in accounts_data:
                accounts_data[account] = {
                    'allocations': [],
                    'bill_total': alloc['bill_total'],
                    'kwh': alloc['kwh']
                }
            accounts_data[account]['allocations'].append(alloc)
        
        # Create a sheet for each account
        for account, data in sorted(accounts_data.items()):
            sheet_name = account[:31]  # Excel sheet name limit
            ws = self.workbook.create_sheet(sheet_name)
            
            # Form header
            ws['A1'] = 'ELECTRICITY COST ALLOCATION FORM'
            ws['A1'].font = Font(bold=True, size=12)
            
            ws['A2'] = f'Account: {account}'
            ws['A2'].font = Font(bold=True)
            
            ws['A3'] = f'Period: {month}'
            
            ws['A4'] = f'Total Bill: HKD {data["bill_total"]:,.2f}'
            ws['A5'] = f'Total KWH: {data["kwh"]:,.0f}'
            
            # Allocation table
            row = 7
            ws[f'A{row}'] = 'Cost Centre'
            ws[f'B{row}'] = 'Percentage'
            ws[f'C{row}'] = 'Allocated Amount (HKD)'
            
            for col in ['A', 'B', 'C']:
                cell = ws[f'{col}{row}']
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')
            
            row = 8
            total_allocated = 0
            
            for alloc in data['allocations']:
                ws[f'A{row}'] = alloc['centre']
                ws[f'B{row}'] = f"{alloc['percentage'] * 100:.2f}%"
                ws[f'C{row}'] = alloc['amount']
                ws[f'C{row}'].number_format = '#,##0.00'
                
                total_allocated += alloc['amount']
                row += 1
            
            # Total
            ws[f'A{row}'] = 'TOTAL'
            ws[f'A{row}'].font = Font(bold=True)
            ws[f'C{row}'] = total_allocated
            ws[f'C{row}'].number_format = '#,##0.00'
            ws[f'C{row}'].font = Font(bold=True)
            
            # Residual check
            row += 2
            residual = data['bill_total'] - total_allocated
            ws[f'A{row}'] = 'Residual:'
            ws[f'B{row}'] = residual
            ws[f'B{row}'].number_format = '#,##0.00'
            
            if abs(residual) <= 0.01:
                ws[f'C{row}'] = '✓ OK'
                ws[f'C{row}'].font = Font(color='008000')
            else:
                ws[f'C{row}'] = '⚠ Check'
                ws[f'C{row}'].font = Font(color='FF0000')
            
            # Adjust column widths
            ws.column_dimensions['A'].width = 30
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 25


def export_to_citybase_excel(allocations: List[Dict],
                             summary: Dict[str, Dict],
                             month: str,
                             output_path: str) -> str:
    """
    Export allocation results to Citybase Excel format.
    
    Args:
        allocations: List of allocation records
        summary: Summary dict by cost centre
        month: Month label
        output_path: Output file path
        
    Returns:
        Output file path
    """
    exporter = CitybaseExporter()
    return exporter.export(allocations, summary, month, output_path)
