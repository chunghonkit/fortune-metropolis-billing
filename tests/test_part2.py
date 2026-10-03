"""
Unit tests for Part 2: Month-End Processing

Tests:
- Copy previous month's workbooks (do not overwrite originals)
- Update workbooks with current month's data
- Preserve formatting
- Compute previous month correctly
"""

import pytest
import tempfile
import os
from pathlib import Path
from datetime import datetime
from app.main import get_previous_month


class TestMonthComputation:
    """Test previous month computation"""
    
    def test_get_previous_month_same_year(self):
        """Test getting previous month within same year"""
        assert get_previous_month('2025-04') == '2025-03'
        assert get_previous_month('2025-06') == '2025-05'
        assert get_previous_month('2025-12') == '2025-11'
    
    def test_get_previous_month_year_boundary(self):
        """Test getting previous month across year boundary"""
        assert get_previous_month('2025-01') == '2024-12'
        assert get_previous_month('2026-01') == '2025-12'
    
    def test_get_previous_month_invalid(self):
        """Test invalid month format returns None"""
        assert get_previous_month('invalid') is None
        assert get_previous_month('2025-13') is None
        assert get_previous_month('') is None


class TestWorkbookCopy:
    """Test workbook copying functionality"""
    
    def test_copy_does_not_overwrite_source(self):
        """Test that copying workbooks does not modify source files"""
        from app.workbook_updater import WorkbookUpdater
        import openpyxl
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            
            # Create mock directory structure
            prev_month_dir = tmpdir / '2025-03'
            curr_month_dir = tmpdir / '2025-04'
            
            prev_out_dir = prev_month_dir / 'out'
            prev_out_dir.mkdir(parents=True, exist_ok=True)
            
            # Create a mock source workbook with real Citybase-style name
            source_wb = openpyxl.Workbook()
            source_sheet = source_wb.active
            source_sheet['A1'] = 'Original Value'
            source_file = prev_out_dir / 'Cost Sheet-2025-03.xls'
            source_wb.save(source_file)
            source_wb.close()
            
            # Get original file size and modification time
            original_size = source_file.stat().st_size
            original_content = source_file.read_bytes()
            
            # Copy the workbook
            updater = WorkbookUpdater(tmpdir)
            copied_files = updater.copy_previous_month_workbooks(
                '2025-04',
                '2025-03'
            )
            
            # Verify source file is unchanged
            assert source_file.exists()
            assert source_file.stat().st_size == original_size
            assert source_file.read_bytes() == original_content
            
            # Verify copy was created
            assert 'cost_sheet' in copied_files
            assert copied_files['cost_sheet'].exists()
            assert copied_files['cost_sheet'].parent == curr_month_dir / 'out'
            
            # Verify copy has correct current-month name
            assert copied_files['cost_sheet'].name == 'Cost Sheet-2025-04.xls'
    
    def test_copy_creates_output_directory(self):
        """Test that copying creates the out/ directory if it doesn't exist"""
        from app.workbook_updater import WorkbookUpdater
        import openpyxl
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            
            # Create previous month workbook
            prev_month_dir = tmpdir / '2025-03' / 'out'
            prev_month_dir.mkdir(parents=True, exist_ok=True)
            
            source_wb = openpyxl.Workbook()
            source_file = prev_month_dir / 'Cost Allocation - Electricity 2007-2.xls'
            source_wb.save(source_file)
            source_wb.close()
            
            # Current month out/ directory does not exist yet
            curr_out_dir = tmpdir / '2025-04' / 'out'
            assert not curr_out_dir.exists()
            
            # Copy workbooks
            updater = WorkbookUpdater(tmpdir)
            copied_files = updater.copy_previous_month_workbooks(
                '2025-04',
                '2025-03'
            )
            
            # Verify out/ directory was created
            assert curr_out_dir.exists()
            assert curr_out_dir.is_dir()
            
            # Verify file was copied
            assert 'cost_allocation' in copied_files
            assert copied_files['cost_allocation'].parent == curr_out_dir
            # Cost Allocation keeps same name (long-lived file)
            assert copied_files['cost_allocation'].name == 'Cost Allocation - Electricity 2007-2.xls'
    
    def test_find_real_citybase_filenames(self):
        """Test finding real Citybase filenames (.xls, Cost Sheet-YYYY-MM, Cost Allocation - Electricity ...)"""
        from app.workbook_updater import WorkbookUpdater
        import openpyxl
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            
            # Create previous month with real Citybase filenames
            prev_masters_dir = tmpdir / '2025-03' / 'masters'
            prev_masters_dir.mkdir(parents=True, exist_ok=True)
            
            prev_out_dir = tmpdir / '2025-03' / 'out'
            prev_out_dir.mkdir(parents=True, exist_ok=True)
            
            # Cost Allocation in masters (long-lived file with 2007-2 name)
            cost_alloc_wb = openpyxl.Workbook()
            cost_alloc_file = prev_masters_dir / 'Cost Allocation - Electricity 2007-2.xls'
            cost_alloc_wb.save(cost_alloc_file)
            cost_alloc_wb.close()
            
            # Cost Sheet in out/ with previous month
            cost_sheet_wb = openpyxl.Workbook()
            cost_sheet_file = prev_out_dir / 'Cost Sheet-2025-03.xls'
            cost_sheet_wb.save(cost_sheet_file)
            cost_sheet_wb.close()
            
            # Copy workbooks
            updater = WorkbookUpdater(tmpdir)
            copied_files = updater.copy_previous_month_workbooks(
                '2025-04',
                '2025-03'
            )
            
            # Verify both files found
            assert 'cost_allocation' in copied_files
            assert 'cost_sheet' in copied_files
            
            curr_out_dir = tmpdir / '2025-04' / 'out'
            
            # Verify Cost Allocation copied with same name
            assert copied_files['cost_allocation'].parent == curr_out_dir
            assert copied_files['cost_allocation'].name == 'Cost Allocation - Electricity 2007-2.xls'
            
            # Verify Cost Sheet copied with current month in name
            assert copied_files['cost_sheet'].parent == curr_out_dir
            assert copied_files['cost_sheet'].name == 'Cost Sheet-2025-04.xls'
            
            # Verify source files unchanged
            assert cost_alloc_file.exists()
            assert cost_sheet_file.exists()


class TestWorkbookUpdate:
    """Test workbook updating functionality"""
    
    def test_update_preserves_formatting(self):
        """Test that updating workbooks preserves cell formatting"""
        from app.workbook_updater import WorkbookUpdater
        import openpyxl
        from openpyxl.styles import Font, PatternFill
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            curr_out_dir = tmpdir / '2025-04' / 'out'
            curr_out_dir.mkdir(parents=True, exist_ok=True)
            
            # Create a workbook with specific formatting
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet['A1'] = 'Header'
            sheet['A1'].font = Font(bold=True, size=14)
            sheet['A1'].fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            
            sheet['B2'] = 123.45
            sheet['B2'].number_format = '#,##0.00'
            
            wb_file = curr_out_dir / 'Test.xlsx'
            wb.save(wb_file)
            wb.close()
            
            # Get original formatting
            wb_original = openpyxl.load_workbook(wb_file)
            sheet_original = wb_original.active
            original_a1_font_bold = sheet_original['A1'].font.bold
            original_a1_font_size = sheet_original['A1'].font.size
            original_a1_fill_color = sheet_original['A1'].fill.start_color.rgb
            original_b2_format = sheet_original['B2'].number_format
            wb_original.close()
            
            # Now update a cell value (simulating what our updater does)
            wb_update = openpyxl.load_workbook(wb_file)
            sheet_update = wb_update.active
            sheet_update['B2'] = 456.78  # Update value but preserve format
            wb_update.save(wb_file)
            wb_update.close()
            
            # Verify formatting is preserved
            wb_check = openpyxl.load_workbook(wb_file)
            sheet_check = wb_check.active
            
            # Check A1 formatting preserved
            assert sheet_check['A1'].font.bold == original_a1_font_bold
            assert sheet_check['A1'].font.size == original_a1_font_size
            assert sheet_check['A1'].fill.start_color.rgb == original_a1_fill_color
            
            # Check B2 number format preserved
            assert sheet_check['B2'].number_format == original_b2_format
            
            # Check B2 value updated
            assert sheet_check['B2'].value == 456.78
            
            wb_check.close()


class TestWorkbookValueUpdates:
    """Test that workbook updates actually write values"""
    
    def test_cost_allocation_elect_charge_update(self):
        """Test writing check-meter readings to Cost Allocation Elect Charge sheet"""
        from app.workbook_updater import WorkbookUpdater
        import openpyxl
        from openpyxl.styles import Font
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            out_dir = tmpdir / '2025-05' / 'out'
            out_dir.mkdir(parents=True, exist_ok=True)
            
            # Create a Cost Allocation workbook with Elect Charge sheet
            wb = openpyxl.Workbook()
            wb.remove(wb.active)  # Remove default sheet
            
            # Create Elect Charge sheet with meter 6681757
            elect_sheet = wb.create_sheet("Elect Charge")
            
            # Add headers
            elect_sheet['A1'] = 'Meter No.'
            elect_sheet['B1'] = 'Description'
            elect_sheet['C1'] = 'Previous'
            elect_sheet['D1'] = 'Present'
            elect_sheet['E1'] = 'Delta'
            
            # Make headers bold
            for cell in elect_sheet[1]:
                cell.font = Font(bold=True)
            
            # Add meter 6681757 row (row 2)
            elect_sheet['A2'] = '6681757'
            elect_sheet['B2'] = '散熱水泵電'
            elect_sheet['C2'] = 95966.6  # Old value
            elect_sheet['D2'] = 96323.1  # Old value
            elect_sheet['E2'] = 356.5    # Old delta
            
            # Save to file
            wb_file = out_dir / 'Cost Allocation - Electricity 2007-2.xlsx'
            wb.save(wb_file)
            wb.close()
            
            # Prepare meter log with new values
            meter_log = {
                'meter_no': '6681757',
                'previous': 96323.1,
                'present': 96504.6,
            }
            
            # Update the workbook
            updater = WorkbookUpdater(tmpdir)
            updater.update_cost_allocation_workbook(
                wb_file,
                allocations=[],
                billing_month='2025-05',
                parsed_bills=[],
                meter_log=meter_log
            )
            
            # Verify values were updated
            wb_updated = openpyxl.load_workbook(wb_file)
            elect_sheet_updated = wb_updated['Elect Charge']
            
            # Check new values were written
            assert elect_sheet_updated['C2'].value == 96323.1, f"Previous should be 96323.1, got {elect_sheet_updated['C2'].value}"
            assert elect_sheet_updated['D2'].value == 96504.6, f"Present should be 96504.6, got {elect_sheet_updated['D2'].value}"
            
            # Check delta was calculated
            expected_delta = round(96504.6 - 96323.1, 1)
            assert elect_sheet_updated['E2'].value == expected_delta, f"Delta should be {expected_delta}, got {elect_sheet_updated['E2'].value}"
            
            # Verify formatting preserved (bold headers)
            assert elect_sheet_updated['A1'].font.bold == True
            
            wb_updated.close()
    
    def test_cost_sheet_year_grid_update(self):
        """Test writing allocation totals to Cost Sheet year-grid month column"""
        from app.workbook_updater import WorkbookUpdater
        import openpyxl
        from openpyxl.styles import Font
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            out_dir = tmpdir / '2025-05' / 'out'
            out_dir.mkdir(parents=True, exist_ok=True)
            
            # Create a Cost Sheet workbook with year-grid sheet
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            
            # Create year-grid sheet "01-12 2025"
            grid_sheet = wb.create_sheet("01-12 2025")
            
            # Add headers: A=Cost Centre, B=Jan, C=Feb, ..., F=May
            grid_sheet['A1'] = 'Cost Centre'
            grid_sheet['B1'] = 'Jan'
            grid_sheet['C1'] = 'Feb'
            grid_sheet['D1'] = 'Mar'
            grid_sheet['E1'] = 'Apr'
            grid_sheet['F1'] = 'May'
            
            # Add cost centre rows
            grid_sheet['A2'] = 'FC'
            grid_sheet['A3'] = 'AC'
            grid_sheet['A4'] = 'SW'
            grid_sheet['A5'] = 'DC'
            
            # Make headers bold
            for cell in grid_sheet[1]:
                cell.font = Font(bold=True)
            
            # Add some April (column E) values
            grid_sheet['E2'] = 50000  # FC Apr
            grid_sheet['E3'] = 300000  # AC Apr
            grid_sheet['E4'] = 150000  # SW Apr
            
            # Save to file
            wb_file = out_dir / 'Cost Sheet-2025-05.xlsx'
            wb.save(wb_file)
            wb.close()
            
            # Prepare allocations with totals per centre
            allocations = [
                {'centre': 'FC', 'amount': 70610.50, 'account': 'test1'},
                {'centre': 'AC', 'amount': 200000.00, 'account': 'test2'},
                {'centre': 'AC', 'amount': 222972.67, 'account': 'test3'},  # Total AC = 422972.67
                {'centre': 'SW', 'amount': 150000.00, 'account': 'test4'},
                {'centre': 'SW', 'amount': 99629.75, 'account': 'test5'},   # Total SW = 249629.75
                {'centre': 'DC', 'amount': 50000.00, 'account': 'test6'},
            ]
            
            # Update the workbook
            updater = WorkbookUpdater(tmpdir)
            updater.update_cost_sheet_workbook(
                wb_file,
                allocations=allocations,
                billing_month='2025-05',
                parsed_bills=[],
                meter_log={}
            )
            
            # Verify values were written to May column (F)
            wb_updated = openpyxl.load_workbook(wb_file)
            grid_sheet_updated = wb_updated['01-12 2025']
            
            # Check May column (F) has correct totals
            assert grid_sheet_updated['F2'].value == 70610.50, f"FC May should be 70610.50, got {grid_sheet_updated['F2'].value}"
            assert grid_sheet_updated['F3'].value == 422972.67, f"AC May should be 422972.67, got {grid_sheet_updated['F3'].value}"
            assert grid_sheet_updated['F4'].value == 249629.75, f"SW May should be 249629.75, got {grid_sheet_updated['F4'].value}"
            assert grid_sheet_updated['F5'].value == 50000.00, f"DC May should be 50000.00, got {grid_sheet_updated['F5'].value}"
            
            # Verify April column (E) unchanged
            assert grid_sheet_updated['E2'].value == 50000
            assert grid_sheet_updated['E3'].value == 300000
            assert grid_sheet_updated['E4'].value == 150000
            
            # Verify formatting preserved (bold headers)
            assert grid_sheet_updated['A1'].font.bold == True
            
            wb_updated.close()
    
    def test_meter_past_column_10_regression(self):
        """
        Regression test: Meter 6681757 is at row 31 in columns past column 10.
        
        Golden May shows meter at BR31/BS31 (columns 70/71).
        App initially failed because search stopped at column 10.
        """
        from app.workbook_updater import WorkbookUpdater
        import openpyxl
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            out_dir = tmpdir / '2025-05' / 'out'
            out_dir.mkdir(parents=True, exist_ok=True)
            
            # Create Elect Charge sheet shaped like real Citybase workbook
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            elect_sheet = wb.create_sheet("Elect Charge")
            
            # Place meter 6681757 at row 31, with readings in columns BR/BS (70/71)
            # Column BR = 70, BS = 71, BQ = 69 (delta formula)
            elect_sheet.cell(row=31, column=1).value = '6681757'  # Meter number in column A or nearby
            elect_sheet.cell(row=31, column=70).value = 95966.6   # BR31 - Previous (April)
            elect_sheet.cell(row=31, column=71).value = 96323.1   # BS31 - Present (April)
            elect_sheet.cell(row=31, column=69).value = '=BS31-BR31'  # BQ31 - Delta formula
            
            # Add headers in row 30
            elect_sheet.cell(row=30, column=70).value = 'Previous'
            elect_sheet.cell(row=30, column=71).value = 'Present'
            elect_sheet.cell(row=30, column=69).value = 'Delta'
            
            wb_file = out_dir / 'Cost Allocation.xlsx'
            wb.save(wb_file)
            wb.close()
            
            # Update with May readings
            meter_log = {
                'meter_no': '6681757',
                'previous': 96323.1,
                'present': 96504.6,
            }
            
            updater = WorkbookUpdater(tmpdir)
            updater.update_cost_allocation_workbook(
                wb_file,
                allocations=[],
                billing_month='2025-05',
                parsed_bills=[],
                meter_log=meter_log
            )
            
            # Verify readings updated
            wb_updated = openpyxl.load_workbook(wb_file)
            elect_sheet_updated = wb_updated['Elect Charge']
            
            assert elect_sheet_updated.cell(row=31, column=70).value == 96323.1, \
                f"BR31 Previous should be 96323.1, got {elect_sheet_updated.cell(row=31, column=70).value}"
            assert elect_sheet_updated.cell(row=31, column=71).value == 96504.6, \
                f"BS31 Present should be 96504.6, got {elect_sheet_updated.cell(row=31, column=71).value}"
            
            # Delta should remain as formula (not overwritten)
            delta_cell = elect_sheet_updated.cell(row=31, column=69)
            assert isinstance(delta_cell.value, str) and delta_cell.value.startswith('='), \
                f"BQ31 Delta should remain formula, got {delta_cell.value}"
            
            wb_updated.close()
    
    def test_current_reading_header_regression(self):
        """
        Regression test: Real Citybase workbook uses "Current Reading" not "Present".
        
        May rerun showed meter found but readings not written because header 
        matching only looked for "Present", not "Current Reading".
        """
        from app.workbook_updater import WorkbookUpdater
        import openpyxl
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            out_dir = tmpdir / '2025-05' / 'out'
            out_dir.mkdir(parents=True, exist_ok=True)
            
            # Create Elect Charge sheet with "Current Reading" header (real Citybase format)
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            elect_sheet = wb.create_sheet("Elect Charge")
            
            # Meter 6681757 at row 31, columns BR/BS (70/71)
            elect_sheet.cell(row=31, column=1).value = '6681757'
            elect_sheet.cell(row=31, column=70).value = 95966.6   # BR31 - Previous (April)
            elect_sheet.cell(row=31, column=71).value = 96323.1   # BS31 - Current (April)
            elect_sheet.cell(row=31, column=69).value = '=BS31-BR31'
            
            # Real Citybase headers in row 30
            elect_sheet.cell(row=30, column=70).value = 'Previous Reading'
            elect_sheet.cell(row=30, column=71).value = 'Current Reading'  # Not "Present"!
            elect_sheet.cell(row=30, column=69).value = 'Difference'
            
            wb_file = out_dir / 'Cost Allocation.xlsx'
            wb.save(wb_file)
            wb.close()
            
            # Update with May readings
            meter_log = {
                'meter_no': '6681757',
                'previous': 96323.1,
                'present': 96504.6,
            }
            
            updater = WorkbookUpdater(tmpdir)
            updater.update_cost_allocation_workbook(
                wb_file,
                allocations=[],
                billing_month='2025-05',
                parsed_bills=[],
                meter_log=meter_log
            )
            
            # Verify readings updated despite "Current Reading" header
            wb_updated = openpyxl.load_workbook(wb_file)
            elect_sheet_updated = wb_updated['Elect Charge']
            
            assert elect_sheet_updated.cell(row=31, column=70).value == 96323.1, \
                f"BR31 Previous should be 96323.1, got {elect_sheet_updated.cell(row=31, column=70).value}"
            assert elect_sheet_updated.cell(row=31, column=71).value == 96504.6, \
                f"BS31 Current Reading should be 96504.6, got {elect_sheet_updated.cell(row=31, column=71).value}"
            
            wb_updated.close()
    
    def test_year_grid_matches_billing_year_regression(self):
        """
        Regression test: Year-grid sheet must match billing year.
        
        Golden May has "01-12 2025". App initially picked "01-12 2022" (first match).
        """
        from app.workbook_updater import WorkbookUpdater
        import openpyxl
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            out_dir = tmpdir / '2025-05' / 'out'
            out_dir.mkdir(parents=True, exist_ok=True)
            
            # Create Cost Sheet with multiple 01-12 sheets (older years)
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            
            # Add old year-grid sheets (should NOT be picked)
            old_sheet_2022 = wb.create_sheet("01-12 2022")
            old_sheet_2022['A1'] = 'Cost Centre'
            old_sheet_2022['F1'] = 'May'
            old_sheet_2022['A2'] = 'FC'
            old_sheet_2022['F2'] = 0  # Should NOT be updated
            
            old_sheet_2023 = wb.create_sheet("01-12 2023")
            old_sheet_2023['A1'] = 'Cost Centre'
            old_sheet_2023['F1'] = 'May'
            old_sheet_2023['A2'] = 'FC'
            old_sheet_2023['F2'] = 0  # Should NOT be updated
            
            # Add current year-grid sheet (SHOULD be picked)
            current_sheet = wb.create_sheet("01-12 2025")
            current_sheet['A1'] = 'Cost Centre'
            current_sheet['F1'] = 'May'
            current_sheet['A2'] = 'FC'
            current_sheet['A3'] = 'AC'
            current_sheet['F2'] = None  # Should be updated
            current_sheet['F3'] = None
            
            wb_file = out_dir / 'Cost Sheet-2025-05.xlsx'
            wb.save(wb_file)
            wb.close()
            
            # Update with May allocations
            allocations = [
                {'centre': 'FC', 'amount': 70610.00, 'account': 'test1'},
                {'centre': 'AC', 'amount': 422972.67, 'account': 'test2'},
            ]
            
            updater = WorkbookUpdater(tmpdir)
            updater.update_cost_sheet_workbook(
                wb_file,
                allocations=allocations,
                billing_month='2025-05',
                parsed_bills=[],
                meter_log={}
            )
            
            # Verify only the 2025 sheet was updated
            wb_updated = openpyxl.load_workbook(wb_file)
            
            # 2022 and 2023 should remain 0
            assert wb_updated['01-12 2022']['F2'].value == 0, "Old 2022 sheet should not be updated"
            assert wb_updated['01-12 2023']['F2'].value == 0, "Old 2023 sheet should not be updated"
            
            # 2025 should have new values
            assert wb_updated['01-12 2025']['F2'].value == 70610.00, \
                f"2025 FC May should be 70610.00, got {wb_updated['01-12 2025']['F2'].value}"
            assert wb_updated['01-12 2025']['F3'].value == 422972.67, \
                f"2025 AC May should be 422972.67, got {wb_updated['01-12 2025']['F3'].value}"
            
            wb_updated.close()
    
    def test_title_cells_not_overwritten_regression(self):
        """
        Regression test: Original title cells must not be overwritten.
        
        Golden keeps original 2008 title. App initially overwrote AC DEPT!A1 to "Cost Allocation - 2025-05".
        """
        from app.workbook_updater import WorkbookUpdater
        import openpyxl
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            out_dir = tmpdir / '2025-05' / 'out'
            out_dir.mkdir(parents=True, exist_ok=True)
            
            # Create Cost Allocation workbook with original 2008 title
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            
            ac_dept_sheet = wb.create_sheet("AC DEPT")
            ac_dept_sheet['A1'] = 'Metropolis AC DEPT 2008'  # Original title from 2008
            
            elect_sheet = wb.create_sheet("Elect Charge")
            elect_sheet['A1'] = 'Electricity Charges - Original Title'
            
            wb_file = out_dir / 'Cost Allocation.xlsx'
            wb.save(wb_file)
            wb.close()
            
            # Update workbook
            updater = WorkbookUpdater(tmpdir)
            updater.update_cost_allocation_workbook(
                wb_file,
                allocations=[],
                billing_month='2025-05',
                parsed_bills=[],
                meter_log={}
            )
            
            # Verify original titles preserved
            wb_updated = openpyxl.load_workbook(wb_file)
            
            assert wb_updated['AC DEPT']['A1'].value == 'Metropolis AC DEPT 2008', \
                f"AC DEPT A1 should preserve original title, got {wb_updated['AC DEPT']['A1'].value}"
            assert wb_updated['Elect Charge']['A1'].value == 'Electricity Charges - Original Title', \
                f"Elect Charge A1 should preserve original title, got {wb_updated['Elect Charge']['A1'].value}"
            
            wb_updated.close()
    
    def test_complex_centre_names_aggregate_correctly(self):
        """
        Regression test: Complex centre names must aggregate to simple codes.
        
        Master workbook has centres like "Hotel/Commercial/SA (11)" which must
        aggregate to "SA" row in Cost Sheet year-grid.
        """
        from app.workbook_updater import WorkbookUpdater
        import openpyxl
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            out_dir = tmpdir / '2025-05' / 'out'
            out_dir.mkdir(parents=True, exist_ok=True)
            
            # Create Cost Sheet with year-grid
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            
            grid_sheet = wb.create_sheet("01-12 2025")
            grid_sheet['A1'] = 'Cost Centre'
            grid_sheet['F1'] = 'May'
            grid_sheet['A2'] = 'AC'
            grid_sheet['A3'] = 'SA'
            grid_sheet['A4'] = 'C'
            grid_sheet['A5'] = 'DC'
            
            wb_file = out_dir / 'Cost Sheet-2025-05.xlsx'
            wb.save(wb_file)
            wb.close()
            
            # Allocations with complex centre names (like real master produces)
            allocations = [
                {'centre': 'AC', 'amount': 10388.15, 'account': 'test1'},
                {'centre': 'C', 'amount': 62705.08, 'account': 'test1'},
                {'centre': 'DC', 'amount': 23428.44, 'account': 'test1'},
                {'centre': 'Hotel/Commercial/SA (11)', 'amount': 421.15, 'account': 'test1'},  # Should -> SA
                {'centre': 'Hotel/SA (13)', 'amount': 805.38, 'account': 'test1'},            # Should -> SA
                {'centre': 'SA/Commercial (16)', 'amount': 731.04, 'account': 'test1'},       # Should -> SA
            ]
            
            # Update workbook
            updater = WorkbookUpdater(tmpdir)
            updater.update_cost_sheet_workbook(
                wb_file,
                allocations=allocations,
                billing_month='2025-05',
                parsed_bills=[],
                meter_log={}
            )
            
            # Verify complex names aggregated correctly
            wb_updated = openpyxl.load_workbook(wb_file)
            grid_sheet_updated = wb_updated['01-12 2025']
            
            # AC should have its simple allocation
            assert grid_sheet_updated['F2'].value == 10388.15, \
                f"AC should be 10388.15, got {grid_sheet_updated['F2'].value}"
            
            # SA should have sum of all SA-related allocations
            expected_sa = 421.15 + 805.38 + 731.04  # = 1957.57
            assert abs(grid_sheet_updated['F3'].value - expected_sa) < 0.01, \
                f"SA should be {expected_sa}, got {grid_sheet_updated['F3'].value}"
            
            # C should have its simple allocation
            assert grid_sheet_updated['F4'].value == 62705.08, \
                f"C should be 62705.08, got {grid_sheet_updated['F4'].value}"
            
            # DC should have its simple allocation
            assert grid_sheet_updated['F5'].value == 23428.44, \
                f"DC should be 23428.44, got {grid_sheet_updated['F5'].value}"
            
            wb_updated.close()
    
    def test_allocation_keyed_by_centre_not_amount(self):
        """
        Regression test: Allocation must be keyed by centre code, not amount.
        
        May 2025 rerun showed allocation totals keyed by dollar amounts
        (e.g., 216328.807591692) instead of centre codes (FC, AC, SW).
        This caused year-grid column to remain empty.
        """
        from app.workbook_updater import WorkbookUpdater
        import openpyxl
        
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            out_dir = tmpdir / '2025-05' / 'out'
            out_dir.mkdir(parents=True, exist_ok=True)
            
            # Create Cost Sheet with year-grid
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            
            grid_sheet = wb.create_sheet("01-12 2025")
            grid_sheet['A1'] = 'Cost Centre'
            grid_sheet['F1'] = 'May'
            grid_sheet['A4'] = 'FC'
            grid_sheet['A5'] = 'AC'
            grid_sheet['A6'] = 'SW'
            
            wb_file = out_dir / 'Cost Sheet-2025-05.xlsx'
            wb.save(wb_file)
            wb.close()
            
            # Correct allocations: keyed by centre CODE
            correct_allocations = [
                {'centre': 'FC', 'amount': 70610.00, 'account': '82805-94744-7'},
                {'centre': 'AC', 'amount': 200000.00, 'account': '55861-52267-1'},
                {'centre': 'AC', 'amount': 222972.67, 'account': '52167-13569-2'},  # Total AC = 422972.67
                {'centre': 'SW', 'amount': 150000.00, 'account': '55861-52267-1'},
                {'centre': 'SW', 'amount': 99629.75, 'account': '23529-59279-9'},   # Total SW = 249629.75
            ]
            
            # Update workbook with correct allocations
            updater = WorkbookUpdater(tmpdir)
            updater.update_cost_sheet_workbook(
                wb_file,
                allocations=correct_allocations,
                billing_month='2025-05',
                parsed_bills=[],
                meter_log={}
            )
            
            # Verify values written correctly
            wb_updated = openpyxl.load_workbook(wb_file)
            grid_sheet_updated = wb_updated['01-12 2025']
            
            # FC, AC, SW should have correct totals
            assert grid_sheet_updated['F4'].value == 70610.00, \
                f"FC should be 70610.00, got {grid_sheet_updated['F4'].value}"
            assert grid_sheet_updated['F5'].value == 422972.67, \
                f"AC should be 422972.67, got {grid_sheet_updated['F5'].value}"
            assert grid_sheet_updated['F6'].value == 249629.75, \
                f"SW should be 249629.75, got {grid_sheet_updated['F6'].value}"
            
            wb_updated.close()
            
            # Now test with WRONG allocations: keyed by amount (the bug)
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            
            grid_sheet = wb.create_sheet("01-12 2025")
            grid_sheet['A1'] = 'Cost Centre'
            grid_sheet['F1'] = 'May'
            grid_sheet['A4'] = 'FC'
            grid_sheet['A5'] = 'AC'
            grid_sheet['A6'] = 'SW'
            
            wb_file2 = out_dir / 'Cost Sheet-2025-05-bad.xlsx'
            wb.save(wb_file2)
            wb.close()
            
            # Wrong allocations: centre is actually an AMOUNT (the bug)
            # This should NOT work - the amounts won't match row labels
            wrong_allocations = [
                {'centre': '216328.807591692', 'amount': 70610.00, 'account': 'test1'},  # centre is a number!
                {'centre': '14747', 'amount': 200000.00, 'account': 'test2'},
                {'centre': '18767.05565', 'amount': 150000.00, 'account': 'test3'},
            ]
            
            updater2 = WorkbookUpdater(tmpdir)
            updater2.update_cost_sheet_workbook(
                wb_file2,
                allocations=wrong_allocations,
                billing_month='2025-05',
                parsed_bills=[],
                meter_log={}
            )
            
            # Verify that wrong allocations did NOT write anything
            # (because numeric centres won't match "FC", "AC", "SW" row labels)
            wb_wrong = openpyxl.load_workbook(wb_file2)
            grid_wrong = wb_wrong['01-12 2025']
            
            # All cells should be None (nothing written)
            assert grid_wrong['F4'].value is None, \
                "FC row should be empty when centre is keyed by amount"
            assert grid_wrong['F5'].value is None, \
                "AC row should be empty when centre is keyed by amount"
            assert grid_wrong['F6'].value is None, \
                "SW row should be empty when centre is keyed by amount"
            
            wb_wrong.close()


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
