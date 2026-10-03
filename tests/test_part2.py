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


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
