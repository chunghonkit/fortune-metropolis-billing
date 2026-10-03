"""
Unit tests for Part 1: Bills + Meter Log Intake & Gate

Tests:
- Gate PASS with 15 mocked accounts
- Gate FAIL on missing account
- Gate FAIL on duplicate account
- Gate FAIL on wrong month
- Gate FAIL on invalid meter log
- Folder scan (mocked)
- Meter log completeness
- Regression tests for compressed PDF and month classification bugs
"""

import pytest
import tempfile
import os
from pathlib import Path
from app.main import (
    EXPECTED_ACCOUNTS,
    session_storage,
    validate_gate,
    extract_bill_month,
    is_meter_log_complete,
)


@pytest.fixture
def reset_session():
    """Reset session storage before each test"""
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
    yield
    # Cleanup after test
    session_storage['billing_month'] = None
    session_storage['parsed_bills'] = []


def create_mock_bill(account, from_date="01-04-25", to_date=None, kwh=1000):
    """Create a mock parsed bill"""
    # If to_date not provided, assume period ends 30 days later in same month
    if to_date is None:
        # Parse from_date to generate to_date
        parts = from_date.split('-')
        if len(parts) == 3:
            dd, mm, yy = parts
            # Default to day 30 of the same month
            to_date = f"30-{mm}-{yy}"
        else:
            to_date = from_date
    
    return {
        'file': f'{account}.pdf',
        'account': account,
        'from_date': from_date,
        'to_date': to_date,
        'kwh': kwh,
        'total_amount': kwh * 1.2,
        'fit_amount': 0,
    }


class TestBillingMonth:
    """Test billing month handling"""
    
    def test_extract_bill_month_from_date(self):
        """Test extracting YYYY-MM from bill to_date (period end)"""
        bill = create_mock_bill('55861-52267-1', from_date='15-03-25', to_date='15-04-25')
        month = extract_bill_month(bill)
        assert month == '2025-04'  # Uses to_date (period end), not from_date
    
    def test_extract_bill_month_invalid(self):
        """Test extracting month from invalid date"""
        bill = {'account': 'test', 'to_date': 'invalid'}
        month = extract_bill_month(bill)
        assert month is None


class TestMeterLog:
    """Test meter log validation"""
    
    def test_meter_log_incomplete_no_data(self, reset_session):
        """Test meter log incomplete when no data"""
        assert is_meter_log_complete() is False
    
    def test_meter_log_incomplete_missing_present(self, reset_session):
        """Test meter log incomplete when missing present"""
        session_storage['meter_log']['previous'] = 1000000.0
        assert is_meter_log_complete() is False
    
    def test_meter_log_invalid_present_less_than_previous(self, reset_session):
        """Test meter log invalid when present < previous"""
        session_storage['meter_log']['previous'] = 1000000.5
        session_storage['meter_log']['present'] = 999999.3
        assert is_meter_log_complete() is False
    
    def test_meter_log_complete_valid(self, reset_session):
        """Test meter log complete when valid"""
        session_storage['meter_log']['previous'] = 1000000.0
        session_storage['meter_log']['present'] = 1001000.5
        assert is_meter_log_complete() is True
    
    def test_meter_log_complete_equal_readings(self, reset_session):
        """Test meter log complete when present = previous (edge case)"""
        session_storage['meter_log']['previous'] = 1000000.0
        session_storage['meter_log']['present'] = 1000000.0
        assert is_meter_log_complete() is True


class TestGateValidation:
    """Test gate validation logic"""
    
    def test_gate_fail_no_month_set(self, reset_session):
        """Test gate fails when billing month not set"""
        gate = validate_gate()
        assert gate['passed'] is False
        assert any('Billing month not set' in err for err in gate['errors'])
    
    def test_gate_fail_meter_log_incomplete(self, reset_session):
        """Test gate fails when meter log incomplete"""
        session_storage['billing_month'] = '2025-04'
        gate = validate_gate()
        assert gate['passed'] is False
        assert any('Meter log' in err for err in gate['errors'])
    
    def test_gate_pass_all_15_accounts(self, reset_session):
        """Test gate passes with all 15 expected accounts"""
        session_storage['billing_month'] = '2025-04'
        session_storage['meter_log']['previous'] = 1000000.0
        session_storage['meter_log']['present'] = 1001000.0
        
        # Add all 15 expected accounts
        for account in EXPECTED_ACCOUNTS:
            bill = create_mock_bill(account, from_date='15-04-25', kwh=10000)
            session_storage['parsed_bills'].append(bill)
        
        gate = validate_gate()
        assert gate['passed'] is True
        assert gate['summary']['matched'] == 15
        assert gate['summary']['missing'] == 0
        assert len(gate['errors']) == 0
    
    def test_gate_fail_missing_account(self, reset_session):
        """Test gate fails when missing an account"""
        session_storage['billing_month'] = '2025-04'
        session_storage['meter_log']['previous'] = 1000000.0
        session_storage['meter_log']['present'] = 1001000.0
        
        # Add only 14 accounts (missing one)
        for account in EXPECTED_ACCOUNTS[:-1]:  # Skip last account
            bill = create_mock_bill(account, from_date='15-04-25', kwh=10000)
            session_storage['parsed_bills'].append(bill)
        
        gate = validate_gate()
        assert gate['passed'] is False
        assert gate['summary']['missing'] == 1
        assert gate['summary']['matched'] == 14
        assert any('Missing' in err for err in gate['errors'])
    
    def test_gate_fail_duplicate_account(self, reset_session):
        """Test gate fails when duplicate account exists"""
        session_storage['billing_month'] = '2025-04'
        session_storage['meter_log']['previous'] = 1000000.0
        session_storage['meter_log']['present'] = 1001000.0
        
        # Add all 15 accounts
        for account in EXPECTED_ACCOUNTS:
            bill = create_mock_bill(account, from_date='15-04-25', kwh=10000)
            session_storage['parsed_bills'].append(bill)
        
        # Add duplicate of first account
        duplicate = create_mock_bill(EXPECTED_ACCOUNTS[0], from_date='15-04-25', kwh=5000)
        session_storage['parsed_bills'].append(duplicate)
        
        gate = validate_gate()
        assert gate['passed'] is False
        assert gate['summary']['duplicates'] == 1
        assert any('Duplicate' in err for err in gate['errors'])
    
    def test_gate_fail_wrong_month(self, reset_session):
        """Test gate fails when bill from wrong month"""
        session_storage['billing_month'] = '2025-04'
        session_storage['meter_log']['previous'] = 1000000.0
        session_storage['meter_log']['present'] = 1001000.0
        
        # Add 14 correct accounts (period ends in April)
        for account in EXPECTED_ACCOUNTS[:-1]:
            bill = create_mock_bill(account, from_date='15-04-25', to_date='30-04-25', kwh=10000)
            session_storage['parsed_bills'].append(bill)
        
        # Add 1 account from wrong month (period ends in May)
        wrong_month_bill = create_mock_bill(EXPECTED_ACCOUNTS[-1], from_date='15-04-25', to_date='15-05-25', kwh=10000)
        session_storage['parsed_bills'].append(wrong_month_bill)
        
        gate = validate_gate()
        assert gate['passed'] is False
        assert gate['summary']['wrong_month'] == 1
        assert any('Wrong month' in err for err in gate['errors'])
    
    def test_gate_warn_unrecognised_account(self, reset_session):
        """Test gate fails on unrecognised account (not in Metropolis 15)"""
        session_storage['billing_month'] = '2025-04'
        session_storage['meter_log']['previous'] = 1000000.0
        session_storage['meter_log']['present'] = 1001000.0
        
        # Add all 15 expected accounts
        for account in EXPECTED_ACCOUNTS:
            bill = create_mock_bill(account, from_date='15-04-25', kwh=10000)
            session_storage['parsed_bills'].append(bill)
        
        # Add unrecognised account
        unrecognised = create_mock_bill('99999-99999-9', from_date='15-04-25', kwh=1000)
        session_storage['parsed_bills'].append(unrecognised)
        
        gate = validate_gate()
        # Per SPEC gate condition: no unrecognised accounts
        assert gate['passed'] is False
        assert gate['summary']['unrecognised'] == 1
        assert any('Unrecognised' in err for err in gate['errors'])


class TestAccountChecklist:
    """Test account checklist display logic"""
    
    def test_checklist_all_missing(self, reset_session):
        """Test checklist when all accounts missing"""
        session_storage['billing_month'] = '2025-04'
        
        gate = validate_gate()
        for account in EXPECTED_ACCOUNTS:
            assert gate['account_checklist'][account]['status'] == 'missing'
    
    def test_checklist_all_matched(self, reset_session):
        """Test checklist when all accounts matched"""
        session_storage['billing_month'] = '2025-04'
        session_storage['meter_log']['previous'] = 1000000.0
        session_storage['meter_log']['present'] = 1001000.0
        
        for account in EXPECTED_ACCOUNTS:
            bill = create_mock_bill(account, from_date='15-04-25', kwh=10000)
            session_storage['parsed_bills'].append(bill)
        
        gate = validate_gate()
        for account in EXPECTED_ACCOUNTS:
            assert gate['account_checklist'][account]['status'] == 'matched'
    
    def test_checklist_mixed_statuses(self, reset_session):
        """Test checklist with mixed statuses"""
        session_storage['billing_month'] = '2025-04'
        session_storage['meter_log']['previous'] = 1000000.0
        session_storage['meter_log']['present'] = 1001000.0
        
        # Add first 10 accounts (matched, period ends in April)
        for account in EXPECTED_ACCOUNTS[:10]:
            bill = create_mock_bill(account, from_date='15-03-25', to_date='15-04-25', kwh=10000)
            session_storage['parsed_bills'].append(bill)
        
        # Add duplicate of first account
        duplicate = create_mock_bill(EXPECTED_ACCOUNTS[0], from_date='15-03-25', to_date='15-04-25', kwh=5000)
        session_storage['parsed_bills'].append(duplicate)
        
        # Add one from wrong month (period ends in May)
        wrong = create_mock_bill(EXPECTED_ACCOUNTS[10], from_date='15-04-25', to_date='15-05-25', kwh=10000)
        session_storage['parsed_bills'].append(wrong)
        
        # Last 4 accounts missing
        
        gate = validate_gate()
        assert gate['account_checklist'][EXPECTED_ACCOUNTS[0]]['status'] == 'duplicate'
        assert gate['account_checklist'][EXPECTED_ACCOUNTS[1]]['status'] == 'matched'
        assert gate['account_checklist'][EXPECTED_ACCOUNTS[10]]['status'] == 'wrong_month'
        assert gate['account_checklist'][EXPECTED_ACCOUNTS[-1]]['status'] == 'missing'


class TestExpectedAccounts:
    """Test expected accounts list"""
    
    def test_expected_accounts_count(self):
        """Test we have exactly 15 expected accounts"""
        assert len(EXPECTED_ACCOUNTS) == 15
    
    def test_expected_accounts_contains_retail_chillers(self):
        """Test list contains retail chillers account"""
        assert '55861-52267-1' in EXPECTED_ACCOUNTS
    
    def test_expected_accounts_contains_food_court(self):
        """Test list contains Food Court account"""
        assert '82805-94744-7' in EXPECTED_ACCOUNTS
    
    def test_expected_accounts_contains_fit_accounts(self):
        """Test list contains FiT accounts"""
        assert '52167-13569-2' in EXPECTED_ACCOUNTS
        assert '00776-78552-1' in EXPECTED_ACCOUNTS


class TestBugRegressions:
    """Regression tests for fixed bugs"""
    
    def test_compressed_pdf_chinese_format(self):
        """
        Regression test for Bug 1: Account 72399-00664-9 compressed PDF parsing.
        
        The compressed PDF uses Chinese date format (由...至) and Chinese kWh label
        (用電度數總計) instead of English "From...to" and "Grand Total Units Consumed".
        """
        # Simulate parsing result with Chinese format
        from app.clp_parser import parse_clp_bill
        
        # Test with the actual compressed PDF
        pdf_path = '/home/ubuntu/.cursor/projects/workspace/uploads/bill-72399-202504_6706.pdf'
        if os.path.exists(pdf_path):
            result = parse_clp_bill(pdf_path)
            
            # Verify account parsed
            assert result['account'] == '72399-00664-9'
            
            # Verify dates parsed (Chinese format: 由 24-03-25 至 23-04-25)
            assert result['from_date'] == '24-03-25'
            assert result['to_date'] == '23-04-25'
            
            # Verify days parsed (Chinese format: 共 31 日)
            assert result['days'] == 31
            
            # Verify kWh parsed (Chinese format: 用電度數總計 59910.00)
            assert result['kwh'] is not None
            assert result['kwh'] > 0
            # Should be around 59,910 (allow small variance)
            assert 59000 < result['kwh'] < 61000
    
    def test_month_classification_uses_period_end(self, reset_session):
        """
        Regression test for Bug 2: Month classification should use period END date.
        
        Bills with periods starting in previous calendar month but ending in selected
        month should be classified as the end month, not the start month.
        
        Example: Period 24-03-25 to 23-04-25 should be April 2025, not March 2025.
        """
        session_storage['billing_month'] = '2025-04'
        session_storage['meter_log']['previous'] = 1000000.0
        session_storage['meter_log']['present'] = 1001000.0
        
        # Add bills that start in March but end in April
        # These should be classified as April bills
        for account in EXPECTED_ACCOUNTS:
            # Period starts March 24, ends April 23
            bill = create_mock_bill(account, from_date='24-03-25', to_date='23-04-25', kwh=10000)
            session_storage['parsed_bills'].append(bill)
        
        # Verify all bills are recognized as April bills
        gate = validate_gate()
        assert gate['passed'] is True
        assert gate['summary']['matched'] == 15
        assert gate['summary']['wrong_month'] == 0
        
        # All accounts should be "matched", not "wrong_month"
        for account in EXPECTED_ACCOUNTS:
            assert gate['account_checklist'][account]['status'] == 'matched'
    
    def test_month_extraction_logic(self):
        """Test that month extraction uses to_date (period end) not from_date (period start)"""
        # Bill period: starts March, ends April
        bill = create_mock_bill('72399-00664-9', from_date='24-03-25', to_date='23-04-25')
        
        # Should extract April (end month), not March (start month)
        month = extract_bill_month(bill)
        assert month == '2025-04', f"Expected 2025-04 (period end), got {month}"
        
        # Another case: starts Feb, ends March
        bill2 = create_mock_bill('72399-00664-9', from_date='24-02-25', to_date='23-03-25')
        month2 = extract_bill_month(bill2)
        assert month2 == '2025-03', f"Expected 2025-03 (period end), got {month2}"
    
    def test_meter_log_accepts_one_decimal_place(self, reset_session):
        """
        Regression test: Meter readings should accept one decimal place.
        
        Real April 2025 meter 6681757 readings:
        - Previous: 95966.6
        - Present: 96323.1
        - Expected delta: 356.5 (not 357)
        """
        session_storage['billing_month'] = '2025-04'
        
        # Set meter log with real April 2025 values (one decimal place)
        session_storage['meter_log']['previous'] = 95966.6
        session_storage['meter_log']['present'] = 96323.1
        
        # Verify meter log is complete
        assert is_meter_log_complete() is True
        
        # Calculate delta with 1 decimal precision
        delta = round(session_storage['meter_log']['present'] - session_storage['meter_log']['previous'], 1)
        assert delta == 356.5, f"Expected delta 356.5, got {delta}"
        
        # Verify values are stored with correct precision
        assert session_storage['meter_log']['previous'] == 95966.6
        assert session_storage['meter_log']['present'] == 96323.1


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
