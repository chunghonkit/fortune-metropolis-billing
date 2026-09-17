"""
Unit tests for allocation engine using July 2026 fixture data.

Tests verify:
- Allocation math correctness
- Residual within ±0.01 HKD tolerance
- 100% single-destination accounts
- Multi-centre splits
- FiT special cases
"""

import unittest
from decimal import Decimal
import json
from pathlib import Path

from app.allocation_engine import AllocationEngine, allocate_costs


class TestAllocationEngine(unittest.TestCase):
    """Test allocation engine with known July 2026 fixture data"""
    
    def setUp(self):
        """Load fixture allocation rules"""
        # Load allocation config (July 2026 snapshot for testing)
        with open('allocation_config.json', 'r') as f:
            config = json.load(f)
        
        # Convert to allocation rules format
        self.allocation_rules = {}
        for key, data in config.items():
            account = data.get('account', key)
            if '-' in account:
                # Use full account number
                self.allocation_rules[account] = [
                    {'centre': a['centre'], 'percentage': a['perc']}
                    for a in data['allocations']
                ]
            else:
                # Use key as-is
                self.allocation_rules[key] = [
                    {'centre': a['centre'], 'percentage': a['perc']}
                    for a in data['allocations']
                ]
    
    def test_single_destination_account(self):
        """Test 100% allocation to single cost centre"""
        # Account 82805-94744-7 (FC 100%)
        bill = {
            'account': '82805-94744-7',
            'total_amount': 88494.0,
            'kwh': 60236
        }
        
        engine = AllocationEngine(self.allocation_rules)
        results = engine.allocate_bills([bill])
        
        # Should have exactly one allocation
        self.assertEqual(len(results['allocations']), 1)
        
        alloc = results['allocations'][0]
        self.assertEqual(alloc['centre'], 'FC')
        self.assertEqual(alloc['percentage'], 1.0)
        self.assertEqual(alloc['amount'], 88494.0)
        
        # Check summary
        self.assertIn('FC', results['summary'])
        self.assertEqual(results['summary']['FC']['total_amount'], 88494.0)
    
    def test_multi_centre_split(self):
        """Test allocation across multiple cost centres"""
        # Account 55861-52267-1 (AC 72.54%, SW 27.46%)
        bill = {
            'account': '55861-52267-1',
            'total_amount': 724078.0,
            'kwh': 489322
        }
        
        engine = AllocationEngine(self.allocation_rules)
        results = engine.allocate_bills([bill])
        
        # Should have 2 allocations
        self.assertEqual(len(results['allocations']), 2)
        
        # Calculate expected amounts
        ac_expected = Decimal('724078.0') * Decimal('0.7253839391172826')
        sw_expected = Decimal('724078.0') * Decimal('0.2746160608827174')
        
        ac_alloc = next(a for a in results['allocations'] if a['centre'] == 'AC')
        sw_alloc = next(a for a in results['allocations'] if a['centre'] == 'SW')
        
        # Check allocations are close (within rounding)
        self.assertAlmostEqual(ac_alloc['amount'], float(ac_expected), places=2)
        self.assertAlmostEqual(sw_alloc['amount'], float(sw_expected), places=2)
        
        # Check residual
        total_allocated = ac_alloc['amount'] + sw_alloc['amount']
        residual = abs(724078.0 - total_allocated)
        self.assertLessEqual(residual, 0.01, f"Residual {residual} exceeds ±0.01 HKD")
    
    def test_residual_tolerance(self):
        """Test residual is within ±0.01 HKD for all July 2026 accounts"""
        # July 2026 fixture data
        july_bills = [
            {'account': '55861-52267-1', 'total_amount': 724078.0, 'kwh': 489322},
            {'account': '24096-78457-6', 'total_amount': 75285.0, 'kwh': 52983},
            {'account': '13639-58422-3', 'total_amount': 194415.0, 'kwh': 126390},
            {'account': '40722-61440-7', 'total_amount': 53708.0, 'kwh': 38036},
            {'account': '35204-69738-4', 'total_amount': 34635.0, 'kwh': 22538},
            {'account': '97968-02236-6', 'total_amount': 1511.0, 'kwh': 3250},
            {'account': '23529-59279-9', 'total_amount': 169084.0, 'kwh': 122733},
            {'account': '79292-23337-6', 'total_amount': 38530.0, 'kwh': 25132},
            {'account': '52167-13569-2', 'total_amount': 182327.0, 'kwh': 176400, 'fit_amount': -72005.0},
            {'account': '08731-83914-5', 'total_amount': 14572.0, 'kwh': 9497},
            {'account': '00776-78552-1', 'total_amount': 53197.0, 'kwh': 35499, 'fit_amount': -3858.0},
            {'account': '88931-57029-6', 'total_amount': 3571.0, 'kwh': 2323},
            {'account': '70873-85471-3', 'total_amount': 129419.0, 'kwh': 91306},
            {'account': '72399-00664-9', 'total_amount': 107150.0, 'kwh': 76563},
            {'account': '82805-94744-7', 'total_amount': 88494.0, 'kwh': 60236},
        ]
        
        engine = AllocationEngine(self.allocation_rules)
        results = engine.allocate_bills(july_bills)
        
        # Check each bill's residual
        for bill in july_bills:
            account = bill['account']
            bill_allocations = [a for a in results['allocations'] if a['account'] == account]
            
            if not bill_allocations:
                continue  # Account may not have rules
            
            total_allocated = sum(a['amount'] for a in bill_allocations)
            residual = abs(bill['total_amount'] - total_allocated)
            
            self.assertLessEqual(
                residual, 0.01,
                f"Account {account}: residual {residual} exceeds ±0.01 HKD"
            )
    
    def test_july_2026_total_allocation(self):
        """Test total allocation matches July 2026 fixture totals (with FiT added to base)"""
        july_bills = [
            {'account': '55861-52267-1', 'total_amount': 724078.0, 'kwh': 489322},
            {'account': '24096-78457-6', 'total_amount': 75285.0, 'kwh': 52983},
            {'account': '13639-58422-3', 'total_amount': 194415.0, 'kwh': 126390},
            {'account': '40722-61440-7', 'total_amount': 53708.0, 'kwh': 38036},
            {'account': '35204-69738-4', 'total_amount': 34635.0, 'kwh': 22538},
            {'account': '97968-02236-6', 'total_amount': 1511.0, 'kwh': 3250},
            {'account': '23529-59279-9', 'total_amount': 169084.0, 'kwh': 122733},
            {'account': '79292-23337-6', 'total_amount': 38530.0, 'kwh': 25132},
            # FiT accounts: allocation base = total + abs(fit_amount)
            {'account': '52167-13569-2', 'total_amount': 182327.0, 'kwh': 176400, 'fit_amount': -72005.0},
            {'account': '08731-83914-5', 'total_amount': 14572.0, 'kwh': 9497},
            {'account': '00776-78552-1', 'total_amount': 53197.0, 'kwh': 35499, 'fit_amount': -3858.0},
            {'account': '88931-57029-6', 'total_amount': 3571.0, 'kwh': 2323},
            {'account': '70873-85471-3', 'total_amount': 129419.0, 'kwh': 91306},
            {'account': '72399-00664-9', 'total_amount': 107150.0, 'kwh': 76563},
            {'account': '82805-94744-7', 'total_amount': 88494.0, 'kwh': 60236},
        ]
        
        # Expected total with FiT added to base:
        # Regular bills sum + FiT amounts (absolute value)
        total_expected = sum(b['total_amount'] for b in july_bills)
        fit_total = sum(abs(b.get('fit_amount', 0)) for b in july_bills)
        allocation_base_expected = total_expected + fit_total
        
        engine = AllocationEngine(self.allocation_rules)
        results = engine.allocate_bills(july_bills)
        
        # Sum all allocations (should match allocation base, not just bill totals)
        total_allocated = sum(a['amount'] for a in results['allocations'])
        
        # Should match allocation base within tolerance
        residual = abs(allocation_base_expected - total_allocated)
        self.assertLessEqual(residual, 0.15, 
                           f"Total residual {residual} too high. "
                           f"Expected allocation base: {allocation_base_expected}, "
                           f"Allocated: {total_allocated}")
        
        # Note: This is expected behavior with FiT - we allocate MORE than the bill totals
        # because FiT is added to the allocation base
        logger.info(f"Bill totals: {total_expected}, FiT: {fit_total}, "
                   f"Allocation base: {allocation_base_expected}, "
                   f"Total allocated: {total_allocated}")
    
    def test_percentage_normalization(self):
        """Test that percentages close to 100% are handled correctly"""
        # Create a rule with percentages that sum to 0.9999 (close to 1.0)
        rules = {
            'TEST-ACCOUNT': [
                {'centre': 'AC', 'percentage': 0.5},
                {'centre': 'DC', 'percentage': 0.4999}
            ]
        }
        
        bill = {
            'account': 'TEST-ACCOUNT',
            'total_amount': 10000.0,
            'kwh': 5000
        }
        
        engine = AllocationEngine(rules)
        results = engine.allocate_bills([bill])
        
        # Should allocate without errors
        total_allocated = sum(a['amount'] for a in results['allocations'])
        residual = abs(10000.0 - total_allocated)
        self.assertLessEqual(residual, 0.01)
    
    def test_fit_added_to_allocation_base(self):
        """
        Test LOCKED rule: FiT is added to allocation base BEFORE applying %.
        
        Example from Kit: 37922 + 4359 = 42281; C = 10% of 42281
        Do NOT post FiT solely to C or DC - apply % to ALL centres.
        """
        # Create rules with C at 10% and other centres
        rules = {
            '00776-78552-1': [
                {'centre': 'C', 'percentage': 0.10},
                {'centre': 'AC', 'percentage': 0.385},
                {'centre': 'DC', 'percentage': 0.23159999999999997},
                {'centre': 'OC', 'percentage': 0.020000000000000004},
                {'centre': 'CP', 'percentage': 0.035},
            ]
        }
        
        # Bill with FiT (negative amount in bill)
        bill = {
            'account': '00776-78552-1',
            'total_amount': 37922.0,  # Simplified from real 53197
            'fit_amount': -4359.0,     # FiT is negative
            'kwh': 35499
        }
        
        engine = AllocationEngine(rules)
        results = engine.allocate_bills([bill])
        
        # Check allocation base = total + abs(fit)
        allocation_base = 37922.0 + 4359.0  # = 42281
        
        # Find C allocation
        c_alloc = next((a for a in results['allocations'] if a['centre'] == 'C'), None)
        self.assertIsNotNone(c_alloc, "C allocation not found")
        
        # C should get 10% of 42281 = 4228.1
        expected_c = allocation_base * 0.10
        self.assertAlmostEqual(c_alloc['amount'], expected_c, places=2,
                             msg=f"C allocation {c_alloc['amount']} != 10% of {allocation_base}")
        
        # All allocations should sum to allocation_base
        total_allocated = sum(a['amount'] for a in results['allocations'])
        residual = abs(allocation_base - total_allocated)
        self.assertLessEqual(residual, 0.01, 
                           f"Residual {residual} exceeds tolerance. "
                           f"Allocation base: {allocation_base}, Allocated: {total_allocated}")
    
    def test_validation_errors(self):
        """Test validation errors are captured"""
        # Bill with no allocation rules
        bill = {
            'account': 'UNKNOWN-ACCOUNT',
            'total_amount': 1000.0,
            'kwh': 500
        }
        
        engine = AllocationEngine(self.allocation_rules)
        results = engine.allocate_bills([bill])
        
        # Should have validation error
        self.assertGreater(len(results['validation']['errors']), 0)
        self.assertIn('UNKNOWN-ACCOUNT', results['validation']['errors'][0])


class TestMasterWorkbookIntegration(unittest.TestCase):
    """Test integration with master workbook parser"""
    
    def test_sample_master_loads(self):
        """Test that sample master workbook loads correctly"""
        from app.master_parser import compute_weights_from_master
        
        master_path = 'data/masters/cost_allocation_master.xlsx'
        if not Path(master_path).exists():
            self.skipTest("Sample master not found")
        
        weights = compute_weights_from_master(master_path)
        
        # Should have some accounts
        self.assertGreater(len(weights), 0, "No weights loaded from master")
        
        # Check structure
        for account, allocations in weights.items():
            self.assertIsInstance(allocations, list)
            self.assertGreater(len(allocations), 0)
            
            for alloc in allocations:
                self.assertIn('centre', alloc)
                self.assertIn('percentage', alloc)
                self.assertIsInstance(alloc['percentage'], (int, float))
                self.assertGreaterEqual(alloc['percentage'], 0)
                self.assertLessEqual(alloc['percentage'], 1)


if __name__ == '__main__':
    unittest.main()
