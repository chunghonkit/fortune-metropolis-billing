#!/usr/bin/env python3
"""
Test script to verify core functionality without pytest
"""

import sys
import json
from pathlib import Path

# Add app to path
sys.path.insert(0, str(Path(__file__).parent))

from app.master_parser import compute_weights_from_master
from app.allocation_engine import allocate_costs

def test_master_parser():
    """Test master parser loads default master"""
    print("\n=== Testing Master Parser ===")
    master_path = 'data/masters/cost_allocation_master.xlsx'
    
    if not Path(master_path).exists():
        print(f"❌ Master not found: {master_path}")
        return False
    
    try:
        weights = compute_weights_from_master(master_path)
        print(f"✓ Loaded {len(weights)} accounts from master")
        
        # Show first few accounts
        for i, (account, rules) in enumerate(list(weights.items())[:3]):
            print(f"  {account}: {len(rules)} centres")
            for rule in rules:
                print(f"    - {rule['centre']}: {rule['percentage']:.4f}")
        
        return True
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_allocation_engine():
    """Test allocation engine with sample bill"""
    print("\n=== Testing Allocation Engine ===")
    
    # Load allocation config for test
    with open('allocation_config.json', 'r') as f:
        config = json.load(f)
    
    # Convert to allocation rules format
    allocation_rules = {}
    for key, data in config.items():
        account = data.get('account', key)
        allocation_rules[account] = [
            {'centre': a['centre'], 'percentage': a['perc']}
            for a in data['allocations']
        ]
    
    # Test bill (FC 100%)
    test_bills = [
        {
            'account': '82805-94744-7',
            'total_amount': 88494.0,
            'kwh': 60236
        }
    ]
    
    try:
        results = allocate_costs(test_bills, allocation_rules)
        
        print(f"✓ Allocated {len(results['allocations'])} entries")
        print(f"  Validation: {results['validation']}")
        
        # Show allocations
        for alloc in results['allocations']:
            print(f"  {alloc['account']} → {alloc['centre']}: ${alloc['amount']:.2f} ({alloc['percentage']*100:.1f}%)")
        
        return True
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_fit_handling():
    """Test FiT added to allocation base"""
    print("\n=== Testing FiT Handling (LOCKED Rule) ===")
    
    # Load rules
    with open('allocation_config.json', 'r') as f:
        config = json.load(f)
    
    allocation_rules = {}
    for key, data in config.items():
        account = data.get('account', key)
        allocation_rules[account] = [
            {'centre': a['centre'], 'percentage': a['perc']}
            for a in data['allocations']
        ]
    
    # Test FiT bill
    test_bills = [
        {
            'account': '00776-78552-1',
            'total_amount': 53197.0,
            'fit_amount': -3858.0,  # Negative in bill
            'kwh': 35499
        }
    ]
    
    try:
        results = allocate_costs(test_bills, allocation_rules)
        
        # Calculate expected allocation base
        allocation_base = 53197.0 + 3858.0  # FiT added to base
        
        # Sum allocations
        total_allocated = sum(a['amount'] for a in results['allocations'])
        
        print(f"  Bill total: $53,197")
        print(f"  FiT amount: $3,858 (added to base)")
        print(f"  Allocation base: ${allocation_base:,.2f}")
        print(f"  Total allocated: ${total_allocated:,.2f}")
        print(f"  Residual: ${abs(allocation_base - total_allocated):.2f}")
        
        if abs(allocation_base - total_allocated) <= 0.01:
            print("✓ FiT correctly added to allocation base")
            return True
        else:
            print("❌ FiT allocation mismatch")
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("=" * 60)
    print("Metropolis Allocation App - Core Functionality Test")
    print("=" * 60)
    
    results = []
    
    results.append(("Master Parser", test_master_parser()))
    results.append(("Allocation Engine", test_allocation_engine()))
    results.append(("FiT Handling", test_fit_handling()))
    
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    
    for name, passed in results:
        status = "✓ PASS" if passed else "❌ FAIL"
        print(f"{status}: {name}")
    
    all_passed = all(r[1] for r in results)
    
    if all_passed:
        print("\n🎉 All tests passed!")
        return 0
    else:
        print("\n⚠️  Some tests failed")
        return 1

if __name__ == '__main__':
    sys.exit(main())
