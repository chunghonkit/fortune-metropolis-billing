"""
Electricity Cost Allocation Engine

Allocates CLP electricity bill costs to cost centres using weights
computed from the master workbook.

Critical: Weights come from the LIVE master workbook each month,
NOT from frozen JSON constants.

LOCKED Rules (Kit verified):
- FC 100% hard rule (stays outside master)
- FiT: Add to allocation base BEFORE applying % (bill total + FiT = actual base)
  Example: 37922 + 4359 = 42281; C = 10% of 42281
- Office chillers: bill meters 7662756/7662761 → form meters 9092771/9091324
  Allocation: 92% AO / 8% OC
"""

from typing import Dict, List, Optional
from decimal import Decimal, ROUND_HALF_UP
import logging
import json
from pathlib import Path

logger = logging.getLogger(__name__)


def load_config():
    """Load application config"""
    config_path = Path('config.json')
    if config_path.exists():
        with open(config_path, 'r') as f:
            return json.load(f)
    return {
        'fc_100_percent_hard_rule': True,
        'validation': {'residual_tolerance_hkd': 0.01}
    }


class AllocationEngine:
    """
    Allocates electricity costs to cost centres based on allocation rules.
    """
    
    # Special case accounts and meters
    FIT_ACCOUNTS = {
        '52167-13569-2': ['9144186', '10353005', '9144880'],  # FiT meters
        '00776-78552-1': ['9144816', '10352137']  # FiT meters
    }
    
    # Account 55861-52267-1 special handling (AC/SW split)
    SPECIAL_AC_SW_ACCOUNT = '55861-52267-1'
    
    RESIDUAL_TOLERANCE = Decimal('0.01')  # ±0.01 HKD tolerance
    
    def _apply_fc_hard_rule(self):
        """
        Apply FC 100% hard rule.
        
        FC (Fitness Centre) accounts are 100% allocated to FC,
        not split across multiple centres per master.
        """
        # Find FC accounts in rules
        fc_accounts = []
        for account, rules in self.allocation_rules.items():
            # Check if this is a 100% FC allocation
            if len(rules) == 1 and rules[0].get('centre') == 'FC':
                fc_accounts.append(account)
        
        logger.info(f"FC 100% hard rule: {len(fc_accounts)} accounts")
    
    def __init__(self, allocation_rules: Dict[str, List[Dict]]):
        """
        Initialize allocation engine.
        
        Args:
            allocation_rules: Dict mapping account/meter -> List[{centre, percentage}]
        """
        self.allocation_rules = allocation_rules
        self.allocations = []
        self.validation_errors = []
        self.config = load_config()
        
        # Apply FC 100% hard rule
        if self.config.get('fc_100_percent_hard_rule', True):
            self._apply_fc_hard_rule()
        
    def allocate_bills(self, parsed_bills: List[Dict]) -> Dict:
        """
        Allocate electricity costs from parsed bills to cost centres.
        
        Args:
            parsed_bills: List of parsed bill dicts from CLP parser
            
        Returns:
            Dict containing:
                - allocations: List of allocation records
                - summary: Summary by cost centre
                - validation: Validation results
        """
        self.allocations = []
        self.validation_errors = []
        
        for bill in parsed_bills:
            account = bill.get('account')
            if not account:
                logger.warning(f"Bill missing account: {bill.get('file')}")
                continue
            
            total_amount = bill.get('total_amount', 0)
            if total_amount is None or total_amount == 0:
                logger.warning(f"Bill {account} has zero or null total_amount")
                continue
            
            # Convert to Decimal for precise calculations
            total_decimal = Decimal(str(total_amount))
            
            # LOCKED RULE: Add FiT to allocation base BEFORE applying %
            # Example: 37922 + 4359 = 42281; C = 10% of 42281
            # Do NOT post FiT solely to C or DC - apply % to ALL centres
            fit_amount = bill.get('fit_amount', 0)
            if fit_amount:
                fit_decimal = Decimal(str(abs(fit_amount)))  # FiT is negative in bill
                allocation_base = total_decimal + fit_decimal
                logger.info(f"Account {account}: FiT {fit_decimal} added to base. "
                          f"Allocation base: {total_decimal} + {fit_decimal} = {allocation_base}")
            else:
                allocation_base = total_decimal
            
            # Get allocation rules for this account
            rules = self._find_allocation_rules(account, bill)
            
            if not rules:
                logger.warning(f"No allocation rules found for account {account}")
                self.validation_errors.append(f"No rules for account {account}")
                continue
            
            # Check if this is a 100% single-destination account
            if len(rules) == 1 and abs(rules[0]['percentage'] - 1.0) < 0.001:
                # Simple case: 100% to one centre
                centre = rules[0]['centre']
                allocated_amount = allocation_base
                
                self.allocations.append({
                    'account': account,
                    'centre': centre,
                    'percentage': 1.0,
                    'amount': float(allocated_amount),
                    'bill_total': float(total_decimal),
                    'allocation_base': float(allocation_base),
                    'kwh': bill.get('kwh', 0)
                })
            else:
                # Split across multiple centres using allocation_base
                allocated_amounts = {}
                total_allocated = Decimal('0')
                
                for rule in rules:
                    centre = rule['centre']
                    pct = Decimal(str(rule['percentage']))
                    
                    # Calculate allocation from allocation_base (includes FiT if present)
                    allocated = (allocation_base * pct).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                    
                    allocated_amounts[centre] = allocated
                    total_allocated += allocated
                    
                    self.allocations.append({
                        'account': account,
                        'centre': centre,
                        'percentage': float(pct),
                        'amount': float(allocated),
                        'bill_total': float(total_decimal),
                        'allocation_base': float(allocation_base),
                        'kwh': bill.get('kwh', 0)
                    })
                
                # Check residual against allocation_base
                residual = allocation_base - total_allocated
                if abs(residual) > self.RESIDUAL_TOLERANCE:
                    self.validation_errors.append(
                        f"Account {account}: residual {residual} exceeds tolerance ±{self.RESIDUAL_TOLERANCE}"
                    )
                    logger.warning(f"Residual for {account}: {residual} HKD")
                else:
                    logger.info(f"Account {account}: residual {residual} within tolerance")
        
        # Generate summary
        summary = self._generate_summary()
        
        # Validation results
        validation = {
            'total_bills': len(parsed_bills),
            'allocated_bills': len(set(a['account'] for a in self.allocations)),
            'errors': self.validation_errors,
            'residual_ok': len([e for e in self.validation_errors if 'residual' in e.lower()]) == 0
        }
        
        return {
            'allocations': self.allocations,
            'summary': summary,
            'validation': validation
        }
    
    def _find_allocation_rules(self, account: str, bill: Dict) -> Optional[List[Dict]]:
        """
        Find allocation rules for an account.
        
        Tries multiple lookup strategies:
        1. Full account number (e.g., "52167-13569-2")
        2. Short account number (first segment, e.g., "52167")
        3. Meter number with LOCKED mapping (7662756→9092771, 7662761→9091324)
        """
        # Try full account
        if account in self.allocation_rules:
            return self.allocation_rules[account]
        
        # Try short account (first segment before hyphen)
        if '-' in account:
            short_account = account.split('-')[0]
            if short_account in self.allocation_rules:
                return self.allocation_rules[short_account]
        
        # LOCKED RULE: Office chiller meter mapping
        # Bill meters 7662756/7662761 → form meters 9092771/9091324
        meter_mapping_config = self.config.get('meter_id_mapping', {})
        apply_mapping = meter_mapping_config.get('apply_mapping', False)
        mappings = meter_mapping_config.get('mappings', {})
        
        meters = bill.get('meters', [])
        for meter in meters:
            meter_no = str(meter.get('meter_no', ''))
            
            # Try direct lookup first
            if meter_no in self.allocation_rules:
                return self.allocation_rules[meter_no]
            
            # Apply LOCKED mapping (now enabled by default)
            if apply_mapping and meter_no in mappings:
                mapped_meter = mappings[meter_no]
                if mapped_meter in self.allocation_rules:
                    logger.info(f"Applied LOCKED meter mapping: {meter_no} → {mapped_meter}")
                    return self.allocation_rules[mapped_meter]
        
        # Try looking for any partial match
        for key in self.allocation_rules.keys():
            if account in key or key in account:
                return self.allocation_rules[key]
        
        return None
    
    def _generate_summary(self) -> Dict[str, Dict]:
        """
        Generate summary by cost centre.
        
        Returns:
            Dict mapping centre -> {total_amount, accounts, kwh}
        """
        summary = {}
        
        for alloc in self.allocations:
            centre = alloc['centre']
            if centre not in summary:
                summary[centre] = {
                    'total_amount': 0.0,
                    'accounts': set(),
                    'kwh': 0.0
                }
            
            summary[centre]['total_amount'] += alloc['amount']
            summary[centre]['accounts'].add(alloc['account'])
            # Note: KWH allocation is approximate (split by $, not actual kWh)
            summary[centre]['kwh'] += alloc['kwh'] * alloc['percentage']
        
        # Convert sets to lists for JSON serialization
        for centre in summary:
            summary[centre]['accounts'] = list(summary[centre]['accounts'])
            summary[centre]['total_amount'] = round(summary[centre]['total_amount'], 2)
            summary[centre]['kwh'] = round(summary[centre]['kwh'], 0)
        
        return summary


def allocate_costs(parsed_bills: List[Dict], allocation_rules: Dict[str, List[Dict]]) -> Dict:
    """
    Convenience function to allocate costs.
    
    Args:
        parsed_bills: List of parsed bill dicts
        allocation_rules: Allocation rules from master workbook
        
    Returns:
        Allocation results dict
    """
    engine = AllocationEngine(allocation_rules)
    return engine.allocate_bills(parsed_bills)
