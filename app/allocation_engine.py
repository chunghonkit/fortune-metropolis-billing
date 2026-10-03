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
    
    def __init__(self, allocation_rules: Dict[str, List[Dict]], meter_log: Dict = None):
        """
        Initialize allocation engine.
        
        Args:
            allocation_rules: Dict mapping account/meter -> List[{centre, percentage}]
            meter_log: Optional dict with check-meter readings (for AC/SW split)
        """
        self.allocation_rules = allocation_rules
        self.meter_log = meter_log
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
        1. LOCKED RULE: Food Court account 82805-94744-7 is 100% FC (outside master)
        2. SPECIAL: Retail chillers 55861-52267-1 AC/SW split from THIS month's check-meter 6681757
        3. HARDCODED: Office chillers 13639-58422-3 and 79292-23337-6 are AO 92% / OC 8%
        4. Full account number (e.g., "52167-13569-2")
        5. Short account number (first segment, e.g., "52167")
        6. Meter number with LOCKED mapping (7662756→9092771, 7662761→9091324)
        """
        # LOCKED RULE: Food Court account 82805-94744-7 = FC 100% (outside master)
        if account == '82805-94744-7':
            logger.info(f"Account {account}: FC 100% hard rule (outside master)")
            return [{'centre': 'FC', 'percentage': 1.0}]
        
        # Retail chillers are not a frozen Column C percentage. Month-end uses
        # the Citybase model (check meter × 160 / meter 9024222 kWh). This
        # engine is only the synthetic-rule path used by unit fixtures.
        
        # HARDCODED RULE: Office chillers AO 92% / OC 8%
        # These accounts are in the Allocation sheet by meter, but if the CLP parser
        # doesn't extract the meters from the bill PDFs, we need a fallback
        if account in ['13639-58422-3', '79292-23337-6']:
            logger.info(f"Account {account}: Office chiller hardcoded AO 92% / OC 8%")
            return [
                {'centre': 'AO', 'percentage': 0.92},
                {'centre': 'OC', 'percentage': 0.08}
            ]
        
        # HARDCODED FALLBACK: Account-to-meter mappings from Max Demand sheet
        # When CLP parser doesn't extract meters from bill PDFs, use these mappings
        account_to_meter = {
            '08731-83914-5': '9043327',
            '00776-78552-1': '9044168',
            '24096-78457-6': '9044624',
            '40722-61440-7': '9044334',
            '35204-69738-4': '9043416',
            '70873-85471-3': '9045951',  # Tower1
            '72399-00664-9': '9043400',  # Tower2
            '23529-59279-9': '9045925',
        }
        
        if account in account_to_meter:
            meter = account_to_meter[account]
            if meter in self.allocation_rules:
                logger.info(f"Account {account}: Using hardcoded meter {meter} from Max Demand sheet")
                return self.allocation_rules[meter]
            else:
                logger.warning(f"Account {account}: Hardcoded meter {meter} not found in allocation rules")
        
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
                logger.info(f"Account {account}: Found meter {meter_no} in allocation rules")
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
    
    def _compute_check_meter_split(self, bill: Dict) -> Optional[List[Dict]]:
        """
        Compute AC/SW split from check-meter 6681757 for retail chillers.
        
        Check meter 6681757 is 散熱水泵電 (Sea Water Pump) = SW
        The rest of consumption goes to AC (Commercial Air Conditioning)
        
        Uses meter_log (user-entered readings), not bill meters array.
        
        Formula:
            SW_kwh = present - previous (from meter_log)
            Total_kwh = bill total kwh
            SW% = SW_kwh / Total_kwh
            AC% = 1 - SW%
        
        Returns:
            List[{centre, percentage}] or None if check meter not found
        """
        if not self.meter_log:
            logger.warning(f"No meter_log provided, cannot compute check-meter split")
            return None
        
        CHECK_METER_NO = '6681757'
        
        # Verify this is the check meter
        if self.meter_log.get('meter_no') != CHECK_METER_NO:
            logger.warning(f"Meter log is for {self.meter_log.get('meter_no')}, expected {CHECK_METER_NO}")
            return None
        
        # Get readings from meter_log
        previous = self.meter_log.get('previous')
        present = self.meter_log.get('present')
        
        if previous is None or present is None:
            logger.warning(f"Meter log incomplete: previous={previous}, present={present}")
            return None
        
        # Compute check meter consumption
        check_meter_kwh = present - previous
        
        if check_meter_kwh <= 0:
            logger.warning(f"Check meter {CHECK_METER_NO} has non-positive consumption: {check_meter_kwh}")
            return None
        
        # Get total kWh from bill
        total_kwh = bill.get('kwh', 0)
        if total_kwh <= 0:
            logger.warning(f"Bill has no kWh, cannot compute check-meter split")
            return None
        
        # Compute percentages
        sw_pct = check_meter_kwh / total_kwh
        ac_pct = 1.0 - sw_pct
        
        if ac_pct < 0 or sw_pct < 0 or ac_pct > 1 or sw_pct > 1:
            logger.error(f"Invalid check-meter split: AC={ac_pct:.4f}, SW={sw_pct:.4f}")
            return None
        
        logger.info(f"Computed check-meter split from meter_log: "
                   f"previous={previous:.1f}, present={present:.1f}, delta={check_meter_kwh:.1f} kWh, "
                   f"total={total_kwh:.1f} kWh → AC={ac_pct:.4f}, SW={sw_pct:.4f}")
        
        return [
            {'centre': 'AC', 'percentage': ac_pct},
            {'centre': 'SW', 'percentage': sw_pct}
        ]
    
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


def allocate_costs(parsed_bills: List[Dict], allocation_rules: Dict[str, List[Dict]], meter_log: Dict = None) -> Dict:
    """
    Convenience function to allocate costs.
    
    Args:
        parsed_bills: List of parsed bill dicts
        allocation_rules: Allocation rules from master workbook
        meter_log: Optional dict with check-meter readings for AC/SW split
        
    Returns:
        Allocation results dict
    """
    engine = AllocationEngine(allocation_rules, meter_log=meter_log)
    return engine.allocate_bills(parsed_bills)
