"""
Evaluate the Citybase Cost Allocation workbook for one billing month.

The April master is a formula model, not a table of frozen percentages:

- Allocation column O is each meter's split. Meter 9046787 (private 9024222)
  is the only line that moves with the check meter:
  O7 = (check_delta × 160) / Elect Charge kWh of meter 9024222
  O6 = 1 − O7
  Meter 9046243 (private 9026183) stays 100% SW.
- AC DEPT column C turns those meter dollars into leaf cost-centre rows.
  Shared names ("Hotel / Commercial / SA", "Commercial Common", …) are
  separate leaves. They are not folded into C, DC, CP, SA, or O.
- Account 55861-52267-1 AC% = C73 / (C73+C74), where C73 sums the AC chiller
  lines and C74 sums the SW lines. That ratio moves when the check meter
  or the chiller kWh mix moves.
- FiT accounts are split on net due + |FiT|.

Dollar amounts are never taken from a golden Cost Sheet. They are computed
from this month's bills, this month's check-meter dial, and the master formulas.
"""

from __future__ import annotations

import logging
import re
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, List, Optional, Tuple

import openpyxl

logger = logging.getLogger(__name__)

KNOWN_CODES = {'AC', 'AO', 'C', 'CP', 'DC', 'FC', 'O', 'OC', 'SA', 'SW'}

# Retail chiller account and FiT retail account. Their Elect Charge amounts
# are one bill shared across several private meters by kWh.
CHILLER_ACCOUNT = '55861-52267-1'
FIT_RETAIL_ACCOUNT = '52167-13569-2'
FOOD_COURT_ACCOUNT = '82805-94744-7'

# Check meter is 散熱水泵電. It is not CLP meter 9132448 (landlord, 100% SW).
CHECK_METER_NO = '6681757'
CHECK_FACTOR = Decimal(160)


def canonical_centre_key(label: str) -> str:
    """
    Key used both for computed centres and for Cost Sheet column A.

    Short codes stay short ("AC", "C", "O"). Shared labels keep every
    segment, with spacing and a trailing "(11)" reference ignored so
    "Hotel/Commercial (10)" matches "Hotel / Commercial".
    """
    text = str(label).strip()
    text = text.replace('／', '/').replace('（', '(').replace('）', ')')
    text = re.sub(r'\s*\(\d+\)\s*$', '', text).strip()
    if '/' in text:
        parts = [re.sub(r'\s+', ' ', part).strip() for part in text.split('/')]
        parts = [part for part in parts if part]
        return ' / '.join(parts)
    if re.fullmatch(r'[A-Za-z]{1,2}', text):
        return text.upper()
    match = re.match(r'^([A-Za-z]{1,2})\s*-\s*', text)
    if match and match.group(1).upper() in KNOWN_CODES:
        return match.group(1).upper()
    return re.sub(r'\s+', ' ', text).strip()


def _dec(value) -> Decimal:
    return Decimal(str(value))


def _excel_round(value: Decimal, places: int) -> Decimal:
    quantum = Decimal('1').scaleb(-places)
    return value.quantize(quantum, rounding=ROUND_HALF_UP)


def _meter_tokens(value) -> List[str]:
    if value is None:
        return []
    return re.findall(r'\d{5,}', str(value))


class _FormulaEvaluator:
    """Evaluate the small formula dialect used on Allocation / AC DEPT."""

    def __init__(self, alloc_d: Dict[int, Decimal], alloc_o: Dict[int, Decimal],
                 n_raw: Dict[int, object], c_raw: Dict[int, object]):
        self.alloc_d = alloc_d
        self.alloc_o = alloc_o
        self.n_raw = n_raw
        self.c_raw = c_raw
        self.n_memo: Dict[int, Decimal] = {}
        self.c_memo: Dict[int, Decimal] = {}
        self.stack = set()

    def cell(self, col: str, row: int) -> Decimal:
        if col == 'C':
            return self._cell(row, self.c_raw, self.c_memo)
        if col == 'N':
            return self._cell(row, self.n_raw, self.n_memo)
        raise ValueError(f'Unsupported column {col}')

    def _cell(self, row: int, raw_map, memo) -> Decimal:
        if row in memo:
            return memo[row]
        token = ('cell', row, id(memo))
        if token in self.stack:
            raise ValueError(f'Circular reference at row {row}')
        self.stack.add(token)
        raw = raw_map.get(row)
        if raw is None or raw == '':
            value = Decimal(0)
        elif isinstance(raw, str) and raw.startswith('='):
            value = self.eval_formula(raw)
        else:
            value = _dec(raw)
        self.stack.remove(token)
        memo[row] = value
        return value

    def eval_formula(self, formula: str) -> Decimal:
        text = formula.strip()
        if text.startswith('='):
            text = text[1:]
        return self._eval_expr(text)

    def _eval_expr(self, expr: str) -> Decimal:
        parts = _split_top(expr, '+-')
        if not parts:
            return Decimal(0)
        total = self._eval_term(parts[0][1])
        # First piece may itself be a leading sign only when the splitter
        # keeps it; _split_top attaches the sign to following terms.
        if parts[0][0] == '-':
            total = -total
        for sign, part in parts[1:]:
            term = self._eval_term(part)
            total = total + term if sign == '+' else total - term
        return total

    def _eval_term(self, expr: str) -> Decimal:
        parts = _split_top(expr, '*/')
        total = self._eval_atom(parts[0][1])
        for op, part in parts[1:]:
            value = self._eval_atom(part)
            total = total * value if op == '*' else total / value
        return total

    def _eval_atom(self, expr: str) -> Decimal:
        expr = expr.strip()
        if not expr:
            return Decimal(0)
        if expr.startswith('SUM(') and expr.endswith(')'):
            inner = expr[4:-1]
            total = Decimal(0)
            for arg in _split_args(inner):
                total += self._eval_range_or_expr(arg)
            return total
        return self._eval_ref_or_number(expr)

    def _eval_range_or_expr(self, expr: str) -> Decimal:
        expr = expr.strip()
        ref = _parse_ref(expr)
        if ref and ref[3] is not None:
            sheet, col, start, end = ref
            total = Decimal(0)
            for row in range(start, end + 1):
                total += self._value_of(sheet, col, row)
            return total
        if any(op in expr for op in ('+', '-', '*', '/')) and not _parse_ref(expr):
            return self._eval_expr(expr)
        return self._eval_ref_or_number(expr)

    def _eval_ref_or_number(self, expr: str) -> Decimal:
        expr = expr.strip()
        ref = _parse_ref(expr)
        if ref:
            sheet, col, start, end = ref
            if end is not None and end != start:
                raise ValueError(f'Range {expr} used outside SUM')
            return self._value_of(sheet, col, start)
        return _dec(expr)

    def _value_of(self, sheet: Optional[str], col: str, row: int) -> Decimal:
        if sheet == 'Allocation':
            if col == 'D':
                return self.alloc_d.get(row, Decimal(0))
            if col == 'O':
                return self.alloc_o.get(row, Decimal(0))
            raise ValueError(f'Unsupported Allocation column {col}')
        if col == 'C':
            return self.cell('C', row)
        if col == 'N':
            return self.cell('N', row)
        raise ValueError(f'Unsupported reference {col}{row}')


def _split_top(expr: str, ops: str) -> List[Tuple[str, str]]:
    """Split on operators that are not inside parentheses. Keeps the operator."""
    parts: List[Tuple[str, str]] = []
    depth = 0
    buf = []
    sign = '+'
    i = 0
    while i < len(expr):
        ch = expr[i]
        if ch == '(':
            depth += 1
            buf.append(ch)
        elif ch == ')':
            depth -= 1
            buf.append(ch)
        elif depth == 0 and ch in ops:
            text = ''.join(buf).strip()
            if text == '' and ch in '+-':
                # Leading sign, or a sign after another operator.
                sign = ch
            else:
                parts.append((sign, text))
                sign = ch
                buf = []
        else:
            buf.append(ch)
        i += 1
    tail = ''.join(buf).strip()
    if tail:
        parts.append((sign, tail))
    return parts


def _split_args(inner: str) -> List[str]:
    args = []
    depth = 0
    buf = []
    for ch in inner:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        if ch == ',' and depth == 0:
            args.append(''.join(buf))
            buf = []
        else:
            buf.append(ch)
    if buf:
        args.append(''.join(buf))
    return args


_REF_RE = re.compile(
    r"^(?:(?P<sheet>[A-Za-z][A-Za-z ]*)!)?"
    r"\$?(?P<col>[A-Z]{1,3})\$?(?P<start>\d+)"
    r"(?::\$?(?P<col2>[A-Z]{1,3})\$?(?P<end>\d+))?$"
)


def _parse_ref(expr: str):
    match = _REF_RE.match(expr.strip())
    if not match:
        return None
    end = int(match.group('end')) if match.group('end') else None
    return match.group('sheet'), match.group('col'), int(match.group('start')), end


class CitybaseModel:
    def __init__(self, path: str):
        self.path = path
        self.alloc_lines: Dict[int, Dict] = {}
        self.c_raw: Dict[int, object] = {}
        self.n_raw: Dict[int, object] = {}
        self.leaf_rows: List[int] = []
        self.leaf_labels: Dict[int, str] = {}
        self.elect_row_account: Dict[int, str] = {}
        self.chiller_groups: List[Tuple[int, set]] = []
        self.fit_groups: List[Tuple[int, set]] = []
        self.check_kwh_meters: set = set()
        self._load()

    def _load(self):
        wb = openpyxl.load_workbook(self.path, data_only=False)
        try:
            allocation = wb['Allocation']
            ac_dept = wb['AC DEPT']
            elect = wb['Elect Charge']
            self._load_allocation(allocation)
            self._load_ac_dept(ac_dept)
            self._load_elect_structure(elect)
            self._load_max_demand(wb['Max Demand'] if 'Max Demand' in wb.sheetnames else None)
        finally:
            wb.close()
        logger.info(
            'Citybase model: %s allocation lines, %s leaf rows, check-meter kWh meters %s',
            len(self.alloc_lines), len(self.leaf_rows), sorted(self.check_kwh_meters)
        )

    def _load_allocation(self, sheet):
        current_meter = None
        for row in range(4, 68):
            meter_val = sheet.cell(row=row, column=2).value
            if meter_val not in (None, ''):
                current_meter = str(meter_val).strip().rstrip('*').strip()
            pct = sheet.cell(row=row, column=15).value
            centre = sheet.cell(row=row, column=17).value
            charge_formula = sheet.cell(row=row, column=4).value
            if pct is None and charge_formula is None:
                continue
            elect_row = None
            if isinstance(charge_formula, str):
                found = re.search(r"Elect Charge'!\$?([A-Z]+)\$?(\d+)", charge_formula)
                if found:
                    elect_row = int(found.group(2))
            kind = 'fixed'
            complement_of = None
            if isinstance(pct, str):
                compact = pct.replace(' ', '')
                if '160' in compact:
                    kind = 'check_sw'
                elif re.match(r'^=1-O\d+$', compact):
                    kind = 'check_ac'
                    complement_of = int(re.search(r'O(\d+)', compact).group(1))
            self.alloc_lines[row] = {
                'meter': current_meter,
                'pct': pct,
                'kind': kind,
                'complement_of': complement_of,
                'centre': centre,
                'elect_row': elect_row,
            }

    def _load_ac_dept(self, sheet):
        for row in range(4, 70):
            self.c_raw[row] = sheet.cell(row=row, column=3).value
            self.n_raw[row] = sheet.cell(row=row, column=14).value
            label = sheet.cell(row=row, column=15).value
            if label:
                self.leaf_labels[row] = canonical_centre_key(str(label))
        total_formula = sheet.cell(row=70, column=3).value
        if isinstance(total_formula, str) and total_formula.startswith('='):
            self.leaf_rows = _expand_sum_rows(total_formula)
        else:
            self.leaf_rows = [row for row in range(4, 70) if row in self.leaf_labels]
        logger.info('AC DEPT leaf rows: %s', self.leaf_rows)

    def _load_elect_structure(self, elect):
        """
        Read the month column that Allocation points at and recover which
        private-meter kWh rows feed the chiller bill and the FiT retail bill.
        """
        col_letter = None
        # Allocation D formulas name the Elect Charge month column, e.g. BC7.
        allocation = elect.parent['Allocation']
        for row in range(4, 68):
            charge_formula = allocation.cell(row=row, column=4).value
            if isinstance(charge_formula, str):
                found = re.search(r"Elect Charge'!\$?([A-Z]+)\$?\d+", charge_formula)
                if found:
                    col_letter = found.group(1)
                    break
        if not col_letter:
            logger.warning('Could not find chiller split column on Elect Charge; using template rows')
            self.chiller_groups = [
                (5, {'9048406', '9046064'}),
                (6, {'9026169', '9046505'}),
                (7, {'9024222', '9046787'}),
                (10, {'9026183', '9046243'}),
            ]
            self.fit_groups = [
                (8, {'9133719', '10220967', '9055260'}),
                (9, {'9134412', '9115936', '9054743'}),
                (13, {'9144186', '9144880', '10353005', '9055919'}),
            ]
        else:
            col = openpyxl.utils.column_index_from_string(col_letter)
            self.chiller_groups = _groups_from_column(elect, col, bill_row=34, total_row=32)
            self.fit_groups = _groups_from_column(elect, col, bill_row=41, total_row=39)
        # Attach legacy meter ids that Allocation sends to the same charge row.
        for groups in (self.chiller_groups, self.fit_groups):
            for idx, (elect_row, meters) in enumerate(groups):
                for line in self.alloc_lines.values():
                    if line['elect_row'] == elect_row and line['meter']:
                        meters.update(_meter_tokens(line['meter']))
                groups[idx] = (elect_row, meters)
        self.check_kwh_meters = set()
        for elect_row, meters in self.chiller_groups:
            if elect_row == 7 or '9024222' in meters or '9046787' in meters:
                self.check_kwh_meters.update(meters)
        if not self.check_kwh_meters:
            self.check_kwh_meters = {'9024222', '9046787'}
        for elect_row, meters in self.chiller_groups:
            self.elect_row_account[elect_row] = CHILLER_ACCOUNT
        for elect_row, meters in self.fit_groups:
            self.elect_row_account[elect_row] = FIT_RETAIL_ACCOUNT
        logger.info('Chiller kWh groups: %s', self.chiller_groups)
        logger.info('FiT retail kWh groups: %s', self.fit_groups)

    def _load_max_demand(self, sheet):
        if sheet is None:
            return
        meter_to_account = {}
        for row in range(1, 40):
            account = sheet.cell(row=row, column=1).value
            meter = sheet.cell(row=row, column=2).value
            if account and meter and re.match(r'\d{5}-\d{5}-\d', str(account)):
                meter_to_account[str(meter).strip()] = str(account).strip()
        for line in self.alloc_lines.values():
            elect_row = line['elect_row']
            meter = line['meter']
            if elect_row and meter and meter in meter_to_account:
                account = meter_to_account[meter]
                # Do not let a chiller/FiT legacy id overwrite the split accounts.
                if elect_row not in self.elect_row_account:
                    self.elect_row_account[elect_row] = account
        logger.info('Elect row → account: %s', dict(sorted(self.elect_row_account.items())))

    def allocate(self, bills: List[Dict], meter_log: Optional[Dict] = None) -> Dict:
        by_account = {}
        for bill in bills:
            account = bill.get('account')
            if not account:
                continue
            base = _bill_base(bill)
            if base == 0:
                continue
            by_account[account] = by_account.get(account, Decimal(0)) + base

        elect, warnings = self._elect_charges(bills, by_account)
        delta, check_kwh = self._check_inputs(bills, meter_log)
        if delta is None or check_kwh is None:
            warnings.append(
                'Check meter split needs dial delta and kWh of private meter 9024222 '
                '(legacy 9046787). Those inputs were missing, so O7 was left at 0.'
            )
        # Keep a real dial even when the private-meter kWh is missing.
        # Overwriting it with 0 made a present meter log look unused.
        if delta is None:
            delta = Decimal(0)
        if check_kwh is None:
            check_kwh = Decimal(1)

        alloc_o = self._percentages(delta, check_kwh)
        alloc_d = {}
        for row, line in self.alloc_lines.items():
            elect_row = line['elect_row']
            charge = elect.get(elect_row, Decimal(0))
            alloc_d[row] = charge * alloc_o.get(row, Decimal(0))

        evaluator = _FormulaEvaluator(alloc_d, alloc_o, self.n_raw, self.c_raw)
        grouped: Dict[Tuple[str, str], Decimal] = {}
        for row in self.leaf_rows:
            label = self.leaf_labels.get(row)
            if not label:
                continue
            amount = evaluator.cell('C', row)
            if amount == 0:
                continue
            account = self._account_for_leaf(row)
            key = (account or '', label)
            grouped[key] = grouped.get(key, Decimal(0)) + amount

        if FOOD_COURT_ACCOUNT in by_account:
            grouped[(FOOD_COURT_ACCOUNT, 'FC')] = (
                grouped.get((FOOD_COURT_ACCOUNT, 'FC'), Decimal(0)) + by_account[FOOD_COURT_ACCOUNT]
            )

        allocations = []
        for (account, centre), amount in sorted(grouped.items()):
            base = by_account.get(account, Decimal(0))
            pct = (amount / base) if base else Decimal(0)
            allocations.append({
                'account': account,
                'centre': centre,
                'percentage': float(pct),
                'amount': float(amount),
                'bill_total': float(base),
                'allocation_base': float(base),
                'kwh': 0,
            })

        self._log_chiller_split(delta, check_kwh, alloc_o, grouped)
        summary = {}
        for alloc in allocations:
            bucket = summary.setdefault(alloc['centre'], {'total_amount': 0.0, 'accounts': []})
            bucket['total_amount'] = round(bucket['total_amount'] + alloc['amount'], 2)
            if alloc['account'] not in bucket['accounts']:
                bucket['accounts'].append(alloc['account'])

        return {
            'allocations': allocations,
            'summary': summary,
            'validation': {'errors': warnings, 'residual_ok': not warnings},
        }

    def _percentages(self, delta: Decimal, check_kwh: Decimal) -> Dict[int, Decimal]:
        sw_pct = (delta * CHECK_FACTOR) / check_kwh if check_kwh else Decimal(0)
        if sw_pct < 0:
            sw_pct = Decimal(0)
        if sw_pct > 1:
            logger.warning('Check-meter SW share %.4f is outside 0–1 (delta=%s kWh=%s)', sw_pct, delta, check_kwh)
        percentages = {}
        for row, line in self.alloc_lines.items():
            if line['kind'] == 'check_sw':
                percentages[row] = sw_pct
            elif line['kind'] == 'check_ac':
                percentages[row] = Decimal(1) - sw_pct
            elif isinstance(line['pct'], str):
                logger.warning('Allocation row %s has unrecognised percentage %s', row, line['pct'])
                percentages[row] = Decimal(0)
            elif line['pct'] is None:
                percentages[row] = Decimal(0)
            else:
                percentages[row] = _dec(line['pct'])
        return percentages

    def _elect_charges(self, bills, by_account):
        warnings = []
        elect: Dict[int, Decimal] = {}

        def fill_groups(account, groups, label):
            if account not in by_account:
                return
            base = by_account[account]
            # Only this account's meters. 9132488 on the Tower bill is not a
            # chiller meter, and another bill must not supply 9024222's units.
            kwh_by_meter = _kwh_index(bills, account)
            weights = []
            for elect_row, meters in groups:
                weight = _group_kwh(meters, kwh_by_meter)
                weights.append((elect_row, weight))
            total_kwh = sum((weight for _, weight in weights), Decimal(0))
            if total_kwh <= 0:
                warnings.append(
                    f'{account}: no private-meter kWh for the {label} split '
                    f'(meters seen: {sorted(kwh_by_meter)})'
                )
                return
            logger.info(
                '%s %s kWh by elect row: %s',
                account, label,
                [(elect_row, str(weight)) for elect_row, weight in weights],
            )
            # Match the master: every line except the first is rounded to 3 d.p.,
            # and the first line is the residual so the parts sum to the bill.
            residual_row = weights[0][0]
            used = Decimal(0)
            for elect_row, weight in weights[1:]:
                amount = _excel_round(base * weight / total_kwh, 3)
                elect[elect_row] = amount
                used += amount
            elect[residual_row] = base - used

        fill_groups(CHILLER_ACCOUNT, self.chiller_groups, 'chiller')
        fill_groups(FIT_RETAIL_ACCOUNT, self.fit_groups, 'FiT retail')

        for elect_row, account in self.elect_row_account.items():
            if account in (CHILLER_ACCOUNT, FIT_RETAIL_ACCOUNT, FOOD_COURT_ACCOUNT):
                continue
            if account in by_account:
                elect[elect_row] = by_account[account]
        return elect, warnings

    def _check_inputs(self, bills, meter_log):
        delta = None
        if meter_log and meter_log.get('previous') is not None and meter_log.get('present') is not None:
            if not meter_log.get('meter_no') or str(meter_log.get('meter_no')) == CHECK_METER_NO:
                delta = _dec(meter_log['present']) - _dec(meter_log['previous'])
        # O7's denominator is 9024222 on the chiller bill only.
        kwh_by_meter = _kwh_index(bills, CHILLER_ACCOUNT)
        check_kwh = _group_kwh(self.check_kwh_meters, kwh_by_meter)
        if check_kwh <= 0:
            check_kwh = None
        return delta, check_kwh

    def _account_for_leaf(self, row: int) -> Optional[str]:
        accounts = set()
        self._collect_accounts(row, accounts, set())
        if len(accounts) == 1:
            return next(iter(accounts))
        if not accounts:
            return None
        logger.warning('AC DEPT row %s mixes accounts %s', row, accounts)
        return sorted(accounts)[0]

    def _collect_accounts(self, row: int, accounts: set, seen: set):
        if row in seen:
            return
        seen.add(row)
        formula = self.c_raw.get(row)
        if not isinstance(formula, str):
            return
        for alloc_row in _allocation_rows_in(formula):
            line = self.alloc_lines.get(alloc_row)
            if not line or not line.get('elect_row'):
                continue
            account = self.elect_row_account.get(line['elect_row'])
            if account:
                accounts.add(account)
        for child in _column_rows_in(formula, 'C'):
            self._collect_accounts(child, accounts, seen)

    def _log_chiller_split(self, delta, check_kwh, alloc_o, grouped):
        sw_row = next((row for row, line in self.alloc_lines.items() if line['kind'] == 'check_sw'), None)
        o7 = alloc_o.get(sw_row, Decimal(0)) if sw_row else Decimal(0)
        ac = sum((amt for (account, centre), amt in grouped.items()
                  if account == CHILLER_ACCOUNT and centre == 'AC'), Decimal(0))
        sw = sum((amt for (account, centre), amt in grouped.items()
                  if account == CHILLER_ACCOUNT and centre == 'SW'), Decimal(0))
        total = ac + sw
        if total:
            logger.info(
                'Check meter %s delta=%s × %s / kWh %s → O7=%.6f. '
                'Account %s AC=%.4f%% SW=%.4f%%',
                CHECK_METER_NO, delta, CHECK_FACTOR, check_kwh, float(o7),
                CHILLER_ACCOUNT, float(ac / total * 100), float(sw / total * 100)
            )


def _bill_base(bill: Dict) -> Decimal:
    total = _dec(bill.get('total_amount') or 0)
    fit = bill.get('fit_amount') or 0
    if fit:
        total += abs(_dec(fit))
    return total


# Allocation still names the pre-2011 meter. The bill names the meter that
# replaced it. A kWh row must not add the legacy id on top of the live ones.
_LEGACY_METERS = {
    '9046064', '9046505', '9046787', '9046243',
    '9055260', '9054743', '9055919',
}


def _kwh_index(bills: List[Dict], account: Optional[str] = None) -> Dict[str, Decimal]:
    index: Dict[str, Decimal] = {}
    for bill in bills:
        if account is not None and bill.get('account') != account:
            continue
        for meter in bill.get('meters') or []:
            number = str(meter.get('meter_no') or '').strip()
            if not number:
                continue
            consumption = meter.get('consumption')
            if consumption is None:
                consumption = meter.get('kwh') or 0
            index[number] = index.get(number, Decimal(0)) + _dec(consumption)
    return index


def _group_kwh(meters, kwh_by_meter: Dict[str, Decimal]) -> Decimal:
    """
    Sum the private meters in one Elect Charge kWh row.

    A legacy id is used only when the bill still carries that id and not
    the meter that replaced it. FiT rows list several live meters; those
    all count. Replacements do not count twice.
    """
    total = Decimal(0)
    present = {meter for meter in meters if kwh_by_meter.get(meter, Decimal(0)) > 0}
    live_present = any(meter not in _LEGACY_METERS for meter in present)
    for meter in meters:
        kwh = kwh_by_meter.get(meter, Decimal(0))
        if kwh <= 0:
            continue
        if meter in _LEGACY_METERS and live_present:
            continue
        total += kwh
    return total


def _expand_sum_rows(formula: str) -> List[int]:
    text = formula.strip()
    if text.startswith('='):
        text = text[1:]
    if text.startswith('SUM(') and text.endswith(')'):
        text = text[4:-1]
    rows = []
    for arg in _split_args(text):
        arg = arg.strip()
        range_match = re.match(r'\$?C\$?(\d+)\s*:\s*\$?C\$?(\d+)', arg)
        cell_match = re.match(r'\$?C\$?(\d+)$', arg)
        if range_match:
            rows.extend(range(int(range_match.group(1)), int(range_match.group(2)) + 1))
        elif cell_match:
            rows.append(int(cell_match.group(1)))
    return rows


def _allocation_rows_in(formula: str) -> List[int]:
    rows = []
    for match in re.finditer(r'Allocation!D(\d+)(?::D(\d+))?', formula):
        start = int(match.group(1))
        end = int(match.group(2) or match.group(1))
        rows.extend(range(start, end + 1))
    return rows


def _column_rows_in(formula: str, col: str) -> List[int]:
    rows = []
    # Avoid matching Allocation!D by requiring the column not be preceded by a letter.
    pattern = rf'(?<![A-Z]){col}(\d+)(?::{col}(\d+))?'
    for match in re.finditer(pattern, formula):
        start = int(match.group(1))
        end = int(match.group(2) or match.group(1))
        rows.extend(range(start, end + 1))
    return rows


def _groups_from_column(elect, col: int, bill_row: int, total_row: int) -> List[Tuple[int, set]]:
    """
    Return (charge_row, meter numbers) for a kWh-proportion block.

    The residual charge row is first, matching Elect Charge:
    charge = bill − sum(other rounded shares).
    """
    residual = None
    weighted = []
    for row in range(5, 25):
        value = elect.cell(row=row, column=col).value
        if not isinstance(value, str) or not value.startswith('='):
            continue
        compact = value.replace(' ', '').replace('$', '')
        if f'(C{row}' in compact:
            continue
        # ROUND(BC34*(BC29/BC32),3) style, or a residual SUM of sibling rows.
        weight_row = _weight_row(compact, bill_row, total_row)
        if weight_row:
            meters = set(_meter_tokens(elect.cell(row=weight_row, column=1).value))
            weighted.append((row, weight_row, meters))
        elif str(bill_row) in compact and 'SUM' in compact.upper() and str(total_row) not in compact:
            residual = row
    # Stable order: residual first, then by kWh row.
    weighted.sort(key=lambda item: item[1])
    used_kwh_rows = {weight_row for _, weight_row, _ in weighted}
    span = _kwh_span(elect, col, total_row)
    residual_meters = set()
    for kwh_row in span:
        if kwh_row not in used_kwh_rows:
            residual_meters.update(_meter_tokens(elect.cell(row=kwh_row, column=1).value))
    groups = []
    if residual:
        groups.append((residual, residual_meters))
    elif weighted:
        # No plug row: every share is rounded. Keep sheet order.
        pass
    for charge_row, _kwh_row, meters in weighted:
        groups.append((charge_row, set(meters)))
    return groups


def _kwh_span(elect, col: int, total_row: int) -> List[int]:
    formula = elect.cell(row=total_row, column=col).value
    if not isinstance(formula, str):
        return []
    match = re.search(r'(\d+)\s*:\s*\$?[A-Z]*\$?(\d+)', formula)
    if not match:
        return []
    return list(range(int(match.group(1)), int(match.group(2)) + 1))


def _weight_row(compact_formula: str, bill_row: int, total_row: int) -> Optional[int]:
    # ROUND(BC34*(BC30/BC32),3) → 30
    match = re.search(rf'\*\([A-Z]+(\d+)/[A-Z]+{total_row}\)', compact_formula)
    if match and str(bill_row) in compact_formula:
        return int(match.group(1))
    return None


def try_load_citybase(path: str) -> Optional[CitybaseModel]:
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        names = set(wb.sheetnames)
        wb.close()
    except Exception as exc:
        logger.info('Not a Citybase workbook %s: %s', path, exc)
        return None
    if not {'Allocation', 'AC DEPT', 'Elect Charge'} <= names:
        return None
    return CitybaseModel(path)
