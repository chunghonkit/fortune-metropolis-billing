"""
Workbook Updater for Cost Sheet and Cost Allocation

Copies previous month's workbooks and updates them with current month's data
while preserving all formatting (fonts, merges, widths, number formats, structure).

Critical rules:
- Copy source files, NEVER overwrite originals
- Write computed VALUES not formulas (no cross-file links that become #REF!)
- Preserve all formatting exactly
- Odd-cent differences of 0.01 are acceptable
"""

import ast
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from datetime import datetime, timedelta
import shutil
import logging
import re
from decimal import Decimal, ROUND_HALF_UP

logger = logging.getLogger(__name__)

_ACCOUNT_RE = re.compile(r'\d{5}-\d{5}-\d')
_MONEY_RE = re.compile(r'\$\s*[\d,]+\.\d{2}')
_ROUND_FORMULA_RE = re.compile(
    r'^=ROUND\((?P<body>.+),2\)(?P<tail>[+-]\d+(?:\.\d+)?)?$'
)
_CHECK_METER_SW_NAME = 'CHECK_METER_SW'
# Citybase's May and June comparison snapshots park the seawater-pump
# share of meter 9046787 in AC. From July the previous column matches
# the year-grid centres, with that share left in SW.
_SEAWATER_RECLASS_BEFORE = '2025-07'


def _excel_round(value, places=2):
    quantum = Decimal('1').scaleb(-places)
    return float(Decimal(str(value)).quantize(quantum, rounding=ROUND_HALF_UP))


def _as_date(value):
    if isinstance(value, datetime):
        return value
    if value is None:
        return None
    text = str(value).strip()
    for fmt in ('%d-%m-%y', '%Y-%m-%d'):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _present_units(bill):
    units = bill.get('billed_units')
    if units is None:
        units = bill.get('kwh')
    if units is None:
        return None
    reversed_units = bill.get('reversed_units') or 0
    units = float(units) - float(reversed_units)
    if abs(units - round(units)) < 1e-6:
        return int(round(units))
    return units


def _present_days(bill):
    reversed_to = bill.get('reversed_to')
    if reversed_to and bill.get('to_date'):
        start = _as_date(reversed_to) + timedelta(days=1)
        end = _as_date(bill['to_date'])
        return (end - start).days + 1
    return bill.get('days')


def _period_dates(bill):
    end = _as_date(bill.get('to_date'))
    if bill.get('reversed_to'):
        start = _as_date(bill['reversed_to']) + timedelta(days=1)
    else:
        start = _as_date(bill.get('from_date'))
    return start, end


def _sheet_account(ws):
    for row in ws.iter_rows(min_row=1, max_row=20, max_col=40):
        for cell in row:
            if isinstance(cell.value, str):
                found = _ACCOUNT_RE.search(cell.value)
                if found:
                    return found.group(0)
    return None


def _period_columns(ws):
    last_col = present_col = None
    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=25, max_col=60):
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            compact = re.sub(r'\s+', ' ', cell.value).strip().lower()
            if compact.startswith('last period'):
                last_col = cell.column
                header_row = cell.row
            elif compact.startswith('present period'):
                present_col = cell.column
                header_row = cell.row
    return last_col, present_col, header_row


def _label_row(ws, snippet, start=15, end=35):
    needle = re.sub(r'\s+', ' ', snippet).lower()
    for row in ws.iter_rows(min_row=start, max_row=end, max_col=20):
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            compact = re.sub(r'\s+', ' ', cell.value).lower()
            if needle in compact:
                return cell.row
    return None


def _stored_units(value):
    if value is None:
        return None
    if isinstance(value, float) and abs(value - round(value)) < 1e-6:
        return int(round(value))
    return value


def _replace_money(text, amount):
    """Swap the dollar amount and keep the width Citybase padded after $."""
    if not isinstance(text, str) or '$' not in text:
        return text
    match = _MONEY_RE.search(text)
    if not match:
        return text
    number = f'{float(amount):,.2f}'
    width = len(match.group(0))
    pad = width - 1 - len(number)
    if pad < 0:
        pad = 0
    replacement = '$' + (' ' * pad) + number
    return text[:match.start()] + replacement + text[match.end():]


def _fit_amount_for_row(ws, row, bill):
    lines = bill.get('fit_lines') or []
    if not lines:
        return abs(float(bill.get('fit_amount') or 0))
    window = []
    for col in range(1, (ws.max_column or 1) + 1):
        value = ws.cell(row, col).value
        if isinstance(value, str):
            window.append(value)
    blob = ' '.join(window)
    for line in lines:
        if line['meter'] in blob:
            return float(line['amount'])
    if len(lines) == 1:
        return float(lines[0]['amount'])
    return None


def _update_bill_amount_labels(ws, bill):
    total = bill.get('total_amount')
    fit_total = abs(float(bill.get('fit_amount') or 0))
    for row in ws.iter_rows(min_row=30, max_row=45, max_col=40):
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            if 'Electricity Bill Amt' in cell.value and total is not None:
                cell.value = _replace_money(cell.value, float(total))
            elif 'Feed-in Tariff' in cell.value:
                amount = _fit_amount_for_row(ws, cell.row, bill)
                if amount is not None:
                    cell.value = _replace_money(cell.value, amount)
            elif 'Actual Electricity' in cell.value and total is not None:
                cell.value = _replace_money(cell.value, float(total) + fit_total)


def _formula_fudge(body):
    match = re.search(r'([+-]\d+(?:\.\d+)?)\s*$', body)
    if not match:
        return 0.0
    return float(match.group(1))


def _format_tail(tail):
    if abs(tail) < 0.0005:
        return ''
    text = f'{tail:.2f}'
    if not text.startswith('-'):
        text = '+' + text
    return text


def _centre_key(label):
    from app.citybase_model import canonical_centre_key
    text = str(label).replace('Commerical', 'Commercial').replace('commerical', 'commercial')
    return canonical_centre_key(text)


def _update_account_amount_cells(ws, account, allocations):
    """
    Keep each centre formula pointed at AC DEPT, and move only the rounding
    plug so the displayed cents follow this month's allocation.
    A literal amount (the food-court sheet) is the allocation itself.
    """
    by_centre = {}
    for alloc in allocations:
        if alloc.get('account') != account:
            continue
        key = _centre_key(alloc['centre'])
        by_centre[key] = by_centre.get(key, 0.0) + float(alloc['amount'])
    if not by_centre:
        return

    amount_col = None
    for row in ws.iter_rows(min_row=30, max_row=45, max_col=70):
        for cell in row:
            if isinstance(cell.value, str) and 'AC DEPT' in cell.value:
                amount_col = cell.column
                break
        if amount_col:
            break
    if amount_col is None:
        amount_col = 56

    lines = []
    for row in range(30, 46):
        label = ws.cell(row, 4).value
        if not isinstance(label, str) or not label.strip():
            continue
        key = _centre_key(label)
        if key not in by_centre:
            continue
        cell = ws.cell(row, amount_col)
        lines.append((row, key, cell, by_centre[key]))
    if not lines:
        return

    # AC DEPT sometimes drops a cent inside its own formulas. The account
    # sheet's existing rounding plug puts that cent back so the lines
    # still add up to the bill (net due, or net due plus FiT).
    bases = [
        float(alloc['allocation_base'])
        for alloc in allocations
        if alloc.get('account') == account and alloc.get('allocation_base') is not None
    ]
    target_total = _excel_round(bases[0] if bases else sum(amount for _, _, _, amount in lines))
    residual = _excel_round(target_total - sum(_excel_round(amount) for _, _, _, amount in lines))
    if abs(residual) > 0.05:
        # A missing centre would dump dollars onto the plug. Leave it.
        residual = 0.0
    # Citybase keeps the rounding plug on the DC line of the account.
    balance_row = None
    for row, key, cell, _amount in lines:
        if key == 'DC' and isinstance(cell.value, str) and cell.value.startswith('=ROUND('):
            balance_row = row
            break

    for row, _key, cell, amount in lines:
        if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
            target = _excel_round(amount)
            cell.value = int(target) if abs(target - round(target)) < 1e-9 else target
            continue
        if row != balance_row or not isinstance(cell.value, str):
            continue
        match = _ROUND_FORMULA_RE.match(cell.value)
        if not match:
            continue
        target = _excel_round(amount + residual)
        fudge = _formula_fudge(match.group('body'))
        tail = _excel_round(target - _excel_round(amount + fudge))
        cell.value = f"=ROUND({match.group('body')},2){_format_tail(tail)}"


def _update_account_sheet(ws, bill):
    last_col, present_col, header_row = _period_columns(ws)
    if not last_col or not present_col:
        logger.warning('Account sheet %s has no Last/Present period columns', ws.title)
        return
    units_row = _label_row(ws, 'totally')
    days_row = _label_row(ws, 'serviced')
    start, end = _period_dates(bill)
    if start and end:
        ws.cell(header_row + 2, last_col).value = start
        ws.cell(header_row + 2, present_col).value = end
    if units_row:
        previous = ws.cell(units_row, present_col).value
        ws.cell(units_row, last_col).value = _stored_units(previous)
        present_units = _present_units(bill)
        if present_units is not None:
            ws.cell(units_row, present_col).value = present_units
    if days_row:
        previous_days = ws.cell(days_row, present_col).value
        ws.cell(days_row, last_col).value = previous_days
        present_days = _present_days(bill)
        if present_days is not None:
            ws.cell(days_row, present_col).value = present_days
    _update_bill_amount_labels(ws, bill)


def _year_grid_centres(ws, column):
    from app.citybase_model import canonical_centre_key as key_of
    centres = {}
    for row in range(1, 40):
        label = ws.cell(row, 1).value
        if not isinstance(label, str) or not label.strip():
            continue
        if label.strip().lower().startswith('total'):
            continue
        value = ws.cell(row, column).value
        if isinstance(value, (int, float)):
            centres[key_of(label)] = float(value)
    return centres


def _defined_number(wb, name):
    defined = wb.defined_names.get(name)
    if defined is None:
        return None
    try:
        return float(str(defined.attr_text).lstrip('='))
    except (TypeError, ValueError):
        return None


def _set_defined_number(wb, name, value):
    existing = wb.defined_names.get(name)
    if existing is not None:
        del wb.defined_names[name]
    wb.defined_names.add(DefinedName(name=name, attr_text=repr(float(value))))


def _cached_check_meter_sw(master_path):
    if not master_path:
        return None
    try:
        wb = openpyxl.load_workbook(master_path, data_only=True)
    except Exception:
        return None
    try:
        if 'AC DEPT' not in wb.sheetnames:
            return None
        value = wb['AC DEPT']['C7'].value
        if isinstance(value, (int, float)):
            return float(value)
        return None
    finally:
        wb.close()


def _comparison_sheet(wb):
    for sheet in wb.worksheets:
        if sheet.title.strip() == 'Comparison':
            return sheet
    return None


def _live_comparison_columns(ws):
    """Previous and current amount columns: the pair sitting left of Difference."""
    for row in ws.iter_rows(min_row=3, max_row=3, max_col=ws.max_column or 1):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip().lower().startswith('difference'):
                return cell.column - 2, cell.column - 1
    return None, None


def _update_comparison_sheet(ws, year_grid, billing_month, seawater):
    previous_col, current_col = _live_comparison_columns(ws)
    if not previous_col or not current_col:
        logger.warning('Comparison sheet has no live date pair')
        return
    year, month = (int(part) for part in billing_month.split('-'))
    current_date = datetime(year, month, 1)
    previous_date = ws.cell(3, current_col).value
    if not isinstance(previous_date, datetime):
        previous_month = month - 1 or 12
        previous_year = year if month > 1 else year - 1
        previous_date = datetime(previous_year, previous_month, 1)
    ws.cell(3, previous_col).value = previous_date
    ws.cell(3, current_col).value = current_date

    previous_month_num = month - 1 or 12
    centres = _year_grid_centres(year_grid, previous_month_num + 1)
    if billing_month < _SEAWATER_RECLASS_BEFORE and seawater:
        if 'AC' in centres and 'SW' in centres:
            centres['AC'] = centres['AC'] + float(seawater)
            centres['SW'] = centres['SW'] - float(seawater)
    from app.citybase_model import canonical_centre_key
    for row in range(4, 32):
        label = ws.cell(row, 1).value
        if not isinstance(label, str) or label.strip().lower().startswith('total'):
            continue
        key = canonical_centre_key(label)
        if key not in centres:
            continue
        amount = _excel_round(centres[key])
        ws.cell(row, previous_col).value = int(amount) if abs(amount - round(amount)) < 1e-9 else amount


def _update_cost_sheet_details(
    wb,
    allocations,
    billing_month,
    parsed_bills,
    year_grid,
    previous_seawater,
    check_meter_sw,
):
    bills_by_account = {}
    for bill in parsed_bills or []:
        account = bill.get('account')
        if account:
            bills_by_account[account] = bill
    for sheet in wb.worksheets:
        account = _sheet_account(sheet)
        if not account or account not in bills_by_account:
            continue
        bill = bills_by_account[account]
        _update_account_sheet(sheet, bill)
        _update_account_amount_cells(sheet, account, allocations)
    comparison = _comparison_sheet(wb)
    if comparison is not None and year_grid is not None:
        _update_comparison_sheet(comparison, year_grid, billing_month, previous_seawater)
    if check_meter_sw is not None:
        _set_defined_number(wb, _CHECK_METER_SW_NAME, check_meter_sw)


# Elect Charge keeps two live charge columns. BB is last month, BC is this month.
_BB_COL = 54
_BC_COL = 55
_CHARGE_ROWS = range(4, 42)
_CELL_REF = re.compile(r'(?<![A-Z])[A-Z]{1,3}\$?\d')
_BC_REF = re.compile(r'(?<![A-Z])BC(?=\$?\d)')
_METER_ID = re.compile(r'\d{7,8}')


def _formula_number(value):
    """Whole units and dollar amounts stay integers, the way Elect Charge stores them."""
    number = float(value)
    if abs(number - round(number)) < 1e-6:
        return str(int(round(number)))
    text = f'{number:.4f}'.rstrip('0').rstrip('.')
    return text


def _stored_number(value):
    number = float(value)
    if abs(number - round(number)) < 1e-6:
        return int(round(number))
    return number


def _is_structural_formula(value):
    return isinstance(value, str) and value.startswith('=') and _CELL_REF.search(value)


def _is_arithmetic_formula(value):
    if not (isinstance(value, str) and value.startswith('=')):
        return False
    if _CELL_REF.search(value):
        return False
    try:
        _eval_arithmetic(value)
    except (ValueError, SyntaxError, TypeError):
        return False
    return True


def _eval_arithmetic(formula):
    """Evaluate a Citybase input formula such as '=27513+7905' or '=88327+3059-0'."""
    tree = ast.parse(formula[1:], mode='eval')

    def evaluate(node):
        if isinstance(node, ast.Expression):
            return evaluate(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return node.value
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.USub):
            return -evaluate(node.operand)
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.UAdd):
            return evaluate(node.operand)
        if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub)):
            left = evaluate(node.left)
            right = evaluate(node.right)
            return left + right if isinstance(node.op, ast.Add) else left - right
        raise ValueError(f'not an input formula: {formula}')

    return evaluate(tree)


def _rewrite_bc_to_bb(formula):
    return _BC_REF.sub('BB', formula)


def _meter_ids(label):
    return _METER_ID.findall(str(label or ''))


def _index_bills(bills):
    by_meter = {}
    for bill in bills or []:
        for meter in bill.get('meters') or []:
            number = str(meter.get('meter_no') or '')
            if number and number not in by_meter:
                by_meter[number] = (bill, meter)
    return by_meter


def _meter_primary(meter):
    if meter.get('primary') is not None:
        return meter['primary']
    return meter.get('consumption') or 0


def _meter_adjustment(meter):
    return meter.get('adjustment') or 0


def _amount_formula(bill):
    """Single-meter rows are the amount due. FiT rows keep net + |FiT| as a formula."""
    total = bill.get('total_amount') or 0
    fit = abs(bill.get('fit_amount') or 0)
    if fit:
        return f'={_formula_number(total)}+{_formula_number(fit)}'
    return _stored_number(total)


def _kwh_formula(meter_numbers, by_meter):
    """
    Citybase writes each meter's first register, then the first meter's
    adjustment, including a zero: =82374+3035-0 or =20674+6878-396.
    """
    if not meter_numbers or any(number not in by_meter for number in meter_numbers):
        missing = [number for number in meter_numbers if number not in by_meter]
        logger.warning('Elect Charge kWh formula missing meters %s', missing)
        return None
    terms = [_formula_number(_meter_primary(by_meter[number][1])) for number in meter_numbers]
    adjustment = abs(_meter_adjustment(by_meter[meter_numbers[0]][1]))
    return '=' + '+'.join(terms) + '-' + _formula_number(adjustment)


def _kwh_blocks(ws):
    """
    Ranges of meter rows that a Total formula sums, plus the Amt row under each.

    Row 32 sums the chiller kWh (28:31) and row 34 is that account's amount.
    Row 39 sums the FiT kWh (36:38) and row 41 is that account's amount.
    """
    blocks = []
    for row in _CHARGE_ROWS:
        value = ws.cell(row=row, column=_BC_COL).value
        if not _is_structural_formula(value):
            continue
        match = re.fullmatch(r'=SUM\([A-Z]{1,3}\$?(\d+):[A-Z]{1,3}\$?(\d+)\)', value.replace(' ', ''))
        if not match:
            continue
        start, end = int(match.group(1)), int(match.group(2))
        summed = list(range(start, end + 1))
        # Dollar subtotals sum charge formulas. A kWh total sums only meter
        # inputs, and the account amount sits on the Amt row just below.
        if any(_is_structural_formula(ws.cell(row=item, column=_BC_COL).value) for item in summed):
            continue
        labels = [ws.cell(row=item, column=1).value for item in summed]
        if not labels or any(not _meter_ids(label) for label in labels):
            continue
        amount_row = None
        for follower in range(row + 1, row + 4):
            if str(ws.cell(row=follower, column=1).value or '').strip().lower() == 'amt':
                amount_row = follower
                break
        if amount_row is None:
            continue
        blocks.append({'rows': summed, 'amount_row': amount_row})
    return blocks


def _roll_and_fill_elect_charge(ws, billing_month, bills):
    """
    Move column BC into column BB, then write this month's meter charges into BC.

    Split and total formulas stay formulas, with the column letter rewritten.
    Input formulas (a sum of meter units or net + |FiT|) are stored in BB as
    the calculated number. Typed charges and kWh copy as values.
    """
    year, month = (int(part) for part in billing_month.split('-'))
    new_date = datetime(year, month, 1)
    rolled = {}
    for row in _CHARGE_ROWS:
        current = ws.cell(row=row, column=_BC_COL).value
        if current is None:
            continue
        if isinstance(current, datetime):
            ws.cell(row=row, column=_BB_COL).value = current
            if row == 4:
                ws.cell(row=row, column=_BC_COL).value = new_date
            rolled[row] = 'date'
            continue
        if _is_structural_formula(current):
            ws.cell(row=row, column=_BB_COL).value = _rewrite_bc_to_bb(current)
            rolled[row] = 'structural'
            continue
        if _is_arithmetic_formula(current):
            ws.cell(row=row, column=_BB_COL).value = _stored_number(_eval_arithmetic(current))
            rolled[row] = 'arithmetic'
            continue
        if isinstance(current, (int, float)):
            ws.cell(row=row, column=_BB_COL).value = current
            rolled[row] = 'number'
            continue

    by_meter = _index_bills(bills)
    blocks = _kwh_blocks(ws)
    kwh_rows = {}
    amount_rows = {}
    for block in blocks:
        for row in block['rows']:
            kwh_rows[row] = block
        if block['amount_row']:
            amount_rows[block['amount_row']] = block

    for row, kind in rolled.items():
        if kind not in ('number', 'arithmetic'):
            continue
        label = ws.cell(row=row, column=1).value
        meters = _meter_ids(label)
        if row in kwh_rows:
            if len(meters) == 1 and kind == 'number':
                found = by_meter.get(meters[0])
                if not found:
                    logger.warning('No parsed kWh for Elect Charge meter %s', meters[0])
                    continue
                ws.cell(row=row, column=_BC_COL).value = _stored_number(_meter_primary(found[1]))
            elif len(meters) >= 1:
                formula = _kwh_formula(meters, by_meter)
                if formula:
                    ws.cell(row=row, column=_BC_COL).value = formula
            continue
        if row in amount_rows or (meters and kind == 'arithmetic'):
            block = amount_rows.get(row)
            owner_meters = meters
            if block:
                owner_meters = []
                for item in block['rows']:
                    owner_meters.extend(_meter_ids(ws.cell(row=item, column=1).value))
            bill = None
            for number in owner_meters:
                if number in by_meter:
                    bill = by_meter[number][0]
                    break
            if bill is None and meters:
                for number in meters:
                    if number in by_meter:
                        bill = by_meter[number][0]
                        break
            if bill is None:
                logger.warning('No bill for Elect Charge row %s (%s)', row, label)
                continue
            ws.cell(row=row, column=_BC_COL).value = _amount_formula(bill)
            continue
        if len(meters) == 1 and kind == 'number':
            found = by_meter.get(meters[0])
            if not found:
                logger.warning('No bill for Elect Charge meter %s', meters[0])
                continue
            ws.cell(row=row, column=_BC_COL).value = _stored_number(found[0].get('total_amount') or 0)


class WorkbookUpdater:
    """
    Updates Cost Sheet and Cost Allocation workbooks with current month's data.
    """
    
    def __init__(self, metropolis_root: Path):
        self.metropolis_root = Path(metropolis_root).expanduser()
    
    def copy_previous_month_workbooks(
        self,
        current_month: str,  # YYYY-MM
        previous_month: str,  # YYYY-MM
    ) -> Dict[str, Path]:
        """
        Copy previous month's workbooks to current month's out/ directory.
        
        Handles real Citybase filenames:
        - Cost Allocation - Electricity 2007-2.xls (long-lived master)
        - Cost Sheet-2025-03.xls (contains previous month)
        
        Looks in:
        - ~/Metropolis/{previous}/masters/
        - ~/Metropolis/{previous}/out/
        - ~/Metropolis/{current}/masters/
        
        Args:
            current_month: Current billing month (YYYY-MM)
            previous_month: Previous month (YYYY-MM)
        
        Returns:
            Dict mapping workbook type to copied file path
        """
        current_month_dir = self.metropolis_root / current_month
        previous_month_dir = self.metropolis_root / previous_month
        
        # Ensure output directory exists
        out_dir = current_month_dir / 'out'
        out_dir.mkdir(parents=True, exist_ok=True)
        
        copied_files = {}
        
        # Previous month's own out/ is the chain (June copies May's outputs).
        # The month folder itself is the start of the chain (May copies April).
        # Never search the month being written, and never its out/ folder.
        search_dirs = [
            previous_month_dir / 'out',
            previous_month_dir,
            previous_month_dir / 'masters',
        ]
        
        # Find Cost Allocation workbook (long-lived file)
        cost_allocation_file = None
        for search_dir in search_dirs:
            if not search_dir.exists():
                continue
            # Look for files containing "Cost Allocation" with .xls or .xlsx
            for pattern in ['*Cost Allocation*.xls', '*Cost Allocation*.xlsx']:
                matches = list(search_dir.glob(pattern))
                if matches:
                    # Take the first match
                    cost_allocation_file = matches[0]
                    logger.info(f"Found Cost Allocation: {cost_allocation_file}")
                    break
            if cost_allocation_file:
                break
        
        if cost_allocation_file and cost_allocation_file.exists():
            # Copy to current month's out/ with same name (long-lived)
            dest_file = out_dir / cost_allocation_file.name
            shutil.copy2(cost_allocation_file, dest_file)
            logger.info(f"Copied Cost Allocation: {cost_allocation_file} -> {dest_file}")
            copied_files['cost_allocation'] = dest_file
        else:
            logger.warning(f"Could not find Cost Allocation workbook")
        
        # Find Cost Sheet workbook (contains previous month in name)
        cost_sheet_file = None
        for search_dir in search_dirs:
            if not search_dir.exists():
                continue
            # Look for files containing "Cost Sheet" with .xls or .xlsx
            for pattern in ['*Cost Sheet*.xls', '*Cost Sheet*.xlsx']:
                matches = list(search_dir.glob(pattern))
                for match in matches:
                    # Prefer files that contain the previous month in the name
                    if previous_month in match.name:
                        cost_sheet_file = match
                        logger.info(f"Found Cost Sheet (with month): {cost_sheet_file}")
                        break
                if cost_sheet_file:
                    break
            if cost_sheet_file:
                break
        
        # If not found with previous month, take any Cost Sheet file
        if not cost_sheet_file:
            for search_dir in search_dirs:
                if not search_dir.exists():
                    continue
                for pattern in ['*Cost Sheet*.xls', '*Cost Sheet*.xlsx']:
                    matches = list(search_dir.glob(pattern))
                    if matches:
                        cost_sheet_file = matches[0]
                        logger.info(f"Found Cost Sheet (generic): {cost_sheet_file}")
                        break
                if cost_sheet_file:
                    break
        
        if cost_sheet_file and cost_sheet_file.exists():
            # Update filename to use current month
            # Cost Sheet-2025-03.xls -> Cost Sheet-2025-04.xls
            original_ext = cost_sheet_file.suffix  # .xls or .xlsx
            dest_name = f"Cost Sheet-{current_month}{original_ext}"
            dest_file = out_dir / dest_name
            shutil.copy2(cost_sheet_file, dest_file)
            logger.info(f"Copied Cost Sheet: {cost_sheet_file} -> {dest_file}")
            copied_files['cost_sheet'] = dest_file
        else:
            logger.warning(f"Could not find Cost Sheet workbook")
        
        return copied_files
    
    def update_cost_allocation_workbook(
        self,
        workbook_path: Path,
        allocations: List[Dict],
        billing_month: str,
        parsed_bills: List[Dict],
        meter_log: Dict
    ) -> Path:
        """
        Update Cost Allocation workbook with current month's allocation data.
        
        Writes computed VALUES (not formulas) to preserve formatting.
        
        Key updates:
        - Elect Charge sheet: check-meter 6681757 previous/present/delta
        - Title row: update month
        
        Args:
            workbook_path: Path to Cost Allocation workbook (already copied)
            allocations: List of allocation dicts from allocation engine
            billing_month: YYYY-MM
            parsed_bills: List of parsed bills
            meter_log: Dict with check-meter readings
            
        Returns:
            Updated workbook path (same as input)
        """
        # Check if .xls - openpyxl only supports .xlsx
        if workbook_path.suffix.lower() == '.xls':
            logger.warning(f"Cost Allocation is .xls format - openpyxl requires .xlsx. "
                          f"Consider converting with LibreOffice first.")
            # For now, just update the .xls file we can
            # In production, you'd use xlrd/xlwt or convert with soffice
        
        try:
            wb = openpyxl.load_workbook(workbook_path)
        except Exception as e:
            logger.error(f"Failed to load Cost Allocation workbook: {e}")
            # If it's .xls and fails, just log and return
            return workbook_path
        
        # DO NOT update title/header cells - preserve original titles from 2008
        # The golden workbooks keep their original titles unchanged
        
        # Find Elect Charge sheet (check-meter 6681757)
        elect_charge_sheet = None
        for sheet in wb.worksheets:
            # Look for sheet with "Elect" or "Charge" in name
            if 'elect' in sheet.title.lower() or 'charge' in sheet.title.lower():
                elect_charge_sheet = sheet
                break
        
        if elect_charge_sheet and parsed_bills:
            _roll_and_fill_elect_charge(elect_charge_sheet, billing_month, parsed_bills)

        if elect_charge_sheet and meter_log.get('previous') is not None and meter_log.get('present') is not None:
            logger.info(f"Found Elect Charge sheet: {elect_charge_sheet.title}")

            # Citybase: meter 6681757 is BP31, previous BR31, present BS31.
            # Delta stays the formula BQ31 =BS31-BR31.
            meter_no = str(meter_log.get('meter_no') or '6681757')
            if meter_no in str(elect_charge_sheet['BP31'].value or ''):
                elect_charge_sheet['BR31'] = meter_log['previous']
                elect_charge_sheet['BS31'] = meter_log['present']
                logger.info(
                    'Check meter %s BR31=%s BS31=%s',
                    meter_no, meter_log['previous'], meter_log['present'],
                )

            # Try to find meter 6681757 row by searching for the meter number
            meter_no = meter_log.get('meter_no', '6681757')
            meter_row = None
            
            # Search for meter number in the sheet - extend search to ALL columns, not just 10
            for row in elect_charge_sheet.iter_rows(min_row=1, max_row=100, min_col=1, max_col=200):
                for cell in row:
                    if cell.value and str(meter_no) in str(cell.value):
                        meter_row = cell.row
                        logger.info(f"Found meter {meter_no} at row {meter_row}")
                        break
                if meter_row:
                    break
            
            # If found meter row, try to write readings
            # Common pattern: Previous, Present, Delta in adjacent columns
            # User mentioned BR31/BS31/BQ31, so columns might be around BQ-BS (columns 69-71)
            if meter_row:
                # Try common column positions for meter readings
                # BQ=69, BR=70, BS=71 (0-indexed: 68, 69, 70)
                # But let's search for "Previous" and "Present" headers first
                
                # Search for column headers in row above meter or in header row - search ALL columns
                prev_col = None
                present_col = None
                delta_col = None
                
                # Check a few rows above meter_row for headers - search ALL columns
                for header_row_idx in range(max(1, meter_row - 5), meter_row):
                    row_cells = list(elect_charge_sheet.iter_rows(
                        min_row=header_row_idx, max_row=header_row_idx, 
                        min_col=1, max_col=200, values_only=False
                    ))[0]
                    
                    for cell in row_cells:
                        if cell.value:
                            val_lower = str(cell.value).lower()
                            # Match: "Previous", "Prev", "上期"
                            if 'previous' in val_lower or 'prev' in val_lower or '上期' in val_lower:
                                prev_col = cell.column
                                logger.info(f"Found Previous column: {openpyxl.utils.get_column_letter(prev_col)}")
                            # Match: "Present", "Current", "Current Reading", "今期"
                            elif 'present' in val_lower or 'current' in val_lower or '今期' in val_lower:
                                present_col = cell.column
                                logger.info(f"Found Present/Current column: {openpyxl.utils.get_column_letter(present_col)}")
                            # Match: "Delta", "Diff", "Difference", "度數"
                            elif 'delta' in val_lower or 'diff' in val_lower or '度數' in val_lower:
                                delta_col = cell.column
                                logger.info(f"Found Delta column: {openpyxl.utils.get_column_letter(delta_col)}")
                
                # If we found the columns, write the values
                if prev_col and present_col:
                    elect_charge_sheet.cell(row=meter_row, column=prev_col).value = meter_log['previous']
                    elect_charge_sheet.cell(row=meter_row, column=present_col).value = meter_log['present']
                    
                    # Check if delta cell has a formula - if so, don't overwrite it
                    if delta_col:
                        delta_cell = elect_charge_sheet.cell(row=meter_row, column=delta_col)
                        # If it's a formula (like =BS31-BR31), leave it alone
                        # Otherwise, write the delta value
                        if not (delta_cell.value and isinstance(delta_cell.value, str) and delta_cell.value.startswith('=')):
                            delta = round(meter_log['present'] - meter_log['previous'], 1)
                            delta_cell.value = delta
                            logger.info(f"Wrote delta value {delta}")
                        else:
                            logger.info(f"Delta cell has formula, leaving it to auto-calculate")
                    
                    logger.info(f"Updated meter readings at row {meter_row}: "
                              f"prev={meter_log['previous']}, present={meter_log['present']}")
                else:
                    logger.warning(f"Could not find Previous/Present columns for meter {meter_no}")
            else:
                logger.warning(f"Could not find meter {meter_no} in Elect Charge sheet")
        
        # Save the updated workbook
        try:
            wb.save(workbook_path)
            wb.close()
            logger.info(f"Updated Cost Allocation workbook: {workbook_path}")
        except Exception as e:
            logger.error(f"Failed to save Cost Allocation workbook: {e}")
            wb.close()
        
        return workbook_path
    
    def update_cost_sheet_workbook(
        self,
        workbook_path: Path,
        allocations: List[Dict],
        billing_month: str,
        parsed_bills: List[Dict],
        meter_log: Dict,
        check_meter_sw: Optional[float] = None,
        master_path: Optional[Path] = None,
    ) -> Path:
        """
        Update Cost Sheet workbook with current month's data.
        
        Updates:
        - Year-grid sheet: month column with allocation totals per cost centre
        - Allocation % for each account/cost centre
        - Separated Amount (computed values)
        - Check-meter 6681757 readings if present
        
        Writes VALUES not formulas to preserve formatting.
        
        Args:
            workbook_path: Path to Cost Sheet workbook (already copied)
            allocations: List of allocation dicts
            billing_month: YYYY-MM (e.g., "2025-05")
            parsed_bills: List of parsed bills
            meter_log: Dict with check-meter readings
            
        Returns:
            Updated workbook path (same as input)
        """
        # Check if .xls - openpyxl only supports .xlsx
        if workbook_path.suffix.lower() == '.xls':
            logger.warning(f"Cost Sheet is .xls format - openpyxl requires .xlsx. "
                          f"Consider converting with LibreOffice first.")
            # For now, just update what we can
        
        try:
            wb = openpyxl.load_workbook(workbook_path)
        except Exception as e:
            logger.error(f"Failed to load Cost Sheet workbook: {e}")
            return workbook_path
        
        # DO NOT update title/header cells - preserve original titles
        # The golden workbooks keep their original titles unchanged
        
        # Find year-grid sheet matching the billing year (e.g., "01-12 2025")
        year_grid_sheet = None
        billing_year = billing_month.split('-')[0]
        
        for sheet in wb.worksheets:
            sheet_name = sheet.title
            # Look for sheets with "01-12" AND the billing year
            if '01-12' in sheet_name and billing_year in sheet_name:
                year_grid_sheet = sheet
                logger.info(f"Found year-grid sheet matching {billing_year}: {sheet.title}")
                break
        
        # If no exact match, try just "01-12" as fallback (but log warning)
        if not year_grid_sheet:
            for sheet in wb.worksheets:
                if '01-12' in sheet.title.lower():
                    year_grid_sheet = sheet
                    logger.warning(f"Using fallback year-grid sheet (no {billing_year} match): {sheet.title}")
                    break
        
        if year_grid_sheet:
            # Calculate which column for this month (1=Jan=B, 2=Feb=C, ..., 5=May=F)
            month_num = int(billing_month.split('-')[1])  # Extract month number
            # Assuming column A is labels, column B is Jan (month 1), C is Feb (month 2), etc.
            month_col = month_num + 1  # B=2, C=3, D=4, E=5, F=6 for May
            month_col_letter = openpyxl.utils.get_column_letter(month_col)
            
            logger.info(f"Writing to month column {month_col_letter} (month {month_num})")
            
            # Shared labels stay on their own rows. "Hotel / Commercial / SA(11)"
            # is not added to SA, and "Commercial Common (2)" is not added to C.
            from app.citybase_model import canonical_centre_key
            centre_totals = {}
            for alloc in allocations:
                centre = canonical_centre_key(str(alloc['centre']))
                centre_totals[centre] = centre_totals.get(centre, 0) + alloc['amount']
            
            logger.info(f"Aggregated allocation totals by centre: {centre_totals}")
            
            if not centre_totals:
                logger.warning("No allocation totals to write - allocations list may be empty or invalid")
            
            # Labels sit in column A on the real year grid. Some shared names
            # are in the next column, so accept the first cell on the row that
            # matches a centre we are about to write.
            wanted = set(centre_totals)
            centre_rows = {}
            for row in year_grid_sheet.iter_rows(min_row=1, max_row=80, min_col=1, max_col=4):
                for cell in row:
                    if not cell.value:
                        continue
                    key = canonical_centre_key(str(cell.value))
                    if key in wanted and key not in centre_rows:
                        centre_rows[key] = cell.row
                        logger.info(f"Found centre {key!r} at {cell.coordinate}")
            
            # Write allocation totals to the month column
            for centre, total in centre_totals.items():
                if centre in centre_rows:
                    row = centre_rows[centre]
                    cell = year_grid_sheet.cell(row=row, column=month_col)
                    cell.value = round(total, 2)
                    logger.info(f"Wrote {centre} total ${total:.2f} to {month_col_letter}{row}")
                else:
                    logger.warning(f"Could not find row for centre {centre}")
        else:
            logger.warning("Could not find year-grid sheet in Cost Sheet")

        previous_seawater = _defined_number(wb, _CHECK_METER_SW_NAME)
        if previous_seawater is None:
            previous_seawater = _cached_check_meter_sw(master_path)
        _update_cost_sheet_details(
            wb,
            allocations,
            billing_month,
            parsed_bills,
            year_grid_sheet,
            previous_seawater,
            check_meter_sw,
        )
        
        # Save the updated workbook
        try:
            wb.save(workbook_path)
            wb.close()
            logger.info(f"Updated Cost Sheet workbook: {workbook_path}")
        except Exception as e:
            logger.error(f"Failed to save Cost Sheet workbook: {e}")
            wb.close()
        
        return workbook_path


def process_month_end(
    metropolis_root: Path,
    current_month: str,
    previous_month: str,
    parsed_bills: List[Dict],
    allocation_rules: Dict[str, List[Dict]],
    meter_log: Dict,
    master_path: Optional[Path] = None,
) -> Dict[str, Path]:
    """
    Process month-end: copy previous workbooks and update with current data.
    
    Args:
        metropolis_root: Root directory (e.g., ~/Metropolis)
        current_month: Current billing month (YYYY-MM)
        previous_month: Previous month (YYYY-MM)
        parsed_bills: List of parsed bill dicts
        allocation_rules: Allocation rules from master workbook
        meter_log: Dict with check-meter readings
        
    Returns:
        Dict mapping workbook type to updated file path
    """
    from app.citybase_model import try_load_citybase

    # Column C of AC DEPT is a historical header ("May-08 Charge"). Do not
    # allocate May by those cached weights. The live split is the Citybase
    # formula model: this month's bills, this month's check-meter dial, and
    # the master formulas.
    if not master_path:
        raise RuntimeError('Citybase master workbook is required for month-end allocation')
    model = try_load_citybase(str(master_path))
    if model is None:
        raise RuntimeError(
            f'Could not read Citybase sheets Allocation / AC DEPT / Elect Charge from {master_path}'
        )
    logger.info(f'Allocating from Citybase master {master_path}')
    allocation_result = model.allocate(parsed_bills, meter_log)
    for warning in allocation_result.get('validation', {}).get('errors') or []:
        logger.warning(warning)
    allocations = allocation_result['allocations']
    
    updater = WorkbookUpdater(metropolis_root)
    
    # Copy previous month's workbooks
    copied_files = updater.copy_previous_month_workbooks(
        current_month,
        previous_month
    )
    
    updated_files = {}
    
    # Update Cost Allocation if copied
    if 'cost_allocation' in copied_files:
        updated_files['cost_allocation'] = updater.update_cost_allocation_workbook(
            copied_files['cost_allocation'],
            allocations,
            current_month,
            parsed_bills,
            meter_log
        )
    
    # Update Cost Sheet if copied
    if 'cost_sheet' in copied_files:
        updated_files['cost_sheet'] = updater.update_cost_sheet_workbook(
            copied_files['cost_sheet'],
            allocations,
            current_month,
            parsed_bills,
            meter_log,
            check_meter_sw=allocation_result.get('check_meter_sw'),
            master_path=master_path,
        )
    
    return updated_files
