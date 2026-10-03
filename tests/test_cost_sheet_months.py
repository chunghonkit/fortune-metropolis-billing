"""
Month-end fills every Cost Sheet Citybase updates, not only the year grid.

May is parsed from the real CLP PDFs. June and July CLP packs are larger
than the file cap, so those bills are rebuilt in the same CLP layout from
that month's Cost Allocation charge column (the inputs the allocation
model already uses) and from the reading dates on the account sheets.
The comparison dollars and the year-grid dollars are not copied off the
golden Cost Sheet; they come back out of the live allocation.
"""

from datetime import datetime
from pathlib import Path
import hashlib
import re
import shutil

import fitz
import openpyxl
import pytest

from app.clp_parser import parse_clp_bill
from app.workbook_updater import process_month_end

MASTER = Path('/home/ubuntu/.cursor/projects/workspace/uploads/april-cost-allocation_f85f.xlsx')
ALLOC = {
    '2025-04': Path('/home/ubuntu/.cursor/projects/workspace/uploads/Cost_Allocation_e50c.xlsx'),
    '2025-05': Path('/home/ubuntu/.cursor/projects/workspace/uploads/Cost_Allocation_e28d.xlsx'),
    '2025-06': Path('/home/ubuntu/.cursor/projects/workspace/uploads/Cost_Allocation_c5a4.xlsx'),
    '2025-07': Path('/home/ubuntu/.cursor/projects/workspace/uploads/Cost_Allocation_0165.xlsx'),
}
SHEETS = {
    '2025-04': Path('/tmp/costsheets/xlsx/Cost-Sheet-2025-04.xlsx'),
    '2025-05': Path('/tmp/costsheets/xlsx/Cost-Sheet-2025-05.xlsx'),
    '2025-06': Path('/tmp/costsheets/xlsx/Cost-Sheet-2025-06.xlsx'),
    '2025-07': Path('/tmp/costsheets/xlsx/Cost-Sheet-2025-07.xlsx'),
}
MAY_PDFS = Path('/tmp/maypdfs')
METER_LOGS = {
    '2025-05': {'meter_no': '6681757', 'previous': 96323.1, 'present': 96504.6},
    '2025-06': {'meter_no': '6681757', 'previous': 96504.6, 'present': 96870.4},
    '2025-07': {'meter_no': '6681757', 'previous': 96870.4, 'present': 97266.7},
}
HISTORICAL = {
    '10-2021 to 9-2022', '1-12 2021', '01-12 2022', '01-12 2023',
    '11-2023-10-2024Budget用', '01-12 2024',
}
ROW_ACCOUNT = {
    11: '35204-69738-4',
    12: '40722-61440-7',
    14: '00776-78552-1',
    15: '13639-58422-3',
    16: '79292-23337-6',
    17: '24096-78457-6',
    18: '08731-83914-5',
    19: '97968-02236-6',
    20: '23529-59279-9',
    22: '88931-57029-6',
    23: '70873-85471-3',
    24: '72399-00664-9',
}

pytestmark = pytest.mark.skipif(
    not MASTER.exists() or not MAY_PDFS.exists() or not SHEETS['2025-07'].exists(),
    reason='Citybase workbooks or May PDFs are not in this workspace',
)


def _norm(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, float) and abs(value - round(value)) < 1e-9:
        return int(round(value))
    return value


def _money_text(value):
    return isinstance(value, str) and ('$' in value or 'Bill' in value or 'Tariff' in value or 'Electricity' in value)


def _bill_cell(value):
    """Dates, readings, amounts, and the AC DEPT rounding formulas."""
    if isinstance(value, (int, float, datetime)):
        return True
    if not isinstance(value, str):
        return False
    return _money_text(value) or 'AC DEPT' in value or value.startswith('=ROUND(')


def _close(actual, expected, tolerance=0.011):
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        return abs(float(actual) - float(expected)) <= tolerance
    return _norm(actual) == _norm(expected)


def _pdf(path, text):
    doc = fitz.open()
    lines = text.splitlines() or [' ']
    for start in range(0, len(lines), 40):
        page = doc.new_page()
        page.insert_text((36, 36), '\n'.join(lines[start:start + 40]), fontsize=8)
    doc.save(path)
    doc.close()


def _sheet_account(ws):
    for row in ws.iter_rows(min_row=1, max_row=20, max_col=40):
        for cell in row:
            if isinstance(cell.value, str):
                found = re.search(r'\d{5}-\d{5}-\d', cell.value)
                if found:
                    return found.group(0)
    return None


def _period(ws):
    last = present = header = None
    for row in ws.iter_rows(min_row=1, max_row=25, max_col=60):
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            compact = re.sub(r'\s+', ' ', cell.value).strip().lower()
            if compact.startswith('last period'):
                last, header = cell.column, cell.row
            elif compact.startswith('present period'):
                present, header = cell.column, cell.row
    return last, present, header


def _face(ws):
    """Reading block Citybase stored, used as the CLP face when the PDF is missing."""
    last, present, header = _period(ws)
    start = ws.cell(header + 2, last).value
    end = ws.cell(header + 2, present).value
    units = days = None
    for row in ws.iter_rows(min_row=15, max_row=35, max_col=15):
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            compact = re.sub(r'\s+', ' ', cell.value).lower()
            if 'totally' in compact:
                units = ws.cell(cell.row, present).value
            elif 'serviced' in compact:
                days = ws.cell(cell.row, present).value
    fit_lines = []
    labels = []
    for row in ws.iter_rows(min_row=30, max_row=45, max_col=45):
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            if 'Feed-in Tariff' in cell.value:
                amount = re.search(r'[\d,]+\.\d{2}', cell.value)
                meter = ''
                for col in range(1, (ws.max_column or 1) + 1):
                    other = ws.cell(cell.row, col).value
                    if isinstance(other, str):
                        found = re.search(r'(\d{7,8})\s*\(FiT\)', other)
                        if found:
                            meter = found.group(1)
                if amount and meter:
                    fit_lines.append((meter, amount.group(0).replace(',', '')))
            elif 'Electricity Bill Amt' in cell.value or 'Actual Electricity' in cell.value:
                labels.append(cell.value)
    return {
        'from': start, 'to': end, 'units': units, 'days': days, 'fit_lines': fit_lines,
    }


def _dmy(value):
    return value.strftime('%d-%m-%y')


def _meter_rows(meter, units):
    units = int(round(float(units)))
    return f'{meter} {units} 0 1 {units}'


def bills_from_workbooks(alloc_path, cost_path, directory):
    """CLP text in the May layout, parsed by the same bill parser."""
    directory.mkdir(parents=True, exist_ok=True)
    elect = openpyxl.load_workbook(alloc_path, data_only=False)['Elect Charge']
    cost = openpyxl.load_workbook(cost_path, data_only=False)
    faces = {}
    for name in cost.sheetnames:
        account = _sheet_account(cost[name])
        if account:
            faces[account] = _face(cost[name])
    # Food court is the literal amount on its account sheet, not an Elect Charge row.
    fc_sheet = cost['9050465']
    fc_total = fc_sheet['BD33'].value
    faces_fc = faces.get('82805-94744-7', {})

    def amount_row(row):
        value = elect.cell(row, 55).value
        if isinstance(value, (int, float)):
            return float(value), 0.0
        text = str(value or '')
        parts = [float(part) for part in re.findall(r'[+-]?\d+(?:\.\d+)?', text)]
        if len(parts) >= 2 and row in (14, 41):
            return parts[0], -abs(parts[1])
        if parts:
            return parts[0], 0.0
        return 0.0, 0.0

    specs = []
    chiller_total, _ = amount_row(34)
    specs.append(('55861-52267-1', chiller_total, 0.0, [
        (str(elect.cell(row, 1).value), elect.cell(row, 55).value)
        for row in range(28, 32)
    ]))
    fit_total, fit_amount = amount_row(41)
    fit_meters = []
    for row in range(36, 39):
        label = str(elect.cell(row, 1).value or '')
        meters = re.findall(r'\d{7,8}', label)
        terms = [float(part) for part in re.findall(r'[+-]?\d+(?:\.\d+)?', str(elect.cell(row, 55).value))]
        positives = [term for term in terms if term > 0]
        negative = sum(term for term in terms if term < 0)
        for index, meter in enumerate(meters):
            units = positives[index] if index < len(positives) else 0
            if index == 0:
                units += negative
            fit_meters.append((meter, units))
    specs.append(('52167-13569-2', fit_total, fit_amount, fit_meters))
    for row, account in ROW_ACCOUNT.items():
        total, fit = amount_row(row)
        meter = re.findall(r'\d{7,8}', str(elect.cell(row, 1).value or ''))
        units = faces.get(account, {}).get('units') or 0
        specs.append((account, total, fit, [(meter[0], units)] if meter else []))
    specs.append(('82805-94744-7', float(fc_total or 0), 0.0, [('9132330', faces_fc.get('units') or 0)]))

    bills = []
    for account, total, fit, meters in specs:
        face = faces.get(account, {})
        start = face.get('from') or datetime(2025, 1, 1)
        end = face.get('to') or datetime(2025, 1, 31)
        days = face.get('days') or 30
        units = face.get('units') if face.get('units') is not None else sum(float(u) for _, u in meters)
        lines = [
            'Account Number',
            account,
            f'From {_dmy(start)} to {_dmy(end)}',
            f'For {int(days)} days of usage',
            'Total Amount',
            f'${float(total):,.2f}',
            'Fuel Cost Adjustment:',
            f'{int(round(float(units)))} units',
        ]
        for meter, meter_units in meters:
            if meter_units is None:
                continue
            lines.append(_meter_rows(meter, meter_units))
        fit_lines = face.get('fit_lines') or []
        if fit and not fit_lines:
            fit_lines = [('10352137', f'{abs(fit):.2f}')]
        if fit_lines or fit:
            lines.append(f'Feed-in Tariff -${abs(fit):,.2f}' if fit else 'Feed-in Tariff -$0.00')
            for meter, amount in fit_lines:
                lines.append(
                    f'{meter}(FiT)\n{_dmy(start)} {_dmy(end)}\n{int(days)}\n100.00\n1000\n-4.00\n-{float(amount):,.2f}'
                )
        path = directory / f'{account}.pdf'
        _pdf(path, '\n'.join(lines))
        bills.append(parse_clp_bill(str(path)))
    elect.parent.close()
    cost.close()
    return bills


def _may_bills():
    return [parse_clp_bill(str(path)) for path in sorted(MAY_PDFS.glob('*.pdf'))]


def _assert_month(actual_path, citybase_path, previous_citybase, our_previous_path, month):
    actual = openpyxl.load_workbook(actual_path, data_only=False)
    gold = openpyxl.load_workbook(citybase_path, data_only=False)
    previous = openpyxl.load_workbook(previous_citybase, data_only=False)
    carried = openpyxl.load_workbook(our_previous_path, data_only=False)
    month_num = int(month.split('-')[1])
    month_col = month_num + 1
    try:
        for name in HISTORICAL:
            left, right = actual[name], gold[name]
            for row in left.iter_rows():
                for cell in row:
                    assert _norm(cell.value) == _norm(right.cell(cell.row, cell.column).value), name

        for name in gold.sheetnames:
            if name in HISTORICAL:
                continue
            gold_ws = gold[name]
            actual_ws = actual[name]
            prev_ws = previous[name]
            carried_ws = carried[name]
            max_row = max(gold_ws.max_row or 1, actual_ws.max_row or 1)
            max_col = max(gold_ws.max_column or 1, actual_ws.max_column or 1)
            for row in range(1, max_row + 1):
                for col in range(1, max_col + 1):
                    gold_value = gold_ws.cell(row, col).value
                    actual_value = actual_ws.cell(row, col).value
                    previous_value = prev_ws.cell(row, col).value
                    carried_value = carried_ws.cell(row, col).value
                    if name == '01-12 2025':
                        if col == month_col:
                            if isinstance(gold_value, (int, float)):
                                # One rounded sum of the live allocation. Citybase's
                                # typed column can sit a few cents away once the
                                # account-sheet DC plugs are included.
                                assert _close(actual_value, gold_value, tolerance=0.05), (
                                    f'{month} year grid {gold_ws.cell(row, 1).value}: {actual_value} vs {gold_value}'
                                )
                            elif month != '2025-06':
                                assert _norm(actual_value) == _norm(gold_value)
                            continue
                        if col > month_col:
                            continue
                        assert _norm(actual_value) == _norm(carried_value), (
                            f'{month} wiped {name} col {col} row {row}'
                        )
                        continue
                    if _norm(gold_value) == _norm(previous_value):
                        assert _norm(actual_value) == _norm(carried_value), (
                            f'{month} {name} {actual_ws.cell(row, col).coordinate} changed a cell Citybase left alone'
                        )
                        continue
                    if not _bill_cell(gold_value) and not _bill_cell(previous_value):
                        assert _norm(actual_value) == _norm(carried_value)
                        continue
                    assert _close(actual_value, gold_value), (
                        f'{month} {name} {actual_ws.cell(row, col).coordinate}: {actual_value!r} vs {gold_value!r}'
                    )
    finally:
        actual.close()
        gold.close()
        previous.close()
        carried.close()


def test_may_june_july_cost_sheet_matches_citybase(tmp_path):
    april_sheet = SHEETS['2025-04']
    april_hash = hashlib.sha256(april_sheet.read_bytes()).hexdigest()
    masters = tmp_path / '2025-04' / 'masters'
    masters.mkdir(parents=True)
    shutil.copy(april_sheet, masters / 'Cost Sheet-2025-04.xlsx')
    shutil.copy(MASTER, masters / 'Cost Allocation.xlsx')

    months = ('2025-05', '2025-06', '2025-07')
    bills_for = {
        '2025-05': _may_bills(),
        '2025-06': bills_from_workbooks(ALLOC['2025-06'], SHEETS['2025-06'], tmp_path / 'bills-06'),
        '2025-07': bills_from_workbooks(ALLOC['2025-07'], SHEETS['2025-07'], tmp_path / 'bills-07'),
    }
    previous = '2025-04'
    our_previous = masters / 'Cost Sheet-2025-04.xlsx'
    for month in months:
        process_month_end(
            tmp_path,
            month,
            previous,
            bills_for[month],
            {},
            METER_LOGS[month],
            master_path=MASTER,
        )
        produced = tmp_path / month / 'out' / f'Cost Sheet-{month}.xlsx'
        _assert_month(produced, SHEETS[month], SHEETS[previous], our_previous, month)
        our_previous = produced
        previous = month

    # June's own Citybase file leaves the June column blank. July's file
    # fills it. The month being processed is still written, and May stays.
    june = openpyxl.load_workbook(tmp_path / '2025-06' / 'out' / 'Cost Sheet-2025-06.xlsx', data_only=False)
    july_gold = openpyxl.load_workbook(SHEETS['2025-07'], data_only=False)
    try:
        grid = june['01-12 2025']
        gold = july_gold['01-12 2025']
        assert grid.cell(4, 6).value not in (None, 0)
        for row in range(4, 31):
            assert _close(grid.cell(row, 7).value, gold.cell(row, 7).value), grid.cell(row, 1).value
            assert _close(grid.cell(row, 6).value, gold.cell(row, 6).value)
    finally:
        june.close()
        july_gold.close()

    assert hashlib.sha256(april_sheet.read_bytes()).hexdigest() == april_hash
