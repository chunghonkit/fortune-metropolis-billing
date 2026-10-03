"""
Elect Charge ledger roll-forward.

May, June, and July are checked against the Citybase workbooks when those
workbooks and that month's bill PDFs are available. The PDFs are parsed;
the golden charge numbers are not written in by hand.
"""

import re
import shutil
from datetime import datetime
from pathlib import Path

import fitz
import openpyxl
import pytest

from app.citybase_model import CitybaseModel, canonical_centre_key
from app.clp_parser import parse_clp_bill
from app.workbook_updater import WorkbookUpdater, process_month_end

UPLOADS = Path('/home/ubuntu/.cursor/projects/workspace/uploads')
REFERENCES = {
    '2025-04': UPLOADS / 'Cost_Allocation_e50c.xlsx',
    '2025-05': UPLOADS / 'Cost_Allocation_e28d.xlsx',
    '2025-06': UPLOADS / 'Cost_Allocation_c5a4.xlsx',
    '2025-07': UPLOADS / 'Cost_Allocation_0165.xlsx',
}
# Real CLP PDFs, one directory per billing month. Not committed.
BILL_DIRS = {
    '2025-05': [Path('/tmp/maypdfs'), UPLOADS / 'bills' / '2025-05'],
    '2025-06': [Path('/tmp/junepdfs'), UPLOADS / 'bills' / '2025-06'],
    '2025-07': [Path('/tmp/julypdfs'), UPLOADS / 'bills' / '2025-07'],
}
METER_LOGS = {
    '2025-05': {'meter_no': '6681757', 'previous': 96323.1, 'present': 96504.6},
    '2025-06': {'meter_no': '6681757', 'previous': 96504.6, 'present': 96870.4},
    '2025-07': {'meter_no': '6681757', 'previous': 96870.4, 'present': 97266.7},
}
CHAIN = ('2025-05', '2025-06', '2025-07')


def _bill_dir(month):
    for path in BILL_DIRS[month]:
        if path.is_dir() and any(path.glob('*.pdf')):
            return path
    return None


def _money(value):
    number = float(value)
    if abs(number - round(number)) < 1e-6:
        number = int(round(number))
    return f'{number:,.2f}' if isinstance(number, float) else f'{number:,.2f}'


def _signed_terms(formula):
    return [float(part) if '.' in part else int(part)
            for part in re.findall(r'[+-]?\d+(?:\.\d+)?', formula.replace(' ', '')[1:])]


def _meter_row(meter, units):
    units = int(units) if float(units) == int(float(units)) else units
    if units < 0:
        delta = abs(int(units))
        return f'{meter}\n{100000 + delta}\n100000\n1\n{units}'
    delta = int(units) if float(units) == int(float(units)) else units
    # Readings stay long enough for the register row, including a 0-unit adjustment.
    present = 5000000 + int(delta)
    previous = 5000000
    shown = int(delta) if float(delta) == int(float(delta)) else delta
    return f'{meter}\n{present}\n{previous}\n1\n{shown}'


def _write_pdf(path, text):
    doc = fitz.open()
    page = doc.new_page(width=800, height=1400)
    page.insert_text((36, 36), text, fontsize=8)
    doc.save(path)
    doc.close()


def _bills_from_citybase_sheet(reference, dest):
    """
    Bill PDFs in the May text shape, using the meter units and amounts that
    month's Elect Charge column records. Used when the CLP zip is not on disk.
    May itself uses the real PDFs.
    """
    ws = openpyxl.load_workbook(reference)['Elect Charge']
    dest.mkdir(parents=True, exist_ok=True)
    serial = 1

    def account():
        nonlocal serial
        serial += 1
        return f'{serial:05d}-10000-1'

    def add(name, body):
        _write_pdf(dest / name, body)

    # Chiller: kWh on 28-31, amount on 34.
    lines = []
    for row in range(28, 32):
        lines.append(_meter_row(ws.cell(row=row, column=1).value, ws.cell(row=row, column=55).value))
    amount = ws.cell(row=34, column=55).value
    add('chiller.pdf', f'Account Number\n{account()}\nTotal Amount\n${_money(amount)}\n' + '\n'.join(lines))

    # FiT retail: component formula on 36-38, net + |FiT| on 41.
    fit_lines = []
    for row in (36, 37, 38):
        meters = re.findall(r'\d{7,8}', str(ws.cell(row=row, column=1).value))
        terms = _signed_terms(str(ws.cell(row=row, column=55).value))
        for meter, units in zip(meters, terms):
            fit_lines.append(_meter_row(meter, units))
        if len(terms) > len(meters):
            fit_lines.append(_meter_row(meters[0], terms[len(meters)]))
    net, fit = _signed_terms(str(ws.cell(row=41, column=55).value))
    add(
        'fit-retail.pdf',
        f'Account Number\n{account()}\nTotal Amount\n${_money(net)}\n'
        f'Feed-in Tariff (Details Attached)\n-${_money(abs(fit))}\n'
        + '\n'.join(fit_lines),
    )

    # 00776 is the arithmetic charge row whose label lists its meters.
    small_meters = re.findall(r'\d{7,8}', str(ws.cell(row=14, column=1).value))
    small_net, small_fit = _signed_terms(str(ws.cell(row=14, column=55).value))
    small_lines = [_meter_row(meter, 1000 + index) for index, meter in enumerate(small_meters)]
    add(
        'fit-small.pdf',
        f'Account Number\n{account()}\nTotal Amount\n${_money(small_net)}\n'
        f'Feed-in Tariff (Details Attached)\n-${_money(abs(small_fit))}\n'
        + '\n'.join(small_lines),
    )

    for row in list(range(11, 13)) + list(range(15, 21)) + list(range(22, 25)):
        meter = ws.cell(row=row, column=1).value
        total = ws.cell(row=row, column=55).value
        if not isinstance(total, (int, float)):
            continue
        add(
            f'row-{row}.pdf',
            f'Account Number\n{account()}\nTotal Amount\n${_money(total)}\n' + _meter_row(meter, 100),
        )
    ws.parent.close()
    return [parse_clp_bill(str(path)) for path in sorted(dest.glob('*.pdf'))]


def _parse_bills(month, work):
    folder = _bill_dir(month)
    if folder is not None:
        return [parse_clp_bill(str(path)) for path in sorted(folder.glob('*.pdf'))]
    # June and July CLP zips are larger than the Drive download cap. Rebuild
    # bill text in the May layout from the units that month's sheet records.
    if month == '2025-05':
        return None
    return _bills_from_citybase_sheet(REFERENCES[month], work / f'bills-{month}')


def _same(got, expected):
    if isinstance(got, datetime) or isinstance(expected, datetime):
        got_day = got.date() if isinstance(got, datetime) else got
        expected_day = expected.date() if isinstance(expected, datetime) else expected
        return got_day == expected_day
    if got == expected:
        return True
    try:
        return abs(float(got) - float(expected)) < 1e-6
    except (TypeError, ValueError):
        return False


def _assert_elect_charge(produced, reference):
    got_book = openpyxl.load_workbook(produced)
    expected_book = openpyxl.load_workbook(reference)
    got = got_book['Elect Charge']
    expected = expected_book['Elect Charge']
    mismatches = []
    for row in range(4, 42):
        for col, name in ((54, 'BB'), (55, 'BC')):
            g = got.cell(row=row, column=col).value
            e = expected.cell(row=row, column=col).value
            if not _same(g, e):
                mismatches.append(f'{name}{row}: got {g!r} expected {e!r}')
    for cell in ('BR31', 'BS31', 'BQ31', 'BP31'):
        if not _same(got[cell].value, expected[cell].value):
            mismatches.append(f'{cell}: got {got[cell].value!r} expected {expected[cell].value!r}')
    got_book.close()
    expected_book.close()
    assert not mismatches, '\n'.join(mismatches)


@pytest.fixture(scope='module')
def citybase_ready():
    if not REFERENCES['2025-04'].exists():
        pytest.skip('Citybase Cost Allocation workbooks are not in this workspace')
    april_bytes = REFERENCES['2025-04'].read_bytes()
    yield april_bytes
    assert REFERENCES['2025-04'].read_bytes() == april_bytes


def test_month_chain_fills_elect_charge_without_touching_april(citybase_ready, tmp_path):
    """May copies April. Each later month copies the previous out/, not its own."""
    root = tmp_path / 'Metropolis'
    april_dir = root / '2025-04'
    april_dir.mkdir(parents=True)
    source = april_dir / 'Cost Allocation - Electricity 2007-2.xlsx'
    shutil.copy2(REFERENCES['2025-04'], source)
    source_bytes = source.read_bytes()

    # A Cost Sheet has to be copied alongside. Dollars are checked separately.
    sheet = openpyxl.Workbook()
    grid = sheet.active
    grid.title = '01-12 2025'
    grid['A4'] = 'FC'
    sheet.save(april_dir / 'Cost Sheet-2025-04.xlsx')
    sheet.close()

    updater = WorkbookUpdater(root)
    previous = '2025-04'
    ran = []
    for month in CHAIN:
        bills = _parse_bills(month, tmp_path)
        if bills is None:
            break
        copied = updater.copy_previous_month_workbooks(month, previous)
        assert (root / month / 'out').resolve() != source.resolve()
        assert source.read_bytes() == source_bytes
        updater.update_cost_allocation_workbook(
            copied['cost_allocation'],
            [],
            month,
            bills,
            METER_LOGS[month],
        )
        _assert_elect_charge(copied['cost_allocation'], REFERENCES[month])
        ran.append(month)
        previous = month

    assert '2025-05' in ran, 'May bills were not found, so Elect Charge was not checked'
    # June and July run in the same chain when their PDFs are on disk.
    # A missing directory stops the chain; it does not undo the months already matched.


def test_elect_charge_fill_does_not_change_cost_sheet_dollars(citybase_ready, tmp_path):
    bills = _parse_bills('2025-05', tmp_path)
    if not bills:
        pytest.skip('May bill PDFs are not in this workspace')
    meter_log = METER_LOGS['2025-05']
    model = CitybaseModel(str(REFERENCES['2025-04']))
    allocated = model.allocate(bills, meter_log)
    totals = {}
    for row in allocated['allocations']:
        key = canonical_centre_key(row['centre'])
        totals[key] = totals.get(key, 0) + row['amount']
    totals = {key: round(amount, 2) for key, amount in totals.items()}

    root = tmp_path / 'Metropolis'
    april_dir = root / '2025-04'
    april_dir.mkdir(parents=True)
    master = april_dir / 'Cost Allocation - Electricity 2007-2.xlsx'
    shutil.copy2(REFERENCES['2025-04'], master)
    before = master.read_bytes()

    book = openpyxl.Workbook()
    grid = book.active
    grid.title = '01-12 2025'
    for index, centre in enumerate(sorted(totals), start=4):
        grid.cell(row=index, column=1).value = centre
    book.save(april_dir / 'Cost Sheet-2025-04.xlsx')
    book.close()

    written = process_month_end(
        root,
        '2025-05',
        '2025-04',
        bills,
        {},
        meter_log,
        master,
    )
    assert master.read_bytes() == before
    assert REFERENCES['2025-04'].read_bytes() == citybase_ready

    saved_book = openpyxl.load_workbook(written['cost_sheet'])
    saved = saved_book['01-12 2025']
    found = {}
    for row in saved.iter_rows(min_row=1, max_row=80, min_col=1, max_col=1):
        label = row[0].value
        if not label:
            continue
        key = canonical_centre_key(str(label))
        if key in totals:
            found[key] = saved.cell(row=row[0].row, column=6).value
    saved_book.close()
    assert found == totals
