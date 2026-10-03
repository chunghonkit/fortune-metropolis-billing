"""
Meter units that the check-meter and FiT retail splits read off the bill.

These are parsing fixtures (Rule 4, Rule 10, a stacked 電錶號碼 block, the
April Tower PDF). They are not May Cost Sheet dollars.
"""

import os
import time
from decimal import Decimal
from pathlib import Path

import pytest

from app.citybase_model import CitybaseModel
from app.clp_parser import _extract_meter_consumptions, parse_clp_bill

MASTER = Path('/home/ubuntu/.cursor/projects/workspace/uploads/april-cost-allocation_f85f.xlsx')
TOWER_PDF = '/home/ubuntu/.cursor/projects/workspace/uploads/bill-72399-202504_6706.pdf'

pytestmark = pytest.mark.skipif(not MASTER.exists(), reason='April master not in this workspace')


def _by_meter(text):
    return {meter['meter_no']: meter['consumption'] for meter in _extract_meter_consumptions(text)}


def test_space_separated_register_row_is_not_read_as_meter_ids():
    """
    Fitz often emits the 55861 table without the printed slashes.
    The factor must not become kWh, and the readings must not become meters.
    """
    text = """
    Meter No. / Present Reading / Previous Reading / Multi Factor / Consumption
    9048406 23012214 22948434 1 63780
    9026169 18000000 17900000 1 100000
    9024222 24570099 24439330 1 130769
    9026183 10001000 10000000 1 1000
    Grand Total Units Consumed 331,091
    """
    found = _by_meter(text)
    assert found['9024222'] == 130769
    assert found['9048406'] == 63780
    assert found['9026169'] == 100000
    assert found['9026183'] == 1000
    assert '23012214' not in found
    assert '22948434' not in found
    assert '24570099' not in found
    assert '24439330' not in found
    assert 331091 not in found.values()
    assert 1 not in found.values()


def test_fit_factor_column_is_kept_for_the_retail_split():
    """
    Renewable meters print FiT in the factor column. Dropping them
    overweights the 9133719 / 10220967 line (Commercial) and underweights
    the Hotel/Commercial and Hotel/SA lines.
    """
    text = """
    9133719 / 20000000 / 19911471 / 1 / 88529
    10220967 / 8003190 / 8000000 / 1 / 3190
    9134412 / 3000000 / 2974497 / 1 / 25503
    9115936 / 5004808 / 5000000 / FiT / 4,808
    9144186 / 9000000 / 8956505.50 / 1 / 43494.50
    9144880 7000000 6989517 FiT 10483
    10353005 / 5002141 / 5000000 / 1 / 2141
    """
    found = _by_meter(text)
    assert found['9115936'] == 4808
    assert found['9144880'] == 10483
    assert found['9133719'] == 88529
    assert found['10220967'] == 3190

    bills = [{
        'account': '52167-13569-2',
        'total_amount': 100000,
        'fit_amount': -20000,
        'meters': _extract_meter_consumptions(text),
    }]
    result = CitybaseModel(str(MASTER)).allocate(
        bills, {'meter_no': '6681757', 'previous': 96323.1, 'present': 96504.6}
    )
    centres = {a['centre']: a['amount'] for a in result['allocations']
               if a['account'] == '52167-13569-2'}
    assert abs(sum(centres.values()) - 120000) < 0.05
    # Group 9 (9134412 + 9115936) feeds Hotel / Commercial. Group 8 does not.
    assert centres['Hotel / Commercial'] > 0
    assert centres['Hotel / SA'] > 0
    assert centres['SA / Commercial'] > 0


def test_may_chiller_units_still_set_o7():
    """The deployed May chiller read must stay on the same O7."""
    text = """
    9048406 18000000 17936220 1 63780
    9026169 15000000 14914242 1 85758
    9024222 24570099 24439330 1 130769
    9026183 10050784 10000000 1 50784
    """
    meters = _extract_meter_consumptions(text)
    by_meter = {m['meter_no']: m['consumption'] for m in meters}
    assert by_meter == {
        '9048406': 63780,
        '9026169': 85758,
        '9024222': 130769,
        '9026183': 50784,
    }
    bills = [{
        'account': '55861-52267-1',
        'total_amount': 500000,
        'meters': meters,
    }]
    result = CitybaseModel(str(MASTER)).allocate(
        bills, {'meter_no': '6681757', 'previous': 96323.1, 'present': 96504.6}
    )
    ac = sum(a['amount'] for a in result['allocations']
             if a['account'] == '55861-52267-1' and a['centre'] == 'AC')
    sw = sum(a['amount'] for a in result['allocations']
             if a['account'] == '55861-52267-1' and a['centre'] == 'SW')
    assert abs(sw / (ac + sw) - 0.241094) < 0.00005
    assert abs(ac / (ac + sw) - 0.758906) < 0.00005


def test_decimal_and_fullwidth_slash_rows():
    decimal = _by_meter('9024222 / 24570099.00 / 24439330.00 / 1 / 130769.00')
    assert decimal['9024222'] == 130769
    wide = _by_meter('9024222 ／ 24,570,099 ／ 24,439,330 ／ 1 ／ 130,769')
    assert wide['9024222'] == 130769


def test_slash_meter_row_finishes_and_uses_the_units_field():
    """55861 prints 9024222 / present / previous / factor / units."""
    line = '9024222 / 24570099 / 24439330 / 1 / 130769'
    # The old suffix scan did not return on a long prefix of this bill.
    prefix = ('電錶號碼 FiT 度數 總用電度數 On-Peak 59,910 \n' * 200)
    started = time.monotonic()
    found = _by_meter(prefix + line + '\nGrand Total Units Consumed 431,769\n' + prefix)
    assert time.monotonic() - started < 1.0
    assert found['9024222'] == 130769
    assert '24570099' not in found
    assert '24439330' not in found
    assert 431769 not in found.values()


def test_chiller_slash_rows_do_not_treat_readings_as_meters():
    text = """
    9048406 / 18000000 / 17800000 / 1 / 200000
    9026169 / 15000000 / 14900000 / 1 / 100000
    9024222 / 24570099 / 24439330 / 1 / 130769
    9026183 / 10001000 / 10000000 / 1 / 1000
    Grand Total Units Consumed 431,769
    """
    assert _by_meter(text) == {
        '9048406': 200000,
        '9026169': 100000,
        '9024222': 130769,
        '9026183': 1000,
    }


def test_parsed_slash_units_set_o7():
    text = """
    9048406 / 18000000 / 17800000 / 1 / 200000
    9026169 / 15000000 / 14900000 / 1 / 100000
    9024222 / 24570099 / 24439330 / 1 / 130769
    9026183 / 10001000 / 10000000 / 1 / 1000
    """
    meters = _extract_meter_consumptions(text)
    units = next(m['consumption'] for m in meters if m['meter_no'] == '9024222')
    assert units == 130769
    bills = [{
        'account': '55861-52267-1',
        'total_amount': 500000,
        'meters': meters,
    }]
    meter_log = {'meter_no': '6681757', 'previous': 96323.1, 'present': 96504.6}
    model = CitybaseModel(str(MASTER))
    result = model.allocate(bills, meter_log)
    o7 = next(
        model._percentages(Decimal('181.5'), Decimal(units))[row]
        for row, line in model.alloc_lines.items()
        if line['kind'] == 'check_sw'
    )
    assert abs(float(o7) - (181.5 * 160 / units)) < 1e-12
    assert abs(float(o7) - 0.2221) < 0.0001
    ac = sum(a['amount'] for a in result['allocations']
             if a['account'] == '55861-52267-1' and a['centre'] == 'AC')
    sw = sum(a['amount'] for a in result['allocations']
             if a['account'] == '55861-52267-1' and a['centre'] == 'SW')
    assert result['validation']['errors'] == []
    assert abs(ac + sw - 500000) < 0.05
    # SW is O7 of the 9024222 line plus the whole 9026183 line.
    assert sw > 0 and ac > sw


def test_parse_clp_bill_on_slash_text_returns(tmp_path):
    import fitz

    doc = fitz.open()
    page = doc.new_page()
    page.insert_text(
        (72, 72),
        "55861-52267-1\n"
        "9024222 / 24570099 / 24439330 / 1 / 130769\n"
        "Grand Total Units Consumed 130,769\n",
    )
    path = tmp_path / '55861.pdf'
    doc.save(path)
    doc.close()
    started = time.monotonic()
    parsed = parse_clp_bill(str(path))
    assert time.monotonic() - started < 2
    assert parsed['account'] == '55861-52267-1'
    by_meter = {m['meter_no']: m['consumption'] for m in parsed['meters']}
    assert by_meter['9024222'] == 130769


def test_fit_slash_row_keeps_the_fit_label():
    text = """
    9133719 / 20000000 / 19911471 / 1 / 88529
    9115936 / 5004808 / 5000000 / 1 / FiT 4808
    10220967 / 8003190 / 8000000 / 1 / 3190
    """
    found = _by_meter(text)
    assert found == {'9133719': 88529, '9115936': 4808, '10220967': 3190}


def test_register_line_keeps_factor_and_allows_commas():
    plain = _by_meter('9086205 1 1711927 1721424')
    assert plain == {'9086205': 9497}
    factored = _by_meter('1045662 10 258213 258538')
    assert factored == {'1045662': 3250}
    # The factor must not be taken as the units, and commas are part of the reading.
    comma = _by_meter('9024222 1 1,200,000 1,330,752')
    assert comma == {'9024222': 130752}


def test_rule4_chiller_units_include_9024222():
    text = '9048406 161,494 + 9026169 126,978 + 9024222 103,307 + 9026183 97,544'
    assert _by_meter(text) == {
        '9048406': 161494,
        '9026169': 126978,
        '9024222': 103307,
        '9026183': 97544,
    }


def test_rule10_fit_meters_keep_fit_word_and_eight_digits():
    text = """
    CLP: 9133719 88,529 + 9134412 25,503 + 9144186 43,494.50 = 157,526.50
    RE System: 9115936 FiT 4,808 + 9144880 FiT 10,483 + 10220967 3,190 + 10353005 2,141
    """
    found = _by_meter(text)
    assert found['9133719'] == 88529
    assert found['9134412'] == 25503
    assert found['9144186'] == 43494.5
    assert found['9115936'] == 4808
    assert found['9144880'] == 10483
    assert found['10220967'] == 3190
    assert found['10353005'] == 2141
    # The grand-total figure is not a meter's units.
    assert 157526.5 not in found.values()


def test_columnar_chinese_block_does_not_read_the_next_meter_as_kwh():
    text = """
    電錶號碼
    度數
    總數
    9048406
    9026169
    9024222
    9026183
    80,000.00
    50,000.00
    130,752.00
    1,000.00
    80,000.00
    50,000.00
    130,752.00
    1,000.00
    總用電度數
    261,752.00
    """
    found = _by_meter(text)
    assert found == {
        '9048406': 80000,
        '9026169': 50000,
        '9024222': 130752,
        '9026183': 1000,
    }


def test_tower_pdf_stays_on_9132488_only():
    if not os.path.exists(TOWER_PDF):
        pytest.skip('sample Tower bill not in this workspace')
    parsed = parse_clp_bill(TOWER_PDF)
    assert parsed['account'] == '72399-00664-9'
    assert parsed['kwh'] == 59910
    assert [(m['meter_no'], m['consumption']) for m in parsed['meters']] == [
        ('9132488', 59911)
    ]


def test_parsed_9024222_sets_o7_and_ignores_tower_meter():
    chiller = """
    Account 55861-52267-1
    9048406 80,000 + 9026169 50,000 + 9024222 130,752 + 9026183 1,000
    """
    tower = """
    Account 72399-00664-9
    電錶號碼
    9132488
    59911.00
    總用電度數
    """
    bills = [
        {
            'account': '55861-52267-1',
            'total_amount': 500000,
            'meters': _extract_meter_consumptions(chiller),
        },
        {
            'account': '72399-00664-9',
            'total_amount': 83897,
            'meters': _extract_meter_consumptions(tower),
        },
    ]
    meter_log = {'meter_no': '6681757', 'previous': 96323.1, 'present': 96504.6}
    model = CitybaseModel(str(MASTER))
    result = model.allocate(bills, meter_log)

    o7 = None
    for row, line in model.alloc_lines.items():
        if line['kind'] == 'check_sw':
            o7 = model._percentages(Decimal('181.5'), Decimal(130752))[row]
    assert abs(float(o7) - (181.5 * 160 / 130752)) < 1e-9

    ac = sum(a['amount'] for a in result['allocations']
             if a['account'] == '55861-52267-1' and a['centre'] == 'AC')
    sw = sum(a['amount'] for a in result['allocations']
             if a['account'] == '55861-52267-1' and a['centre'] == 'SW')
    assert abs(ac + sw - 500000) < 0.05
    # Live O7 on 130,752 units is about 0.2221, not April's 0.3406 and not 0.
    assert ac / (ac + sw) > 0.80

    tower_centres = {a['centre'] for a in result['allocations'] if a['account'] == '72399-00664-9'}
    assert 'SA' in tower_centres
    assert abs(sum(a['amount'] for a in result['allocations']
                    if a['account'] == '72399-00664-9') - 83897) < 0.05


def test_fit_retail_split_uses_meter_units_and_net_plus_fit():
    text = """
    9133719 80,000
    9134412 20,000
    9144186 40,000.00
    9115936 FiT 5,000
    9144880 FiT 10,000
    10220967 3,000
    10353005 2,000
    """
    meters = _extract_meter_consumptions(text)
    # base = net due + |FiT|
    bills = [{
        'account': '52167-13569-2',
        'total_amount': 100000,
        'fit_amount': -20000,
        'meters': meters,
    }]
    result = CitybaseModel(str(MASTER)).allocate(
        bills, {'meter_no': '6681757', 'previous': 1, 'present': 1}
    )
    rows = [a for a in result['allocations'] if a['account'] == '52167-13569-2']
    assert abs(sum(a['amount'] for a in rows) - 120000) < 0.05
    centres = {a['centre'] for a in rows}
    assert 'Hotel / Commercial / SA' in centres
    assert 'Hotel / Commercial' in centres
    assert 'C' in centres
    assert 'DC' in centres
    assert all(a['amount'] > 0 for a in rows)


def test_legacy_id_is_not_added_twice_and_direct_accounts_stay_put():
    bills = [
        {
            'account': '55861-52267-1',
            'total_amount': 400000,
            'meters': [
                {'meter_no': '9048406', 'consumption': 80000},
                {'meter_no': '9046064', 'consumption': 80000},
                {'meter_no': '9026169', 'consumption': 50000},
                {'meter_no': '9024222', 'consumption': 130752},
                {'meter_no': '9046787', 'consumption': 130752},
                {'meter_no': '9026183', 'consumption': 1000},
            ],
        },
        {'account': '13639-58422-3', 'total_amount': 10000, 'meters': [
            {'meter_no': '9092771', 'consumption': 9000},
        ]},
        {'account': '82805-94744-7', 'total_amount': 70610, 'meters': []},
    ]
    result = CitybaseModel(str(MASTER)).allocate(
        bills, {'meter_no': '6681757', 'previous': 96323.1, 'present': 96504.6}
    )

    def total(account, centre=None):
        return sum(
            a['amount'] for a in result['allocations']
            if a['account'] == account and (centre is None or a['centre'] == centre)
        )

    assert abs(total('55861-52267-1') - 400000) < 0.05
    assert abs(total('13639-58422-3', 'AO') - 9200) < 0.02
    assert abs(total('13639-58422-3', 'OC') - 800) < 0.02
    assert abs(total('82805-94744-7', 'FC') - 70610) < 0.02
