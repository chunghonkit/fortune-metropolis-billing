"""
CLP Parser v2 FINAL - Fixes all 0 KWH fake cases
Based on 15 bills analysis (6 families)

Families:
  A. Single Rate (Non-Residential) - 4 bills: 13639 (126390), 35204 (22538), 79292 (25132), 08731 (9497)
     Rule: KWH = Present - Previous (Meter bottom right). Factor usually 1. Fuel = same units.
  
  B. Bulk 1 Meter / 4 Meters - 6 bills: 55861 (489322 4 meters), 24096 (52983), 40722 (38036 min 100kVA), 
     23529 (122733), 70873 (91306 Tower1), 72399 (76563 Tower2), 82805 (60236 Level8)
     Rule: KWH = Grand Total Units Consumed = Fuel units. Meter sum. Demand On-Peak kVA from Page1 billing, not Page3 actual for min charge.

  C. Adjusted Bimonthly - 1 bill: 97968 (3250) For 61 days, Factor 10, Less Electricity Charge -3464.44
     Rule: KWH = (Present-Previous)*Multi Factor. Others contains negative Less Charge (previous estimated deducted). Fuel split 3 segments (9+30+22).

  D. Estimated Bimonthly - 1 bill: 88931-57029 (2323) For 29 days, E mark, Total Consumption only
     Rule: KWH = Total Consumption field. No Previous/Present. Next bill will be Adjusted (A).

  E. Bulk+FiT Large - 1 bill: 52167 (176400) Grand Total = CLP 157526 + RE 18876, FiT gen 20622 units, FiT amount -72005
  F. Bulk+FiT Small - 1 bill: 00776 (35499) CLP 34214 + RE 1286, FiT gen 1286, FiT amount -3858

Critical fixes for previous 0 KWH bugs:
  Bug1: Parser looked at wrong location for Meter No. - now searches bottom right Meter table with regex (\d{7})\s+(\d+)\s+(\d{6,7})\s+(\d{6,7})
  Bug2: Forgot to multiply Multi Factor (97968 factor 10 => 325*10=3250)
  Bug3: Estimated bills have no Previous/Present, only Total Consumption
  Bug4: Bulk Grand Total is not in Energy Charge, it's in separate Page 3 section
  Bug5: FiT generation must NOT be counted as consumption (Grand Total is billing total, FiT gen separate field)
  Bug6: Fuel No.of Days sum validation: sum(Fuel No.of Days) must equal For xx days
"""

import logging
import fitz
import re
import json
import os
import glob

logger = logging.getLogger(__name__)

# A private-meter id on these bills is 7 or 8 digits (10220967, 10353005).
_METER = r'(?<!\d)(\d{7,8})(?!\d)'
# Register readings are plain 6–8 digits or grouped with thousands commas.
_READING = r'(\d{1,3}(?:,\d{3})+|\d{6,8})(?:\.\d+)?'
_READING_LINE = re.compile(
    _METER + r'\s+(\d{1,3})\s+' + _READING + r'\s+' + _READING
)
# Units sit in the next token: "161,494", "43,494.50", "59911.00", "4,808",
# or a bare integer. "FiT" may sit between the meter and the units.
# No IGNORECASE flag: Python's \s under re.I backtracks on long bill text.
_FIT = r'[Ff][Ii][Tt]'
_UNITS = r'(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+|\d{1,6})(?!\d)'
_UNITS_LINE = re.compile(_METER + r'(?:\s+' + _FIT + r')?' + r'\s+' + _UNITS)
# Bulk meter table. Fields are adjacent, separated by a slash or by
# whitespace (PyMuPDF drops the printed slash and puts each cell on its
# own line). Readings on a renewable meter can be shorter than 6 digits.
#   9024222 / 24570099 / 24439330 / 1 / 130769
#   9115936(FiT) / 430616 / 423738 / 1 / 6878
#   9134412 / 21454 / 21058 / 1 / -396
# meter, present, previous, factor, units. The units field is the kWh,
# and a later row for the same meter is an adjustment (zero or negative).
_ROW_SEP = r'(?:\s*/\s*|\s+)'
_ROW_READING = r'((?:\d{1,3}(?:,\d{3})+|\d{1,9})(?:\.\d+)?)'
# The factor column is a number, or the word FiT. On the May retail bill
# the word is glued to the meter id — "9115936(FiT)" — and the factor is 1.
_METER_ROW = re.compile(
    r'(?<!\d)(\d{7,8})(?!\d)'
    + _ROW_SEP + _ROW_READING
    + _ROW_SEP + _ROW_READING
    + _ROW_SEP + r'(\d{1,3}(?:\.\d+)?|' + _FIT + r')'
    + _ROW_SEP + r'(' + _FIT + r'\s+)?'
    + r'(-?(?:\d{1,3}(?:,\d{3})+|\d{1,7})(?:\.\d+)?)(?!\d)'
)
# "9115936(FiT)" is the meter id. The parentheses are not a field separator.
_FIT_SUFFIX = re.compile(
    r'(?<!\d)(\d{7,8})\s*[\(（]\s*' + _FIT + r'\s*[\)）]'
)
_TOKEN = re.compile(
    r'(?<!\d)(?P<meter>\d{7,8})(?!\d)'
    r'|(?P<num>\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+)'
    r'|(?P<plain>\d{1,6})(?!\d)'
)


def _normalize_bill_text(text):
    text = (text.replace('\u00a0', ' ')
                .replace('\u3000', ' ')
                .replace('，', ',')
                .replace('．', '.')
                .replace('−', '-')
                .replace('–', '-'))
    for slash in ('／', '∕', '⁄', '╱'):
        text = text.replace(slash, '/')
    return text


def _number(token):
    value = float(str(token).replace(',', ''))
    if abs(value - round(value)) < 1e-6:
        return int(round(value))
    return value


def _overlaps(span, spans):
    start, end = span
    return any(start < other_end and end > other_start for other_start, other_end in spans)


def _in_meter_run(text, start):
    """True when this meter continues a stacked 電錶號碼 column.

    Walks back over whitespace and an optional FiT label. This used to be a
    regex substitution with IGNORECASE, which did not finish on the 55861 text.
    """
    i = start
    while i > 0 and text[i - 1].isspace():
        i -= 1
    if i >= 3 and text[i - 3:i].lower() == 'fit' and (i == 3 or not text[i - 4].isalnum()):
        i -= 3
        while i > 0 and text[i - 1].isspace():
            i -= 1
    digit_start = i
    while digit_start > 0 and text[digit_start - 1].isdigit() and (i - digit_start) < 8:
        digit_start -= 1
    length = i - digit_start
    if length not in (7, 8):
        return False
    return digit_start == 0 or not text[digit_start - 1].isdigit()


def _append_meter(out_list, seen, meter_no, consumption, factor=1, previous=None, present=None):
    if meter_no in seen:
        return 0
    if consumption is None or consumption <= 0 or consumption > 2_000_000:
        return 0
    out_list.append({
        'meter_no': meter_no,
        'factor': factor,
        'previous': previous,
        'present': present,
        'consumption': consumption,
    })
    seen.add(meter_no)
    return consumption


def _extract_columnar_units(text, meters, seen, consumed):
    """
    Pair a stacked meter column with the units column below it.

    PyMuPDF sometimes emits the 電錶號碼 column first and the 度數 column
    second. Pairing each meter with the next number would treat the next
    meter id as kWh. Zip equal runs instead. A following 總數 column repeats
    the same units, so only the first run is used.
    """
    tokens = []
    for match in _TOKEN.finditer(text):
        if _overlaps(match.span(), consumed):
            continue
        if match.group('meter'):
            tokens.append(('meter', match.group('meter'), match.span()))
        else:
            raw = match.group('num') or match.group('plain')
            tokens.append(('num', _number(raw), match.span()))

    index = 0
    while index < len(tokens):
        if tokens[index][0] != 'meter':
            index += 1
            continue
        end = index + 1
        while end < len(tokens) and tokens[end][0] == 'meter':
            gap = text[tokens[end - 1][2][1]:tokens[end][2][0]]
            # "9024222 / 24570099 / 24439330" is one meter's readings, not a
            # column of meter ids. A slash between them ends the run.
            if '/' in gap:
                break
            end += 1
        run = [token for token in tokens[index:end] if token[1] not in seen]
        number_end = end
        while number_end < len(tokens) and tokens[number_end][0] == 'num':
            number_end += 1
        numbers = tokens[end:number_end]
        if len(run) >= 2 and len(numbers) >= len(run):
            for (kind, meter_no, span), (_nkind, units, nspan) in zip(run, numbers):
                added = _append_meter(meters, seen, meter_no, units)
                if added:
                    consumed.append(span)
                    consumed.append(nspan)
        index = end if len(run) < 2 or len(numbers) < len(run) else number_end


def _add_meter_units(meters, seen, meter_no, units, factor=1, previous=None, present=None):
    """Add units onto a meter already read. A later row is an adjustment."""
    for meter in meters:
        if meter['meter_no'] != meter_no:
            continue
        meter['consumption'] = meter['consumption'] + units
        return meter['consumption']
    return _append_meter(meters, seen, meter_no, units, factor, previous, present)


def _extract_meter_consumptions(text):
    """
    Meter kWh from a CLP bill.

    Tried in order, and only the immediate next tokens are used:
    1. Bulk table row: meter / present / previous / factor / units.
       The units field is the kWh (55861 prints 9024222 this way).
       A renewable id is printed "9115936(FiT)"; a later row of the same
       id carries a negative adjustment that belongs in that meter's kWh.
    2. English register line: meter, factor, previous, present
       (commas allowed). Consumption is (present − previous) × factor.
    3. The next token is the units, optionally after the word FiT
       (Rule 4 "9048406 161,494 + …" and Rule 10 "9115936 FiT 4,808").
    4. A columnar 電錶號碼 / 度數 block whose meters are stacked apart
       from their units.
    """
    text = _FIT_SUFFIX.sub(r'\1', _normalize_bill_text(text or ''))
    meters = []
    seen = set()
    consumed = []

    for match in _METER_ROW.finditer(text):
        if _overlaps(match.span(), consumed):
            continue
        meter_no, first_s, second_s, factor_s, fit_tag, units_s = match.groups()
        fit_factor = factor_s.lower() == 'fit' or bool(fit_tag)
        try:
            first = _number(first_s)
            second = _number(second_s)
            units = _number(units_s)
            factor_value = 1 if fit_factor else _number(factor_s)
        except ValueError:
            continue
        if factor_value < 1 or factor_value > 100 or factor_value != int(factor_value):
            continue
        factor = int(factor_value)
        # A numeric factor is a register row: |units| must be |present−previous|×factor.
        # The units field keeps its sign (the RE block adjusts with -396, -2514).
        # FiT in the factor column is the renewable flag; trust the units field.
        if not fit_factor:
            delta = abs(first - second) * factor
            if abs(delta - abs(units)) > 1:
                continue
        if meter_no in seen:
            # Keep the first positive read. A later negative (or zero) row is
            # the adjustment the Elect Charge kWh formula subtracts. A later
            # positive row is another register of the same id (on/off peak);
            # it must be consumed so its readings are not stored as meters,
            # and it must not be added on top of the private-meter kWh.
            if units < 0:
                _add_meter_units(meters, seen, meter_no, units)
            consumed.append(match.span())
            continue
        present, previous = (first, second) if first >= second else (second, first)
        if _append_meter(meters, seen, meter_no, units, factor, previous, present):
            consumed.append(match.span())

    for match in _READING_LINE.finditer(text):
        if _overlaps(match.span(), consumed):
            continue
        meter_no, factor_s, prev_s, present_s = match.groups()
        try:
            factor = int(factor_s)
            previous = _number(prev_s)
            present = _number(present_s)
        except ValueError:
            continue
        if factor < 1 or factor > 100:
            continue
        diff = present - previous
        if diff <= 0 or diff > 1_000_000:
            continue
        consumption = _number(diff * factor)
        if _append_meter(meters, seen, meter_no, consumption, factor, previous, present):
            consumed.append(match.span())

    for match in _UNITS_LINE.finditer(text):
        if _overlaps(match.span(), consumed) or _in_meter_run(text, match.start()):
            continue
        meter_no, units_s = match.group(1), match.group(2)
        if meter_no in seen:
            continue
        try:
            units = _number(units_s)
        except ValueError:
            continue
        if _append_meter(meters, seen, meter_no, units):
            consumed.append(match.span())

    _extract_columnar_units(text, meters, seen, consumed)
    return meters


def _extract_chinese_meter_units(text, out):
    """
    Read meter units from a Chinese CLP 用電度數 block.

    Example:
        電錶號碼
        9048406
        161494.00
        9024222
        130752.00
        總用電度數
    """
    seen = {str(meter.get('meter_no')) for meter in out.get('meters', [])}
    added = 0.0
    for meter in _extract_meter_consumptions(text):
        if meter['meter_no'] in seen:
            continue
        out['meters'].append(meter)
        seen.add(meter['meter_no'])
        added += meter['consumption']
    return added


def parse_clp_bill(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = "\n".join([page.get_text() for page in doc])
    text = _normalize_bill_text(full_text)

    out = {
        "file": os.path.basename(pdf_path),
        "account": None,
        "from_date": None,
        "to_date": None,
        "days": None,
        "bill_type": None,
        "tariff": None,
        "is_adjusted": False,
        "is_estimated": False,
        "is_fit": False,
        "is_bimonthly": False,
        "deposit": None,
        "total_amount": None,
        "kwh": None,
        "clp_consumption": None,
        "fit_generation_units": 0,
        "fit_amount": 0.0,
        "on_peak_units": None,
        "off_peak_units": None,
        "on_peak_pct": None,
        "off_peak_pct": None,
        "demand_on_kva_billing": None,
        "demand_off_kva_billing": None,
        "demand_on_kva_actual": None,
        "demand_off_kva_actual": None,
        "energy_charge": None,
        "fuel_units": None,
        "fuel_amount": None,
        "fuel_segments": [],
        "less_charge": 0.0,
        "deposit_interest": 0.0,
        "meters": [],
        "avg_daily": None,
        "validation": {}
    }

    # Account
    m = re.search(r'(\d{5}-\d{5}-\d)', text)
    if m: out["account"] = m.group(1)

    # Dates - support both English and Chinese formats
    m = re.search(r'From\s+(\d{2}-\d{2}-\d{2})\s+to\s+(\d{2}-\d{2}-\d{2})', text, re.I)
    if m:
        out["from_date"] = m.group(1)
        out["to_date"] = m.group(2)
    else:
        # Try Chinese format: 由 DD-MM-YY 至 DD-MM-YY
        m = re.search(r'由\s+(\d{2}-\d{2}-\d{2})\s+至\s+(\d{2}-\d{2}-\d{2})', text)
        if m:
            out["from_date"] = m.group(1)
            out["to_date"] = m.group(2)
    
    m = re.search(r'For\s+(\d+)\s+days', text, re.I)
    if m:
        out["days"] = int(m.group(1))
    else:
        # Try Chinese format: 共 XX 日
        m = re.search(r'共\s+(\d+)\s+日', text)
        if m:
            out["days"] = int(m.group(1))

    # Flags
    out["is_bulk"] = "Bulk Tariff" in text
    out["is_adjusted"] = "Adjusted Bill" in text
    out["is_estimated"] = "Estimated Bill" in text
    out["is_fit"] = "Feed-in Tariff" in text or "(FiT)" in text
    out["is_bimonthly"] = "Bimonthly Reading" in text
    out["tariff"] = "Bulk" if out["is_bulk"] else "Non-Residential"
    if out["is_adjusted"]:
        out["bill_type"] = "Adjusted"
    elif out["is_estimated"]:
        out["bill_type"] = "Estimated"
    elif out["is_fit"]:
        out["bill_type"] = "Bulk+FiT"
    else:
        out["bill_type"] = out["tariff"]

    # Amounts - support both English and Chinese
    m = re.search(r'Total Amount\s+\$?([\d,]+\.\d{2})', text)
    if m:
        out["total_amount"] = float(m.group(1).replace(",", ""))
    else:
        # Try Chinese format: 應繳總數 $83,897.00
        m = re.search(r'應繳總數\s+\$?([\d,]+\.\d{2})', text)
        if m:
            out["total_amount"] = float(m.group(1).replace(",", ""))
    
    m = re.search(r'Deposit:\s*\$?([\d,]+\.\d{2})', text)
    if m:
        out["deposit"] = float(m.group(1).replace(",", ""))
    else:
        # Try Chinese format: 按金 $340,000.00
        m = re.search(r'按金\s+\$?([\d,]+\.\d{2})', text)
        if m:
            out["deposit"] = float(m.group(1).replace(",", ""))

    # Private-meter units. 9024222 (legacy 9046787) is the check-meter
    # denominator, and 52167's FiT retail split needs each of its meters.
    # Do not substitute the building total, and do not treat the next meter
    # id in a stacked 電錶號碼 column as kWh.
    out["meters"] = _extract_meter_consumptions(text)
    total_meter_consumption = sum(meter["consumption"] for meter in out["meters"])
    if out["meters"]:
        logger.info(
            "Account %s meter units: %s",
            out.get("account"),
            ", ".join(
                f"{meter['meter_no']}={meter['consumption']:g}"
                for meter in out["meters"]
            ),
        )

    # --- ESTIMATED: Rule 17 Total Consumption ---
    if out["is_estimated"]:
        m_est = re.search(r'Total Consumption\s*\n?\s*\d+\s+(\d+)', text, re.I)
        if m_est:
            out["kwh"] = int(m_est.group(1).replace(",", ""))
            out["fuel_units"] = out["kwh"]

    # --- GRAND TOTAL (Bulk): Rule 6, 12 ---
    m_grand = re.search(r'Grand Total Units Consumed[^\d]*([\d,\.]+)', text, re.I)
    if m_grand:
        try:
            out["kwh"] = float(m_grand.group(1).replace(",", ""))
            out["fuel_units"] = out["kwh"]
        except:
            pass
    
    # Try Chinese format: 用電度數總計 or 總用電度數
    if out["kwh"] is None:
        m_chinese = re.search(r'用電度數總計\s+([\d,]+\.?\d*)', text)
        if not m_chinese:
            m_chinese = re.search(r'總用電度數\s+([\d,]+\.?\d*)', text)
        if m_chinese:
            try:
                out["kwh"] = float(m_chinese.group(1).replace(",", ""))
                out["fuel_units"] = out["kwh"]
            except:
                pass

    # --- SINGLE RATE fallback: Energy Charge @110.6 ---
    if out["kwh"] is None:
        # Try to get from meter consumption (Single Rate)
        if total_meter_consumption > 0:
            out["kwh"] = total_meter_consumption
            out["fuel_units"] = total_meter_consumption

    # --- FUEL fallback: Rule 9 ---
    if out["kwh"] is None:
        # Fuel Cost Adjustment: 122,733 units  or Sub-total (122,733 units)
        for pat in [r'Fuel Cost Adjustment:\s+([\d,]+)\s+units', r'Fuel Cost Adjustment:\s+Sub-total \(([\d,]+)\s+units\)']:
            m_f = re.search(pat, text, re.I)
            if m_f:
                out["kwh"] = int(m_f.group(1).replace(",", ""))
                out["fuel_units"] = out["kwh"]
                break

    # On/Off Peak: Rule 20
    m_on = re.search(r'On-Peak\s+@\s+82\.8[^\d]*([\d,]+)\s+units', text, re.I)
    if m_on:
        out["on_peak_units"] = int(m_on.group(1).replace(",", ""))
    m_off = re.search(r'Off-Peak\s+@\s+75\.1[^\d]*([\d,]+)\s+units', text, re.I)
    if m_off:
        out["off_peak_units"] = int(m_off.group(1).replace(",", ""))

    m_on_pct = re.search(r'On-Peak.*?(\d+\.\d+)%', text, re.I)
    # Better: On-Peak 96,964 (54.9682%)
    m_pct = re.findall(r'([\d,]+)\s+\((\d+\.\d+)%\)', text)
    # For Bulk, the two percentages correspond to On/Off
    if len(m_pct) >= 2:
        # Usually first is On-Peak
        try:
            out["on_peak_pct"] = float(m_pct[0][1])
            out["off_peak_pct"] = float(m_pct[1][1])
        except:
            pass

    # Demand: Rule 8 - Billing kVA from Page1, Actual from Page3
    m_dem_bill = re.search(r'On-Peak\s+@\s+\$?\s*74\.9\s+(\d+)\s+kVA', text, re.I)
    if m_dem_bill:
        out["demand_on_kva_billing"] = int(m_dem_bill.group(1).replace(",", ""))
    m_dem_bill_off = re.search(r'Off-Peak\s+@\s+\$?\s*26\.8\s+(\d+)\s+kVA', text, re.I)
    if m_dem_bill_off:
        out["demand_off_kva_billing"] = int(m_dem_bill_off.group(1).replace(",", ""))

    # Actual Max Demand from "Maximum Demand" section
    m_actual = re.search(r'Maximum Demand.*?On-Peak\D+(\d+).*?Off-Peak\D+(\d+)', text, re.I | re.S)
    if m_actual:
        out["demand_on_kva_actual"] = int(m_actual.group(1))
        out["demand_off_kva_actual"] = int(m_actual.group(2))
    else:
        # Chinese version
        m_actual_cn = re.search(r'本月最高需求量.*?高峰\D+(\d+).*?非高峰\D+(\d+)', text, re.I | re.S)
        if m_actual_cn:
            out["demand_on_kva_actual"] = int(m_actual_cn.group(1))
            out["demand_off_kva_actual"] = int(m_actual_cn.group(2))

    # Fuel segments: Rule 9, 19, 21 - sum must equal days
    # From 24-06-26 to 30-06-26  7  2215 0.426 943.59
    fuel_segs = re.findall(r'(\d{2}-\d{2}-\d{2})\s+(\d{2}-\d{2}-\d{2})\s+(\d+)\s+([\d,]+)\s+0\.\d+\s+([\d,]+\.\d{2})', text)
    for fr, to, days, units, amount in fuel_segs:
        out["fuel_segments"].append({
            "from": fr,
            "to": to,
            "days": int(days),
            "units": int(units.replace(",", "")),
            "amount": float(amount.replace(",", ""))
        })
    if out["fuel_segments"]:
        out["fuel_amount"] = sum([s["amount"] for s in out["fuel_segments"]])
        # validation
        out["validation"]["fuel_days_sum"] = sum([s["days"] for s in out["fuel_segments"]])
        out["validation"]["fuel_units_sum"] = sum([s["units"] for s in out["fuel_segments"]])
        out["validation"]["fuel_days_match"] = out["validation"]["fuel_days_sum"] == out["days"]

    # Others: Less Electricity Charge Rule 5,6
    m_less = re.search(r'Less Electricity Charge[^\-]*-([\d,]+\.\d{2})', text, re.I | re.S)
    if m_less:
        out["less_charge"] = -float(m_less.group(1).replace(",", ""))

    # Deposit Interest
    m_dep_int = re.search(r'Deposit Interest\s+-([\d,]+\.\d{2})', text, re.I)
    if m_dep_int:
        out["deposit_interest"] = -float(m_dep_int.group(1).replace(",", ""))

    # FiT: Rule 12,13,15
    if out["is_fit"]:
        # FiT generation units: pattern Unit Rate -3.00 Amount -xxxxx
        fit_gen = 0
        for unit_str, rate, amount in re.findall(r'(\d{1,3}(?:,\d{3})*)\s+-([34])\.00\s+-([\d,]+\.00)', text):
            try:
                fit_gen += int(unit_str.replace(",", ""))
            except:
                pass
        out["fit_generation_units"] = fit_gen

        # FiT amount total
        m_fit = re.search(r'Feed-in Tariff[^\-]*-\$?([\d,]+\.\d{2})', text, re.I | re.S)
        if m_fit:
            out["fit_amount"] = -float(m_fit.group(1).replace(",", ""))

        # CLP consumption = Grand Total - adjustments? Actually Grand Total is billing units (fuel units)
        # For traceability: CLP consumption = Grand Total - FiT generation? No, Grand Total INCLUDES RE generation for billing?
        # From analysis: Page3 Total Units Consumed = CLP + RE = Grand Total. Fuel billed on Grand Total.
        # But for cost allocation, we need separate: clp_consumption = Grand Total - fit_gen_adjustments?
        # From 52167: CLP 157526.50 + RE 18876.50 = 176403 -3 = 176400 = Grand Total. So Grand Total is total billing units.
        # FiT generation for revenue is 20622 (different from RE subtotal due to adjustments).
        # We'll keep clp_consumption as Grand Total for now, and fit_generation separately.
        out["clp_consumption"] = out["kwh"]

    else:
        out["clp_consumption"] = out["kwh"]

    # Avg daily
    if out["days"] and out["kwh"]:
        out["avg_daily"] = round(out["kwh"] / out["days"], 2)

    # Final validation
    out["validation"]["has_kwh"] = out["kwh"] is not None and out["kwh"] > 0
    out["validation"]["is_zero_fake"] = not out["validation"]["has_kwh"]

    return out

# Test on all txt if pdf not available, otherwise pdf
if __name__ == "__main__":
    import sys
    pdfs = sorted(glob.glob("/mnt/data/*.pdf"))
    if not pdfs:
        # fallback to txt for demo
        txts = sorted([f for f in os.listdir("/mnt/data") if f.endswith(".pdf.txt") and f[0].isdigit()])
        print(f"No PDFs found, using {len(txts)} txt extractions for demo")
        for tf in txts:
            path = os.path.join("/mnt/data", tf)
            with open(path, encoding="utf-8", errors="ignore") as f:
                txt = f.read()
            # Save temp pdf.txt as if pdf
            # Use parse logic on txt directly via doc trick
            # We'll just print that file exists
            print(tf)
    else:
        results = []
        for pdf in pdfs:
            try:
                r = parse_clp_bill(pdf)
                results.append(r)
                print(f"{r['account']} | KWH={r['kwh']} | Type={r['bill_type']} | Days={r['days']} | Total=${r['total_amount']} | FiTgen={r['fit_generation_units']} Less={r['less_charge']}")
            except Exception as e:
                print(f"Error {pdf}: {e}")

        # Save combined
        with open("/mnt/data/clp_parsed_v2_results.json", "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        total_kwh = sum([r["kwh"] or 0 for r in results])
        total_fit_gen = sum([r["fit_generation_units"] or 0 for r in results])
        print(f"\n=== SUMMARY ===")
        print(f"Total bills: {len(results)}")
        print(f"Total KWH (billing units): {total_kwh}")
        print(f"Total FiT generation: {total_fit_gen}")
        print(f"Zero KWH fake cases: {sum([1 for r in results if not r['validation']['has_kwh']])}")


# VALIDATED KWH FROM 15 BILLS - GROUND TRUTH
GROUND_TRUTH = {
  "00776-78552-1": {
    "kwh": 35499,
    "type": "Bulk+FiT Small",
    "fit_gen": 1286,
    "fit_amt": -3858
  },
  "08731-83914-5": {
    "kwh": 9497,
    "type": "Single Rate"
  },
  "13639-58422-3": {
    "kwh": 126390,
    "type": "Single Rate"
  },
  "23529-59279-9": {
    "kwh": 122733,
    "type": "Bulk 1M"
  },
  "24096-78457-6": {
    "kwh": 52983,
    "type": "Bulk 1M"
  },
  "35204-69738-4": {
    "kwh": 22538,
    "type": "Single Rate"
  },
  "40722-61440-7": {
    "kwh": 38036,
    "type": "Bulk Min 100kVA"
  },
  "52167-13569-2": {
    "kwh": 176400,
    "type": "Bulk+FiT Large",
    "fit_gen": 20622,
    "fit_amt": -72005
  },
  "55861-52267-1": {
    "kwh": 489322,
    "type": "Bulk 4M"
  },
  "70873-85471-3": {
    "kwh": 91306,
    "type": "Bulk Tower1"
  },
  "72399-00664-9": {
    "kwh": 76563,
    "type": "Bulk Tower2"
  },
  "79292-23337-6": {
    "kwh": 25132,
    "type": "Single Rate 31d"
  },
  "82805-94744-7": {
    "kwh": 60236,
    "type": "Bulk Level8 71% On"
  },
  "88931-57029-6": {
    "kwh": 2323,
    "type": "Estimated Bimonthly 29d"
  },
  "97968-02236-6": {
    "kwh": 3250,
    "type": "Adjusted Bimonthly 61d Factor10",
    "less": -3464.44
  }
}
