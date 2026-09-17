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

import fitz
import re
import json
import os
import glob

def parse_clp_bill(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = "\n".join([page.get_text() for page in doc])
    text = full_text.replace("−", "-").replace("–", "-")

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

    # Dates
    m = re.search(r'From\s+(\d{2}-\d{2}-\d{2})\s+to\s+(\d{2}-\d{2}-\d{2})', text, re.I)
    if m:
        out["from_date"] = m.group(1)
        out["to_date"] = m.group(2)
    m = re.search(r'For\s+(\d+)\s+days', text, re.I)
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

    # Amounts
    m = re.search(r'Total Amount\s+\$?([\d,]+\.\d{2})', text)
    if m: out["total_amount"] = float(m.group(1).replace(",", ""))
    m = re.search(r'Deposit:\s*\$?([\d,]+\.\d{2})', text)
    if m: out["deposit"] = float(m.group(1).replace(",", ""))

    # --- METERS: Rule 3 Multi Factor ---
    # Pattern: MeterNo (7 digits) + Factor + Previous + Present on consecutive lines/spaces
    meter_pattern = re.findall(r'(\d{7})\s+(\d+)\s+(\d{6,7})\s+(\d{6,7})', text)
    total_meter_consumption = 0
    for meter_no, factor_s, prev_s, present_s in meter_pattern:
        try:
            factor = int(factor_s)
            prev = int(prev_s)
            present = int(present_s)
            diff = present - prev
            if diff <= 0 or diff > 1_000_000:  # invalid
                continue
            if factor > 100:  # not a factor
                continue
            cons = diff * factor
            out["meters"].append({
                "meter_no": meter_no,
                "factor": factor,
                "previous": prev,
                "present": present,
                "consumption": cons
            })
            total_meter_consumption += cons
        except:
            continue

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
