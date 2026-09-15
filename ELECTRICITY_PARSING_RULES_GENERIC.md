# CLP Electricity Bill Parsing Rules (Generic / Portable)
## 中電電費單解析規則 - 通用版

> **Portable**: Drop this file into ANY project that parses Hong Kong CLP bills. No Fortune Metropolis dependencies.
> **通用**: 可直接放入任何解析香港中電賬單的項目，無需依賴置富都會配置。

**Version**: v2.0 - 2026-09-15
**Validated**: 15 bills, 6 families, 1,332,208 KWH ground truth
**Language**: Python + regex, works with PyMuPDF (fitz)

---

## Quick Start for Other Projects

1. Copy `ELECTRICITY_PARSING_RULES.md` + `clp_parser_v2_final.py` into your project
2. `pip install PyMuPDF openpyxl pandas`
3. ```python
   from clp_parser_v2_final import parse_clp_bill
   result = parse_clp_bill("bill.pdf")
   print(result["kwh"], result["account"], result["bill_type"])
   ```
4. If new format appears, update rules below and re-validate with `fuel_days_sum == days`

No `allocation_config.json` needed for generic parsing. Only KWH / amounts / FiT.

---

## 6 Bill Families (Applies to all CLP commercial accounts)

| Family | How to Detect | KWH Source |
|--------|---------------|------------|
| **Single Rate Non-Residential** | Text contains `Non-Residential Tariff` AND NOT `Bulk`, Rate `@110.6¢` | Page1 bottom-right `Meter No. / Factor / Previous / Present` => `(Present-Previous)*Factor` |
| **Bulk 1M / 4M** | Text contains `Bulk Tariff` | Page3 `Grand Total Units Consumed` = Fuel units. If 4 meters, sum them. |
| **Adjusted Bimonthly** | `Adjusted Bill` + `Bimonthly Reading` + `For 61 days` (or 60-62) | `(Present-Previous)*Factor` + check `Less Electricity Charge -$xxx` in Others |
| **Estimated Bimonthly** | `Estimated Bill` + `Bimonthly Reading` + chart shows `E` + message "could not read the meter" | `Meter No. / Total Consumption` => `Total Consumption` field only |
| **Bulk+FiT** | `Bulk Tariff` + `Feed-in Tariff` + `Renewable Energy (RE) System` | Grand Total = billing units. FiT generation separate from FiT Details table. |

---

## 22 Rules (Copy-Paste Ready)

### 1. Meter Table Regex (Fixes 0 KWH fake)
```python
meter_matches = re.findall(r'(\d{7})\s+(\d+)\s+(\d{6,7})\s+(\d{6,7})', text)
# Each: meter_no, factor, previous, present
# Consumption = (present - previous) * factor
# Filter: 0 < diff < 1_000_000 and factor <= 100
```
**Bug fixed**: Old parsers only looked at Energy Charge line, missed bottom table.

### 2. Multi Factor
`consumption = diff * factor` - e.g., Factor 10: (258538-258213)*10=3250

### 3. Grand Total (Bulk)
`Grand Total Units Consumed` = `Total Units Consumed + Odd Figure Adjustment`
Regex: `Grand Total Units Consumed[^\d]*([\d,\.]+)`
Chinese: `用電度數總計`

### 4. Bulk 4 Meters
Sum all: 161,494+126,978+103,307+97,544=489,322

### 5. Estimated Bill
Detect: `Estimated Bill` in text
KWH: `Total Consumption` field, regex `Total Consumption\s*\n?\s*\d+\s+(\d+)`
No Previous/Present available.

### 6. Adjusted Bill
Detect: `Adjusted Bill`
KWH: Use Rule 2
Others: Has `Less Electricity Charge from xx to yy -$3,464.44`
Fuel: 3 segments (9+30+22=61 days)

### 7. Fuel Days Validation (Critical)
```python
segs = re.findall(r'(\d{2}-\d{2}-\d{2})\s+(\d{2}-\d{2}-\d{2})\s+(\d+)\s+([\d,]+)\s+0\.\d+', text)
assert sum(int(d) for _,_,d,_ in segs) == days, "Fuel days mismatch"
```
Examples: 7+23=30, 5+25=30, 8+23=31, 5+24=29, 9+30+22=61

### 8. Demand Billing vs Actual
Page1 billing kVA (what CLP charges) may differ from Page3 actual max demand.
Min charge: actual 83 but billing 100(min)
Off-Peak 0: actual 253 but billing 0 (contract only charges On-Peak)

### 9. On/Off-Peak %
`On-Peak 62,981 (51.3154%)` => On% + Off% = 100%
Use to validate: `On_units / On% ≈ Grand Total`

### 10. FiT - Grand Total Includes RE
Page3: CLP sum + RE sum = Total Units Consumed -> Odd Adjustment -> Grand Total
Grand Total = fuel billing units

### 11. FiT Generation Separate (Do NOT count as consumption)
```python
fit_gen = sum(int(u.replace(",","")) for u,_,_ in re.findall(r'(\d{1,3}(?:,\d{3})*)\s+-([34])\.00\s+-([\d,]+\.00)', text))
# Rate: $3.00 (150kW+) or $4.00 (<100kW old)
```
Example: 10,483+4,808+2,141+3,190=20,622

### 12. FiT Amount
`Feed-in Tariff (Details Attached) -$72,005.00` in Others
`Total Amount = Energy+Fuel+Others` where Others includes negative FiT

### 13. Less Charge (Adjusted)
`Less Electricity Charge ... -$3,464.44`

### 14. Single Rate Rate
`@110.6¢ = $1.106/kWh`

### 15. Bulk Rates
Energy: On @82.8¢, Off @75.1¢; Demand: On @$74.9/kVA, Off @$26.8/kVA (often 0)

### 16. Fuel Rate Pro-rata
Rates: 40.4¢ (May), 42.6¢ (late June), 43.5¢ (July) - pro-rata by days

### 17. Odd Cents
`Odd Cents Brought Forward/Carried Forward` small but needed for total validation

### 18. Deposit Interest
`Deposit Interest -$xxx` based on deposit $61k-$390k

### 19. Avg Daily Validation
`Avg Daily = KWH / Days` should match chart last bar

### 20. Account Format
`XXXXX-XXXXX-X` e.g., 23529-59279-9

### 21. Date Format
`From 24-06-26 to 23-07-26` DD-MM-YY, `For 30 days`

### 22. Zero Fake Detection
If `kwh == 0 or None` => check Rule 1,2,3,5 - previous bug gave 7 fake zeros

---

## Minimal Parser (Copy to any project)

```python
import fitz, re

def parse_clp_bill_generic(pdf_path):
    txt = "\n".join([p.get_text() for p in fitz.open(pdf_path)]).replace("−","-")
    
    # Account & days
    acct = re.search(r'(\d{5}-\d{5}-\d)', txt)
    acct = acct.group(1) if acct else None
    days_m = re.search(r'For\s+(\d+)\s+days', txt, re.I)
    days = int(days_m.group(1)) if days_m else None
    
    # Flags
    is_bulk = "Bulk Tariff" in txt
    is_adj = "Adjusted Bill" in txt
    is_est = "Estimated Bill" in txt
    is_fit = "Feed-in Tariff" in txt
    
    # KWH - Try Grand Total, then meter diff*factor, then estimated Total Consumption, then fuel fallback
    kwh = None
    m = re.search(r'Grand Total Units Consumed[^\d]*([\d,\.]+)', txt, re.I)
    if m:
        kwh = float(m.group(1).replace(",",""))
    
    meters = re.findall(r'(\d{7})\s+(\d+)\s+(\d{6,7})\s+(\d{6,7})', txt)
    total_meter = 0
    for _, factor, prev, pres in meters:
        try:
            f=int(factor); diff=int(pres)-int(prev)
            if 0 < diff < 1_000_000 and f <=100:
                total_meter += diff*f
        except: pass
    if kwh is None and total_meter>0:
        kwh = total_meter
    
    if is_est:
        m2 = re.search(r'Total Consumption\s*\n?\s*\d+\s+(\d+)', txt, re.I)
        if m2: kwh = int(m2.group(1).replace(",",""))
    
    if kwh is None:
        for pat in [r'Fuel Cost Adjustment:\s+([\d,]+)\s+units', r'Fuel Cost Adjustment:\s+Sub-total \(([\d,]+)\s+units\)']:
            mf = re.search(pat, txt, re.I)
            if mf:
                kwh = int(mf.group(1).replace(",","")); break
    
    # FiT
    fit_gen = sum(int(u.replace(",","")) for u,_,_ in re.findall(r'(\d{1,3}(?:,\d{3})*)\s+-([34])\.00\s+-([\d,]+\.00)', txt))
    m_fit = re.search(r'Feed-in Tariff[^\-]*-\$?([\d,]+\.\d{2})', txt, re.I|re.S)
    fit_amt = -float(m_fit.group(1).replace(",","")) if m_fit else 0
    
    # Total Amount
    m_tot = re.search(r'Total Amount\s+\$?([\d,]+\.\d{2})', txt)
    total_amt = float(m_tot.group(1).replace(",","")) if m_tot else None
    
    return {
        "account": acct,
        "kwh": kwh,
        "days": days,
        "is_bulk": is_bulk,
        "is_adjusted": is_adj,
        "is_estimated": is_est,
        "is_fit": is_fit,
        "fit_generation": fit_gen,
        "fit_amount": fit_amt,
        "total_amount": total_amt,
        "meters_found": len(meters)
    }
```

---

## Testing

```bash
# Put 15 PDFs in ./bills/
python -c "from clp_parser_v2_final import parse_clp_bill; import glob; print([parse_clp_bill(p)['kwh'] for p in glob.glob('bills/*.pdf')])"
# Expected: No zeros, total 1,332,208 for July 2026 sample
```

---

## For GitHub

Add to your repo root:
- `ELECTRICITY_PARSING_RULES.md` (this file)
- `clp_parser_v2_final.py` (full parser)
- `tests/` folder with sample PDFs (optional)

In `README.md` add:
```markdown
## CLP Parsing
See [ELECTRICITY_PARSING_RULES.md](./ELECTRICITY_PARSING_RULES.md) for 22 rules.
Validated on 15 bills, 6 families. Zero fake detection fixed.
```

Future agents: Update this MD when new CLP format appears (e.g., new FiT rate $2.50).

---

## License

MIT - Free to reuse in any project (commercial/residential, HK CLP or other utilities with similar format).

---

*Portable version - No Fortune Metropolis allocation_config.json dependency*
*Full version with allocation: See ELECTRICITY_PARSING_RULES.md in Fortune Metropolis repo*
