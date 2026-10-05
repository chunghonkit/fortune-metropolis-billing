"""
Month-end figures for the remote dashboard.

Consumption is the parsed bill kWh. Centre dollars are the allocation
result: each line is that account's allocation base times the live
percent from that month's master. Elect Charge ledger cells are not
read back out of the workbook.
"""

import json
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional

DASHBOARD_FILE = 'dashboard.json'
_MONTH = re.compile(r'^\d{4}-(0[1-9]|1[0-2])$')
_KINDS = {
    'cost-allocation': 'cost_allocation',
    'cost-sheet': 'cost_sheet',
}
_GLOBS = {
    'cost_allocation': ('*Cost Allocation*.xlsx', '*Cost Allocation*.xls'),
    'cost_sheet': ('*Cost Sheet*.xlsx', '*Cost Sheet*.xls'),
}


def valid_month(month: str) -> bool:
    return bool(month and _MONTH.fullmatch(month))


def allocation_base(total_amount, fit_amount) -> float:
    """Same base the Citybase model uses: net due plus absolute FiT."""
    total = float(total_amount or 0)
    fit = float(fit_amount or 0)
    if fit:
        total += abs(fit)
    return total


def accounts_from_bills(parsed_bills: List[Dict]) -> List[Dict]:
    merged: Dict[str, Dict] = {}
    order: List[str] = []
    for bill in parsed_bills or []:
        account = bill.get('account')
        if not account:
            continue
        if account not in merged:
            merged[account] = {
                'account': account,
                'kwh': 0.0,
                'total_amount': 0.0,
                'fit_amount': 0.0,
            }
            order.append(account)
        row = merged[account]
        row['kwh'] += float(bill.get('kwh') or 0)
        row['total_amount'] += float(bill.get('total_amount') or 0)
        row['fit_amount'] += float(bill.get('fit_amount') or 0)
    rows = []
    for account in order:
        row = merged[account]
        row['allocation_base'] = allocation_base(row['total_amount'], row['fit_amount'])
        rows.append(row)
    return rows


def build_month_report(
    billing_month: str,
    parsed_bills: List[Dict],
    allocation_result: Dict,
    files: Optional[Dict] = None,
) -> Dict:
    allocations = []
    for row in (allocation_result or {}).get('allocations') or []:
        allocations.append({
            'account': row.get('account') or '',
            'centre': row.get('centre') or '',
            'percentage': float(row.get('percentage') or 0),
            'amount': float(row.get('amount') or 0),
            'allocation_base': float(
                row.get('allocation_base') if row.get('allocation_base') is not None
                else row.get('bill_total') or 0
            ),
        })
    stored_files = {}
    for key, path in (files or {}).items():
        if path:
            stored_files[key] = Path(path).name
    summary = {}
    for centre, bucket in ((allocation_result or {}).get('summary') or {}).items():
        summary[centre] = {
            'total_amount': float(bucket.get('total_amount') or 0),
            'accounts': list(bucket.get('accounts') or []),
        }
    return {
        'billing_month': billing_month,
        'accounts': accounts_from_bills(parsed_bills),
        'allocations': allocations,
        'summary': summary,
        'files': stored_files,
    }


def save_month_report(
    metropolis_root: Path,
    billing_month: str,
    parsed_bills: List[Dict],
    allocation_result: Dict,
    files: Optional[Dict] = None,
) -> Path:
    if not valid_month(billing_month):
        raise ValueError(f'billing month must be YYYY-MM, got {billing_month!r}')
    out_dir = Path(metropolis_root) / billing_month / 'out'
    out_dir.mkdir(parents=True, exist_ok=True)
    report = build_month_report(billing_month, parsed_bills, allocation_result, files)
    path = out_dir / DASHBOARD_FILE
    path.write_text(json.dumps(report, indent=2), encoding='utf-8')
    from app.consumption_history import upsert_history
    upsert_history(metropolis_root, billing_month, parsed_bills)
    return path


def load_month_report(metropolis_root: Path, billing_month: str) -> Optional[Dict]:
    if not valid_month(billing_month):
        return None
    path = Path(metropolis_root) / billing_month / 'out' / DASHBOARD_FILE
    if not path.is_file():
        return None
    return json.loads(path.read_text(encoding='utf-8'))


def load_all_reports(metropolis_root: Path) -> List[Dict]:
    root = Path(metropolis_root)
    if not root.is_dir():
        return []
    reports = []
    for child in sorted(root.iterdir()):
        if not child.is_dir() or not valid_month(child.name):
            continue
        report = load_month_report(root, child.name)
        if report:
            reports.append(report)
    return reports


def _out_dir(metropolis_root: Path, month: str) -> Path:
    if not valid_month(month):
        raise ValueError('month must be YYYY-MM')
    return Path(metropolis_root) / month / 'out'


def _inside(out_dir: Path, path: Path) -> bool:
    try:
        path.resolve().relative_to(out_dir.resolve())
    except ValueError:
        return False
    return path.is_file()


def find_output_workbook(metropolis_root: Path, month: str, kind: str) -> Optional[Path]:
    """
    Locate one workbook under {month}/out.

    kind is the URL slug cost-allocation or cost-sheet. A name stored in
    dashboard.json wins. Otherwise the first xlsx match, then xls.
    """
    file_key = _KINDS.get(kind)
    if not file_key:
        raise ValueError('kind must be cost-allocation or cost-sheet')
    out_dir = _out_dir(metropolis_root, month)
    report = load_month_report(metropolis_root, month)
    recorded = ((report or {}).get('files') or {}).get(file_key)
    if recorded:
        candidate = out_dir / Path(recorded).name
        if _inside(out_dir, candidate):
            return candidate
    matches = []
    if out_dir.is_dir():
        for pattern in _GLOBS[file_key]:
            matches.extend(path for path in out_dir.glob(pattern) if path.is_file())
    xlsx = [path for path in matches if path.suffix.lower() == '.xlsx']
    chosen = sorted(xlsx or matches)
    return chosen[0] if chosen else None


def list_processed_months(metropolis_root: Path) -> List[Dict]:
    root = Path(metropolis_root)
    if not root.is_dir():
        return []
    months = []
    for child in sorted(child for child in root.iterdir() if child.is_dir() and valid_month(child.name)):
        report = load_month_report(root, child.name)
        allocation = find_output_workbook(root, child.name, 'cost-allocation')
        sheet = find_output_workbook(root, child.name, 'cost-sheet')
        if not report and not allocation and not sheet:
            continue
        months.append({
            'month': child.name,
            'has_report': report is not None,
            'cost_allocation': allocation is not None,
            'cost_sheet': sheet is not None,
        })
    return months


def _account_kwh(report: Dict) -> Dict[str, float]:
    return {row['account']: float(row.get('kwh') or 0) for row in report.get('accounts') or []}


def _matching_lines(report: Dict, account: Optional[str], centre: Optional[str]) -> List[Dict]:
    lines = []
    for line in report.get('allocations') or []:
        if account and line.get('account') != account:
            continue
        if centre and line.get('centre') != centre:
            continue
        lines.append(line)
    return lines


def _month_kwh(report: Dict, account: Optional[str], centre: Optional[str]) -> float:
    """
    Bill kWh when no centre is selected. With a centre, attribute each
    account's kWh by that line's live percent. No meter charge is invented.
    """
    kwh_by_account = _account_kwh(report)
    if not centre:
        rows = report.get('accounts') or []
        if account:
            rows = [row for row in rows if row.get('account') == account]
        return sum(float(row.get('kwh') or 0) for row in rows)
    total = 0.0
    for line in _matching_lines(report, account, centre):
        total += kwh_by_account.get(line.get('account'), 0.0) * float(line.get('percentage') or 0)
    return total


def _month_cost(report: Dict, account: Optional[str], centre: Optional[str]) -> float:
    return sum(float(line.get('amount') or 0) for line in _matching_lines(report, account, centre))


def chart_series(
    reports: Iterable[Dict],
    months: Optional[Iterable[str]] = None,
    account: Optional[str] = None,
    centre: Optional[str] = None,
) -> Dict:
    """
    Series for the dashboard charts.

    kwh and cost are parallel to months. by_centre and by_account cover the
    same filter. Rows are the allocation lines that produced the cost.
    """
    chosen = set(months) if months else None
    selected = []
    for report in reports:
        month = report.get('billing_month')
        if chosen is not None and month not in chosen:
            continue
        selected.append(report)
    selected.sort(key=lambda report: report.get('billing_month') or '')

    month_labels = []
    kwh_series = []
    cost_series = []
    bill_rows = []
    centre_totals: Dict[str, Dict[str, float]] = {}
    account_totals: Dict[str, Dict[str, float]] = {}
    rows = []

    for report in selected:
        month = report['billing_month']
        month_labels.append(month)
        kwh_series.append(_month_kwh(report, account, centre))
        cost_series.append(_month_cost(report, account, centre))
        for row in report.get('accounts') or []:
            if account and row.get('account') != account:
                continue
            bill_rows.append({
                'month': month,
                'account': row.get('account'),
                'kwh': float(row.get('kwh') or 0),
                'total_amount': float(row.get('total_amount') or 0),
                'fit_amount': float(row.get('fit_amount') or 0),
                'allocation_base': float(row.get('allocation_base') or 0),
            })
        kwh_by_account = _account_kwh(report)
        lines = _matching_lines(report, account, centre)
        for line in lines:
            line_account = line.get('account') or ''
            line_centre = line.get('centre') or ''
            percentage = float(line.get('percentage') or 0)
            amount = float(line.get('amount') or 0)
            attributed = kwh_by_account.get(line_account, 0.0) * percentage
            rows.append({
                'month': month,
                'account': line_account,
                'centre': line_centre,
                'kwh': kwh_by_account.get(line_account, 0.0),
                'attributed_kwh': attributed,
                'percentage': percentage,
                'amount': amount,
                'allocation_base': float(line.get('allocation_base') or 0),
            })
            centre_bucket = centre_totals.setdefault(line_centre, {'kwh': 0.0, 'cost': 0.0})
            centre_bucket['kwh'] += attributed
            centre_bucket['cost'] += amount
            account_bucket = account_totals.setdefault(line_account, {'kwh': 0.0, 'cost': 0.0})
            account_bucket['cost'] += amount
            if centre:
                account_bucket['kwh'] += attributed
        if not centre:
            # Bill kWh once per account. Sharing it across centres would double-count.
            seen_accounts = {line.get('account') or '' for line in lines}
            if account and account not in seen_accounts:
                seen_accounts.add(account)
                account_totals.setdefault(account, {'kwh': 0.0, 'cost': 0.0})
            for line_account in seen_accounts:
                account_totals.setdefault(line_account, {'kwh': 0.0, 'cost': 0.0})
                account_totals[line_account]['kwh'] += kwh_by_account.get(line_account, 0.0)
        if not centre and not account:
            for row in report.get('accounts') or []:
                if row['account'] not in account_totals:
                    account_totals[row['account']] = {
                        'kwh': float(row.get('kwh') or 0),
                        'cost': 0.0,
                    }

    def _breakdown(totals):
        return [
            {'name': name, 'kwh': totals[name]['kwh'], 'cost': totals[name]['cost']}
            for name in sorted(totals)
        ]

    return {
        'months': month_labels,
        'kwh': kwh_series,
        'cost': cost_series,
        'by_centre': _breakdown(centre_totals),
        'by_account': _breakdown(account_totals),
        'rows': rows,
        'bill_rows': bill_rows,
        'source': 'parsed_bills+allocation',
    }
