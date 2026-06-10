"""
Excel Reader & Export — v2
==========================
Read bond parameters from any Excel layout.
Write professional formatted bond reports.
"""

import io
from datetime import date, datetime
from typing import Dict, List, Tuple, Any

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from engine.calculator import parse_date, safe_float, derive_accrued_interest, fmt_date, calc_discount_premium

# ─── LABEL MAP ────────────────────────────────────────────────────────────────

LABEL_MAP = {
    'par value': 'par_value', 'face value': 'par_value', 'par val': 'par_value',
    'nominal': 'par_value',
    'clean trade price': 'clean_price', 'clean price': 'clean_price',
    'clean tradeprice': 'clean_price', 'purchase price': 'clean_price',
    'coupon rate': 'coupon_rate', 'coupon %': 'coupon_rate',
    'coupon amount': 'coupon_amount',
    'accrued interest': 'accrued_interest', 'accrued': 'accrued_interest',
    'interest frequency': 'interest_frequency', 'coupon frequency': 'interest_frequency',
    'frequency': 'interest_frequency',
    'settle date': 'settle_date', 'settlement date': 'settle_date',
    'last interest date': 'last_interest_date', 'last coupon date': 'last_interest_date',
    'next interest date': 'next_interest_date', 'next coupon date': 'next_interest_date',
    'maturity date': 'maturity_date', 'maturity': 'maturity_date',
    'discount': 'discount', 'premium': 'premium',
    'isin': 'isin', 'cusip': 'cusip', 'issuer': 'issuer',
}


# ─── READER ───────────────────────────────────────────────────────────────────

def read_from_excel(filepath_or_buffer) -> Tuple[Dict, List[str]]:
    errors = []
    try:
        wb = openpyxl.load_workbook(filepath_or_buffer, data_only=True)
    except Exception as e:
        return {}, [f"Cannot open file: {e}"]

    ws_name = wb.sheetnames[0]
    for pref in ['Automated', 'Manual', 'Sheet1', 'Bond', 'Input']:
        if pref in wb.sheetnames:
            ws_name = pref
            break
    ws = wb[ws_name]

    params = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            cs = str(cell.value).strip().lower()
            for label, key in LABEL_MAP.items():
                if label in cs and key not in params:
                    vc = ws.cell(row=cell.row, column=cell.column + 1)
                    val = vc.value
                    if val is None:
                        vc = ws.cell(row=cell.row, column=cell.column + 2)
                        val = vc.value
                    params[key] = val
                    break

    # Parse dates
    for dk in ['settle_date', 'last_interest_date', 'next_interest_date', 'maturity_date']:
        if dk in params:
            params[dk] = parse_date(params[dk])

    # Safe-float numerics
    for fk in ['par_value', 'clean_price', 'coupon_rate', 'coupon_amount',
                'accrued_interest', 'interest_frequency', 'discount', 'premium']:
        if fk in params:
            params[fk] = safe_float(params[fk])

    params.setdefault('interest_frequency', 2)
    params.setdefault('discount', 0)
    params.setdefault('premium', 0)
    params.setdefault('isin', '')
    params.setdefault('cusip', '')
    params.setdefault('issuer', '')

    # Derive missing values
    if not params.get('coupon_amount') and params.get('par_value') and params.get('coupon_rate'):
        freq = params.get('interest_frequency') or 2
        params['coupon_amount'] = params['par_value'] * params['coupon_rate'] / 100.0

    if not params.get('accrued_interest'):
        sd, li, ni = params.get('settle_date'), params.get('last_interest_date'), params.get('next_interest_date')
        pv, cr     = params.get('par_value'), params.get('coupon_rate')
        fr         = params.get('interest_frequency', 2)
        if all([sd, li, ni, pv, cr]):
            params['accrued_interest'] = derive_accrued_interest(pv, cr, fr, sd, li, ni)

    # Scan Transaction Table — reads multiple buy/sell rows
    transactions = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and 'transaction' in str(cell.value).lower():
                hr = cell.row + 1
                col_map = {}
                for hcell in ws.iter_rows(min_row=hr, max_row=hr,
                                           min_col=cell.column, max_col=cell.column + 20):
                    for hc in hcell:
                        if not hc.value:
                            continue
                        ht = str(hc.value).strip().lower()
                        if 'date' in ht:              col_map.setdefault('date', hc.column)
                        elif 'nominal' in ht:         col_map['nominal'] = hc.column
                        elif 'price' in ht:           col_map['price'] = hc.column
                        elif 'type' in ht:            col_map['type'] = hc.column
                        elif 'accrued' in ht:         col_map['accrued'] = hc.column
                        elif 'buy discount' in ht:    col_map['buy_discount'] = hc.column
                        elif 'buy premium' in ht:     col_map['buy_premium'] = hc.column
                        elif 'note' in ht:            col_map['note'] = hc.column

                for r in range(hr + 1, hr + 30):
                    if 'date' not in col_map:
                        break
                    d = parse_date(ws.cell(row=r, column=col_map['date']).value)
                    if not d:
                        continue
                    base = col_map['date']
                    tx_type_raw = ws.cell(row=r, column=col_map.get('type', base + 4)).value
                    tx_type = str(tx_type_raw or 'BUY').upper()
                    if 'SELL' in tx_type and 'PARTIAL' in tx_type:
                        tx_type = 'SELL_PARTIAL'
                    elif 'SELL' in tx_type or tx_type in ('1', '2'):
                        tx_type = 'SELL_FULL' if tx_type == '1' else 'SELL_PARTIAL'
                    else:
                        tx_type = 'BUY'

                    transactions.append({
                        'date':             d,
                        'type':             tx_type,
                        'nominal':          safe_float(ws.cell(row=r, column=col_map.get('nominal', base+1)).value) or 0,
                        'clean_price':      safe_float(ws.cell(row=r, column=col_map.get('price', base+2)).value) or params.get('clean_price', 100),
                        'accrued_interest': safe_float(ws.cell(row=r, column=col_map.get('accrued', base+3)).value) or 0,
                        'note':             ws.cell(row=r, column=col_map.get('note', base+5)).value or '',
                    })
                break

    params['transactions'] = transactions
    params['coupon_dates'] = []

    required = ['par_value', 'clean_price', 'coupon_rate', 'settle_date',
                'last_interest_date', 'next_interest_date', 'maturity_date']
    missing = [k for k in required if not params.get(k)]
    if missing:
        errors.append(f"Missing required fields: {', '.join(missing)}")

    return params, errors


# ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────

def _fill(color): return PatternFill('solid', start_color=color, end_color=color)
def _border():
    s = Side(style='thin', color='BFBFBF')
    return Border(left=s, right=s, top=s, bottom=s)

FIN  = '#,##0.00;(#,##0.00);"-"'
DATE = 'DD-MMM-YY'
PCT8 = '0.00000000'
NAVY = '1F4E79'; BLUE = '2E75B6'; ALT = 'EBF3FB'
GOLD = 'FFF2CC'; GREEN_BG = 'E2EFDA'; RED_BG = 'FDECEA'; BUY_BG = 'D6E4F0'


def export_to_excel(params: Dict, rows: List[Dict]) -> bytes:
    wb = Workbook()
    _write_report(wb, params, rows)
    _write_params(wb, params)
    _write_transactions(wb, params.get('transactions', []))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _write_report(wb, params, rows):
    ws = wb.active
    ws.title = 'Bond Report'

    widths = [14,14,14,8,16,20,14,18,20,16,14,12,12,12,14,18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    disc, prem = calc_discount_premium(params['par_value'], params['clean_price'])
    bond_type  = 'Discount' if disc < 0 else ('Premium' if prem > 0 else 'Par')

    # Title row
    ws.merge_cells('A1:P1')
    c = ws['A1']
    c.value     = 'BOND AMORTIZATION & CASHFLOW SCHEDULE  —  MULTI-TRANSACTION'
    c.font      = Font(name='Arial', bold=True, size=13, color='FFFFFF')
    c.fill      = _fill(NAVY)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 26

    ws.merge_cells('A2:P2')
    ws['A2'].value = f"ISIN: {params.get('isin','—')}   |   Issuer: {params.get('issuer','—')}   |   Type: {bond_type}   |   Generated: {datetime.now().strftime('%d %b %Y %H:%M')}"
    ws['A2'].font  = Font(name='Arial', size=9, italic=True, color='D9E1F2')
    ws['A2'].fill  = _fill(NAVY)
    ws['A2'].alignment = Alignment(horizontal='left', indent=2, vertical='center')
    ws.row_dimensions[2].height = 14

    # Table title
    ws.merge_cells('A3:P3')
    ws['A3'].value = 'CASHFLOW & AMORTIZATION SCHEDULE'
    ws['A3'].font  = Font(name='Arial', bold=True, size=10, color='FFFFFF')
    ws['A3'].fill  = _fill(BLUE)
    ws['A3'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[3].height = 18

    headers = ['Date / Event','Cashflow','Nominal','Days',
               'Bond Discount','Cum Amort Disc',
               'Bond Premium','Cum Amort Prem',
               'Carrying Value','Accrued Int',
               'Nominal Balance','Price',
               'MTM','OCI G/L','NAV','Transaction Detail']
    hr = 4
    ws.row_dimensions[hr].height = 30
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=hr, column=ci, value=h)
        c.font      = Font(name='Arial', bold=True, size=9, color='FFFFFF')
        c.fill      = _fill(BLUE)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border    = _border()

    for i, row in enumerate(rows):
        r = hr + 1 + i
        ws.row_dimensions[r].height = 14

        if row.get('is_maturity'):   bg = GREEN_BG
        elif row.get('is_sell'):     bg = RED_BG
        elif row.get('is_buy') and not row.get('is_header'): bg = BUY_BG
        elif row.get('is_header'):   bg = 'D6E4F0'
        elif row.get('is_coupon'):   bg = GOLD
        elif i % 2 == 1:             bg = ALT
        else:                        bg = 'FFFFFF'

        label = row.get('label') or (fmt_date(row['date']) if row.get('date') else '—')
        vals = [
            label, row['cashflow'], row['nominal'],
            row['num_days'] if row.get('num_days') else None,
            row['bond_discount'], row['cum_amort_disc'],
            row['bond_premium'],  row['cum_amort_prem'],
            row['carrying_value'], row['accrued_int'],
            row['nominal_balance'], row['price'],
            row['mtm'], row['oci_gl'], row['nav'],
            row.get('tx_detail', ''),
        ]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=val)
            c.font   = Font(name='Arial', size=9)
            c.fill   = _fill(bg)
            c.border = _border()
            c.alignment = Alignment(horizontal='left' if ci in (1,16) else 'right')
            if ci == 1: pass
            elif ci in (2,3,5,6,7,8,9,10,11,13,14,15): c.number_format = FIN
            elif ci == 12: c.number_format = PCT8
            if ci == 5 and isinstance(val,(int,float)) and val < 0:
                c.font = Font(name='Arial', size=9, color='C00000', bold=True)
            if ci == 7 and isinstance(val,(int,float)) and val > 0:
                c.font = Font(name='Arial', size=9, color='375623', bold=True)
            if row.get('is_sell') and ci in (2,14,15):
                c.font = Font(name='Arial', size=9, color='C00000', bold=True)
            if row.get('is_buy') and ci == 2:
                c.font = Font(name='Arial', size=9, color='375623', bold=True)

    ws.freeze_panes = ws.cell(row=5, column=2)


def _write_params(wb, params):
    ws = wb.create_sheet('Parameters')
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 22
    ws.merge_cells('A1:B1')
    c = ws['A1']
    c.value = 'BOND INPUT PARAMETERS'
    c.font  = Font(name='Arial', bold=True, size=12, color='FFFFFF')
    c.fill  = _fill(NAVY)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 22

    disc, prem = calc_discount_premium(params['par_value'], params['clean_price'])
    items = [
        ('ISIN', params.get('isin','—')), ('CUSIP', params.get('cusip','—')),
        ('Issuer', params.get('issuer','—')),
        ('Par Value', params.get('par_value')), ('Clean Price', params.get('clean_price')),
        ('Coupon Rate (%)', params.get('coupon_rate')),
        ('Frequency', params.get('interest_frequency', 2)),
        ('Accrued Interest', params.get('accrued_interest', 0)),
        ('Settle Date', params.get('settle_date')),
        ('Last Int. Date', params.get('last_interest_date')),
        ('Next Int. Date', params.get('next_interest_date')),
        ('Maturity Date', params.get('maturity_date')),
        ('Discount', disc), ('Premium', prem),
    ]
    for i, (lbl, val) in enumerate(items, 2):
        lc = ws.cell(row=i, column=1, value=lbl)
        lc.font   = Font(name='Arial', bold=True, size=10, color=NAVY)
        lc.border = _border()
        lc.fill   = _fill(ALT) if i % 2 == 0 else _fill('FFFFFF')
        vc = ws.cell(row=i, column=2, value=val)
        vc.font   = Font(name='Arial', size=10, color='0000FF')
        vc.border = _border()
        vc.fill   = _fill(GOLD)
        vc.alignment = Alignment(horizontal='right')
        if isinstance(val, date): vc.number_format = DATE
        elif isinstance(val, float) and abs(val) > 100: vc.number_format = FIN
        elif isinstance(val, float): vc.number_format = PCT8


def _write_transactions(wb, transactions):
    if not transactions:
        return
    ws = wb.create_sheet('Transactions')
    for w, col in zip([14,12,16,14,16,10,30], 'ABCDEFG'):
        ws.column_dimensions[col].width = w

    ws.merge_cells('A1:G1')
    c = ws['A1']
    c.value = 'TRANSACTION LOG'
    c.font  = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    c.fill  = _fill(NAVY)
    c.alignment = Alignment(horizontal='center')

    hdrs = ['Date','Type','Nominal','Clean Price (%)','Accrued Interest','Note','—']
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
        c.fill = _fill(BLUE); c.border = _border()
        c.alignment = Alignment(horizontal='center')

    for i, tx in enumerate(transactions, 3):
        bg = RED_BG if 'SELL' in str(tx.get('type','')).upper() else BUY_BG
        vals = [
            fmt_date(tx.get('date')), tx.get('type','BUY'),
            tx.get('nominal',0), tx.get('clean_price',0),
            tx.get('accrued_interest',0), tx.get('note',''), '',
        ]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=i, column=ci, value=val)
            c.font = Font(name='Arial', size=9)
            c.fill = _fill(bg); c.border = _border()
            if ci in (3,4,5): c.number_format = FIN
