"""
Bond Calculation Engine — Multi-Transaction Edition
====================================================
Handles any number of buys and sells on the same ISIN.

Transaction types:
  'BUY'          — adds a new lot to the position
  'SELL_PARTIAL' — reduces position by a given nominal (FIFO)
  'SELL_FULL'    — exits the entire remaining position

Each BUY lot tracks its own:
  - settle date (for amortization start)
  - nominal face value
  - clean price paid
  - initial discount / premium

On each period date the engine:
  1. Sums all remaining lots
  2. Amortizes each lot's discount/premium from its own settle date
  3. Blends into a single schedule row
"""

import calendar
from datetime import datetime, date, timedelta
from typing import List, Dict, Optional, Any


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def parse_date(val) -> Optional[date]:
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, (int, float)) and 40000 < val < 60000:
        return (datetime(1899, 12, 30) + timedelta(days=int(val))).date()
    if isinstance(val, str):
        for fmt in ('%Y-%m-%d','%d-%b-%y','%d-%b-%Y','%m/%d/%Y','%d/%m/%Y'):
            try:
                return datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                pass
    return None


def safe_float(val) -> Optional[float]:
    if val is None:
        return None
    if isinstance(val, str):
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def add_months(d: date, months: int) -> date:
    month = d.month + months
    year  = d.year + (month - 1) // 12
    month = (month - 1) % 12 + 1
    last  = calendar.monthrange(year, month)[1]
    return d.replace(year=year, month=month, day=min(d.day, last))


def days_between(a: date, b: date) -> int:
    return (b - a).days


def fmt_date(d: Optional[date]) -> str:
    if not d:
        return ""
    m = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    return f"{d.day:02d}-{m[d.month-1]}-{d.year}"


# ─── DISCOUNT / PREMIUM ───────────────────────────────────────────────────────

def calc_discount_premium(par_value: float, clean_price: float):
    """Returns (discount<=0, premium>=0)."""
    diff = par_value * (clean_price / 100.0 - 1.0)
    return (diff if diff < 0 else 0.0), (diff if diff > 0 else 0.0)


def derive_accrued_interest(par_value, coupon_rate, freq,
                             settle_date, last_int_date, next_int_date) -> float:
    if not all([settle_date, last_int_date, next_int_date]):
        return 0.0
    pd_  = days_between(last_int_date, next_int_date)
    ad   = days_between(last_int_date, settle_date)
    cpp  = par_value * coupon_rate / 100.0 / freq
    return cpp * (ad / pd_) if pd_ > 0 else 0.0


# ─── DATE GRID ────────────────────────────────────────────────────────────────

def _coupon_dates(next_int_date: date, maturity_date: date, freq: int) -> List[date]:
    result, d = [], next_int_date
    ms = 12 // freq
    while d <= maturity_date:
        result.append(d)
        d = add_months(d, ms)
    return result


def _all_dates(settle: date, maturity: date, next_int: date,
               freq: int, extra: List[date] = None) -> List[date]:
    s = {settle, maturity}
    d, ms = next_int, 12 // freq
    while d <= maturity:
        s.add(d); d = add_months(d, ms)
    cur = date(settle.year, settle.month, 1)
    while cur <= maturity:
        ld  = calendar.monthrange(cur.year, cur.month)[1]
        eom = date(cur.year, cur.month, ld)
        if settle < eom < maturity:
            s.add(eom)
        cur = add_months(cur, 1)
    if extra:
        s.update(extra)
    return sorted(s)


# ─── LOT HELPERS ──────────────────────────────────────────────────────────────

def _lot_carrying_value(lot: Dict, on_date: date, maturity: date) -> float:
    """Carrying value of a single lot on a given date."""
    total_days  = days_between(lot['settle_date'], maturity)
    elapsed     = days_between(lot['settle_date'], on_date)
    frac        = elapsed / total_days if total_days > 0 else 1.0
    rem         = 1.0 - frac
    return lot['nominal'] + lot['disc'] * rem + lot['prem'] * rem


def _fifo_unwind(lots: List[Dict], sell_nom: float) -> List[Dict]:
    """Remove sell_nom from lots FIFO. Returns the remaining lots."""
    remaining = []
    to_sell   = sell_nom
    for lot in lots:
        if to_sell <= 0:
            remaining.append(lot)
        elif lot['nominal'] <= to_sell + 0.01:   # absorb rounding
            to_sell -= lot['nominal']
        else:
            frac = (lot['nominal'] - to_sell) / lot['nominal']
            remaining.append({**lot,
                'nominal': lot['nominal'] - to_sell,
                'disc':    lot['disc'] * frac,
                'prem':    lot['prem'] * frac,
            })
            to_sell = 0
    return remaining


def _realized_pl(lots: List[Dict], sell_nom: float,
                 sell_price: float, sell_date: date, maturity: date) -> float:
    """Realized P&L = sale proceeds − carrying value of FIFO lots sold."""
    proceeds  = sell_nom * sell_price / 100.0
    cv_sold   = 0.0
    remaining = sell_nom
    for lot in lots:
        if remaining <= 0:
            break
        matched     = min(lot['nominal'], remaining)
        frac_of_lot = matched / lot['nominal']
        cv_sold    += _lot_carrying_value(
            {**lot, 'nominal': matched,
             'disc': lot['disc'] * frac_of_lot,
             'prem': lot['prem'] * frac_of_lot},
            sell_date, maturity
        )
        remaining  -= matched
    return round(proceeds - cv_sold, 2)


# ─── MAIN ENGINE ──────────────────────────────────────────────────────────────

def build_cashflow_table(params: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Build the complete cashflow + amortization schedule.
    Returns a list of row-dicts — one per date event.
    """
    par_value     = float(params['par_value'])
    coupon_rate   = float(params.get('coupon_rate', 0) or 0)
    freq          = int(params.get('interest_frequency', 2) or 2)
    settle_date   = params['settle_date']
    last_int_date = params['last_interest_date']
    next_int_date = params['next_interest_date']
    maturity_date = params['maturity_date']
    transactions  = sorted(
        [t for t in (params.get('transactions') or []) if t.get('date') and t.get('nominal')],
        key=lambda t: t['date']
    )

    initial_price   = float(params.get('clean_price', 100))
    accrued_initial = float(params.get('accrued_interest') or 0)
    if not accrued_initial and coupon_rate:
        accrued_initial = derive_accrued_interest(
            par_value, coupon_rate, freq, settle_date, last_int_date, next_int_date
        )

    disc0, prem0 = calc_discount_premium(par_value, initial_price)

    # Transaction lookup by date
    tx_by_date: Dict[date, List] = {}
    for t in transactions:
        tx_by_date.setdefault(t['date'], []).append(t)

    # Date grid
    all_d   = _all_dates(settle_date, maturity_date, next_int_date,
                         freq, [t['date'] for t in transactions])
    cpn_set = set(_coupon_dates(next_int_date, maturity_date, freq))
    cpn_sorted = sorted(cpn_set)
    cpp_initial = par_value * coupon_rate / 100.0 / freq

    # ── Initial lot ───────────────────────────────────────────────────────────
    lots: List[Dict] = [{
        'settle_date': settle_date,
        'nominal':     par_value,
        'clean_price': initial_price,
        'disc':        disc0,
        'prem':        prem0,
    }]

    rows: List[Dict] = []

    def _blank_row():
        return {
            'date': None, 'label': None,
            'cashflow': None, 'nominal_change': None,
            'num_days': None,
            'bond_discount': None, 'cum_amort_disc': None,
            'bond_premium': None,  'cum_amort_prem': None,
            'carrying_value': None,
            'accrued_int': None,
            'nominal_balance': None,
            'price': None,
            'mtm': None, 'oci_gl': None, 'nav': None,
            'realized_pl': None,
            'is_header': False, 'is_settle': False,
            'is_coupon': False,  'is_maturity': False,
            'is_buy': False,     'is_sell': False,
            'tx_detail': None,
        }

    # ── Opening row (purchase) ────────────────────────────────────────────────
    purchase_cf = -(par_value * initial_price / 100.0 + accrued_initial)
    r = _blank_row()
    r.update({
        'label':           f'Buy #1 — {par_value:,.0f} @ {initial_price:.6f}%',
        'cashflow':        purchase_cf,
        'nominal_change':  par_value,
        'num_days':        0,
        'bond_discount':   disc0,
        'cum_amort_disc':  0.0,
        'bond_premium':    prem0,
        'cum_amort_prem':  0.0,
        'carrying_value':  par_value * initial_price / 100.0,
        'accrued_int':     accrued_initial,
        'nominal_balance': par_value,
        'price':           initial_price,
        'mtm':             0.0,
        'oci_gl':          0.0,
        'nav':             par_value * initial_price / 100.0,
        'is_header':       True,
        'is_buy':          True,
        'tx_detail':       f'Initial purchase: {par_value:,.0f} @ {initial_price:.6f}%  |  accrued paid: {accrued_initial:,.2f}',
    })
    rows.append(r)

    # ── Main date loop ────────────────────────────────────────────────────────
    buy_counter = 1   # we already have Buy #1

    for d in all_d:
        if d < settle_date:
            continue

        # ── Process any transactions on this date ─────────────────────────────
        if d in tx_by_date:
            for tx in tx_by_date[d]:
                tx_type  = str(tx.get('type', 'BUY')).upper().strip()
                tx_nom   = float(tx.get('nominal', 0) or 0)
                tx_price = float(tx.get('clean_price', initial_price) or initial_price)
                tx_acc   = float(tx.get('accrued_interest', 0) or 0)
                tx_note  = tx.get('note', '')

                if tx_type == 'BUY':
                    buy_counter += 1
                    disc_tx, prem_tx = calc_discount_premium(tx_nom, tx_price)
                    lots.append({
                        'settle_date': d,
                        'nominal':     tx_nom,
                        'clean_price': tx_price,
                        'disc':        disc_tx,
                        'prem':        prem_tx,
                    })
                    cf_tx = -(tx_nom * tx_price / 100.0 + tx_acc)
                    new_balance = sum(l['nominal'] for l in lots)

                    r = _blank_row()
                    r.update({
                        'date':            d,
                        'label':           f'Buy #{buy_counter} — {tx_nom:,.0f} @ {tx_price:.6f}%',
                        'cashflow':        cf_tx,
                        'nominal_change':  tx_nom,
                        'num_days':        days_between(settle_date, d),
                        'bond_discount':   disc_tx,
                        'cum_amort_disc':  0.0,
                        'bond_premium':    prem_tx,
                        'cum_amort_prem':  0.0,
                        'carrying_value':  tx_nom * tx_price / 100.0,
                        'accrued_int':     tx_acc,
                        'nominal_balance': new_balance,
                        'price':           tx_price,
                        'mtm':             0.0,
                        'oci_gl':          0.0,
                        'nav':             tx_nom * tx_price / 100.0,
                        'is_buy':          True,
                        'tx_detail':       f'Buy #{buy_counter}: {tx_nom:,.0f} face @ {tx_price:.6f}%  |  accrued paid: {tx_acc:,.2f}'
                                           + (f'  |  {tx_note}' if tx_note else ''),
                    })
                    rows.append(r)

                elif tx_type in ('SELL_FULL', 'SELL_PARTIAL'):
                    total_nom_before = sum(l['nominal'] for l in lots)
                    sell_nom = total_nom_before if tx_type == 'SELL_FULL' else min(tx_nom, total_nom_before)

                    if sell_nom <= 0:
                        continue

                    pl = _realized_pl(lots, sell_nom, tx_price, d, maturity_date)
                    proceeds = sell_nom * tx_price / 100.0 + tx_acc
                    lots = _fifo_unwind(lots, sell_nom)
                    new_balance = sum(l['nominal'] for l in lots)

                    r = _blank_row()
                    r.update({
                        'date':            d,
                        'label':           f'{"Full Sell" if tx_type=="SELL_FULL" else "Partial Sell"} — {sell_nom:,.0f} @ {tx_price:.6f}%',
                        'cashflow':        proceeds,
                        'nominal_change':  -sell_nom,
                        'num_days':        days_between(settle_date, d),
                        'accrued_int':     tx_acc,
                        'nominal_balance': new_balance,
                        'price':           tx_price,
                        'realized_pl':     pl,
                        'is_sell':         True,
                        'tx_detail':       f'Sell {sell_nom:,.0f} face @ {tx_price:.6f}%  |  proceeds: {proceeds:,.2f}  |  realized P&L: {pl:+,.2f}'
                                           + (f'  |  {tx_note}' if tx_note else ''),
                    })
                    rows.append(r)

        # Skip settle date period row (already in opening row)
        if d == settle_date:
            continue

        # Skip if fully sold out
        total_nominal = sum(l['nominal'] for l in lots)
        if total_nominal <= 0:
            if d == maturity_date:
                break
            continue

        # ── Period snapshot ───────────────────────────────────────────────────
        num_days = days_between(settle_date, d)

        # Aggregate amortized disc/prem across all lots
        total_disc = total_prem = cum_adisc = cum_aprem = 0.0
        for lot in lots:
            lot_total = days_between(lot['settle_date'], maturity_date)
            lot_elap  = days_between(lot['settle_date'], d)
            lot_frac  = lot_elap / lot_total if lot_total > 0 else 1.0
            lot_rem   = 1.0 - lot_frac
            total_disc += lot['disc'] * lot_rem
            total_prem += lot['prem'] * lot_rem
            cum_adisc  += -(lot['disc'] * lot_frac) if lot['disc'] != 0 else 0.0
            cum_aprem  +=   lot['prem'] * lot_frac  if lot['prem'] != 0 else 0.0

        carrying_val = total_nominal + total_disc + total_prem

        # Accrued interest on current nominal balance
        cpp = total_nominal * coupon_rate / 100.0 / freq
        last_cpn = last_int_date
        for cd in cpn_sorted:
            if cd <= d: last_cpn = cd
            else:       break
        next_cpn = next((cd for cd in cpn_sorted if cd > last_cpn), None)
        accrued_row = 0.0
        if next_cpn:
            pd_ = days_between(last_cpn, next_cpn)
            ad  = days_between(last_cpn, d)
            accrued_row = cpp * (ad / pd_) if pd_ > 0 else 0.0

        is_coupon   = d in cpn_set
        is_maturity = d == maturity_date
        if is_coupon and not is_maturity:
            accrued_row = cpp   # full coupon paid on coupon date

        # WAC price for MTM proxy
        wac = sum(l['nominal'] * l['clean_price'] for l in lots) / total_nominal
        mtm    = total_nominal * wac / 100.0 - carrying_val
        oci_gl = mtm - cum_adisc
        nav    = carrying_val + mtm

        r = _blank_row()
        r.update({
            'date':            d,
            'num_days':        num_days,
            'bond_discount':   round(total_disc, 4),
            'cum_amort_disc':  round(cum_adisc, 4),
            'bond_premium':    round(total_prem, 4),
            'cum_amort_prem':  round(cum_aprem, 4),
            'carrying_value':  round(carrying_val, 4),
            'accrued_int':     round(accrued_row, 4),
            'nominal_balance': total_nominal,
            'price':           round(wac, 8),
            'mtm':             round(mtm, 2),
            'oci_gl':          round(oci_gl, 2),
            'nav':             round(nav, 2),
            'is_coupon':       is_coupon,
            'is_maturity':     is_maturity,
        })

        if is_maturity:
            r['cashflow']       = total_nominal
            r['nominal_change'] = -total_nominal
            r['accrued_int']    = cpp
            r['label']          = 'Maturity — Principal Repayment'

        rows.append(r)

    return rows


# ─── SUMMARY ──────────────────────────────────────────────────────────────────

def get_summary(params: Dict, rows: List[Dict]) -> Dict:
    par_value    = float(params['par_value'])
    clean_price  = float(params.get('clean_price', 100))
    coupon_rate  = float(params.get('coupon_rate', 0) or 0)
    freq         = int(params.get('interest_frequency', 2) or 2)

    disc, prem = calc_discount_premium(par_value, clean_price)
    accrued    = float(params.get('accrued_interest') or 0)
    if not accrued and coupon_rate:
        accrued = derive_accrued_interest(
            par_value, coupon_rate, freq,
            params['settle_date'], params['last_interest_date'], params['next_interest_date']
        )

    bond_type      = 'Discount' if disc < 0 else ('Premium' if prem > 0 else 'Par')
    purchase_cost  = par_value * clean_price / 100.0 + accrued
    annual_income  = par_value * coupon_rate / 100.0
    years_to_mat   = days_between(params['settle_date'], params['maturity_date']) / 365.25

    total_buys        = sum(1 for r in rows if r.get('is_buy'))
    total_sells       = sum(1 for r in rows if r.get('is_sell'))
    total_realized_pl = sum(r['realized_pl'] for r in rows if r.get('realized_pl') is not None)
    coupon_payments   = sum(1 for r in rows if r.get('is_coupon') and not r.get('is_maturity'))

    return {
        'bond_type':          bond_type,
        'discount':           disc,
        'premium':            prem,
        'purchase_cost':      purchase_cost,
        'annual_income':      annual_income,
        'years_to_maturity':  years_to_mat,
        'total_rows':         len(rows),
        'coupon_payments':    coupon_payments,
        'total_buys':         total_buys,
        'total_sells':        total_sells,
        'total_realized_pl':  total_realized_pl,
    }
