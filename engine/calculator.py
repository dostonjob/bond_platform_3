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

Coupon income is recorded as a cashflow on every coupon date for the
nominal held on that date; maturity pays principal + final coupon.
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
        s = val.strip()
        try:
            return datetime.fromisoformat(s).date()
        except ValueError:
            pass
        for fmt in ('%Y-%m-%d', '%d-%b-%y', '%d-%b-%Y', '%m/%d/%Y', '%d/%m/%Y',
                    '%d.%m.%Y', '%Y/%m/%d', '%d %b %Y', '%b %d, %Y', '%d-%m-%Y'):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                pass
    return None


def safe_float(val) -> Optional[float]:
    if val is None:
        return None
    if isinstance(val, bool):
        return None
    if isinstance(val, str):
        s = val.strip().replace(',', '').replace('%', '').replace(' ', '')
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
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


# ─── YIELD ────────────────────────────────────────────────────────────────────

def calc_ytm(par_value, clean_price, coupon_rate, freq,
             settle_date, last_int_date, next_int_date, maturity_date,
             accrued: Optional[float] = None) -> Optional[float]:
    """
    Yield to maturity (annual %, compounded `freq` times per year) of the
    initial purchase, solved by bisection on the dirty price.
    Returns None if inputs are insufficient or no solution exists.
    """
    if not all([par_value, clean_price, settle_date, last_int_date,
                next_int_date, maturity_date]) or par_value <= 0 or clean_price <= 0:
        return None
    freq = int(freq or 2)
    cpn_dates = [d for d in _coupon_dates(next_int_date, maturity_date, freq)
                 if d > settle_date]
    if not cpn_dates:
        return None
    cpp = par_value * coupon_rate / 100.0 / freq
    if accrued is None:
        accrued = derive_accrued_interest(par_value, coupon_rate, freq,
                                          settle_date, last_int_date, next_int_date)
    dirty = par_value * clean_price / 100.0 + (accrued or 0.0)

    pd_ = days_between(last_int_date, next_int_date)
    w   = days_between(settle_date, next_int_date) / pd_ if pd_ > 0 else 1.0
    times = [w + i for i in range(len(cpn_dates))]
    last_t = times[-1]

    def pv(y):
        per = y / freq
        total = 0.0
        for t in times:
            total += cpp / (1.0 + per) ** t
        total += par_value / (1.0 + per) ** last_t
        return total

    lo, hi = -0.90, 10.0
    f_lo, f_hi = pv(lo) - dirty, pv(hi) - dirty
    if f_lo < 0 or f_hi > 0:        # price outside achievable range
        return None
    for _ in range(200):
        mid = (lo + hi) / 2.0
        if pv(mid) - dirty > 0:
            lo = mid
        else:
            hi = mid
        if hi - lo < 1e-10:
            break
    return round((lo + hi) / 2.0 * 100.0, 6)


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
        s.update(x for x in extra if settle <= x <= maturity)
    return sorted(s)


# ─── LOT HELPERS ──────────────────────────────────────────────────────────────

def _lot_carrying_value(lot: Dict, on_date: date, maturity: date) -> float:
    """Carrying value of a single lot on a given date."""
    total_days  = days_between(lot['settle_date'], maturity)
    elapsed     = days_between(lot['settle_date'], on_date)
    frac        = elapsed / total_days if total_days > 0 else 1.0
    frac        = min(max(frac, 0.0), 1.0)
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


# ─── VALIDATION ───────────────────────────────────────────────────────────────

def validate_params(params: Dict[str, Any]) -> List[str]:
    """
    Sanity-check bond parameters and transactions.
    Returns a list of human-readable warnings (empty = all good).
    The engine still runs with warnings — it clamps/skips bad transactions.
    """
    w: List[str] = []
    sd = parse_date(params.get('settle_date'))
    md = parse_date(params.get('maturity_date'))
    li = parse_date(params.get('last_interest_date'))
    ni = parse_date(params.get('next_interest_date'))
    pv = safe_float(params.get('par_value')) or 0.0
    cp = safe_float(params.get('clean_price')) or 0.0

    if pv <= 0:
        w.append("Par value must be positive.")
    if cp <= 0:
        w.append("Clean price must be positive.")
    if sd and md and md <= sd:
        w.append("Maturity date must be after the settlement date.")
    if li and ni and ni <= li:
        w.append("Next interest date must be after the last interest date.")
    if sd and li and ni and not (li <= sd < ni):
        w.append("Settlement date lies outside the last→next interest period — "
                 "derived accrued interest may be wrong.")

    txs = [t for t in (params.get('transactions') or []) if t.get('date')]
    txs = sorted(txs, key=lambda t: parse_date(t['date']) or sd or date.min)
    balance = pv
    for i, t in enumerate(txs, start=2):
        d   = parse_date(t.get('date'))
        tp  = str(t.get('type', 'BUY')).upper().strip()
        nom = safe_float(t.get('nominal')) or 0.0
        prc = safe_float(t.get('clean_price'))
        label = f"Transaction #{i} ({tp} on {fmt_date(d)})"
        if d and sd and d < sd:
            w.append(f"{label} is dated before settlement — it will be ignored.")
            continue
        if d and md and d > md:
            w.append(f"{label} is dated after maturity — it will be ignored.")
            continue
        if prc is not None and prc <= 0:
            w.append(f"{label} has a non-positive price.")
        if tp == 'BUY':
            if nom <= 0:
                w.append(f"{label} has no nominal — it will be ignored.")
            else:
                balance += nom
        elif 'SELL' in tp:
            if balance <= 0:
                w.append(f"{label}: nothing left to sell — it will be skipped.")
            elif tp == 'SELL_PARTIAL' and nom > balance + 0.01:
                w.append(f"{label}: sell nominal {nom:,.0f} exceeds holding "
                         f"{balance:,.0f} — it will be clamped to the holding.")
            balance = 0.0 if tp == 'SELL_FULL' else max(0.0, balance - min(nom, balance))
    return w


# ─── MAIN ENGINE ──────────────────────────────────────────────────────────────

def build_cashflow_table(params: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Build the complete cashflow + amortization schedule.
    Returns a list of row-dicts — one per date event.
    """
    par_value     = float(params['par_value'])
    coupon_rate   = float(params.get('coupon_rate', 0) or 0)
    freq          = int(params.get('interest_frequency', 2) or 2)
    settle_date   = parse_date(params['settle_date'])
    last_int_date = parse_date(params['last_interest_date'])
    next_int_date = parse_date(params['next_interest_date'])
    maturity_date = parse_date(params['maturity_date'])

    # ── Credit-risk parameters (fractions, like the reference sheet) ─────────
    # LGD = 1 − CDS recovery (standard recovery 40% → LGD 0.6).
    # PD: explicit DRSK value if supplied, otherwise implied from the price
    # discount via the credit triangle: PD ≈ (1 − price/100) / (LGD × years).
    lgd_frac = safe_float(params.get('lgd'))
    if lgd_frac is None:
        rec = safe_float(params.get('recovery_rate'))
        if rec is not None:
            if rec > 1:
                rec /= 100.0
            lgd_frac = 1.0 - rec
    if lgd_frac is None:
        lgd_frac = 0.60
    elif lgd_frac > 1:        # entered as a percent (e.g. 60)
        lgd_frac /= 100.0

    pd_frac = safe_float(params.get('default_probability'))
    if pd_frac is not None and pd_frac > 1:   # entered as a percent
        pd_frac /= 100.0
    if pd_frac is None:
        _price = safe_float(params.get('clean_price')) or 100.0
        _years = days_between(settle_date, maturity_date) / 365.25
        pd_frac = (max(0.0, (1.0 - _price / 100.0) / (lgd_frac * _years))
                   if lgd_frac > 0 and _years > 0 else 0.0)
    el_rate = pd_frac * lgd_frac

    transactions = []
    for t in (params.get('transactions') or []):
        d = parse_date(t.get('date'))
        nom = safe_float(t.get('nominal')) or 0.0
        tp  = str(t.get('type', 'BUY')).upper().strip()
        # keep SELL_FULL even without a nominal — it sells the whole position
        if not d or not (settle_date <= d <= maturity_date):
            continue
        if nom <= 0 and tp != 'SELL_FULL':
            continue
        transactions.append({**t, 'date': d, 'nominal': nom, 'type': tp})
    transactions.sort(key=lambda t: t['date'])

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

    def _accrued_at(nominal: float, d: date) -> float:
        """Accrued interest on `nominal` face at date d (0 on coupon dates)."""
        if not coupon_rate or nominal <= 0:
            return 0.0
        cpp_ = nominal * coupon_rate / 100.0 / freq
        last_cpn = last_int_date
        for cd in cpn_sorted:
            if cd <= d: last_cpn = cd
            else:       break
        nxt = next((cd for cd in cpn_sorted if cd > last_cpn), None)
        if not nxt:
            return 0.0
        pd_ = days_between(last_cpn, nxt)
        ad  = days_between(last_cpn, d)
        return cpp_ * (ad / pd_) if pd_ > 0 else 0.0

    def _agg(current_lots: List[Dict], d: date):
        """Aggregate position snapshot across all remaining lots on date d."""
        total_nominal = sum(l['nominal'] for l in current_lots)
        total_disc = total_prem = cum_adisc = cum_aprem = 0.0
        for lot in current_lots:
            lot_total = days_between(lot['settle_date'], maturity_date)
            lot_elap  = days_between(lot['settle_date'], d)
            lot_frac  = lot_elap / lot_total if lot_total > 0 else 1.0
            lot_frac  = min(max(lot_frac, 0.0), 1.0)
            lot_rem   = 1.0 - lot_frac
            total_disc += lot['disc'] * lot_rem
            total_prem += lot['prem'] * lot_rem
            cum_adisc  += -(lot['disc'] * lot_frac)
            cum_aprem  +=   lot['prem'] * lot_frac
        carrying = total_nominal + total_disc + total_prem
        wac = (sum(l['nominal'] * l['clean_price'] for l in current_lots) / total_nominal
               if total_nominal > 0 else None)
        return total_nominal, total_disc, total_prem, cum_adisc, cum_aprem, carrying, wac

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
            'check': None,
            'realized_pl': None,
            'realized_interest_income': None,
            'total_pl': None,
            'wac': None,
            'check2': None,
            'default_prob': round(pd_frac, 8),
            'lgd': round(lgd_frac, 6),
            'expected_loss': None,
            'el_change': None,
            'interest_paid': None,
            'is_header': False, 'is_settle': False,
            'is_coupon': False,  'is_maturity': False,
            'is_buy': False,     'is_sell': False,
            'tx_detail': None,
        }

    # Running realized totals across the whole schedule
    cum_sell_pl  = 0.0                  # Σ realized P&L from sells
    cum_interest = -accrued_initial     # accrued paid at purchase is negative income
    prev_el      = None                 # for the Expected Loss "Change" column

    def _credit(row: Dict, cost_basis: float):
        """Expected Loss = invested cost × PD × LGD (as in the reference sheet);
        Change is non-zero only when the position changes."""
        nonlocal prev_el
        el = round((cost_basis or 0.0) * el_rate, 2)
        row['expected_loss'] = el
        row['el_change']     = round(el - prev_el, 2) if prev_el is not None else 0.0
        prev_el = el

    # ── Opening row (purchase) ────────────────────────────────────────────────
    purchase_cf = -(par_value * initial_price / 100.0 + accrued_initial)
    carrying0   = par_value * initial_price / 100.0
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
        'carrying_value':  carrying0,
        'accrued_int':     accrued_initial,
        'nominal_balance': par_value,
        'price':           initial_price,
        'mtm':             0.0,
        'oci_gl':          0.0,
        'nav':             carrying0,
        'check':           0.0,
        'check2':          0.0,
        'realized_interest_income': round(cum_interest, 2),
        'total_pl':        round(cum_sell_pl + cum_interest, 2),
        'wac':             initial_price,
        'is_header':       True,
        'is_buy':          True,
        'tx_detail':       f'Initial purchase: {par_value:,.0f} @ {initial_price:.6f}%  |  accrued paid: {accrued_initial:,.2f}',
    })
    _credit(r, carrying0)
    rows.append(r)

    # ── Main date loop ────────────────────────────────────────────────────────
    buy_counter = 1   # we already have Buy #1

    for d in all_d:
        if d < settle_date:
            continue

        # ── Process any transactions on this date ─────────────────────────────
        if d in tx_by_date:
            for tx in tx_by_date[d]:
                tx_type  = tx['type']
                tx_nom   = float(tx.get('nominal', 0) or 0)
                tx_price = float(safe_float(tx.get('clean_price')) or initial_price)
                tx_acc   = float(safe_float(tx.get('accrued_interest')) or 0)
                tx_note  = tx.get('note', '')

                if tx_type == 'BUY':
                    buy_counter += 1
                    if not tx_acc:
                        tx_acc = _accrued_at(tx_nom, d)
                    disc_tx, prem_tx = calc_discount_premium(tx_nom, tx_price)
                    lots.append({
                        'settle_date': d,
                        'nominal':     tx_nom,
                        'clean_price': tx_price,
                        'disc':        disc_tx,
                        'prem':        prem_tx,
                    })
                    cf_tx = -(tx_nom * tx_price / 100.0 + tx_acc)
                    cum_interest -= tx_acc

                    # Full position snapshot AFTER the buy (all lots combined)
                    tn, td, tp_, ca, cp_, cv, wac = _agg(lots, d)
                    mtm = tn * wac / 100.0 - cv
                    nav = cv + mtm

                    r = _blank_row()
                    r.update({
                        'date':            d,
                        'label':           f'Buy #{buy_counter} — {tx_nom:,.0f} @ {tx_price:.6f}%',
                        'cashflow':        cf_tx,
                        'nominal_change':  tx_nom,
                        'num_days':        days_between(settle_date, d),
                        'bond_discount':   round(td, 4),
                        'cum_amort_disc':  round(ca, 4),
                        'bond_premium':    round(tp_, 4),
                        'cum_amort_prem':  round(cp_, 4),
                        'carrying_value':  round(cv, 4),
                        'accrued_int':     tx_acc,
                        'nominal_balance': tn,
                        'price':           tx_price,
                        'mtm':             round(mtm, 2),
                        'oci_gl':          round(mtm, 2),
                        'nav':             round(nav, 2),
                        'check':           round(cv - (tn + td + tp_), 2),
                        'check2':          round(nav - (cv + mtm), 2),
                        'realized_interest_income': round(cum_interest, 2),
                        'total_pl':        round(cum_sell_pl + cum_interest, 2),
                        'wac':             round(wac, 8),
                        'is_buy':          True,
                        'tx_detail':       f'Buy #{buy_counter}: {tx_nom:,.0f} face @ {tx_price:.6f}%  |  accrued paid: {tx_acc:,.2f}'
                                           + (f'  |  {tx_note}' if tx_note else ''),
                    })
                    _credit(r, tn * wac / 100.0)
                    rows.append(r)

                elif tx_type in ('SELL_FULL', 'SELL_PARTIAL'):
                    total_nom_before = sum(l['nominal'] for l in lots)
                    sell_nom = total_nom_before if tx_type == 'SELL_FULL' else min(tx_nom, total_nom_before)

                    if sell_nom <= 0:
                        continue

                    if not tx_acc:
                        tx_acc = _accrued_at(sell_nom, d)
                    pl = _realized_pl(lots, sell_nom, tx_price, d, maturity_date)
                    proceeds = sell_nom * tx_price / 100.0 + tx_acc
                    lots = _fifo_unwind(lots, sell_nom)
                    cum_sell_pl  += pl
                    cum_interest += tx_acc

                    # Full position snapshot AFTER the sell (remaining lots)
                    tn, td, tp_, ca, cp_, cv, wac = _agg(lots, d)
                    mtm = (tn * wac / 100.0 - cv) if wac is not None else 0.0
                    nav = cv + mtm

                    r = _blank_row()
                    r.update({
                        'date':            d,
                        'label':           f'{"Full Sell" if tx_type=="SELL_FULL" else "Partial Sell"} — {sell_nom:,.0f} @ {tx_price:.6f}%',
                        'cashflow':        proceeds,
                        'nominal_change':  -sell_nom,
                        'num_days':        days_between(settle_date, d),
                        'bond_discount':   round(td, 4),
                        'cum_amort_disc':  round(ca, 4),
                        'bond_premium':    round(tp_, 4),
                        'cum_amort_prem':  round(cp_, 4),
                        'carrying_value':  round(cv, 4),
                        'accrued_int':     tx_acc,
                        'nominal_balance': tn,
                        'price':           tx_price,
                        'mtm':             round(mtm, 2),
                        'oci_gl':          round(mtm, 2),
                        'nav':             round(nav, 2),
                        'check':           round(cv - (tn + td + tp_), 2),
                        'check2':          round(nav - (cv + mtm), 2),
                        'realized_pl':     pl,
                        'realized_interest_income': round(cum_interest, 2),
                        'total_pl':        round(cum_sell_pl + cum_interest, 2),
                        'wac':             round(wac, 8) if wac is not None else None,
                        'is_sell':         True,
                        'tx_detail':       f'Sell {sell_nom:,.0f} face @ {tx_price:.6f}%  |  proceeds: {proceeds:,.2f}  |  realized P&L: {pl:+,.2f}'
                                           + (f'  |  {tx_note}' if tx_note else ''),
                    })
                    _credit(r, tn * wac / 100.0 if wac is not None else 0.0)
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
        total_nominal, total_disc, total_prem, cum_adisc, cum_aprem, carrying_val, wac = _agg(lots, d)

        # Accrued interest on current nominal balance
        cpp = total_nominal * coupon_rate / 100.0 / freq
        accrued_row = _accrued_at(total_nominal, d)

        is_coupon   = d in cpn_set
        is_maturity = d == maturity_date

        # Coupon cashflow: full coupon paid on coupon dates; accrued shows the
        # full coupon that accumulated and is being paid (matches the cashflow)
        coupon_cf     = None
        interest_paid = None
        if is_coupon and not is_maturity:
            coupon_cf     = cpp
            interest_paid = cpp
            accrued_row   = cpp
            cum_interest += cpp

        # WAC price for MTM proxy (fair value at cost vs amortized carrying value)
        mtm    = total_nominal * wac / 100.0 - carrying_val
        oci_gl = mtm
        nav    = carrying_val + mtm

        r = _blank_row()
        r.update({
            'date':            d,
            'cashflow':        coupon_cf,
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
            'check':           round(carrying_val - (total_nominal + total_disc + total_prem), 2),
            'check2':          round(nav - (carrying_val + mtm), 2),
            'realized_interest_income': round(cum_interest, 2),
            'total_pl':        round(cum_sell_pl + cum_interest, 2),
            'wac':             round(wac, 8),
            'interest_paid':   interest_paid,
            'is_coupon':       is_coupon,
            'is_maturity':     is_maturity,
        })
        _credit(r, total_nominal * wac / 100.0)

        if is_maturity:
            # Redemption at par: principal + final coupon (or stub accrued)
            final_int = cpp if is_coupon else accrued_row
            cum_interest += final_int
            r['cashflow']       = total_nominal + final_int
            r['nominal_change'] = -total_nominal
            r['accrued_int']    = round(final_int, 4)
            r['interest_paid']  = final_int
            r['mtm']            = 0.0
            r['oci_gl']         = 0.0
            r['nav']            = round(carrying_val, 4)
            r['realized_interest_income'] = round(cum_interest, 2)
            r['total_pl']       = round(cum_sell_pl + cum_interest, 2)
            r['label']          = ('Maturity — Principal + Final Coupon'
                                   if final_int else 'Maturity — Principal Repayment')

        rows.append(r)

    return rows


# ─── SUMMARY ──────────────────────────────────────────────────────────────────

def get_summary(params: Dict, rows: List[Dict]) -> Dict:
    par_value    = float(params['par_value'])
    clean_price  = float(params.get('clean_price', 100))
    coupon_rate  = float(params.get('coupon_rate', 0) or 0)
    freq         = int(params.get('interest_frequency', 2) or 2)
    settle_date  = parse_date(params['settle_date'])
    last_int     = parse_date(params['last_interest_date'])
    next_int     = parse_date(params['next_interest_date'])
    maturity     = parse_date(params['maturity_date'])

    disc, prem = calc_discount_premium(par_value, clean_price)
    accrued    = float(params.get('accrued_interest') or 0)
    if not accrued and coupon_rate:
        accrued = derive_accrued_interest(
            par_value, coupon_rate, freq, settle_date, last_int, next_int
        )

    bond_type      = 'Discount' if disc < 0 else ('Premium' if prem > 0 else 'Par')
    purchase_cost  = par_value * clean_price / 100.0 + accrued
    annual_income  = par_value * coupon_rate / 100.0
    years_to_mat   = days_between(settle_date, maturity) / 365.25

    total_buys        = sum(1 for r in rows if r.get('is_buy'))
    total_sells       = sum(1 for r in rows if r.get('is_sell'))
    total_realized_pl = sum(r['realized_pl'] for r in rows if r.get('realized_pl') is not None)
    coupon_payments   = sum(1 for r in rows if r.get('is_coupon') and not r.get('is_maturity'))

    # Cashflow aggregates
    total_invested      = -sum(r['cashflow'] for r in rows
                               if r.get('is_buy') and r.get('cashflow') is not None)
    total_sell_proceeds = sum(r['cashflow'] for r in rows
                              if r.get('is_sell') and r.get('cashflow') is not None)
    total_coupon_income = sum(r['interest_paid'] for r in rows
                              if r.get('interest_paid') is not None)
    redemption          = sum(-r['nominal_change'] for r in rows
                              if r.get('is_maturity') and r.get('nominal_change') is not None)
    net_cashflow        = sum(r['cashflow'] for r in rows if r.get('cashflow') is not None)

    # Final running totals from the schedule (net of accrued paid/received)
    realized_interest_income = 0.0
    total_pl                 = 0.0
    expected_loss            = None
    for r in reversed(rows):
        if r.get('realized_interest_income') is not None:
            realized_interest_income = r['realized_interest_income']
            total_pl                 = r.get('total_pl') or 0.0
            expected_loss            = r.get('expected_loss')
            break

    # Current holding = last known nominal balance in the schedule
    current_nominal = 0.0
    for r in reversed(rows):
        if r.get('nominal_balance') is not None:
            current_nominal = r['nominal_balance']
            break
    matured = any(r.get('is_maturity') for r in rows)
    if matured:
        status, current_nominal = 'Matured', 0.0
    elif current_nominal <= 0:
        status = 'Closed'
    else:
        status = 'Active'

    ytm = calc_ytm(par_value, clean_price, coupon_rate, freq,
                   settle_date, last_int, next_int, maturity, accrued=accrued)
    current_yield = (coupon_rate / clean_price * 100.0) if clean_price > 0 else None

    return {
        'bond_type':           bond_type,
        'discount':            disc,
        'premium':             prem,
        'purchase_cost':       purchase_cost,
        'annual_income':       annual_income,
        'years_to_maturity':   years_to_mat,
        'total_rows':          len(rows),
        'coupon_payments':     coupon_payments,
        'total_buys':          total_buys,
        'total_sells':         total_sells,
        'total_realized_pl':   total_realized_pl,
        'total_invested':      total_invested,
        'total_sell_proceeds': total_sell_proceeds,
        'total_coupon_income': total_coupon_income,
        'redemption':          redemption,
        'net_cashflow':        net_cashflow,
        'current_nominal':     current_nominal,
        'status':              status,
        'ytm':                 ytm,
        'current_yield':       current_yield,
        'realized_interest_income': realized_interest_income,
        'total_pl':            total_pl,
        'expected_loss':       expected_loss,
        'default_probability': rows[0].get('default_prob') if rows else None,
        'lgd':                 rows[0].get('lgd') if rows else None,
    }
