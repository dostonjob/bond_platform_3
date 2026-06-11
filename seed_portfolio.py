"""
Seed the portfolio with synthetic bonds for stress testing.
Usage: python seed_portfolio.py [count]   (default 100000)
Existing records are kept; synthetic ones get ISINs like STRESS000001.
"""
import json
import os
import random
import sys
import time
import uuid
from datetime import date, datetime, timedelta

sys.path.insert(0, os.path.dirname(__file__))
from engine.calculator import add_months
from engine.portfolio import STORE_PATH, _load, _DateEncoder

ISSUERS = ['Ministry of Finance', 'Central Bank', 'NavoiAzot', 'UzAuto', 'Ipoteka Bank',
           'AgroBank', 'Uztelecom', 'Thames Water', 'Acme Corp', 'Global Energy']

def make_record(i: int, rng: random.Random) -> dict:
    freq    = rng.choice([1, 2, 2, 2, 4, 12])
    settle  = date(2020, 1, 1) + timedelta(days=rng.randrange(0, 2300))
    years   = rng.randrange(2, 11)
    maturity = add_months(settle, years * 12)
    next_int = add_months(settle, 12 // freq)
    par     = rng.choice([100_000, 500_000, 1_000_000, 5_000_000, 10_000_000]) * rng.randrange(1, 20)
    price   = round(rng.uniform(88.0, 112.0), 6)
    coupon  = round(rng.uniform(2.0, 14.0), 3)

    txs = []
    n_tx = rng.choice([0, 0, 0, 1, 1, 2, 3])
    bal = par
    for _ in range(n_tx):
        span = (maturity - settle).days
        d = settle + timedelta(days=rng.randrange(30, max(31, span)))
        if rng.random() < 0.5:
            nom = par * rng.choice([0.25, 0.5, 1.0])
            bal += nom
            txs.append({'date': d, 'type': 'BUY', 'nominal': nom,
                        'clean_price': round(price + rng.uniform(-3, 3), 6),
                        'accrued_interest': 0, 'note': 'stress buy'})
        elif bal > 0:
            nom = min(bal, par * rng.choice([0.25, 0.5]))
            bal -= nom
            txs.append({'date': d, 'type': 'SELL_PARTIAL', 'nominal': nom,
                        'clean_price': round(price + rng.uniform(-3, 3), 6),
                        'accrued_interest': 0, 'note': 'stress sell'})
    txs.sort(key=lambda t: t['date'])

    created = datetime(2024, 1, 1) + timedelta(minutes=i)
    return {
        'id': uuid.uuid4().hex[:8],
        'isin': f'STRESS{i:06d}',
        'issuer': ISSUERS[i % len(ISSUERS)],
        'created_at': created.isoformat(),
        'updated_at': created.isoformat(),
        'params': {
            'isin': f'STRESS{i:06d}', 'issuer': ISSUERS[i % len(ISSUERS)],
            'par_value': float(par), 'clean_price': price,
            'coupon_rate': coupon, 'interest_frequency': freq,
            'settle_date': settle, 'last_interest_date': settle,
            'next_interest_date': next_int, 'maturity_date': maturity,
            'discount': 0, 'premium': 0, 'coupon_dates': [],
        },
        'transactions': txs,
        'notes': '',
    }

def main():
    count = int(sys.argv[1]) if len(sys.argv) > 1 else 100_000
    rng = random.Random(42)
    existing = _load()
    print(f'{len(existing)} existing records kept')
    t0 = time.perf_counter()
    records = existing + [make_record(i, rng) for i in range(1, count + 1)]
    t1 = time.perf_counter()
    print(f'generated {count:,} records in {t1-t0:.1f}s')
    with open(STORE_PATH, 'w') as f:
        json.dump(records, f, cls=_DateEncoder, separators=(',', ':'))
    t2 = time.perf_counter()
    sz = os.path.getsize(STORE_PATH) / 1e6
    print(f'wrote {len(records):,} records, {sz:.1f} MB in {t2-t1:.1f}s')

if __name__ == '__main__':
    main()
