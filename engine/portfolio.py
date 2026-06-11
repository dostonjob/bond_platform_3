"""
Portfolio Store
===============
Persistent bond portfolio using a local JSON file.
Supports: save, load, search by ISIN, edit, delete, add transactions.

Each bond record:
{
  "id":           "uuid",
  "isin":         "XYZ123",
  "issuer":       "...",
  "created_at":   "ISO datetime",
  "updated_at":   "ISO datetime",
  "params":       { ... full bond params ... },
  "transactions": [ { date, type, nominal, clean_price, accrued_interest, note } ],
  "notes":        "free text"
}
"""

import json
import uuid
import os
from datetime import date, datetime
from typing import List, Dict, Optional, Any
from copy import deepcopy


STORE_PATH = os.path.join(os.path.dirname(__file__), '..', 'data', 'portfolio.json')


# ─── JSON SERIALISATION ───────────────────────────────────────────────────────

class _DateEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (date, datetime)):
            return obj.isoformat()
        return super().default(obj)


def _date_hook(dct):
    date_keys = {
        'settle_date', 'last_interest_date', 'next_interest_date',
        'maturity_date', 'date'
    }
    for k in date_keys:
        if k in dct and isinstance(dct[k], str):
            try:
                dct[k] = date.fromisoformat(dct[k])
            except ValueError:
                pass
    return dct


# ─── LOAD / SAVE ──────────────────────────────────────────────────────────────
# The parsed file is cached in-process and invalidated by file mtime+size, so
# repeated reads (sidebar stats, page reruns) don't re-parse a large JSON.

_CACHE: Dict[str, Any] = {'stamp': None, 'records': None}


def _file_stamp():
    try:
        st = os.stat(STORE_PATH)
        return (st.st_mtime_ns, st.st_size)
    except OSError:
        return None


def _load() -> List[Dict]:
    os.makedirs(os.path.dirname(STORE_PATH), exist_ok=True)
    stamp = _file_stamp()
    if stamp is None:
        return []
    if _CACHE['stamp'] == stamp and _CACHE['records'] is not None:
        return _CACHE['records']
    try:
        with open(STORE_PATH, 'r') as f:
            records = json.load(f, object_hook=_date_hook)
    except (json.JSONDecodeError, OSError):
        return []
    _CACHE['stamp']   = stamp
    _CACHE['records'] = records
    return records


def _save(records: List[Dict]):
    os.makedirs(os.path.dirname(STORE_PATH), exist_ok=True)
    tmp = STORE_PATH + '.tmp'
    with open(tmp, 'w') as f:
        json.dump(records, f, cls=_DateEncoder, separators=(',', ':'))
    os.replace(tmp, STORE_PATH)          # atomic — no torn file on crash
    _CACHE['stamp']   = _file_stamp()
    _CACHE['records'] = records
    for k in ('sorted_for', 'index_for', 'stats_for'):   # derived views rebuild
        _CACHE.pop(k, None)


def _by_id() -> Dict[str, Dict]:
    """Lazy id→record index, rebuilt only when the record list changes."""
    records = _load()
    if _CACHE.get('index_for') is not records:
        _CACHE['by_id']    = {r['id']: r for r in records}
        _CACHE['index_for'] = records
    return _CACHE['by_id']


# ─── PUBLIC API ───────────────────────────────────────────────────────────────

def save_bond(params: Dict, transactions: List[Dict], notes: str = '') -> str:
    """
    Save a new bond to the portfolio. Returns the new record id.
    transactions should already be embedded in params if coming from calculator,
    but are also stored separately for easy editing.
    """
    records = _load()
    now = datetime.now().isoformat()
    record_id = str(uuid.uuid4())[:8]

    # Separate out transactions from params for clean storage
    clean_params = {k: v for k, v in params.items() if k != 'transactions'}
    tx_list = deepcopy(transactions or params.get('transactions', []))

    record = {
        'id':         record_id,
        'isin':       (params.get('isin') or '').upper().strip(),
        'issuer':     params.get('issuer', '') or '',
        'created_at': now,
        'updated_at': now,
        'params':     clean_params,
        'transactions': tx_list,
        'notes':      notes,
    }

    records.append(record)
    _save(records)
    return record_id


def get_all() -> List[Dict]:
    """Return all bond records, newest first (sorted view is cached)."""
    records = _load()
    if _CACHE.get('sorted_for') is not records:
        _CACHE['sorted'] = sorted(records, key=lambda r: r.get('created_at', ''), reverse=True)
        _CACHE['sorted_for'] = records
    return _CACHE['sorted']


def get_by_id(record_id: str) -> Optional[Dict]:
    return _by_id().get(record_id)


def search_by_isin(isin: str) -> List[Dict]:
    """Return all records matching the ISIN (case-insensitive, partial match)."""
    q = isin.upper().strip()
    return [r for r in _load() if q in r.get('isin', '').upper()]


def update_bond(record_id: str, params: Optional[Dict] = None,
                transactions: Optional[List] = None,
                notes: Optional[str] = None) -> bool:
    """Update an existing bond record. Pass only the fields you want to change."""
    records = _load()
    for i, r in enumerate(records):
        if r['id'] == record_id:
            if params is not None:
                records[i]['params'] = {k: v for k, v in params.items() if k != 'transactions'}
                records[i]['isin']   = (params.get('isin') or records[i]['isin']).upper().strip()
                records[i]['issuer'] = params.get('issuer') or records[i]['issuer']
            if transactions is not None:
                records[i]['transactions'] = deepcopy(transactions)
            if notes is not None:
                records[i]['notes'] = notes
            records[i]['updated_at'] = datetime.now().isoformat()
            _save(records)
            return True
    return False


def add_transaction(record_id: str, tx: Dict) -> bool:
    """Append a single transaction to an existing bond record."""
    records = _load()
    rec = _by_id().get(record_id)
    if rec is None:
        return False
    rec.setdefault('transactions', []).append(deepcopy(tx))
    rec['updated_at'] = datetime.now().isoformat()
    _save(records)
    return True


def delete_bond(record_id: str) -> bool:
    records = _load()
    new_records = [r for r in records if r['id'] != record_id]
    if len(new_records) == len(records):
        return False
    _save(new_records)
    return True


def get_full_params(record: Dict) -> Dict:
    """Merge params + transactions into a single dict for the calculator."""
    p = deepcopy(record['params'])
    p['transactions'] = deepcopy(record.get('transactions', []))
    return p


def get_portfolio_stats() -> Dict:
    records = _load()
    if _CACHE.get('stats_for') is not records:
        isins = list({r['isin'] for r in records if r.get('isin')})
        _CACHE['stats'] = {
            'total_bonds':  len(records),
            'unique_isins': len(isins),
            'isins':        isins,
        }
        _CACHE['stats_for'] = records
    return _CACHE['stats']
