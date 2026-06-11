"""
Bond Analytics Platform v2
===========================
Multi-transaction bonds + persistent portfolio history.
No dashboards — just the workflow.
Run: streamlit run app.py
"""

import streamlit as st
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

st.set_page_config(page_title="Bond Analytics Platform", page_icon="🏦",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
#MainMenu,footer,header{visibility:hidden;}

/* ── Banner ── */
.banner{background:linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%);color:white;
  padding:1.2rem 1.8rem;border-radius:12px;margin-bottom:1.4rem;
  display:flex;align-items:center;gap:1rem;}
.banner h1{margin:0;font-size:1.4rem;font-weight:700;color:white;}
.banner p{margin:0;font-size:0.82rem;opacity:.85;}

/* ── Section headers ── */
.sec{font-size:.72rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
  color:#2E75B6;border-bottom:2px solid #D6E4F0;padding-bottom:4px;margin:1rem 0 .7rem;}

/* ── Transaction pills (Updated to Dark/Black Background) ── */
.tx-buy{background:#0A0A0A;border:1px solid #222;border-left:4px solid #2E75B6;border-radius:6px;
  padding:.55rem .9rem;margin:.35rem 0;font-size:.83rem;line-height:1.5;color:#FFFFFF !important;}
.tx-sell-partial{background:#0A0A0A;border:1px solid #222;border-left:4px solid #E8A000;border-radius:6px;
  padding:.55rem .9rem;margin:.35rem 0;font-size:.83rem;line-height:1.5;color:#FFFFFF !important;}
.tx-sell-full{background:#0A0A0A;border:1px solid #222;border-left:4px solid #C00000;border-radius:6px;
  padding:.55rem .9rem;margin:.35rem 0;font-size:.83rem;line-height:1.5;color:#FFFFFF !important;}

/* ── Bond card in portfolio ── */
.bcard{background:white;border:1px solid #D6E4F0;border-radius:10px;
  padding:.9rem 1.1rem;margin-bottom:.5rem;}
.bcard:hover{border-color:#2E75B6;}

/* ── Badges ── */
.badge{display:inline-block;padding:2px 10px;border-radius:20px;font-size:.74rem;font-weight:700;}
.disc{background:#FDECEA;color:#B71C1C;}
.prem{background:#E8F5E9;color:#1B5E20;}
.par{background:#E3F2FD;color:#0D47A1;}

/* ── Summary metric cards ── */
.mcard{background:#F7FAFF;border:1px solid #D6E4F0;border-radius:9px;
  padding:.8rem 1rem;text-align:center;height:100%;}
.mcard .v{font-size:1.3rem;font-weight:700;color:#1F4E79;}
.mcard .l{font-size:.74rem;color:#888;margin-top:3px;}

/* ── Black sidebar ── */
[data-testid="stSidebar"]{background:#0A0A0A !important;border-right:1px solid #1a1a1a !important;}
[data-testid="stSidebar"] *{color:#E8E8E8 !important;}
[data-testid="stSidebar"] .stMarkdown h3,
[data-testid="stSidebar"] .stMarkdown strong{color:#FFFFFF !important;}
[data-testid="stSidebar"] hr{border-color:#2a2a2a !important;}
[data-testid="stSidebar"] .stButton>button{
  background:#1C1C1C !important;color:#E8E8E8 !important;
  border:1px solid #333 !important;border-radius:8px !important;}
[data-testid="stSidebar"] .stButton>button:hover{
  background:#2E75B6 !important;color:#fff !important;border-color:#2E75B6 !important;}
[data-testid="stSidebar"] .stSuccess{
  background:#0d2a0d !important;color:#4ade80 !important;border-color:#166534 !important;}
[data-testid="stSidebar"] small,[data-testid="stSidebar"] .stCaption{color:#888 !important;}

.stButton>button{border-radius:8px;}
</style>
""", unsafe_allow_html=True)

import pandas as pd
from datetime import date as dt_date
from copy import deepcopy

from engine.calculator import (
    build_cashflow_table, get_summary, calc_discount_premium,
    derive_accrued_interest, fmt_date, parse_date, validate_params
)
from engine.portfolio import (
    save_bond, get_all, get_by_id, search_by_isin,
    update_bond, add_transaction, delete_bond, get_full_params, get_portfolio_stats
)
from engine.excel_io import read_from_excel, export_to_excel

# ── Session state ──────────────────────────────────────────────────────────────
for k, v in [('page','home'),('params',None),('rows',None),('summary',None),
              ('edit_id',None),('calc_txs',[]),('warnings',[])]:
    if k not in st.session_state:
        st.session_state[k] = v

def go(page, **kw):
    st.session_state.page = page
    for k,v in kw.items():
        st.session_state[k] = v
    st.rerun()

# ── UI helpers ─────────────────────────────────────────────────────────────────
def banner(title, sub, icon='🏦'):
    st.markdown(f'<div class="banner"><div style="font-size:2rem">{icon}</div>'
                f'<div><h1>{title}</h1><p>{sub}</p></div></div>', unsafe_allow_html=True)

def sec(t):
    st.markdown(f'<div class="sec">{t}</div>', unsafe_allow_html=True)

def mcard_row(items):
    cols = st.columns(len(items))
    for col,(lbl,val,*rest) in zip(cols,items):
        color = rest[0] if rest else '#1F4E79'
        with col:
            st.markdown(f'<div class="mcard"><div class="v" style="color:{color}">{val}</div>'
                        f'<div class="l">{lbl}</div></div>', unsafe_allow_html=True)

def tx_pill(tx, num):
    tp = str(tx.get('type','BUY')).upper()
    if tp == 'BUY':
        cls, icon = 'tx-buy', '🔵'
    elif tp == 'SELL_FULL':
        cls, icon = 'tx-sell-full', '🔴'
    else:
        cls, icon = 'tx-sell-partial', '🟡'
    note_part = f' &nbsp;|&nbsp; {tx.get("note","")}' if tx.get('note') else ''
    return (f'<div class="{cls}">'
            f'{icon} <b>#{num} {tp}</b> &nbsp;|&nbsp; '
            f'📅 {fmt_date(tx.get("date"))} &nbsp;|&nbsp; '
            f'Nominal: <b>{tx.get("nominal",0):,.0f}</b> &nbsp;|&nbsp; '
            f'Price: {tx.get("clean_price",0):.6f}% &nbsp;|&nbsp; '
            f'Accrued: {tx.get("accrued_interest",0):,.2f}'
            f'{note_part}</div>')

def run_calc(params):
    warnings = validate_params(params)
    rows     = build_cashflow_table(params)
    summary  = get_summary(params, rows)
    st.session_state.params   = params
    st.session_state.rows     = rows
    st.session_state.summary  = summary
    st.session_state.warnings = warnings
    return rows, summary

def build_df(rows, filter_opt='All rows'):
    disp = rows
    if filter_opt == 'Coupon dates only':
        disp = [r for r in rows if r.get('is_coupon') or r.get('is_header') or r.get('is_buy') or r.get('is_sell')]
    elif filter_opt == 'Year-end only':
        disp = [r for r in rows if r.get('is_header') or r.get('is_maturity') or r.get('is_buy') or r.get('is_sell')
                or (r.get('date') and r['date'].month==12 and r['date'].day==31)]
    elif filter_opt == 'Transactions only':
        disp = [r for r in rows if r.get('is_header') or r.get('is_buy') or r.get('is_sell') or r.get('is_maturity')]

    data = []
    for r in disp:
        lbl = r.get('label') or (fmt_date(r['date']) if r.get('date') else '—')
        data.append({
            'Date / Event':    lbl,
            'Cashflow':        r.get('cashflow'),
            'Nominal Δ':       r.get('nominal_change'),
            'Days':            r.get('num_days'),
            'Bond Discount':   r.get('bond_discount'),
            'Cum Amort Disc':  r.get('cum_amort_disc'),
            'Bond Premium':    r.get('bond_premium'),
            'Cum Amort Prem':  r.get('cum_amort_prem'),
            'Carrying Value':  r.get('carrying_value'),
            'Accrued Int':     r.get('accrued_int'),
            'Nominal Balance': r.get('nominal_balance'),
            'Price':           r.get('price'),
            'MTM':             r.get('mtm'),
            'OCI G/L':         r.get('oci_gl'),
            'NAV':             r.get('nav'),
            'Check':           r.get('check'),
            'P&L (Sell)':      r.get('realized_pl'),
            'Realized Int Income': r.get('realized_interest_income'),
            'Total P&L':       r.get('total_pl'),
            'WAC':             r.get('wac'),
            'Check 2':         r.get('check2'),
            'Default Prob (DRSK)': r.get('default_prob'),
            'LGD':             r.get('lgd'),
            'Expected Loss':   r.get('expected_loss'),
            'Change':          r.get('el_change'),
        })
    df = pd.DataFrame(data)
    num_cols = [c for c in df.columns if c not in ('Date / Event', 'Days')]
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors='coerce')
    return df, disp, num_cols

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏦 Bond Analytics")
    st.markdown("---")
    stats = get_portfolio_stats()
    st.markdown(f"**Portfolio:** {stats['total_bonds']} bonds · {stats['unique_isins']} ISINs")
    st.markdown("---")
    nav = {
        "🏠 Home":              "home",
        "📤 Upload Excel":      "upload",
        "🔢 New Bond":          "calculator",
        "📋 Portfolio":         "portfolio",
        "📊 Results":           "results",
        "📥 Export":            "export",
    }
    for lbl, key in nav.items():
        dis = key in ('results','export') and st.session_state.rows is None
        if st.button(lbl, use_container_width=True, key=f"nav_{key}", disabled=dis):
            go(key)
    st.markdown("---")
    if st.session_state.rows:
        p = st.session_state.params
        st.success(f"✓ {p.get('isin','—')}  loaded")
        if st.button("🗑 Clear", use_container_width=True):
            for k in ['params','rows','summary']:
                st.session_state[k] = None
            go("home")

page = st.session_state.page

# ════════════════════════════════════════════════════════════════════════════════
# HOME
# ════════════════════════════════════════════════════════════════════════════════
if page == 'home':
    banner("Bond Analytics Platform", "Multi-buy · Multi-sell · Portfolio history · Excel export", "🏦")
    c1,c2,c3 = st.columns(3, gap="large")
    with c1:
        st.markdown("#### 📤 Upload Excel")
        st.markdown("Drop your bond Excel file. Auto-reads all parameters and any transaction table.")
        if st.button("Upload Excel →", use_container_width=True, type="primary"): go("upload")
    with c2:
        st.markdown("#### 🔢 Manual Entry")
        st.markdown("Enter bond terms and add unlimited buy/sell transactions one by one.")
        if st.button("New Bond →", use_container_width=True): go("calculator")
    with c3:
        st.markdown("#### 📋 Portfolio")
        st.markdown("Browse saved bonds. Search by ISIN. Add transactions to existing bonds.")
        if st.button("View Portfolio →", use_container_width=True): go("portfolio")

# ════════════════════════════════════════════════════════════════════════════════
# UPLOAD
# ════════════════════════════════════════════════════════════════════════════════
elif page == 'upload':
    banner("Upload Excel File", "Auto-parse bond parameters and transaction table", "📤")

    uploaded = st.file_uploader("Drop your bond Excel file", type=['xlsx','xls'])
    if not uploaded:
        st.info("Supports any layout — the reader scans all cells for labels like "
                "'Par Value', 'Clean Trade Price', 'Coupon Rate', 'Settle Date', etc.")
    else:
        with st.spinner("Reading file..."):
            params, errors = read_from_excel(uploaded)

        if errors:
            for e in errors: st.error(e)
            st.markdown("**Required labels** (in any cell, value to its right):\n"
                        "`Par Value` · `Clean Trade Price` · `Coupon Rate` · "
                        "`Settle Date` · `Last Interest Date` · `Next Interest Date` · `Maturity Date`")
        else:
            st.success("✅ File read successfully")
            sec("Detected Parameters")
            c1,c2,c3 = st.columns(3)
            with c1:
                st.metric("Par Value",   f"{params.get('par_value',0):,.2f}")
                st.metric("Clean Price", f"{params.get('clean_price',0):.6f}%")
                st.metric("Coupon Rate", f"{params.get('coupon_rate',0):.4f}%")
            with c2:
                st.metric("Settle Date",   str(params.get('settle_date','—')))
                st.metric("Maturity Date", str(params.get('maturity_date','—')))
                st.metric("ISIN / Issuer", f"{params.get('isin','—')} / {params.get('issuer','—')}")
            with c3:
                st.metric("Frequency",       f"{params.get('interest_frequency',2)}x / yr")
                st.metric("Accrued Interest",f"{params.get('accrued_interest',0) or 0:,.4f}")
                st.metric("Transactions in file", len(params.get('transactions',[])))

            if params.get('transactions'):
                sec(f"Transactions Found in File ({len(params['transactions'])})")
                for i,tx in enumerate(params['transactions']):
                    st.markdown(tx_pill(tx, i+2), unsafe_allow_html=True)

            st.markdown("---")
            col1, col2 = st.columns([3,1])
            with col1:
                notes = st.text_input("Notes (optional)", placeholder="e.g. Quarterly review position")
            with col2:
                st.markdown("<br>", unsafe_allow_html=True)
                calc_btn = st.button("▶ Calculate & Save", type="primary", use_container_width=True)

            if calc_btn:
                with st.spinner("Calculating..."):
                    rows, summary = run_calc(params)
                rid = save_bond(params, params.get('transactions',[]), notes)
                st.session_state.edit_id = rid
                st.success(f"✅ Saved to portfolio (ID: {rid})")
                go("results")

# ════════════════════════════════════════════════════════════════════════════════
# CALCULATOR — New Bond + Transactions
# ════════════════════════════════════════════════════════════════════════════════
elif page == 'calculator':
    banner("New Bond", "Define bond terms and add buy / sell transactions", "🔢")

    # ── Step 1: Base bond terms ───────────────────────────────────────────────
    sec("Step 1 — Bond Terms (Initial Buy)")
    with st.form("base_form"):
        c1,c2 = st.columns(2)
        with c1: isin   = st.text_input("ISIN",   placeholder="e.g. XYZ001")
        with c2: issuer = st.text_input("Issuer", placeholder="e.g. Ministry of Finance")

        c1,c2,c3,c4 = st.columns(4)
        with c1: par_value   = st.number_input("Par / Face Value *",       value=1_000_000.0, step=100_000.0, format="%.2f")
        with c2: clean_price = st.number_input("Clean Trade Price (%) *",  value=95.0, step=0.0001, format="%.6f")
        with c3: coupon_rate = st.number_input("Coupon Rate (% annual) *", value=8.0,  step=0.001,  format="%.4f")
        with c4: freq = st.selectbox("Frequency *", [1,2,4,12],
                            format_func=lambda x:{1:"Annual",2:"Semi-annual",4:"Quarterly",12:"Monthly"}[x],
                            index=1)

        c1,c2,c3,c4 = st.columns(4)
        with c1: settle_date   = st.date_input("Settlement Date *",    value=dt_date(2024,1,15))
        with c2: maturity_date = st.date_input("Maturity Date *",      value=dt_date(2029,1,15))
        with c3: last_int_date = st.date_input("Last Interest Date *", value=dt_date(2023,7,15))
        with c4: next_int_date = st.date_input("Next Interest Date *", value=dt_date(2024,7,15))

        notes = st.text_area("Notes (optional)", height=50)
        save_base = st.form_submit_button("✅ Save Bond Terms", use_container_width=True, type="primary")

    if save_base:
        if maturity_date <= settle_date:
            st.error("Maturity date must be after settlement date.")
        elif next_int_date <= last_int_date:
            st.error("Next interest date must be after last interest date.")
        else:
            disc, prem = calc_discount_premium(float(par_value), float(clean_price))
            diff = float(par_value) * (float(clean_price)/100 - 1)
            st.session_state._bp = dict(
                isin=isin, issuer=issuer,
                par_value=float(par_value), clean_price=float(clean_price),
                coupon_rate=float(coupon_rate), interest_frequency=freq,
                settle_date=settle_date, last_interest_date=last_int_date,
                next_interest_date=next_int_date, maturity_date=maturity_date,
                discount=disc, premium=prem, coupon_dates=[],
            )
            st.session_state._notes = notes
            st.session_state.calc_txs = []
            st.rerun()

    # Show bond type hint live
    try:
        diff = float(par_value) * (float(clean_price)/100 - 1)
        if diff < 0:   st.markdown(f'🔴 **Discount Bond** — discount = `{diff:,.2f}`')
        elif diff > 0: st.markdown(f'🟢 **Premium Bond** — premium = `+{diff:,.2f}`')
        else:          st.markdown(f'🔵 **Par Bond**')
    except: pass

    # ── Step 2: Additional Transactions ──────────────────────────────────────
    if '_bp' in st.session_state:
        bp = st.session_state._bp
        st.markdown("---")
        sec(f"Step 2 — Additional Transactions  (ISIN: {bp.get('isin','—')})")

        # Show initial buy
        accrued0 = derive_accrued_interest(
            bp['par_value'], bp['coupon_rate'], bp['interest_frequency'],
            bp['settle_date'], bp['last_interest_date'], bp['next_interest_date']
        )
        st.markdown(
            f'<div class="tx-buy">🔵 <b>#1 INITIAL BUY</b> &nbsp;|&nbsp; '
            f'📅 {fmt_date(bp["settle_date"])} &nbsp;|&nbsp; '
            f'Nominal: <b>{bp["par_value"]:,.0f}</b> &nbsp;|&nbsp; '
            f'Price: {bp["clean_price"]:.6f}% &nbsp;|&nbsp; '
            f'Accrued at buy: {accrued0:,.4f}</div>',
            unsafe_allow_html=True
        )

        # Show existing extra transactions
        for i, tx in enumerate(st.session_state.calc_txs):
            col_tx, col_del = st.columns([10,1])
            with col_tx:
                st.markdown(tx_pill(tx, i+2), unsafe_allow_html=True)
            with col_del:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("✕", key=f"del_{i}", help="Remove this transaction"):
                    st.session_state.calc_txs.pop(i)
                    st.rerun()

        # Add transaction form
        st.markdown("##### ➕ Add Transaction")
        with st.form("add_tx", clear_on_submit=True):
            c1,c2,c3,c4,c5 = st.columns([2,2,2,2,3])
            with c1: tx_date  = st.date_input("Date",         value=dt_date(2024,7,15))
            with c2: tx_type  = st.selectbox("Type",          ["BUY","SELL_PARTIAL","SELL_FULL"])
            with c3: tx_nom   = st.number_input("Nominal",    value=1_000_000.0, step=100_000.0, format="%.2f")
            with c4: tx_price = st.number_input("Clean Price (%)", value=float(bp.get('clean_price',95)), step=0.0001, format="%.6f")
            with c5: tx_note  = st.text_input("Note",         placeholder="e.g. Buy #2 — secondary market")
            c1, c2 = st.columns([2,5])
            with c1: tx_acc = st.number_input("Accrued Interest (0 = auto-calculate)", value=0.0, format="%.4f")
            add_tx = st.form_submit_button("➕ Add", type="primary", use_container_width=True)

        if add_tx:
            st.session_state.calc_txs.append({
                'date':             tx_date,
                'type':             tx_type,
                'nominal':          float(tx_nom),
                'clean_price':      float(tx_price),
                'accrued_interest': float(tx_acc),
                'note':             tx_note,
            })
            st.rerun()

        # Summary before calculate
        n_buys  = sum(1 for t in st.session_state.calc_txs if t['type']=='BUY')
        n_sells = len(st.session_state.calc_txs) - n_buys
        total_face = bp['par_value'] + sum(t['nominal'] for t in st.session_state.calc_txs if t['type']=='BUY')
        sold_face  = sum(t['nominal'] for t in st.session_state.calc_txs if 'SELL' in t['type'])

        st.markdown("---")
        mcard_row([
            ("Total Transactions",  f"{1 + len(st.session_state.calc_txs)}",   "#1F4E79"),
            ("Total Buys",          f"{1 + n_buys}",                            "#2E75B6"),
            ("Total Sells",         f"{n_sells}",                               "#C00000"),
            ("Total Face Bought",   f"{total_face:,.0f}",                       "#1F4E79"),
            ("Total Face Sold",     f"{sold_face:,.0f}",                        "#C00000"),
        ])
        st.markdown("")

        if st.button("▶ Calculate Full Schedule", type="primary", use_container_width=True):
            params = {**bp, 'transactions': st.session_state.calc_txs}
            with st.spinner("Building multi-transaction schedule..."):
                rows, summary = run_calc(params)
            rid = save_bond(params, st.session_state.calc_txs,
                            st.session_state.get('_notes',''))
            st.session_state.edit_id = rid
            st.success(f"✅ Saved — ID: {rid}")
            go("results")

# ════════════════════════════════════════════════════════════════════════════════
# PORTFOLIO
# ════════════════════════════════════════════════════════════════════════════════
elif page == 'portfolio':
    banner("Portfolio History", "All saved bonds — search by ISIN, view results, add transactions", "📋")

    # Search bar
    c1,c2 = st.columns([5,1])
    with c1: isin_q = st.text_input("Search by ISIN", placeholder="Type ISIN or partial...", label_visibility="collapsed")
    with c2: st.button("🔍 Search", use_container_width=True)

    records = search_by_isin(isin_q) if isin_q else get_all()
    sec(f"{'Search Results' if isin_q else 'All Bonds'}  ({len(records)})")

    if not records:
        if isin_q:
            st.warning(f"No bonds found matching '{isin_q}'")
        else:
            st.info("No bonds saved yet. Use Upload or New Bond to get started.")
    else:
        for rec in records:
            p    = rec['params']
            disc, prem = calc_discount_premium(p.get('par_value',0), p.get('clean_price',100))
            bt   = 'Discount' if disc<0 else ('Premium' if prem>0 else 'Par')
            bcls = {'Discount':'disc','Premium':'prem','Par':'par'}[bt]
            txs  = rec.get('transactions',[])
            n_tx = len(txs) + 1
            n_b  = sum(1 for t in txs if t.get('type')=='BUY') + 1
            n_s  = sum(1 for t in txs if 'SELL' in str(t.get('type','')).upper())

            with st.container():
                st.markdown(f"""
                <div class="bcard">
                  <span style="font-weight:700;font-size:1rem;color:#1F4E79">{rec.get('isin','—')}</span>
                  &nbsp;<span class="badge {bcls}">{bt}</span>
                  &nbsp;<span style="font-size:.85rem;color:#555">{rec.get('issuer','')}</span>
                  &nbsp;&nbsp;<span style="font-size:.8rem;color:#888">
                    Par: {p.get('par_value',0):,.0f} &nbsp;·&nbsp;
                    Price: {p.get('clean_price',0):.4f}% &nbsp;·&nbsp;
                    Coupon: {p.get('coupon_rate',0):.3f}% &nbsp;·&nbsp;
                    {n_tx} transaction(s) ({n_b} buys · {n_s} sells)
                  </span><br>
                  <span style="font-size:.75rem;color:#aaa">
                    Settle: {fmt_date(p.get('settle_date'))} &nbsp;·&nbsp;
                    Maturity: {fmt_date(p.get('maturity_date'))} &nbsp;·&nbsp;
                    Saved: {rec.get('created_at','')[:10]} &nbsp;·&nbsp;
                    ID: <code>{rec['id']}</code>
                  </span>
                </div>
                """, unsafe_allow_html=True)

                c1,c2,c3,c4 = st.columns(4)
                with c1:
                    if st.button("📊 Results", key=f"v_{rec['id']}", use_container_width=True):
                        fp = get_full_params(rec)
                        with st.spinner("Calculating..."):
                            rows, summary = run_calc(fp)
                        st.session_state.edit_id = rec['id']
                        go("results")
                with c2:
                    if st.button("✏️ Add Transaction", key=f"e_{rec['id']}", use_container_width=True):
                        st.session_state.edit_id = rec['id']
                        go("edit_bond")
                with c3:
                    # Inline export
                    fp   = get_full_params(rec)
                    rows2 = build_cashflow_table(fp)
                    xls2  = export_to_excel(fp, rows2)
                    st.download_button("📥 Excel", data=xls2,
                        file_name=f"bond_{rec.get('isin','x')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{rec['id']}", use_container_width=True)
                with c4:
                    if st.button("🗑 Delete", key=f"d_{rec['id']}", use_container_width=True):
                        delete_bond(rec['id'])
                        st.rerun()

                st.markdown("<hr style='margin:.3rem 0;border-color:#F0F0F0'>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════════
# EDIT BOND — add / remove transactions
# ════════════════════════════════════════════════════════════════════════════════
elif page == 'edit_bond':
    rec = get_by_id(st.session_state.edit_id) if st.session_state.edit_id else None
    if not rec:
        st.error("Bond not found."); go("portfolio")

    p  = rec['params']
    txs = rec.get('transactions', [])

    banner(f"Edit Bond — {rec.get('isin','—')}",
           f"Issuer: {rec.get('issuer','—')} · Par: {p.get('par_value',0):,.0f} · "
           f"Settle: {fmt_date(p.get('settle_date'))} · Maturity: {fmt_date(p.get('maturity_date'))}", "✏️")

    sec("Current Transaction History")
    accrued0 = derive_accrued_interest(
        p['par_value'], p.get('coupon_rate',0), p.get('interest_frequency',2),
        p['settle_date'], p['last_interest_date'], p['next_interest_date']
    )
    st.markdown(
        f'<div class="tx-buy">🔵 <b>#1 INITIAL BUY (fixed)</b> &nbsp;|&nbsp; '
        f'📅 {fmt_date(p.get("settle_date"))} &nbsp;|&nbsp; '
        f'Nominal: <b>{p.get("par_value",0):,.0f}</b> &nbsp;|&nbsp; '
        f'Price: {p.get("clean_price",0):.6f}% &nbsp;|&nbsp; '
        f'Accrued: {accrued0:,.4f}</div>',
        unsafe_allow_html=True
    )

    for i, tx in enumerate(txs):
        col_tx, col_del = st.columns([10,1])
        with col_tx:
            st.markdown(tx_pill(tx, i+2), unsafe_allow_html=True)
        with col_del:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("✕", key=f"edel_{i}"):
                new_txs = [t for j,t in enumerate(txs) if j != i]
                update_bond(rec['id'], transactions=new_txs)
                st.rerun()

    st.markdown("---")
    sec("Add New Transaction")
    with st.form("edit_tx", clear_on_submit=True):
        c1,c2,c3,c4,c5 = st.columns([2,2,2,2,3])
        with c1: nt_date  = st.date_input("Date",        value=dt_date(2024,7,15))
        with c2: nt_type  = st.selectbox("Type",         ["BUY","SELL_PARTIAL","SELL_FULL"])
        with c3: nt_nom   = st.number_input("Nominal",   value=1_000_000.0, step=100_000.0, format="%.2f")
        with c4: nt_price = st.number_input("Clean Price (%)", value=float(p.get('clean_price',95)), step=0.0001, format="%.6f")
        with c5: nt_note  = st.text_input("Note",        placeholder="e.g. Buy #3 — tap")
        c1,_ = st.columns([2,5])
        with c1: nt_acc = st.number_input("Accrued Interest (0 = auto-calculate)", value=0.0, format="%.4f")
        add_btn = st.form_submit_button("➕ Add Transaction", type="primary", use_container_width=True)

    if add_btn:
        add_transaction(rec['id'], {
            'date': nt_date, 'type': nt_type,
            'nominal': float(nt_nom), 'clean_price': float(nt_price),
            'accrued_interest': float(nt_acc), 'note': nt_note,
        })
        st.success(f"Added {nt_type} {float(nt_nom):,.0f} on {nt_date}")
        st.rerun()

    st.markdown("---")
    c1,c2 = st.columns(2)
    with c1:
        if st.button("▶ Recalculate & View Results", type="primary", use_container_width=True):
            fresh = get_by_id(rec['id'])
            fp    = get_full_params(fresh)
            with st.spinner("Calculating..."):
                rows, summary = run_calc(fp)
            st.session_state.edit_id = rec['id']
            go("results")
    with c2:
        if st.button("← Back to Portfolio", use_container_width=True):
            go("portfolio")

# ════════════════════════════════════════════════════════════════════════════════
# RESULTS — schedule table + export, NO charts
# ════════════════════════════════════════════════════════════════════════════════
elif page == 'results':
    if not st.session_state.rows:
        st.warning("No bond loaded."); go("home")

    params  = st.session_state.params
    rows    = st.session_state.rows
    summary = st.session_state.summary
    bt      = summary['bond_type']
    bcls    = {'Discount':'disc','Premium':'prem','Par':'par'}[bt]

    banner(
        f"Results — {params.get('isin','Bond')}",
        f"Issuer: {params.get('issuer','—') or '—'}  ·  "
        f"Settle: {fmt_date(params.get('settle_date'))}  ·  "
        f"Maturity: {fmt_date(params.get('maturity_date'))}",
        "📊"
    )

    # ── Validation warnings ───────────────────────────────────────────────────
    for w in (st.session_state.warnings or []):
        st.warning(f"⚠️ {w}")

    # ── Key metrics ───────────────────────────────────────────────────────────
    sec("Key Metrics")
    mcard_row([
        ("Par Value (Buy #1)",    f"{params['par_value']:,.0f}",             "#1F4E79"),
        ("Initial Purchase Cost", f"{summary['purchase_cost']:,.2f}",         "#1F4E79"),
        ("Clean Price",           f"{params['clean_price']:.6f}%",            "#1F4E79"),
        ("Bond Type",             bt,
            "#C00000" if bt=="Discount" else "#375623" if bt=="Premium" else "#2E75B6"),
        ("Total Buys",            str(summary['total_buys']),                 "#2E75B6"),
        ("Total Sells",           str(summary['total_sells']),                "#C00000"),
        ("P&L (Sell)",            f"{summary['total_realized_pl']:+,.2f}",
            "#375623" if summary['total_realized_pl']>=0 else "#C00000"),
        ("Years to Maturity",     f"{summary['years_to_maturity']:.2f}",      "#1F4E79"),
    ])

    # ── P&L breakdown & credit risk ───────────────────────────────────────────
    st.markdown("")
    rii   = summary.get('realized_interest_income', 0) or 0
    tpl   = summary.get('total_pl', 0) or 0
    el    = summary.get('expected_loss')
    pd_in = summary.get('default_probability') or 0
    lgd_v = summary.get('lgd') or 0
    mcard_row([
        ("Realized Interest Income", f"{rii:+,.2f}",
            "#375623" if rii >= 0 else "#C00000"),
        ("Total P&L (Sells + Interest)", f"{tpl:+,.2f}",
            "#375623" if tpl >= 0 else "#C00000"),
        ("Default Probability (DRSK)", f"{pd_in:.6f}",                    "#1F4E79"),
        ("LGD (1 − CDS recovery)",     f"{lgd_v:.2f}",                    "#1F4E79"),
        ("Expected Loss",       f"{el:,.2f}" if el is not None else "—",  "#C00000"),
    ])

    # ── Yield & cashflow metrics ──────────────────────────────────────────────
    st.markdown("")
    ytm_v = summary.get('ytm')
    cy_v  = summary.get('current_yield')
    ncf_v = summary.get('net_cashflow', 0) or 0
    mcard_row([
        ("YTM at Purchase",     f"{ytm_v:.4f}%" if ytm_v is not None else "—", "#1F4E79"),
        ("Current Yield",       f"{cy_v:.4f}%"  if cy_v  is not None else "—", "#1F4E79"),
        ("Total Coupon Income", f"{summary.get('total_coupon_income',0):,.2f}", "#375623"),
        ("Total Invested",      f"{summary.get('total_invested',0):,.2f}",      "#1F4E79"),
        ("Sell Proceeds",       f"{summary.get('total_sell_proceeds',0):,.2f}", "#1F4E79"),
        ("Net Cashflow",        f"{ncf_v:+,.2f}",
            "#375623" if ncf_v >= 0 else "#C00000"),
        ("Current Holding",     f"{summary.get('current_nominal',0):,.0f}",     "#1F4E79"),
    ])

    d_val = summary.get('discount',0) or 0
    p_val = summary.get('premium',0)  or 0
    if d_val < 0 or p_val > 0:
        st.markdown("")
        mcard_row([
            (f"Initial {'Discount' if d_val<0 else 'Premium'}",
             f"{d_val:+,.2f}" if d_val<0 else f"+{p_val:,.2f}",
             "#C00000" if d_val<0 else "#375623"),
            ("Annual Coupon Income",  f"{summary['annual_income']:,.2f}",    "#1F4E79"),
            ("Coupon / Period",
             f"{summary['annual_income']/params['interest_frequency']:,.2f}", "#1F4E79"),
            ("Coupon Payments",       str(summary['coupon_payments']),        "#1F4E79"),
        ])

    # ── Transaction log ───────────────────────────────────────────────────────
    tx_rows = [r for r in rows if r.get('is_buy') or r.get('is_sell')]
    if tx_rows:
        st.markdown("")
        sec(f"Transaction Log  ({len(tx_rows)} transactions)")
        for r in tx_rows:
            if r.get('is_buy'):
                cls = 'tx-buy'; icon = '🔵'
            elif r.get('realized_pl',0) == 0 or r.get('label','').startswith('Full'):
                cls = 'tx-sell-full'; icon = '🔴'
            else:
                cls = 'tx-sell-partial'; icon = '🟡'
            detail = r.get('tx_detail') or r.get('label') or ''
            st.markdown(f'<div class="{cls}">{icon} {detail}</div>', unsafe_allow_html=True)

    # ── Schedule table ────────────────────────────────────────────────────────
    st.markdown("")
    sec("Cashflow & Amortization Schedule")

    c1,c2 = st.columns([3,2])
    with c1:
        f_opt = st.radio("Show rows:",
            ["All rows","Coupon dates only","Year-end only","Transactions only"],
            horizontal=True)
    with c2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"🟡 coupon &nbsp;·&nbsp; 🟢 maturity &nbsp;·&nbsp; 🔵 buy &nbsp;·&nbsp; 🔴 sell")

    df, disp, num_cols = build_df(rows, f_opt)
    col_fmt = {c:'{:,.4f}' for c in num_cols}
    col_fmt['Default Prob (DRSK)'] = '{:.6f}'
    col_fmt['LGD'] = '{:.2f}'
    st.dataframe(
        df.style.format(col_fmt, na_rep='—'),
        use_container_width=True, height=520
    )
    st.caption(f"{len(disp)} rows shown  ·  {len(rows)} total rows in schedule")

    # ── Actions ───────────────────────────────────────────────────────────────
    st.markdown("---")
    c1,c2,c3,c4 = st.columns(4)
    with c1:
        xls = export_to_excel(params, rows)
        isin_s = (params.get('isin') or 'bond').replace(' ','_')
        st.download_button("📥 Download Excel", data=xls,
            file_name=f"bond_{isin_s}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary")
    with c2:
        table_data = []
        for r in rows:
            lbl = r.get('label') or (fmt_date(r['date']) if r.get('date') else '')
            table_data.append({'Date':lbl,'Cashflow':r.get('cashflow'),
                'Nominal_Delta':r.get('nominal_change'),'Days':r.get('num_days'),
                'Bond_Discount':r.get('bond_discount'),'Cum_Amort_Disc':r.get('cum_amort_disc'),
                'Bond_Premium':r.get('bond_premium'),'Cum_Amort_Prem':r.get('cum_amort_prem'),
                'Carrying_Value':r.get('carrying_value'),'Accrued_Int':r.get('accrued_int'),
                'Nominal_Balance':r.get('nominal_balance'),'Price':r.get('price'),
                'MTM':r.get('mtm'),'OCI_GL':r.get('oci_gl'),'NAV':r.get('nav'),
                'Check':r.get('check'),'PL_Sell':r.get('realized_pl'),
                'Realized_Interest_Income':r.get('realized_interest_income'),
                'Total_PL':r.get('total_pl'),'WAC':r.get('wac'),'Check_2':r.get('check2'),
                'Default_Prob_DRSK':r.get('default_prob'),'LGD':r.get('lgd'),
                'Expected_Loss':r.get('expected_loss'),'Change':r.get('el_change')})
        df_csv = pd.DataFrame(table_data)
        st.download_button("📄 Download CSV", data=df_csv.to_csv(index=False),
            file_name=f"bond_{isin_s}.csv", mime="text/csv", use_container_width=True)
    with c3:
        if st.session_state.edit_id:
            if st.button("✏️ Add Transaction", use_container_width=True):
                go("edit_bond")
    with c4:
        if st.button("📋 Portfolio", use_container_width=True):
            go("portfolio")

# ════════════════════════════════════════════════════════════════════════════════
# EXPORT — standalone page kept for direct nav
# ════════════════════════════════════════════════════════════════════════════════
elif page == 'export':
    if not st.session_state.rows:
        st.warning("No bond loaded."); go("home")
    else:
        go("results")   # redirect — export is now inline on results page

else:
    go("home")
