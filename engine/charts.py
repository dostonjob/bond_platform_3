"""Charts — v2 (multi-transaction aware)"""
import plotly.graph_objects as go
from engine.calculator import fmt_date

NAVY='#1F4E79'; BLUE='#2E75B6'; GREEN='#375623'; RED='#C00000'; AMBER='#E8A000'; GRAY='#888'

def _layout(title=''):
    return dict(
        title=dict(text=title, font=dict(size=13, color=NAVY), x=0.01),
        plot_bgcolor='white', paper_bgcolor='white',
        font=dict(family='Arial', size=11, color='#333'),
        margin=dict(l=50, r=20, t=50, b=50),
        hovermode='x unified', legend=dict(orientation='h', y=-0.18),
        xaxis=dict(showgrid=False, zeroline=False, tickfont=dict(size=10)),
        yaxis=dict(showgrid=True, gridcolor='#F0F0F0', zeroline=False, tickfont=dict(size=10)),
    )

def carrying_value_chart(rows):
    sched = [r for r in rows if r.get('date') and not r.get('is_header') and not r.get('is_buy') and not r.get('is_sell')]
    buy_r = [r for r in rows if r.get('is_buy') and r.get('date')]
    sell_r= [r for r in rows if r.get('is_sell') and r.get('date')]
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=[r['date'] for r in sched], y=[r['carrying_value'] for r in sched],
        name='Carrying Value', line=dict(color=BLUE, width=2), mode='lines',
        hovertemplate='<b>%{x|%d-%b-%Y}</b><br>Carrying Value: %{y:,.2f}<extra></extra>',
    ))
    fig.add_trace(go.Scatter(
        x=[r['date'] for r in sched],
        y=[r['nav'] for r in sched if r.get('nav') is not None],
        name='NAV', line=dict(color=GREEN, width=1.5, dash='dot'), mode='lines',
        hovertemplate='NAV: %{y:,.2f}<extra></extra>',
    ))
    if buy_r:
        fig.add_trace(go.Scatter(
            x=[r['date'] for r in buy_r], y=[r.get('carrying_value') or 0 for r in buy_r],
            name='Buy', mode='markers',
            marker=dict(symbol='triangle-up', size=12, color=GREEN),
            hovertemplate='<b>BUY</b> %{x|%d-%b-%Y}<extra></extra>',
        ))
    if sell_r:
        fig.add_trace(go.Scatter(
            x=[r['date'] for r in sell_r], y=[0]*len(sell_r),
            name='Sell', mode='markers',
            marker=dict(symbol='triangle-down', size=12, color=RED),
            hovertemplate='<b>SELL</b> %{x|%d-%b-%Y}<extra></extra>',
        ))
    fig.update_layout(**_layout('Carrying Value & NAV — All Transactions'))
    return fig

def nominal_balance_chart(rows):
    dates, noms = [], []
    for r in rows:
        if r.get('date') and r.get('nominal_balance') is not None:
            dates.append(r['date']); noms.append(r['nominal_balance'])
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=dates, y=noms, name='Nominal Balance',
        fill='tozeroy', fillcolor='rgba(46,117,182,0.08)',
        line=dict(color=BLUE, width=2), mode='lines',
        hovertemplate='<b>%{x|%d-%b-%Y}</b><br>Nominal: %{y:,.0f}<extra></extra>',
    ))
    fig.update_layout(**_layout('Nominal Position Over Time'))
    return fig

def amortization_chart(rows, bond_type):
    sched = [r for r in rows if r.get('date') and not r.get('is_header') and not r.get('is_buy') and not r.get('is_sell')]
    if bond_type == 'Discount':
        vals  = [r['bond_discount'] for r in sched]
        color = RED; label = 'Bond Discount Remaining'
    else:
        vals  = [r['bond_premium'] for r in sched]
        color = GREEN; label = 'Bond Premium Remaining'
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=[r['date'] for r in sched], y=vals, name=label,
        marker_color=color, opacity=0.7,
        hovertemplate='<b>%{x|%d-%b-%Y}</b><br>' + label + ': %{y:,.4f}<extra></extra>',
    ))
    fig.update_layout(**_layout(f'{label} Over Time'))
    return fig

def accrued_interest_chart(rows):
    sched = [r for r in rows if r.get('date') and not r.get('is_header')]
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=[r['date'] for r in sched], y=[r.get('accrued_int', 0) for r in sched],
        name='Accrued Interest', fill='tozeroy',
        fillcolor='rgba(46,117,182,0.1)', line=dict(color=BLUE, width=1.5),
        hovertemplate='<b>%{x|%d-%b-%Y}</b><br>Accrued: %{y:,.4f}<extra></extra>',
    ))
    fig.update_layout(**_layout('Accrued Interest — Coupon Cycle'))
    return fig

def cashflow_timeline(rows):
    """All cashflows on a scatter timeline — buys red, coupons blue, sells green, maturity gold."""
    fig = go.Figure()
    groups = {
        'Buy':      ([], [], RED,   'triangle-up'),
        'Coupon':   ([], [], BLUE,  'circle'),
        'Sell':     ([], [], GREEN, 'triangle-down'),
        'Maturity': ([], [], AMBER, 'star'),
    }
    for r in rows:
        if not r.get('date') or r.get('cashflow') is None:
            continue
        if r.get('is_buy'):      g = 'Buy'
        elif r.get('is_sell'):   g = 'Sell'
        elif r.get('is_maturity'): g = 'Maturity'
        elif r.get('is_coupon'): g = 'Coupon'
        else: continue
        groups[g][0].append(r['date'])
        groups[g][1].append(r['cashflow'])

    for name, (x, y, color, sym) in groups.items():
        if x:
            fig.add_trace(go.Scatter(
                x=x, y=y, name=name, mode='markers',
                marker=dict(symbol=sym, size=11, color=color, line=dict(width=1, color='white')),
                hovertemplate=f'<b>{name}</b><br>%{{x|%d-%b-%Y}}<br>CF: %{{y:,.2f}}<extra></extra>',
            ))
    fig.update_layout(**_layout('Cashflow Timeline — All Events'))
    return fig

def realized_pl_chart(rows):
    sell_rows = [r for r in rows if r.get('is_sell') and r.get('realized_pl') is not None]
    if not sell_rows:
        return None
    fig = go.Figure()
    colors = [GREEN if r['realized_pl'] >= 0 else RED for r in sell_rows]
    fig.add_trace(go.Bar(
        x=[fmt_date(r['date']) for r in sell_rows],
        y=[r['realized_pl'] for r in sell_rows],
        marker_color=colors, name='Realized P&L',
        hovertemplate='<b>Sell %{x}</b><br>Realized P&L: %{y:,.2f}<extra></extra>',
    ))
    fig.add_hline(y=0, line_dash='dot', line_color=GRAY)
    fig.update_layout(**_layout('Realized P&L by Sell Transaction'))
    return fig

def summary_gauge(clean_price):
    fig = go.Figure(go.Indicator(
        mode='gauge+number+delta',
        value=clean_price,
        delta={'reference': 100, 'valueformat': '.4f'},
        title={'text': 'Clean Price vs Par (100)', 'font': {'size': 12, 'color': NAVY}},
        number={'suffix': '%', 'valueformat': '.4f'},
        gauge={
            'axis': {'range': [max(80, clean_price-20), min(120, clean_price+20)]},
            'bar': {'color': BLUE},
            'steps': [{'range':[80,100],'color':'#FDECEA'},{'range':[100,120],'color':'#E8F5E9'}],
            'threshold': {'line':{'color':GRAY,'width':2},'thickness':0.8,'value':100},
        },
    ))
    fig.update_layout(paper_bgcolor='white', margin=dict(l=20,r=20,t=60,b=20), height=210,
                      font=dict(family='Arial',size=11))
    return fig
