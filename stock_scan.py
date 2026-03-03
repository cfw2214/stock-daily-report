#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
美股每日開盤掃描報告 Powered by 黑叔 — v5.2
量比 + EMA15/50/100 合併欄 | SPY/QQQ EMA技術分析建議
使用 yfinance 抓取真實股價、期權 OI、IV、P/C Ratio、EMA 數據
執行方式：python3 stock_scan.py
依賴套件：pip3 install yfinance pandas requests
"""

import yfinance as yf
import pandas as pd
import math
import os
import sys
from datetime import datetime, timedelta

# ═══════════════════════════════════════════════════════════════
#  設定區 — 依需求修改
# ═══════════════════════════════════════════════════════════════

TICKERS = ['NVDA', 'GOOGL', 'AAPL', 'MSFT', 'AMZN', 'META', 'TSLA',
           'TSM', 'AMD', 'MU', 'AVGO', 'SNDK', 'NFLX', 'LITE', 'COHR']

MAG7 = {'AAPL', 'MSFT', 'NVDA', 'AMZN', 'GOOGL', 'META', 'TSLA'}

MORNINGSTAR = {
    'AAPL':  {'fv': 260,  'stars': 3, 'moat': 'Wide'},
    'MSFT':  {'fv': 600,  'stars': 4, 'moat': 'Wide'},
    'NVDA':  {'fv': 232,  'stars': 4, 'moat': 'Wide'},
    'AMZN':  {'fv': 260,  'stars': 4, 'moat': 'Wide'},
    'GOOGL': {'fv': 340,  'stars': 3, 'moat': 'Wide'},
    'META':  {'fv': 850,  'stars': 4, 'moat': 'Wide'},
    'TSLA':  {'fv': 400,  'stars': 3, 'moat': 'Narrow'},
    'TSM':   {'fv': 428,  'stars': 3, 'moat': 'Wide'},
    'MU':    {'fv': 115,  'stars': 4, 'moat': 'None'},
    'SNDK':  {'fv': 138,  'stars': 5, 'moat': 'None'},
    'NFLX':  {'fv': 79,   'stars': 3, 'moat': 'Narrow'},
    'AMD':   {'fv': 145,  'stars': 3, 'moat': 'Narrow'},
    'AVGO':  {'fv': 258,  'stars': 3, 'moat': 'Wide'},
    'LITE':  {'fv': 90,   'stars': 3, 'moat': 'Narrow'},
    'COHR':  {'fv': 85,   'stars': 3, 'moat': 'Narrow'},
}

OUTPUT_DIR = os.environ.get('OUTPUT_DIR', os.path.expanduser('~/Desktop'))


# ═══════════════════════════════════════════════════════════════
#  工具函數
# ═══════════════════════════════════════════════════════════════

def get_four_weekly_expiries():
    today = datetime.now().date()
    days_to_friday = (4 - today.weekday()) % 7
    this_friday = today + timedelta(days=days_to_friday)
    results = []
    for i in range(4):
        friday = this_friday + timedelta(weeks=i)
        label  = ['本週', '下週', '下下週', '下下下週'][i]
        y, m   = friday.year, friday.month
        first  = datetime(y, m, 1).date()
        first_fri = first + timedelta(days=(4 - first.weekday()) % 7)
        third_fri = first_fri + timedelta(weeks=2)
        results.append({'label': label, 'date': friday.strftime('%Y-%m-%d'),
                        'is_monthly': (friday == third_fri)})
    return results


def calc_gex(calls_df, puts_df, current_price):
    try:
        import math as _m
        S = current_price
        if not S or S <= 0:
            return None
        def bs_gamma(S, K, T, sigma):
            if T <= 0 or sigma <= 0: return 0.0
            try:
                d1 = (_m.log(S/K) + 0.5*sigma**2*T) / (sigma*_m.sqrt(T))
                return _m.exp(-0.5*d1**2) / (_m.sqrt(2*_m.pi)*S*sigma*_m.sqrt(T))
            except: return 0.0
        T = 5/252
        def side(df):
            tot = 0.0
            for _, row in df.iterrows():
                K  = float(row.get('strike', 0) or 0)
                oi = float(row.get('openInterest', 0) or 0)
                iv = float(row.get('impliedVolatility', 0) or 0)
                if K <= 0 or oi <= 0: continue
                if iv <= 0: iv = 0.50
                tot += oi * bs_gamma(S, K, T, iv) * 100 * S * S
            return tot
        lo, hi = S*0.70, S*1.30
        c = calls_df[(calls_df['strike']>=lo)&(calls_df['strike']<=hi)].copy()
        p = puts_df[(puts_df['strike']>=lo)&(puts_df['strike']<=hi)].copy()
        return round((side(c) - side(p)) / 1e6, 1)
    except: return None


def calc_max_pain(calls_df, puts_df, current_price=None):
    try:
        c, p = calls_df.copy(), puts_df.copy()
        if current_price and current_price > 0:
            lo, hi = current_price*0.80, current_price*1.20
            c = c[(c['strike']>=lo)&(c['strike']<=hi)]
            p = p[(p['strike']>=lo)&(p['strike']<=hi)]
        strikes = sorted(set(c['strike'].tolist() + p['strike'].tolist()))
        if not strikes: return None
        cm = dict(zip(c['strike'], c['openInterest'].fillna(0)))
        pm = dict(zip(p['strike'], p['openInterest'].fillna(0)))
        min_pain, best = float('inf'), strikes[len(strikes)//2]
        for s in strikes:
            tot = (sum(max(0,s-k)*oi for k,oi in cm.items()) +
                   sum(max(0,k-s)*oi for k,oi in pm.items()))
            if tot < min_pain: min_pain, best = tot, s
        return best
    except: return None


def calc_ivr(ticker_obj, current_iv):
    try:
        hist = ticker_obj.history(period='1y')
        if hist.empty or len(hist) < 30: return 50
        ret = hist['Close'].pct_change().dropna()
        vs  = (ret.rolling(21).std() * math.sqrt(252) * 100).dropna()
        vmin, vmax = vs.min(), vs.max()
        if vmax == vmin: return 50
        return max(0, min(100, int((current_iv - vmin)/(vmax - vmin)*100)))
    except: return 50


def fmt_cap(mc):
    if not mc or math.isnan(mc): return '—'
    if mc >= 1e12: return f'${mc/1e12:.2f}T'
    if mc >= 1e9:  return f'${mc/1e9:.1f}B'
    return f'${mc/1e6:.0f}M'


def fmt_pe(val):
    if val is None or (isinstance(val, float) and math.isnan(val)): return None
    return round(float(val), 1)


# ═══════════════════════════════════════════════════════════════
#  EMA 計算 + 訊號引擎（共用）
# ═══════════════════════════════════════════════════════════════

def _compute_emas(closes):
    """計算 EMA15/50/100 及前一日值，回傳 dict"""
    n    = len(closes)
    e15  = closes.ewm(span=15,  adjust=False).mean()
    e50  = closes.ewm(span=50,  adjust=False).mean()
    e100 = closes.ewm(span=100, adjust=False).mean()
    return {
        'ema15':      round(float(e15.iloc[-1]),  2),
        'ema15_prev': round(float(e15.iloc[-2]),  2) if n >= 2  else round(float(e15.iloc[-1]),  2),
        'ema50':      round(float(e50.iloc[-1]),  2) if n >= 50 else None,
        'ema50_prev': round(float(e50.iloc[-2]),  2) if n >= 51 else None,
        'ema100':     round(float(e100.iloc[-1]), 2) if n >= 100 else None,
        'prev_close': round(float(closes.iloc[-2]),2) if n >= 2  else round(float(closes.iloc[-1]),2),
    }


def _ema_signals(price, prev_close, e15, e15_prev, e50, e50_prev, e100):
    """
    回傳 list of (priority, icon, css_class, title, note)，按優先度排序
    Priority: 1=黃金/死亡交叉, 2=多/空頭排列+短強長弱, 3=突破/跌破, 4=貼近
    """
    if price is None or e15 is None:
        return []
    sigs = []

    # 1. EMA15 × EMA50 交叉
    if e50 and e15_prev and e50_prev:
        if e15_prev < e50_prev and e15 >= e50:
            sigs.append((1,'🌟','sig-bull','EMA15 黃金交叉 EMA50','短線轉多，動能增強'))
        elif e15_prev > e50_prev and e15 <= e50:
            sigs.append((1,'💀','sig-bear','EMA15 死亡交叉 EMA50','短線轉空，留意下行'))

    # 2. 排列
    if e50 and e100:
        if e15 > e50 > e100:
            sigs.append((2,'📈','sig-bull','多頭排列','EMA15 > EMA50 > EMA100'))
        elif e15 < e50 < e100:
            sigs.append((2,'📉','sig-bear','空頭排列','EMA15 < EMA50 < EMA100'))
        elif e15 > e50 and e50 < e100:
            sigs.append((2,'⚡','sig-warn','短強長弱','短期反彈但中長期仍偏弱'))
        elif e15 < e50 and e50 > e100:
            sigs.append((2,'⚡','sig-warn','短弱長強','短期回落，中長期趨勢偏多'))

    # 3. 突破 / 跌破 EMA15
    if prev_close and e15_prev:
        if prev_close < e15_prev and price >= e15:
            sigs.append((3,'⬆','sig-bull','突破 EMA15','短期轉強訊號'))
        elif prev_close > e15_prev and price < e15:
            sigs.append((3,'⬇','sig-bear','跌破 EMA15','短期轉弱，留意加速'))

    # 4. 貼近 EMA15（±0.8%）—— 無更強訊號才顯示
    if not sigs and e15:
        pct = (price - e15) / e15 * 100
        if abs(pct) <= 0.8:
            chg = price - (prev_close or price)
            if chg >= 0:
                sigs.append((4,'🔄','sig-bull','回測EMA15回彈','貼近EMA15且收正，關注確認'))
            else:
                sigs.append((4,'🔄','sig-bear','貼近EMA15回落','貼近EMA15收黑，注意支撐有效性'))

    return sorted(sigs, key=lambda x: x[0])


def _ema_advice(e15, e50, e100):
    """根據 EMA 排列生成文字建議，回傳 (css_class, text)"""
    if not e15 or not e50 or not e100:
        return ('etf-neu', '數據不足，無法完整判斷趨勢排列。')
    if e15 > e50 > e100:
        return ('etf-bull',
                '✅ 多頭排列完整（EMA15 > EMA50 > EMA100），趨勢偏強。'
                ' EMA15 為短線支撐，持多宜守 EMA15 上方；若回測 EMA50 且守穩，可視為加倉機會。')
    if e15 < e50 < e100:
        return ('etf-bear',
                '🔴 空頭排列（EMA15 < EMA50 < EMA100），趨勢偏弱。'
                ' 反彈至 EMA15～EMA50 區間視為壓力帶，逢反彈輕多重空；'
                ' 跌破近期低點需嚴格止損。')
    if e15 > e50 and e50 < e100:
        return ('etf-warn',
                '⚠️ 短期走強但中長期仍偏弱。EMA50 尚未站回 EMA100，存在假突破風險。'
                ' 需等待 EMA50 有效突破 EMA100 後再確認多頭趨勢。')
    if e15 < e50 and e50 > e100:
        return ('etf-warn',
                '⚠️ 短期轉弱但中長期偏強。觀察 EMA15 能否重新站回 EMA50，'
                ' 若站回可視為回檔買點；若持續走弱並跌破 EMA100，需重新評估方向。')
    return ('etf-neu',
            '🔘 EMA 均線纏繞，方向未明。等待有效突破或跌破形成方向，'
            ' 暫以 EMA50 為多空分界線，區間震盪操作。')


# ═══════════════════════════════════════════════════════════════
#  數據抓取
# ═══════════════════════════════════════════════════════════════

def fetch_vix():
    result = {'value': 18.5, 'change': 0.0, 'pct': 0.0}
    try:
        hist = yf.Ticker('^VIX').history(period='2d')
        if len(hist) >= 2:
            now = round(hist['Close'].iloc[-1], 2)
            prev = round(hist['Close'].iloc[-2], 2)
            chg = round(now - prev, 2)
            result.update({'value': now, 'change': chg, 'pct': round(chg/prev*100, 2)})
    except: pass
    return result


def fetch_stock(ticker):
    print(f'  [{ticker}] 抓取中...')
    d = {'ticker': ticker, 'ok': False}
    try:
        t  = yf.Ticker(ticker)
        info = t.info
        hs   = t.history(period='5d')
        hl   = t.history(period='1y')
        if hs.empty: return d

        price = float(hs['Close'].iloc[-1])
        prev  = float(hs['Close'].iloc[-2]) if len(hs) > 1 else price
        d.update({
            'price':      price,
            'change':     round(price - prev, 2),
            'change_pct': round((price - prev)/prev*100, 2),
            'company':    info.get('shortName', ticker),
            'market_cap': info.get('marketCap'),
            'pe_ttm':     fmt_pe(info.get('trailingPE')),
            'pe_fwd':     fmt_pe(info.get('forwardPE')),
        })
        avg_vol = hs['Volume'].iloc[:-1].mean()
        d['vol_ratio'] = round(hs['Volume'].iloc[-1]/avg_vol, 2) if avg_vol else 1.0

        closes = hl['Close'] if not hl.empty else hs['Close']
        d.update(_compute_emas(closes))

        WEEK_KEYS = ['w0','w1','w2','w3']
        for k in WEEK_KEYS:
            d.update({f'{k}_expiry':None, f'{k}_put_wall':None, f'{k}_call_wall':None,
                      f'{k}_max_pain':None, f'{k}_gex':None,
                      f'{k}_is_monthly':False, f'{k}_label':''})
        d.update({'iv':None,'ivr':None,'pc_ratio':None})
        for wk in WEEK_KEYS:
            d.update({f'{wk}_target_call_strike':None, f'{wk}_target_call_expiry':None,
                      f'{wk}_target_put_strike':None,  f'{wk}_target_put_expiry':None})
        d.update({'target_call_strike':None,'target_call_expiry':None,
                  'target_put_strike':None, 'target_put_expiry':None})

        def _fetch_oi(target_str, avail):
            if not avail or not target_str: raise ValueError('無到期日')
            tdt  = datetime.strptime(target_str,'%Y-%m-%d')
            best = min(avail, key=lambda e: abs((datetime.strptime(e,'%Y-%m-%d')-tdt).days))
            ch   = t.option_chain(best)
            c, p = ch.calls.copy(), ch.puts.copy()
            c['openInterest'] = pd.to_numeric(c['openInterest'],errors='coerce').fillna(0)
            p['openInterest'] = pd.to_numeric(p['openInterest'],errors='coerce').fillna(0)
            res = {'expiry': best}
            pb = p[(p['strike']<=price*1.02)&(p['strike']>=price*0.70)]
            ca = c[(c['strike']>=price*0.98)&(c['strike']<=price*1.30)]
            if not pb.empty: res['put_wall']  = float(pb.loc[pb['openInterest'].idxmax(),'strike'])
            if not ca.empty: res['call_wall'] = float(ca.loc[ca['openInterest'].idxmax(),'strike'])
            res['max_pain'] = calc_max_pain(c, p, current_price=price)
            return res, c, p

        last_calls = None
        last_puts  = pd.DataFrame()

        try:
            avail      = t.options
            four_weeks = get_four_weekly_expiries()
            for i, wk in enumerate(four_weeks):
                k = WEEK_KEYS[i]
                d[f'{k}_label']      = wk['label']
                d[f'{k}_is_monthly'] = wk['is_monthly']
                try:
                    res, cdf, pdf = _fetch_oi(wk['date'], avail)
                    d[f'{k}_expiry']    = res.get('expiry')
                    d[f'{k}_put_wall']  = res.get('put_wall')
                    d[f'{k}_call_wall'] = res.get('call_wall')
                    d[f'{k}_max_pain']  = res.get('max_pain')
                    d[f'{k}_gex']       = calc_gex(cdf, pdf, price)
                    last_calls = cdf; last_puts = pdf
                    mo = '★月結' if wk['is_monthly'] else ''
                    print(f'    {wk["label"]}{mo} ({d[f"{k}_expiry"]})  '
                          f'Put牆:{d[f"{k}_put_wall"]}  Pain:{d[f"{k}_max_pain"]}  Call牆:{d[f"{k}_call_wall"]}')
                    try:
                        es = res.get('expiry','')
                        wp = WEEK_KEYS[i]+'_'
                        cv = cdf.copy()
                        cv['volume'] = pd.to_numeric(cv['volume'],errors='coerce').fillna(0)
                        cv['openInterest'] = pd.to_numeric(cv['openInterest'],errors='coerce').fillna(0)
                        oc = cv[(cv['strike']>price*1.01)&(cv['strike']<=price*1.30)&(cv['openInterest']>0)]
                        if not oc.empty:
                            oc = oc.copy()
                            oc['score'] = (oc['volume']/oc['openInterest']*oc['volume']**0.5
                                           if oc['volume'].sum()>0 else oc['openInterest'])
                            bc = oc.loc[oc['score'].idxmax()]
                            d[f'{wp}target_call_strike'] = float(bc['strike'])
                            d[f'{wp}target_call_expiry'] = es
                            if i==0: d['target_call_strike']=d[f'{wp}target_call_strike']; d['target_call_expiry']=es
                        pv = pdf.copy()
                        pv['volume'] = pd.to_numeric(pv['volume'],errors='coerce').fillna(0)
                        pv['openInterest'] = pd.to_numeric(pv['openInterest'],errors='coerce').fillna(0)
                        op = pv[(pv['strike']<price*0.99)&(pv['strike']>=price*0.70)&(pv['openInterest']>0)]
                        if not op.empty:
                            op = op.copy()
                            op['score'] = (op['volume']/op['openInterest']*op['volume']**0.5
                                           if op['volume'].sum()>0 else op['openInterest'])
                            bp = op.loc[op['score'].idxmax()]
                            d[f'{wp}target_put_strike'] = float(bp['strike'])
                            d[f'{wp}target_put_expiry'] = es
                            if i==0: d['target_put_strike']=d[f'{wp}target_put_strike']; d['target_put_expiry']=es
                        print(f'    {wk["label"]}OTM Call:{d[f"{wp}target_call_strike"]}  Put:{d[f"{wp}target_put_strike"]}')
                    except Exception as e: print(f'    OTM錯誤[{i}]: {e}')
                except Exception as e: print(f'    {wk["label"]}期權錯誤 {ticker}: {e}')

            if last_calls is not None:
                atm = last_calls.iloc[(last_calls['strike']-price).abs().argsort()[:3]]
                ivv = pd.to_numeric(atm['impliedVolatility'],errors='coerce').dropna()
                if not ivv.empty:
                    d['iv']  = round(float(ivv.mean())*100, 1)
                    d['ivr'] = calc_ivr(t, d['iv'])
                cv = pd.to_numeric(last_calls['volume'],errors='coerce').sum()
                pv = pd.to_numeric(last_puts['volume'], errors='coerce').sum() if not last_puts.empty else 0
                if cv and cv>0: d['pc_ratio'] = round(float(pv)/float(cv), 2)
        except Exception as e: print(f'    期權錯誤 {ticker}: {e}')

        d['ok'] = True
    except Exception as e: print(f'  錯誤 {ticker}: {e}')
    return d


def fetch_spy_qqq():
    """SPY / QQQ — EMA15/50/100 + 四週期權 OI"""
    result     = {}
    four_weeks = get_four_weekly_expiries()
    WEEK_KEYS  = ['w0','w1','w2','w3']

    for sym in ['SPY','QQQ']:
        try:
            t    = yf.Ticker(sym)
            hist = t.history(period='1y')
            if hist.empty: continue
            closes = hist['Close']
            price  = float(closes.iloc[-1])
            prev   = float(closes.iloc[-2])
            entry  = {
                'price':      round(price,2),
                'change_pct': round((price-prev)/prev*100,2),
            }
            entry.update(_compute_emas(closes))

            for k in WEEK_KEYS:
                entry.update({f'{k}_expiry':None, f'{k}_put_wall':None,
                               f'{k}_call_wall':None, f'{k}_max_pain':None,
                               f'{k}_is_monthly':False, f'{k}_label':''})
            avail = t.options

            def _fetch_etf(tgt):
                if not avail or not tgt: raise ValueError('無到期日')
                tdt  = datetime.strptime(tgt,'%Y-%m-%d')
                best = min(avail,key=lambda e:abs((datetime.strptime(e,'%Y-%m-%d')-tdt).days))
                ch   = t.option_chain(best)
                c, p = ch.calls.copy(), ch.puts.copy()
                c['openInterest']=pd.to_numeric(c['openInterest'],errors='coerce').fillna(0)
                p['openInterest']=pd.to_numeric(p['openInterest'],errors='coerce').fillna(0)
                res = {'expiry':best}
                pb = p[(p['strike']<=price*1.02)&(p['strike']>=price*0.70)]
                ca = c[(c['strike']>=price*0.98)&(c['strike']<=price*1.30)]
                if not pb.empty: res['put_wall']  = float(pb.loc[pb['openInterest'].idxmax(),'strike'])
                if not ca.empty: res['call_wall'] = float(ca.loc[ca['openInterest'].idxmax(),'strike'])
                res['max_pain'] = calc_max_pain(c,p,current_price=price)
                return res,c,p

            for i,wk in enumerate(four_weeks):
                k=WEEK_KEYS[i]
                entry[f'{k}_label']=wk['label']; entry[f'{k}_is_monthly']=wk['is_monthly']
                try:
                    res,_,_=_fetch_etf(wk['date'])
                    entry[f'{k}_expiry']=res.get('expiry')
                    entry[f'{k}_put_wall']=res.get('put_wall')
                    entry[f'{k}_call_wall']=res.get('call_wall')
                    entry[f'{k}_max_pain']=res.get('max_pain')
                    mo='★月結' if wk['is_monthly'] else ''
                    print(f'    {sym} {wk["label"]}{mo} ({entry[f"{k}_expiry"]}) Put:{entry[f"{k}_put_wall"]} Pain:{entry[f"{k}_max_pain"]} Call:{entry[f"{k}_call_wall"]}')
                except Exception as e: print(f'    {sym} {wk["label"]}錯誤:{e}')
            result[sym]=entry
        except: pass
    return result


# ═══════════════════════════════════════════════════════════════
#  HTML 渲染函數
# ═══════════════════════════════════════════════════════════════

def vol_cell_html(d):
    """量比數值 + 建議文字，用於現價 td"""
    vol = d.get('vol_ratio', 1.0)
    if vol >= 2.0:
        v_html = (f'<span class="vol-hi">{vol:.2f}x</span>'
                  f'<span class="vb">放量</span>')
        tip = '<div class="vol-tip vol-tip-hi">異常放量，注意突破或出貨</div>'
    elif vol >= 1.5:
        v_html = (f'<span class="vol-hi">{vol:.2f}x</span>'
                  f'<span class="vb">放量</span>')
        tip = '<div class="vol-tip vol-tip-hi">量能偏強，動能確認中</div>'
    elif vol >= 1.2:
        v_html = f'<span class="vol-mid">{vol:.2f}x</span>'
        tip = '<div class="vol-tip vol-tip-mid">略微放量，觀察方向</div>'
    else:
        v_html = f'<span class="vol-norm">{vol:.2f}x</span>'
        tip = '<div class="vol-tip vol-tip-norm">量能平淡，趨勢延續</div>'
    return (f'<div style="font-size:0.62em;color:#8b949e;letter-spacing:0.8px;margin-bottom:2px">量比</div>'
            f'<div class="vol-val">{v_html}</div>'
            f'{tip}')


def ema_cell_html(d):
    """EMA15/50/100 三行 + 趨勢訊號，用於獨立 td"""
    price      = d.get('price')
    e15        = d.get('ema15')
    e15_prev   = d.get('ema15_prev', e15)
    e50        = d.get('ema50')
    e50_prev   = d.get('ema50_prev', e50)
    e100       = d.get('ema100')
    prev_close = d.get('prev_close', price)

    def ema_row(label, ema):
        if ema is None:
            return (f'<div class="er">'
                    f'<span class="el">{label}</span>'
                    f'<span class="man">N/A</span></div>')
        pct   = (price - ema) / ema * 100
        cls   = 'mau' if pct >= 0 else 'mad'
        arrow = '▲' if pct >= 0 else '▼'
        sign  = '+' if pct >= 0 else ''
        return (f'<div class="er">'
                f'<span class="el">{label}</span>'
                f'<span class="{cls} ev">${ema:.2f}</span>'
                f'<span class="{cls} ep">{arrow}{sign}{pct:.1f}%</span>'
                f'</div>')

    ema_rows = (ema_row('EMA15', e15) +
                ema_row('EMA50', e50) +
                ema_row('EMA100', e100))

    sigs = _ema_signals(price, prev_close, e15, e15_prev, e50, e50_prev, e100)
    if sigs:
        _, icon, cls, title, note = sigs[0]
        sig_html = (f'<div class="ema-sig">'
                    f'<span class="{cls}">{icon} {title}</span>'
                    f'<span class="sig-note">{note}</span>'
                    f'</div>')
    else:
        sig_html = ''

    return f'{ema_rows}{sig_html}'


def vol_ema_cell_html(d):
    """舊版相容，不再使用"""
    return vol_cell_html(d)


def etf_ema_block_html(etf_data):
    """SPY/QQQ 卡片：EMA15/50/100 數值 + 最強訊號 + 建議文字"""
    price      = etf_data.get('price', 0)
    e15        = etf_data.get('ema15')
    e15_prev   = etf_data.get('ema15_prev', e15)
    e50        = etf_data.get('ema50')
    e50_prev   = etf_data.get('ema50_prev', e50)
    e100       = etf_data.get('ema100')
    prev_close = etf_data.get('prev_close', price)

    def ema_row(label, ema):
        if ema is None:
            return (f'<div class="mkr">'
                    f'<span class="k">{label}</span>'
                    f'<span class="v" style="color:#484f58">N/A</span></div>')
        pct   = (price - ema) / ema * 100
        cls   = 'pos' if pct >= 0 else 'neg'
        arrow = '▲ 站上' if pct >= 0 else '▼ 跌破'
        sign  = '+' if pct >= 0 else ''
        return (f'<div class="mkr">'
                f'<span class="k">{label}</span>'
                f'<span class="v {cls}">${ema:.2f} {arrow} ({sign}{pct:.1f}%)</span>'
                f'</div>')

    rows = ema_row('EMA15', e15) + ema_row('EMA50', e50) + ema_row('EMA100', e100)

    # 訊號
    sigs = _ema_signals(price, prev_close, e15, e15_prev, e50, e50_prev, e100)
    if sigs:
        _, icon, cls, title, note = sigs[0]
        sig_div = (f'<div class="etf-sig">'
                   f'<span class="{cls}">{icon} {title}</span>'
                   f'<span class="sig-note" style="margin-left:6px">{note}</span>'
                   f'</div>')
    else:
        sig_div = ''

    # 建議文字
    advice_cls, advice_txt = _ema_advice(e15, e50, e100)

    return (f'<div class="etf-ema-blk">'
            f'<div class="oi-lbl" style="margin-bottom:7px;letter-spacing:0.6px">📐 EMA 均線分析</div>'
            f'{rows}'
            f'{sig_div}'
            f'<div class="etf-advice {advice_cls}">{advice_txt}</div>'
            f'</div>')


def pe_html(pe_ttm, pe_fwd):
    if pe_ttm is None:
        ttm = '<div class="pe pe-l">虧損中</div><div class="pe-note neg">EPS 為負</div>'
    elif pe_ttm > 200:
        ttm = f'<div class="pe pe-w">{pe_ttm}x</div><div class="pe-note" style="color:#e3b341">⚠ 極高估值</div>'
    else:
        ttm = f'<div class="pe">{pe_ttm}x</div><div class="pe-note" style="color:#484f58">TTM</div>'
    if pe_fwd is None:
        fwd = '<div class="fpe-fl">—</div>'
    elif pe_ttm is not None and pe_fwd < pe_ttm:
        col = '#e3b341' if pe_fwd > 100 else '#3fb950'
        fwd = f'<div class="fpe-dn" style="color:{col}">{pe_fwd}x</div><div class="pe-note pos">↓ 獲利成長</div>'
    else:
        fwd = f'<div class="fpe-fl">{pe_fwd}x</div><div class="pe-note gf">≈ 平穩</div>'
    return ttm, fwd


def iv_html(iv, ivr):
    if iv is None or ivr is None:
        return '<div class="ivr-lbl">N/A</div>', '<span class="ivb ivn">—</span>'
    if ivr < 25:   bc,bd,bt = 'ivrl','ivs','低 IVR'
    elif ivr < 50: bc,bd,bt = 'ivrm','ivn','正常'
    elif ivr < 75: bc,bd,bt = 'ivrm','ive','偏高 ⚠'
    else:          bc,bd,bt = 'ivrh','ivh','極高 🔥'
    bar   = f'<div class="ivr-bg"><div class="ivr-f {bc}" style="width:{ivr}%"></div></div><div class="ivr-lbl">{iv:.1f}% · IVR {ivr}%</div>'
    badge = f'<span class="ivb {bd}">{bt}</span>'
    return bar, badge


def pc_html(pc):
    if pc is None: return '<div class="pc-n">—</div>'
    cls,note = (('pc-b','看多傾向') if pc<0.7 else ('pc-r','避險需求') if pc>1.0 else ('pc-n','中性'))
    return f'<div class="{cls}">{pc}</div><div class="pc-l">{note}</div>'


def ms_html(ticker, price):
    ms = MORNINGSTAR.get(ticker, {})
    if not ms: return '<div class="ms-fv">—</div>'
    fv,stars,moat = ms['fv'], ms['stars'], ms['moat']
    disc = round((1 - price/fv)*100)
    ph   = (f'<span class="ms-disc">▼{disc}%折</span>'  if disc>0 else
            f'<span class="ms-prem">▲{abs(disc)}%溢</span>' if disc<0 else
            '<span class="ms-neu">合理</span>')
    mc = {'Wide':'ms-mw','Narrow':'ms-mn','None':'ms-m0'}.get(moat,'ms-m0')
    mt = {'Wide':'Wide 護城河','Narrow':'Narrow 護城河','None':'No Moat'}.get(moat,'—')
    return (f'<div class="ms-fv">${fv} {ph}</div>'
            f'<div class="ms-str">{"★"*stars}{"☆"*(5-stars)}</div>'
            f'<span class="{mc}">{mt}</span>')


def target_zone_html(d, pfx='w0_'):
    cs = d.get(f'{pfx}target_call_strike')
    ps = d.get(f'{pfx}target_put_strike')
    if cs is None and ps is None:
        return '<div class="tz-wrap"><div style="color:#484f58;font-size:0.7em">目標：無訊號</div></div>'
    fmt = lambda v: f'${v:g}' if v is not None else '—'
    ch = f'<div class="tz-call">🎯↑ {fmt(cs)}</div>' if cs else '<div class="tz-call" style="color:#484f58">🎯↑ —</div>'
    ph = f'<div class="tz-put">🎯↓ {fmt(ps)}</div>'  if ps else '<div class="tz-put"  style="color:#484f58">🎯↓ —</div>'
    return f'<div class="tz-wrap">{ch}{ph}</div>'


def oi_html_single(d, pfx, label=''):
    cw    = d.get(f'{pfx}_call_wall')
    mp    = d.get(f'{pfx}_max_pain')
    pw    = d.get(f'{pfx}_put_wall')
    exp   = d.get(f'{pfx}_expiry','') or ''
    is_mo = d.get(f'{pfx}_is_monthly',False)
    lbl   = d.get(f'{pfx}_label',label)
    exps  = exp[5:] if exp else lbl
    fmt   = lambda v: f'${v:g}' if v is not None else '—'
    mob   = ('<span style="font-size:0.65em;color:#e3b341;margin-left:3px;'
             'background:rgba(227,179,65,.12);border:1px solid rgba(227,179,65,.3);'
             'border-radius:3px;padding:0 3px">月結</span>' if is_mo else '')
    if cw is None and mp is None and pw is None:
        return f'<div style="color:#484f58;font-size:0.75em">{lbl}{mob}<br>N/A</div>'
    return (f'<div class="oi-lbl">{lbl} {exps}{mob}</div>'
            f'<div class="oi-call">▼ {fmt(cw)}</div>'
            f'<div class="oi-pain">⚡ {fmt(mp)}</div>'
            f'<div class="oi-put">▲ {fmt(pw)}</div>')


def gex_html_single(d, pfx):
    gex = d.get(f'{pfx}_gex')
    if gex is None:
        return '<div style="color:#484f58;font-size:0.75em">GEX<br>N/A</div>'
    col = '#3fb950' if gex >= 0 else '#f85149'
    arr = '▲' if gex >= 0 else '▼'
    lbl = 'Long γ' if gex >= 0 else 'Short γ'
    bp  = min(int(abs(gex)/5000*100),100)
    return (f'<div style="font-size:0.65em;color:#484f58;margin-bottom:2px;'
            f'text-transform:uppercase;letter-spacing:0.5px">GEX</div>'
            f'<div style="font-family:monospace;font-size:0.88em;font-weight:700;color:{col}">{arr} ${abs(gex):,.0f}M</div>'
            f'<div style="font-size:0.65em;color:{col};margin-bottom:3px">{lbl}</div>'
            f'<div style="background:#21262d;border-radius:2px;height:4px;width:60px">'
            f'<div style="background:{col};height:4px;border-radius:2px;width:{bp}%"></div></div>')


def stock_row(d):
    if not d.get('ok'):
        mag = "<span class='m7'>Mag 7</span>" if d['ticker'] in MAG7 else ''
        return (f'<tr><td><div class="tk"><span class="tb">{d["ticker"]}</span>{mag}</div></td>'
                f'<td colspan="14" style="color:#f85149;font-size:0.8em">⚠ 數據抓取失敗</td></tr>')

    ticker  = d['ticker']
    price   = d['price']
    chg     = d['change']
    chg_pct = d['change_pct']

    if chg > 0:   cc, cs, ps2 = 'pos', f'+{chg:.2f}', f'+{chg_pct:.2f}%'
    elif chg < 0: cc, cs, ps2 = 'neg', f'{chg:.2f}',  f'{chg_pct:.2f}%'
    else:         cc, cs, ps2 = 'neu', '—', '—'

    oi_w = [oi_html_single(d, f'w{i}') + target_zone_html(d, f'w{i}_') for i in range(4)]
    gx_w = [gex_html_single(d, f'w{i}') for i in range(4)]

    ttm_h, fwd_h     = pe_html(d['pe_ttm'], d['pe_fwd'])
    iv_bar, iv_badge = iv_html(d['iv'], d['ivr'])
    pc_h = pc_html(d['pc_ratio'])
    ms_h = ms_html(ticker, price)
    mc    = fmt_cap(d.get('market_cap'))
    vol_h = vol_cell_html(d)
    ema_h = ema_cell_html(d)
    mb    = f'<span class="m7">Mag 7</span>' if ticker in MAG7 else ''

    return f'''<tr>
  <td><div class="tk"><span class="tb">{ticker}</span>{mb}</div></td>
  <td class="pr-td">
    <div class="pr">${price:.2f}</div>
    <div class="chg-row {cc}">{cs} <span class="pct">{ps2}</span></div>
    <div class="vol-inline">{vol_h}</div>
  </td>
  <td class="oi">{oi_w[0]}<div class="gex-inline">{gx_w[0]}</div></td>
  <td class="oi">{oi_w[1]}<div class="gex-inline">{gx_w[1]}</div></td>
  <td class="oi">{oi_w[2]}<div class="gex-inline">{gx_w[2]}</div></td>
  <td class="oi">{oi_w[3]}<div class="gex-inline">{gx_w[3]}</div></td>
  <td class="ema-td">{ema_h}</td>
  <td>{ms_h}</td>
  <td>{ttm_h}</td><td>{fwd_h}</td>
  <td>{iv_bar}</td><td>{iv_badge}</td>
  <td>{pc_h}</td>
  <td class="mc">{mc}</td>
</tr>'''


def summary_cards(stocks):
    ok   = [s for s in stocks if s.get('ok')]
    up   = sum(1 for s in ok if s['change'] > 0)
    down = sum(1 for s in ok if s['change'] < 0)
    flat = len(stocks) - up - down
    best = max(ok, key=lambda s: s['change_pct'], default=None)
    bstr = f'{best["ticker"]} {best["change_pct"]:+.2f}%' if best else '—'
    ivv  = [s['ivr'] for s in ok if s.get('ivr') is not None]
    avg_ivr = int(sum(ivv)/len(ivv)) if ivv else 0
    return f'''
  <div class="sc"><div class="l">上漲</div><div class="v gu">▲ {up}</div></div>
  <div class="sc"><div class="l">下跌</div><div class="v gd">▼ {down}</div></div>
  <div class="sc"><div class="l">數據不完整</div><div class="v gf">— {flat}</div></div>
  <div class="sc"><div class="l">市場情緒</div><div class="v gb">{"偏多 🟢" if up>down else "偏空 🔴" if down>up else "中性 ⚪"}</div></div>
  <div class="sc"><div class="l">最強漲幅</div><div class="v gu" style="font-size:1em">{bstr}</div></div>
  <div class="sc"><div class="l">整體 IV 水位</div><div class="v gmx">IVR≈{avg_ivr}%</div></div>
  <div class="sc"><div class="l">掃描股票</div><div class="v gbl">{len(stocks)}</div></div>
'''


def vix_bar_pct(val): return min(int(val/40*100),100)


def generate_report(stocks, vix, spy_qqq, date_str):
    rows_html  = '\n'.join(stock_row(s) for s in stocks)
    cards_html = summary_cards(stocks)

    vix_val  = vix['value']
    vix_chg  = vix['change']
    vix_pct  = vix['pct']
    vix_sign = '▼' if vix_chg <= 0 else '▲'

    spy = spy_qqq.get('SPY', {})
    qqq = spy_qqq.get('QQQ', {})
    spy_cc = 'pos' if spy.get('change_pct', 0) >= 0 else 'neg'
    qqq_cc = 'pos' if qqq.get('change_pct', 0) >= 0 else 'neg'

    def oi_grid(ed):
        cells = []
        for k in ['w0','w1','w2','w3']:
            lbl   = ed.get(f'{k}_label','')
            exp   = ed.get(f'{k}_expiry','') or ''
            is_mo = ed.get(f'{k}_is_monthly',False)
            cw, mp, pw = ed.get(f'{k}_call_wall'), ed.get(f'{k}_max_pain'), ed.get(f'{k}_put_wall')
            exps  = exp[5:] if exp else ''
            mob   = ('<span style="font-size:0.6em;color:#e3b341;background:rgba(227,179,65,.12);'
                     'border:1px solid rgba(227,179,65,.3);border-radius:3px;padding:0 3px;margin-left:2px">月結</span>' if is_mo else '')
            fmt   = lambda v: f'${v:g}' if v is not None else '—'
            cells.append(f'<div style="background:#21262d;border-radius:6px;padding:7px 10px">'
                         f'<div class="oi-lbl">{lbl} {exps}{mob}</div>'
                         f'<div class="oi-call" style="font-size:0.85em">▼ {fmt(cw)}</div>'
                         f'<div class="oi-pain" style="font-size:0.85em">⚡ {fmt(mp)}</div>'
                         f'<div class="oi-put"  style="font-size:0.85em">▲ {fmt(pw)}</div></div>')
        return '\n'.join(cells)

    now_tw = datetime.now().strftime('%Y-%m-%d %H:%M')
    wd = {0:'一',1:'二',2:'三',3:'四',4:'五',5:'六',6:'日'}[datetime.now().weekday()]
    display_date = datetime.now().strftime(f'%Y年%m月%d日（週{wd}）')

    return f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>美股每日開盤掃描報告 Powered by 黑叔 - {date_str}</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','Noto Sans TC','PingFang TC',sans-serif;background:#0d1117;color:#c9d1d9;min-height:100vh;padding:16px;line-height:1.5;font-size:16px}}
.wrap{{max-width:1700px;margin:0 auto}}
.hdr{{background:linear-gradient(135deg,#161b22,#1c2128);border:1px solid #30363d;border-radius:12px;padding:18px 26px;margin-bottom:14px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}}
.hdr h1{{font-size:1.35em;font-weight:700;color:#e6edf3}}
.hdr .sub{{color:#8b949e;font-size:0.76em;margin-top:4px}}
.hdr-r{{text-align:right;font-size:0.78em;color:#8b949e;line-height:1.7}}
.hdr-r .dt{{font-size:1em;color:#e6edf3;font-weight:600}}
.vix{{background:#161b22;border:1px solid #30363d;border-radius:10px;padding:12px 22px;margin-bottom:14px;display:flex;align-items:center;gap:24px;flex-wrap:wrap}}
.vix-val{{font-size:2em;font-weight:800;color:#3fb950;font-family:monospace}}
.vix-chg{{font-size:0.85em;color:#3fb950}}
.vix-lbl{{font-size:0.7em;color:#8b949e;text-transform:uppercase;letter-spacing:1px}}
.vix-gauge{{flex:1;min-width:200px}}
.vix-g-lbl{{display:flex;justify-content:space-between;font-size:0.62em;color:#484f58;margin-bottom:5px}}
.vix-bar{{background:linear-gradient(90deg,#3fb950 0%,#d29922 40%,#f85149 100%);border-radius:4px;height:7px;position:relative}}
.vix-dot{{position:absolute;top:-4px;width:15px;height:15px;background:#e6edf3;border:2px solid #0d1117;border-radius:50%;transform:translateX(-50%)}}
.vix-ctx{{font-size:0.76em;color:#8b949e;line-height:1.6}}
.vix-ctx strong{{color:#c9d1d9}}
.sum{{display:flex;gap:10px;margin-bottom:14px;flex-wrap:wrap}}
.sc{{background:#161b22;border:1px solid #30363d;border-radius:10px;padding:12px 16px;flex:1;min-width:110px;text-align:center}}
.sc .l{{font-size:0.65em;color:#8b949e;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px}}
.sc .v{{font-size:1.3em;font-weight:700}}
.gu{{color:#3fb950}}.gd{{color:#f85149}}.gf{{color:#8b949e}}.gb{{color:#3fb950}}.gmx{{color:#d29922}}.gbl{{color:#79c0ff}}
.stitle{{font-size:0.72em;font-weight:700;color:#8b949e;text-transform:uppercase;letter-spacing:1.2px;border-left:3px solid #30363d;padding-left:10px;margin:16px 0 10px}}
.tbl-wrap{{background:#161b22;border:1px solid #30363d;border-radius:12px;overflow:hidden;overflow-x:auto;margin-bottom:10px}}
table{{width:100%;border-collapse:collapse;font-size:0.92em}}
thead th{{background:#1c2128;color:#8b949e;font-weight:600;text-transform:uppercase;font-size:0.72em;letter-spacing:0.7px;padding:11px 13px;text-align:left;border-bottom:1px solid #30363d;white-space:nowrap}}
.th-grp{{text-align:center!important;font-size:0.63em;letter-spacing:0.5px}}
.th-val{{color:#3fb950}}.th-opt{{color:#d29922}}
tbody tr{{border-bottom:1px solid #21262d;transition:background .12s}}
tbody tr:hover{{background:#1c2128}}
tbody tr:last-child{{border-bottom:none}}
td{{padding:11px 13px;vertical-align:top}}
.tk{{display:flex;flex-direction:column;gap:3px}}
.tb{{background:#21262d;border:1px solid #30363d;border-radius:5px;padding:2px 8px;font-weight:700;font-size:0.86em;color:#79c0ff;font-family:monospace;letter-spacing:.5px;width:fit-content}}
.m7{{font-size:0.58em;background:rgba(121,192,255,.08);border:1px solid rgba(121,192,255,.15);color:#79c0ff;border-radius:3px;padding:1px 5px;width:fit-content}}
.pr{{font-weight:700;color:#e6edf3;font-family:monospace;white-space:nowrap;vertical-align:middle}}
.pos{{color:#3fb950}}.neg{{color:#f85149}}.neu{{color:#8b949e}}.pct{{font-weight:600}}
/* ═══ 量比 + EMA 合併欄樣式 ═══ */
.ve-td{{min-width:160px;max-width:195px;vertical-align:top}}
.vol-lbl{{font-size:0.64em;color:#8b949e;text-transform:uppercase;letter-spacing:1.2px;margin-bottom:2px}}
.vol-val{{margin-bottom:7px;line-height:1.2}}
/* 量比數值：三個檔位的字體大小 */
.vol-hi{{font-size:1.45em;font-weight:800;color:#e3b341;font-family:monospace}}
.vol-mid{{font-size:1.28em;font-weight:700;color:#c9d1d9;font-family:monospace}}
.vol-norm{{font-size:1.25em;font-weight:600;color:#8b949e;font-family:monospace}}
.vb{{display:inline-block;background:rgba(227,179,65,.12);border:1px solid rgba(227,179,65,.25);color:#e3b341;border-radius:4px;padding:1px 5px;font-size:0.66em;margin-left:4px;vertical-align:middle}}
/* EMA 小字三行 */
.ema-block{{border-top:1px dashed #21262d;padding-top:5px;margin-top:1px}}
.er{{display:flex;align-items:baseline;gap:4px;margin-bottom:3px;flex-wrap:wrap}}
.el{{font-size:0.63em;color:#484f58;min-width:46px;text-transform:uppercase;letter-spacing:0.4px;flex-shrink:0}}
.ev{{font-family:monospace;font-size:0.82em;font-weight:700}}
.ep{{font-size:0.70em;font-weight:600}}
.mau{{color:#3fb950}}.mad{{color:#f85149}}.man{{color:#484f58}}
/* 訊號 */
.ema-sig{{margin-top:5px;padding-top:4px;border-top:1px dashed #21262d}}
.sig-bull{{color:#3fb950;font-weight:700;font-size:0.74em;display:block;line-height:1.4}}
.sig-bear{{color:#f85149;font-weight:700;font-size:0.74em;display:block;line-height:1.4}}
.sig-warn{{color:#e3b341;font-weight:700;font-size:0.74em;display:block;line-height:1.4}}
.sig-note{{color:#484f58;font-size:0.68em;display:block;margin-top:1px;line-height:1.3}}
/* 其他原有樣式 */
.oi{{min-width:110px;max-width:140px;vertical-align:top}}
.pr-td{{min-width:120px;max-width:160px;vertical-align:top}}
.chg-row{{font-size:0.85em;margin-top:3px;white-space:nowrap}}
.vol-inline{{margin-top:6px;padding-top:5px;border-top:1px dashed #30363d}}
.gex-inline{{margin-top:6px;padding-top:5px;border-top:1px dashed #30363d}}
.ema-td{{min-width:175px;max-width:210px;vertical-align:top}}
.vol-tip{{font-size:0.68em;margin-top:4px;padding:2px 6px;border-radius:4px;line-height:1.4;display:inline-block}}
.vol-tip-hi{{background:rgba(227,179,65,.1);color:#e3b341;border:1px solid rgba(227,179,65,.25)}}
.vol-tip-mid{{background:rgba(121,192,255,.08);color:#79c0ff;border:1px solid rgba(121,192,255,.2)}}
.vol-tip-norm{{background:rgba(139,148,158,.07);color:#8b949e;border:1px solid rgba(139,148,158,.15)}}
.oi-lbl{{font-size:0.7em;color:#484f58;margin-bottom:3px;letter-spacing:0.5px;text-transform:uppercase}}
.oi-call{{font-family:monospace;font-size:0.92em;color:#f85149;white-space:nowrap}}
.oi-pain{{font-family:monospace;font-size:0.92em;color:#d29922;white-space:nowrap;margin:3px 0}}
.oi-put{{font-family:monospace;font-size:0.92em;color:#3fb950;white-space:nowrap}}
.tz-wrap{{margin-top:6px;padding-top:5px;border-top:1px dashed #30363d}}
.tz-call{{font-family:monospace;font-size:0.82em;color:#a371f7;white-space:nowrap}}
.tz-put{{font-family:monospace;font-size:0.82em;color:#79c0ff;white-space:nowrap;margin-top:2px}}
.ms-fv{{font-family:monospace;font-size:0.95em;font-weight:700;color:#e6edf3}}
.ms-disc{{font-size:0.78em;color:#3fb950;font-weight:600}}
.ms-prem{{font-size:0.78em;color:#f85149;font-weight:600}}
.ms-neu{{font-size:0.78em;color:#8b949e;font-weight:600}}
.ms-str{{color:#e3b341;font-size:1em;letter-spacing:2px;margin:3px 0;line-height:1}}
.ms-mw{{display:inline-block;font-size:0.6em;padding:1px 5px;border-radius:3px;background:rgba(121,192,255,.1);color:#79c0ff;border:1px solid rgba(121,192,255,.2)}}
.ms-mn{{display:inline-block;font-size:0.6em;padding:1px 5px;border-radius:3px;background:rgba(210,153,34,.1);color:#d29922;border:1px solid rgba(210,153,34,.2)}}
.ms-m0{{display:inline-block;font-size:0.6em;padding:1px 5px;border-radius:3px;background:rgba(72,79,88,.1);color:#484f58;border:1px solid rgba(72,79,88,.2)}}
.pe{{font-family:monospace;font-size:0.95em;font-weight:700;white-space:nowrap}}
.pe-w{{color:#e3b341}}.pe-l{{color:#f85149;font-size:0.85em}}
.fpe-dn{{color:#3fb950;font-family:monospace;font-size:0.92em;font-weight:700}}
.fpe-fl{{color:#8b949e;font-family:monospace;font-size:0.92em;font-weight:700}}
.pe-note{{font-size:0.7em;margin-top:2px}}
.ivb{{display:inline-block;padding:2px 7px;border-radius:8px;font-size:0.72em;font-weight:600}}
.ivs{{background:rgba(63,185,80,.1);color:#3fb950;border:1px solid rgba(63,185,80,.2)}}
.ivn{{background:rgba(139,148,158,.1);color:#8b949e;border:1px solid rgba(139,148,158,.2)}}
.ive{{background:rgba(210,153,34,.12);color:#d29922;border:1px solid rgba(210,153,34,.25)}}
.ivh{{background:rgba(248,81,73,.12);color:#f85149;border:1px solid rgba(248,81,73,.25)}}
.ivr-bg{{background:#21262d;border-radius:3px;height:6px;margin-bottom:4px;overflow:hidden;width:78px}}
.ivr-f{{height:100%;border-radius:3px}}
.ivrl{{background:#3fb950}}.ivrm{{background:#d29922}}.ivrh{{background:#f85149}}
.ivr-lbl{{font-size:0.75em;color:#8b949e}}
.pc-b{{color:#3fb950;font-weight:600;font-family:monospace}}
.pc-r{{color:#f85149;font-weight:600;font-family:monospace}}
.pc-n{{color:#8b949e;font-weight:600;font-family:monospace}}
.pc-l{{font-size:0.65em;color:#484f58;margin-top:2px}}
.mc{{color:#8b949e;font-size:0.8em;font-family:monospace;white-space:nowrap}}
.leg{{background:#161b22;border:1px solid #30363d;border-radius:10px;padding:10px 16px;margin-bottom:12px;display:flex;gap:14px;flex-wrap:wrap;align-items:center;font-size:0.73em;color:#8b949e}}
.leg strong{{color:#c9d1d9}}
/* ═══ SPY/QQQ 卡片 ═══ */
.mkt-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}}
@media(max-width:780px){{.mkt-grid{{grid-template-columns:1fr}}}}
.mkc{{background:#161b22;border:1px solid #30363d;border-radius:12px;padding:18px 20px}}
.mkr{{display:flex;justify-content:space-between;align-items:center;font-size:0.78em;padding:5px 0;border-bottom:1px solid #21262d}}
.mkr:last-child{{border-bottom:none}}
.mkr .k{{color:#8b949e}}.mkr .v{{font-family:monospace;font-weight:600}}
.mkc-tk{{font-size:1.3em;font-weight:800;color:#e6edf3;font-family:monospace}}
.mkc-sub{{font-size:0.7em;color:#8b949e;margin-top:2px}}
.mkc-p{{font-size:1.25em;font-weight:700;font-family:monospace;color:#e6edf3;text-align:right}}
.mkc-hdr{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px}}
/* ETF EMA 分析塊 */
.etf-ema-blk{{margin-top:12px;border-top:1px solid #21262d;padding-top:10px}}
.etf-sig{{margin:7px 0;padding:5px 9px;background:#1c2128;border-radius:6px;font-size:0.79em;line-height:1.5}}
.etf-advice{{margin-top:8px;padding:9px 12px;border-radius:7px;font-size:0.77em;line-height:1.65}}
.etf-bull{{background:rgba(63,185,80,.07);border:1px solid rgba(63,185,80,.2);color:#c9d1d9}}
.etf-bear{{background:rgba(248,81,73,.07);border:1px solid rgba(248,81,73,.2);color:#c9d1d9}}
.etf-warn{{background:rgba(210,153,34,.07);border:1px solid rgba(210,153,34,.2);color:#c9d1d9}}
.etf-neu{{background:rgba(72,79,88,.1);border:1px solid #30363d;color:#8b949e}}
/* OI 四週 grid */
.oi4grid{{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:12px}}
.ftr{{background:#161b22;border:1px solid #30363d;border-radius:10px;padding:12px 18px;text-align:center;color:#484f58;font-size:0.7em;line-height:2;margin-top:14px}}
</style>
</head>
<body>
<div class="wrap">

<div class="hdr">
  <div>
    <h1>📊 美股每日開盤掃描報告 Powered by 黑叔</h1>
    <div class="sub">Magnificent 7 + TSM · MU · SNDK · NFLX · AMD · AVGO · LITE · COHR　｜　真實數據 yfinance　｜　晨星估值 · OI · IV · P/C Ratio · EMA15/50/100 趨勢訊號</div>
  </div>
  <div class="hdr-r">
    <div class="dt">{display_date}</div>
    <div>美股交易時段　｜　22:00 台灣時間　｜　本機生成</div>
  </div>
</div>

<div class="vix">
  <div>
    <div class="vix-lbl">恐慌指數 VIX</div>
    <div style="display:flex;align-items:baseline;gap:10px">
      <div class="vix-val">{vix_val}</div>
      <div class="vix-chg">{vix_sign} {abs(vix_chg)} ({vix_pct:+.2f}%)</div>
    </div>
  </div>
  <div class="vix-gauge">
    <div class="vix-g-lbl"><span>極低 &lt;12</span><span>正常 12–20</span><span>警戒 20–30</span><span>恐慌 &gt;30</span></div>
    <div class="vix-bar"><div class="vix-dot" style="left:{vix_bar_pct(vix_val)}%"></div></div>
  </div>
  <div class="vix-ctx">
    <strong>{"低波動 ✓" if vix_val<20 else "警戒 ⚠" if vix_val<30 else "恐慌 🔥"}</strong>　VIX {vix_val}
    {"位於正常區間（12–20），期權保護成本偏低。" if vix_val<20 else "進入警戒區間，市場波動加劇，留意風險。" if vix_val<30 else "極度恐慌，考慮逢低布局或持有避險部位。"}
  </div>
</div>

<div class="sum">{cards_html}</div>

<div class="leg">
  <strong>欄位說明：</strong>
  <span><span class="ivb ivs">低 IVR</span> &lt;25%</span>
  <span><span class="ivb ivn">正常</span> 25–50%</span>
  <span><span class="ivb ive">偏高</span> 50–75%</span>
  <span><span class="ivb ivh">極高</span> &gt;75%</span>
  <span>｜</span>
  <span class="oi-call">▼ Call Wall</span>=壓力
  <span class="oi-pain">⚡ Max Pain</span>=痛苦點
  <span class="oi-put">▲ Put Wall</span>=支撐
  <span style="color:#e3b341">月結</span>=當月第三週五
  <span>｜</span>
  <span style="color:#a371f7">🎯↑</span> OTM Call　<span style="color:#79c0ff">🎯↓</span> OTM Put
  <span>｜ EMA訊號：</span>
  <span class="sig-bull" style="display:inline">🌟黃金交叉</span>
  <span class="sig-bear" style="display:inline">💀死亡交叉</span>
  <span class="sig-bull" style="display:inline">📈多頭排列</span>
  <span class="sig-bear" style="display:inline">📉空頭排列</span>
  <span class="sig-warn" style="display:inline">⚡短強長弱/短弱長強</span>
  <span class="sig-bull" style="display:inline">⬆突破EMA15</span>
  <span class="sig-bear" style="display:inline">⬇跌破EMA15</span>
  <span class="sig-bull" style="display:inline">🔄回彈</span>
  <span class="sig-bear" style="display:inline">🔄回落</span>
</div>

<div class="stitle">📋 個股全覽</div>
<div class="tbl-wrap">
<table>
<thead>
<tr>
  <th rowspan="2">代碼</th>
  <th rowspan="2" style="min-width:130px">現價 / 漲跌 / 量比</th>
  <th colspan="4" class="th-grp" style="color:#58a6ff;min-width:520px">── OI 支撐壓力 + GEX + OTM目標（本週 / 下週 / 下下週 / 下下三週）──</th>
  <th rowspan="2" style="min-width:175px;color:#58a6ff">EMA15 / 50 / 100</th>
  <th rowspan="2" style="min-width:100px">晨星估值</th>
  <th colspan="2" class="th-grp th-val">── 估 值 ──</th>
  <th colspan="3" class="th-grp th-opt">── 期 權 分 析 ──</th>
  <th rowspan="2">市值</th>
</tr>
<tr>
  <th style="color:#58a6ff;font-size:0.7em">本週</th>
  <th style="color:#58a6ff;font-size:0.7em">下週</th>
  <th style="color:#58a6ff;font-size:0.7em">下下週</th>
  <th style="color:#58a6ff;font-size:0.7em">下下三週</th>
  <th>P/E TTM</th><th>Fwd P/E</th>
  <th>IV / IVR</th><th>IV狀態</th><th>P/C Ratio</th>
</tr>
</thead>
<tbody>
{rows_html}
</tbody>
</table>
</div>

<div class="stitle">🏦 大盤技術面 — SPY & QQQ</div>
<div class="mkt-grid">
  <div class="mkc">
    <div class="mkc-hdr">
      <div><div class="mkc-tk">SPY</div><div class="mkc-sub">S&P 500 ETF</div></div>
      <div>
        <div class="mkc-p">${spy.get('price','—')}</div>
        <div style="font-size:0.82em;text-align:right" class="{spy_cc}">{spy.get('change_pct',0):+.2f}%</div>
      </div>
    </div>
    {etf_ema_block_html(spy)}
    <div class="oi4grid">{oi_grid(spy)}</div>
  </div>
  <div class="mkc">
    <div class="mkc-hdr">
      <div><div class="mkc-tk">QQQ</div><div class="mkc-sub">Nasdaq 100 ETF</div></div>
      <div>
        <div class="mkc-p">${qqq.get('price','—')}</div>
        <div style="font-size:0.82em;text-align:right" class="{qqq_cc}">{qqq.get('change_pct',0):+.2f}%</div>
      </div>
    </div>
    {etf_ema_block_html(qqq)}
    <div class="oi4grid">{oi_grid(qqq)}</div>
  </div>
</div>

<div class="ftr">
  ⚠️ 本報告數據來自 Yahoo Finance（yfinance），OI / Max Pain 為真實期權鏈計算，晨星估值為靜態手動更新。僅供參考，不構成投資建議。<br>
  數據來源：Yahoo Finance · yfinance　｜　報告產生：{now_tw} 台灣時間<br>
  <span style="color:#21262d">本機版 v5.2 · python3 stock_scan.py · 四週 OI + GEX + EMA15/50/100 趨勢訊號</span>
</div>

</div>
</body>
</html>'''


# ═══════════════════════════════════════════════════════════════
#  主程式
# ═══════════════════════════════════════════════════════════════

def main():
    print('=' * 50)
    print('美股每日掃描報告 Powered by 黑叔 — v5.2')
    print('=' * 50)

    date_str    = datetime.now().strftime('%Y%m%d')
    output_path = os.path.join(OUTPUT_DIR, 'index.html')
    dated_path  = os.path.join(OUTPUT_DIR, f'stock_report_{date_str}.html')
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print('\n[1/4] 抓取 VIX...')
    vix = fetch_vix()
    print(f'  VIX: {vix["value"]} ({vix["change"]:+.2f})')

    print(f'\n[2/4] 抓取個股數據（共 {len(TICKERS)} 支）...')
    stocks = [fetch_stock(t) for t in TICKERS]
    ok_count = sum(1 for s in stocks if s.get('ok'))
    print(f'  完成：{ok_count}/{len(TICKERS)} 支成功')

    print('\n[3/4] 抓取 SPY / QQQ...')
    spy_qqq = fetch_spy_qqq()

    print('\n[4/4] 生成 HTML 報告...')
    html = generate_report(stocks, vix, spy_qqq, date_str)

    for path in [output_path, dated_path]:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(html)

    print(f'\n✅ 報告已存至：{output_path}')
    print(f'   日期版本：{dated_path}')
    print('=' * 50)


if __name__ == '__main__':
    main()
