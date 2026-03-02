#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
美股每日開盤掃描報告 — 本機版 v3.0
使用 yfinance 抓取真實股價、期權 OI、IV、P/C Ratio、MA 數據
執行方式：python3 stock_scan_local.py
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

TICKERS = ['AAPL', 'MSFT', 'NVDA', 'AMZN', 'GOOGL', 'META', 'TSLA',
           'TSM', 'MU', 'SNDK', 'NFLX', 'AMD', 'AVGO', 'COST']

MAG7 = {'AAPL', 'MSFT', 'NVDA', 'AMZN', 'GOOGL', 'META', 'TSLA'}

# 晨星估值 (靜態，建議每月手動更新一次)
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
    'COST':  {'fv': 700,  'stars': 2, 'moat': 'Wide'},
}

# 輸出資料夾：預設桌面，有 Google Drive 請改為同步路徑
# 例：OUTPUT_DIR = os.path.expanduser('~/Library/CloudStorage/GoogleDrive-你的帳號/My Drive/股票報告')
OUTPUT_DIR = os.environ.get('OUTPUT_DIR', os.path.expanduser('~/Desktop'))


# ═══════════════════════════════════════════════════════════════
#  工具函數
# ═══════════════════════════════════════════════════════════════

def get_four_weekly_expiries():
    """取得本週、下週、下下週、下下下週的週五到期日，並標記是否為月結算週（當月第三個週五）"""
    today = datetime.now().date()
    days_to_friday = (4 - today.weekday()) % 7
    if days_to_friday == 0:
        days_to_friday = 0  # 今天就是週五，算本週
    this_friday = today + timedelta(days=days_to_friday)

    results = []
    for i in range(4):
        friday = this_friday + timedelta(weeks=i)
        label = ['本週', '下週', '下下週', '下下下週'][i]

        # 判斷是否為當月第三個週五（月結算日）
        y, m = friday.year, friday.month
        first = datetime(y, m, 1).date()
        first_fri = first + timedelta(days=(4 - first.weekday()) % 7)
        third_fri = first_fri + timedelta(weeks=2)
        is_monthly = (friday == third_fri)

        results.append({
            'label': label,
            'date': friday.strftime('%Y-%m-%d'),
            'is_monthly': is_monthly
        })
    return results


def get_weekly_expiry():
    """找最近的週結到期日（下一個週五，至少 1 天後）【舊版，保留相容性】"""
    today = datetime.now()
    days_to_friday = (4 - today.weekday()) % 7
    if days_to_friday == 0:
        days_to_friday = 7
    candidate = today + timedelta(days=days_to_friday)
    return candidate.strftime('%Y-%m-%d')


def get_monthly_expiry():
    """找最近的月結到期日（當月或下月第三個週五）【舊版，保留相容性】"""
    today = datetime.now()
    for offset in range(3):
        m = today.month + offset
        y = today.year + (m - 1) // 12
        m = ((m - 1) % 12) + 1
        first = datetime(y, m, 1)
        first_fri = first + timedelta(days=(4 - first.weekday()) % 7)
        third_fri = first_fri + timedelta(weeks=2)
        if (third_fri - today).days >= 7:
            return third_fri.strftime('%Y-%m-%d')
    return None


def calc_max_pain(calls_df, puts_df, current_price=None):
    """計算 Max Pain：option writers 受益最多的 strike 價格。
    改進：
      1. 只計算現價 ±20% 以內的 strike（過濾遠 OTM 雜訊）
      2. 過濾 OI < 50 口的 strike（低流動性）
    """
    try:
        c = calls_df.copy()
        p = puts_df.copy()

        # 過濾低 OI
        c = c[c['openInterest'] >= 50]
        p = p[p['openInterest'] >= 50]

        # 過濾遠 OTM（現價 ±20%）
        if current_price and current_price > 0:
            lo = current_price * 0.80
            hi = current_price * 1.20
            c = c[(c['strike'] >= lo) & (c['strike'] <= hi)]
            p = p[(p['strike'] >= lo) & (p['strike'] <= hi)]

        all_strikes = sorted(set(c['strike'].tolist() + p['strike'].tolist()))
        if not all_strikes:
            return None

        calls_map = dict(zip(c['strike'], c['openInterest'].fillna(0)))
        puts_map  = dict(zip(p['strike'], p['openInterest'].fillna(0)))

        min_pain = float('inf')
        max_pain_strike = all_strikes[len(all_strikes) // 2]

        for s in all_strikes:
            call_pain = sum(max(0, s - k) * oi for k, oi in calls_map.items())
            put_pain  = sum(max(0, k - s) * oi for k, oi in puts_map.items())
            total = call_pain + put_pain
            if total < min_pain:
                min_pain = total
                max_pain_strike = s

        return max_pain_strike
    except Exception:
        return None


def calc_ivr(ticker_obj, current_iv):
    """用過去一年實現波動率估算 IVR（0–100）"""
    try:
        hist = ticker_obj.history(period='1y')
        if hist.empty or len(hist) < 30:
            return 50
        returns = hist['Close'].pct_change().dropna()
        vol_series = (returns.rolling(21).std() * math.sqrt(252) * 100).dropna()
        v_min, v_max = vol_series.min(), vol_series.max()
        if v_max == v_min:
            return 50
        ivr = int((current_iv - v_min) / (v_max - v_min) * 100)
        return max(0, min(100, ivr))
    except Exception:
        return 50


def fmt_cap(mc):
    if not mc or math.isnan(mc):
        return '—'
    if mc >= 1e12:
        return f'${mc/1e12:.2f}T'
    elif mc >= 1e9:
        return f'${mc/1e9:.1f}B'
    return f'${mc/1e6:.0f}M'


def fmt_pe(val):
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    return round(float(val), 1)


# ═══════════════════════════════════════════════════════════════
#  數據抓取
# ═══════════════════════════════════════════════════════════════

def fetch_vix():
    """抓取 VIX 數據 + S&P500 成份股站上 20日/50日均線比例"""
    result = {'value': 18.5, 'change': 0.0, 'pct': 0.0, 's5tw': None, 's5fi': None}
    try:
        v = yf.Ticker('^VIX')
        hist = v.history(period='2d')
        if len(hist) >= 2:
            now  = round(hist['Close'].iloc[-1], 2)
            prev = round(hist['Close'].iloc[-2], 2)
            chg  = round(now - prev, 2)
            pct  = round(chg / prev * 100, 2)
            result.update({'value': now, 'change': chg, 'pct': pct})
    except Exception:
        pass

    # $S5TW: S&P 500 Stocks Above 20-Day MA
    try:
        t = yf.Ticker('^SP500-20MA')  # 嘗試常見 ticker
        h = t.history(period='2d')
        if not h.empty:
            result['s5tw'] = round(float(h['Close'].iloc[-1]), 1)
    except Exception:
        pass

    # 若 ^SP500-20MA 取不到，試 $S5TW via alternative symbol
    if result['s5tw'] is None:
        for sym in ['$S5TW', 'S5TW']:
            try:
                t = yf.Ticker(sym)
                h = t.history(period='2d')
                if not h.empty:
                    result['s5tw'] = round(float(h['Close'].iloc[-1]), 1)
                    break
            except Exception:
                pass

    # $S5FI: S&P 500 Stocks Above 50-Day MA
    for sym in ['$S5FI', 'S5FI', '^SP500-50MA']:
        try:
            t = yf.Ticker(sym)
            h = t.history(period='2d')
            if not h.empty:
                result['s5fi'] = round(float(h['Close'].iloc[-1]), 1)
                break
        except Exception:
            pass

    return result


def fetch_stock(ticker):
    """抓取單支股票所有數據"""
    print(f'  [{ticker}] 抓取中...')
    d = {'ticker': ticker, 'ok': False}

    try:
        t = yf.Ticker(ticker)
        info = t.info

        # 歷史價格
        hist_short = t.history(period='5d')
        hist_long  = t.history(period='1y')

        if hist_short.empty:
            return d

        price = float(hist_short['Close'].iloc[-1])
        prev  = float(hist_short['Close'].iloc[-2]) if len(hist_short) > 1 else price
        d['price']      = price
        d['change']     = round(price - prev, 2)
        d['change_pct'] = round((price - prev) / prev * 100, 2)
        d['company']    = info.get('shortName', ticker)
        d['market_cap'] = info.get('marketCap')

        # 量比
        avg_vol = hist_short['Volume'].iloc[:-1].mean()
        d['vol_ratio'] = round(hist_short['Volume'].iloc[-1] / avg_vol, 2) if avg_vol else 1.0

        # MA
        closes = hist_long['Close'] if not hist_long.empty else hist_short['Close']
        d['ma20']  = float(closes.rolling(20).mean().iloc[-1])  if len(closes) >= 20  else None
        d['ma50']  = float(closes.rolling(50).mean().iloc[-1])  if len(closes) >= 50  else None
        d['ma200'] = float(closes.rolling(200).mean().iloc[-1]) if len(closes) >= 200 else None

        # P/E
        d['pe_ttm'] = fmt_pe(info.get('trailingPE'))
        d['pe_fwd'] = fmt_pe(info.get('forwardPE'))

        # 期權數據（四週：本週、下週、下下週、下下下週）
        WEEK_KEYS = ['w0', 'w1', 'w2', 'w3']
        for k in WEEK_KEYS:
            d[f'{k}_expiry']    = None
            d[f'{k}_put_wall']  = None
            d[f'{k}_call_wall'] = None
            d[f'{k}_max_pain']  = None
            d[f'{k}_is_monthly'] = False
            d[f'{k}_label']     = ''
        d['iv']       = None
        d['ivr']      = None
        d['pc_ratio'] = None
        # 期權流量訊號（模擬 Barchart 流量法）
        d['flow_signal']     = None   # 'CALL_SWEEP'|'PUT_SWEEP'|'NEUTRAL'|'MIXED'
        d['flow_call_vol']   = None
        d['flow_put_vol']    = None
        d['flow_call_oi']    = None
        d['flow_put_oi']     = None
        d['flow_call_ratio'] = None   # vol/OI > 1 = 激進買方
        d['flow_put_ratio']  = None
        d['flow_premium']    = None   # 正=多頭錢流，負=空頭
        d['flow_note']       = ''

        def _fetch_oi(target_expiry_str, available_dates):
            """通用：找最接近 target_expiry_str 的到期日並抓 OI。
            回傳 (result_dict, calls_df, puts_df)，失敗時拋出例外。"""
            if not available_dates or not target_expiry_str:
                raise ValueError('無可用到期日或目標日期')
            target_dt = datetime.strptime(target_expiry_str, '%Y-%m-%d')
            best = min(available_dates,
                       key=lambda e: abs((datetime.strptime(e, '%Y-%m-%d') - target_dt).days))
            chain  = t.option_chain(best)
            calls  = chain.calls.copy()
            puts   = chain.puts.copy()
            calls['openInterest'] = pd.to_numeric(calls['openInterest'], errors='coerce').fillna(0)
            puts['openInterest']  = pd.to_numeric(puts['openInterest'],  errors='coerce').fillna(0)

            result = {'expiry': best}

            # Put Wall（現價以下 OI 最大的 Put strike）
            puts_below = puts[puts['strike'] <= price * 1.02]
            if not puts_below.empty:
                result['put_wall'] = float(puts_below.loc[puts_below['openInterest'].idxmax(), 'strike'])

            # Call Wall（現價以上 OI 最大的 Call strike）
            calls_above = calls[calls['strike'] >= price * 0.98]
            if not calls_above.empty:
                result['call_wall'] = float(calls_above.loc[calls_above['openInterest'].idxmax(), 'strike'])

            # Max Pain（傳入現價做範圍過濾）
            result['max_pain'] = calc_max_pain(calls, puts, current_price=price)

            return result, calls, puts  # 永遠回傳 3-tuple

        # 初始化 — 避免 NameError
        last_calls = None
        last_puts  = pd.DataFrame()

        # 四週累積流量用容器
        all_calls_list = []
        all_puts_list  = []

        try:
            available = t.options
            four_weeks = get_four_weekly_expiries()

            # ── 四週期權 OI ─────────────────────────────
            for i, wk_info in enumerate(four_weeks):
                k = WEEK_KEYS[i]
                d[f'{k}_label']      = wk_info['label']
                d[f'{k}_is_monthly'] = wk_info['is_monthly']
                try:
                    res, calls_df, puts_df = _fetch_oi(wk_info['date'], available)
                    d[f'{k}_expiry']    = res.get('expiry')
                    d[f'{k}_put_wall']  = res.get('put_wall')
                    d[f'{k}_call_wall'] = res.get('call_wall')
                    d[f'{k}_max_pain']  = res.get('max_pain')
                    last_calls = calls_df
                    last_puts  = puts_df
                    monthly_note = '★月結' if wk_info['is_monthly'] else ''
                    print(f'    {wk_info["label"]}{monthly_note} ({d[f"{k}_expiry"]})  Put牆:{d[f"{k}_put_wall"]}  Pain:{d[f"{k}_max_pain"]}  Call牆:{d[f"{k}_call_wall"]}')

                    # 累積四週數據供流量計算
                    all_calls_list.append(calls_df)
                    all_puts_list.append(puts_df)

                except Exception as e:
                    print(f'    {wk_info["label"]}期權錯誤 {ticker}: {e}')

            # ── 流量訊號：四週全部加總計算 ──────────────────
            if all_calls_list:
                try:
                    # 合併四週期權鏈
                    all_calls = pd.concat(all_calls_list, ignore_index=True)
                    all_puts  = pd.concat(all_puts_list,  ignore_index=True) if all_puts_list else pd.DataFrame()

                    c_vol = pd.to_numeric(all_calls['volume'],       errors='coerce').fillna(0)
                    p_vol = pd.to_numeric(all_puts['volume'],        errors='coerce').fillna(0) if not all_puts.empty else pd.Series([0])
                    c_oi  = pd.to_numeric(all_calls['openInterest'], errors='coerce').fillna(0)
                    p_oi  = pd.to_numeric(all_puts['openInterest'],  errors='coerce').fillna(0) if not all_puts.empty else pd.Series([0])

                    total_cv  = float(c_vol.sum())
                    total_pv  = float(p_vol.sum())
                    total_coi = float(c_oi.sum())
                    total_poi = float(p_oi.sum())

                    d['flow_call_vol'] = int(total_cv)
                    d['flow_put_vol']  = int(total_pv)
                    d['flow_call_oi']  = int(total_coi)
                    d['flow_put_oi']   = int(total_poi)

                    # vol/OI 比（四週加總，>1 = 今日建倉超過歷史累積）
                    d['flow_call_ratio'] = round(total_cv / total_coi, 2) if total_coi > 0 else None
                    d['flow_put_ratio']  = round(total_pv / total_poi, 2) if total_poi > 0 else None

                    # Premium 估算：全部到期日的 ATM ±10% 範圍內 mid × volume
                    def calc_premium_allexp(df, vol_series):
                        df = df.copy()
                        df['mid'] = (pd.to_numeric(df.get('bid', 0), errors='coerce').fillna(0) +
                                     pd.to_numeric(df.get('ask', 0), errors='coerce').fillna(0)) / 2
                        df['vol'] = vol_series.values
                        # 過濾現價 ±10% 範圍（排除極遠 OTM 噪音）
                        in_range = (df['strike'] >= price * 0.90) & (df['strike'] <= price * 1.10)
                        df = df[in_range]
                        return float((df['mid'] * df['vol']).sum())

                    call_prem = calc_premium_allexp(all_calls, c_vol)
                    put_prem  = calc_premium_allexp(all_puts,  p_vol) if not all_puts.empty else 0.0
                    d['flow_premium'] = round(call_prem - put_prem, 0)

                    # 判斷訊號
                    cr   = d['flow_call_ratio'] or 0
                    pr   = d['flow_put_ratio']  or 0
                    prem = d['flow_premium']    or 0

                    # 閾值：四週合計 vol/OI > 0.3 就算異常（比單週標準寬鬆）
                    THRESH    = 0.3
                    call_active = cr > THRESH
                    put_active  = pr > THRESH
                    call_dom    = total_cv > total_pv * 1.5
                    put_dom     = total_pv > total_cv * 1.5

                    if call_active and call_dom and prem > 0:
                        d['flow_signal'] = 'CALL_SWEEP'
                        d['flow_note']   = f'Call 掃單 4週加總 (vol/OI={cr:.2f}，淨溢價 ${prem:,.0f})'
                    elif put_active and put_dom and prem < 0:
                        d['flow_signal'] = 'PUT_SWEEP'
                        d['flow_note']   = f'Put 掃單 4週加總 (vol/OI={pr:.2f}，淨溢價 ${abs(prem):,.0f})'
                    elif call_active and put_active:
                        d['flow_signal'] = 'MIXED'
                        d['flow_note']   = f'雙向異常 4週 (C={cr:.2f} P={pr:.2f})'
                    else:
                        d['flow_signal'] = 'NEUTRAL'
                        d['flow_note']   = '無異常流量'

                    print(f'    流量訊號(4週): {d["flow_signal"]} | {d["flow_note"]}')
                except Exception as fe:
                    print(f'    流量計算錯誤: {fe}')

            # ── IV / IVR / P/C（用最近一週期鏈）─────────
            if last_calls is not None:
                atm = last_calls.iloc[(last_calls['strike'] - price).abs().argsort()[:3]]
                iv_vals = pd.to_numeric(atm['impliedVolatility'], errors='coerce').dropna()
                if not iv_vals.empty:
                    d['iv']  = round(float(iv_vals.mean()) * 100, 1)
                    d['ivr'] = calc_ivr(t, d['iv'])

                cv = pd.to_numeric(last_calls['volume'], errors='coerce').sum()
                pv = pd.to_numeric(last_puts['volume'],  errors='coerce').sum() if not last_puts.empty else 0
                if cv and cv > 0:
                    d['pc_ratio'] = round(float(pv) / float(cv), 2)

        except Exception as e:
            print(f'    期權錯誤 {ticker}: {e}')

        d['ok'] = True

    except Exception as e:
        print(f'  錯誤 {ticker}: {e}')

    return d


def fetch_spy_qqq():
    """抓取 SPY / QQQ 技術面數據 + 四週期權 OI & Max Pain & 四週流量"""
    result = {}
    four_weeks = get_four_weekly_expiries()
    WEEK_KEYS = ['w0', 'w1', 'w2', 'w3']

    for sym in ['SPY', 'QQQ']:
        try:
            t    = yf.Ticker(sym)
            hist = t.history(period='1y')
            if hist.empty:
                continue
            closes = hist['Close']
            price  = float(closes.iloc[-1])
            prev   = float(closes.iloc[-2])
            entry = {
                'price':      round(price, 2),
                'change_pct': round((price - prev) / prev * 100, 2),
                'ma50':       round(float(closes.rolling(50).mean().iloc[-1]), 2),
                'ma200':      round(float(closes.rolling(200).mean().iloc[-1]), 2),
                # 流量欄位
                'flow_signal':     None,
                'flow_call_vol':   None,
                'flow_put_vol':    None,
                'flow_call_ratio': None,
                'flow_put_ratio':  None,
                'flow_premium':    None,
                'flow_note':       '',
            }

            # 初始化四週期權欄位
            for k in WEEK_KEYS:
                entry[f'{k}_expiry']     = None
                entry[f'{k}_put_wall']   = None
                entry[f'{k}_call_wall']  = None
                entry[f'{k}_max_pain']   = None
                entry[f'{k}_is_monthly'] = False
                entry[f'{k}_label']      = ''

            # 取得可用到期日
            available = t.options

            def _fetch_oi_etf(target_date_str):
                if not available or not target_date_str:
                    raise ValueError('無到期日')
                target_dt = datetime.strptime(target_date_str, '%Y-%m-%d')
                best = min(available, key=lambda e: abs((datetime.strptime(e, '%Y-%m-%d') - target_dt).days))
                chain = t.option_chain(best)
                calls = chain.calls.copy()
                puts  = chain.puts.copy()
                calls['openInterest'] = pd.to_numeric(calls['openInterest'], errors='coerce').fillna(0)
                puts['openInterest']  = pd.to_numeric(puts['openInterest'],  errors='coerce').fillna(0)
                res = {'expiry': best}
                puts_below = puts[puts['strike'] <= price * 1.02]
                if not puts_below.empty:
                    res['put_wall'] = float(puts_below.loc[puts_below['openInterest'].idxmax(), 'strike'])
                calls_above = calls[calls['strike'] >= price * 0.98]
                if not calls_above.empty:
                    res['call_wall'] = float(calls_above.loc[calls_above['openInterest'].idxmax(), 'strike'])
                res['max_pain'] = calc_max_pain(calls, puts, current_price=price)
                return res, calls, puts

            all_calls_list = []
            all_puts_list  = []

            for i, wk_info in enumerate(four_weeks):
                k = WEEK_KEYS[i]
                entry[f'{k}_label']      = wk_info['label']
                entry[f'{k}_is_monthly'] = wk_info['is_monthly']
                try:
                    res, calls_df, puts_df = _fetch_oi_etf(wk_info['date'])
                    entry[f'{k}_expiry']    = res.get('expiry')
                    entry[f'{k}_put_wall']  = res.get('put_wall')
                    entry[f'{k}_call_wall'] = res.get('call_wall')
                    entry[f'{k}_max_pain']  = res.get('max_pain')
                    all_calls_list.append(calls_df)
                    all_puts_list.append(puts_df)
                    mo_note = '★月結' if wk_info['is_monthly'] else ''
                    print(f'    {sym} {wk_info["label"]}{mo_note} ({entry[f"{k}_expiry"]})  Put牆:{entry[f"{k}_put_wall"]}  Pain:{entry[f"{k}_max_pain"]}  Call牆:{entry[f"{k}_call_wall"]}')
                except Exception as e:
                    print(f'    {sym} {wk_info["label"]}期權錯誤: {e}')

            # ── 四週流量加總 ──────────────────────────────
            if all_calls_list:
                try:
                    ac = pd.concat(all_calls_list, ignore_index=True)
                    ap = pd.concat(all_puts_list,  ignore_index=True) if all_puts_list else pd.DataFrame()

                    c_vol = pd.to_numeric(ac['volume'],       errors='coerce').fillna(0)
                    p_vol = pd.to_numeric(ap['volume'],       errors='coerce').fillna(0) if not ap.empty else pd.Series([0])
                    c_oi  = pd.to_numeric(ac['openInterest'], errors='coerce').fillna(0)
                    p_oi  = pd.to_numeric(ap['openInterest'], errors='coerce').fillna(0) if not ap.empty else pd.Series([0])

                    tcv = float(c_vol.sum()); tpv = float(p_vol.sum())
                    tco = float(c_oi.sum());  tpo = float(p_oi.sum())

                    entry['flow_call_vol']   = int(tcv)
                    entry['flow_put_vol']    = int(tpv)
                    entry['flow_call_ratio'] = round(tcv / tco, 2) if tco > 0 else None
                    entry['flow_put_ratio']  = round(tpv / tpo, 2) if tpo > 0 else None

                    def etf_prem(df, vol_s):
                        df = df.copy(); df['mid'] = (pd.to_numeric(df.get('bid',0),errors='coerce').fillna(0)+pd.to_numeric(df.get('ask',0),errors='coerce').fillna(0))/2
                        df['vol'] = vol_s.values; rng = (df['strike']>=price*0.90)&(df['strike']<=price*1.10)
                        return float((df[rng]['mid']*df[rng]['vol']).sum())

                    cp = etf_prem(ac, c_vol)
                    pp = etf_prem(ap, p_vol) if not ap.empty else 0.0
                    entry['flow_premium'] = round(cp - pp, 0)

                    cr = entry['flow_call_ratio'] or 0
                    pr = entry['flow_put_ratio']  or 0
                    prem = entry['flow_premium']  or 0
                    THRESH = 0.3

                    if cr > THRESH and tcv > tpv * 1.5 and prem > 0:
                        entry['flow_signal'] = 'CALL_SWEEP'
                        entry['flow_note']   = f'Call 掃單 4週 (vol/OI={cr:.2f}，淨溢價 ${prem:,.0f})'
                    elif pr > THRESH and tpv > tcv * 1.5 and prem < 0:
                        entry['flow_signal'] = 'PUT_SWEEP'
                        entry['flow_note']   = f'Put 掃單 4週 (vol/OI={pr:.2f}，淨溢價 ${abs(prem):,.0f})'
                    elif cr > THRESH and pr > THRESH:
                        entry['flow_signal'] = 'MIXED'
                        entry['flow_note']   = f'雙向異常 4週 (C={cr:.2f} P={pr:.2f})'
                    else:
                        entry['flow_signal'] = 'NEUTRAL'
                        entry['flow_note']   = '無異常流量'

                    print(f'    {sym} 流量訊號(4週): {entry["flow_signal"]} | {entry["flow_note"]}')
                except Exception as fe:
                    print(f'    {sym} 流量計算錯誤: {fe}')

            result[sym] = entry
        except Exception:
            pass
    return result


# ═══════════════════════════════════════════════════════════════
#  HTML 生成
# ═══════════════════════════════════════════════════════════════

def ma_html(price, ma, label):
    if ma is None:
        return f'<div class="man">N/A</div>'
    pct = (price - ma) / ma * 100
    above = pct >= 0
    cls   = 'mau' if above else 'mad'
    arrow = '▲' if above else '▼'
    tag   = '站上' if above else '跌破'
    sign  = '+' if above else ''
    return (f'<div class="{cls} ma-price">${ma:.2f}</div>'
            f'<div class="{cls} ma-tag">{arrow} {tag} {sign}{pct:.1f}%</div>')


def pe_html(pe_ttm, pe_fwd):
    if pe_ttm is None:
        ttm_html = '<div class="pe pe-l">虧損中</div><div class="pe-note neg">EPS 為負</div>'
    elif pe_ttm > 200:
        ttm_html = f'<div class="pe pe-w">{pe_ttm}x</div><div class="pe-note" style="color:#e3b341">⚠ 極高估值</div>'
    else:
        ttm_html = f'<div class="pe">{pe_ttm}x</div><div class="pe-note" style="color:#484f58">TTM</div>'

    if pe_fwd is None:
        fwd_html = '<div class="fpe-fl">—</div>'
    elif pe_ttm is not None and pe_fwd < pe_ttm:
        color = '#e3b341' if pe_fwd > 100 else '#3fb950'
        fwd_html = f'<div class="fpe-dn" style="color:{color}">{pe_fwd}x</div><div class="pe-note pos">↓ 獲利成長</div>'
    else:
        fwd_html = f'<div class="fpe-fl">{pe_fwd}x</div><div class="pe-note gf">≈ 平穩</div>'

    return ttm_html, fwd_html


def iv_html(iv, ivr):
    if iv is None or ivr is None:
        return '<div class="ivr-lbl">N/A</div>', '<span class="ivb ivn">—</span>'

    if ivr < 25:
        bar_cls, badge_cls, badge_txt = 'ivrl', 'ivs', '低 IVR'
    elif ivr < 50:
        bar_cls, badge_cls, badge_txt = 'ivrm', 'ivn', '正常'
    elif ivr < 75:
        bar_cls, badge_cls, badge_txt = 'ivrm', 'ive', '偏高 ⚠'
    else:
        bar_cls, badge_cls, badge_txt = 'ivrh', 'ivh', '極高 🔥'

    bar = f'''<div class="ivr-bg"><div class="ivr-f {bar_cls}" style="width:{ivr}%"></div></div>
    <div class="ivr-lbl">{iv:.1f}% · IVR {ivr}%</div>'''
    badge = f'<span class="ivb {badge_cls}">{badge_txt}</span>'
    return bar, badge


def pc_html(pc):
    if pc is None:
        return '<div class="pc-n">—</div>'
    if pc < 0.7:
        cls, note = 'pc-b', '看多傾向'
    elif pc > 1.0:
        cls, note = 'pc-r', '避險需求'
    else:
        cls, note = 'pc-n', '中性'
    return f'<div class="{cls}">{pc}</div><div class="pc-l">{note}</div>'


def flow_html(d):
    """渲染期權流量訊號欄位"""
    sig  = d.get('flow_signal')
    note = d.get('flow_note', '')
    cv   = d.get('flow_call_vol')
    pv   = d.get('flow_put_vol')
    cr   = d.get('flow_call_ratio')
    pr   = d.get('flow_put_ratio')
    prem = d.get('flow_premium')

    if sig is None:
        return '<div style="color:#484f58;font-size:0.75em">N/A</div>'

    def fmt_k(v):
        if v is None: return '—'
        return f'{v/1000:.1f}K' if v >= 1000 else str(v)

    def fmt_prem(v):
        if v is None: return '—'
        sign = '+' if v >= 0 else ''
        return f'{sign}${abs(v)/1000:.1f}K' if abs(v) >= 1000 else f'{sign}${v:.0f}'

    if sig == 'CALL_SWEEP':
        badge_cls  = 'flow-call'
        badge_text = '🟢 Call Sweep'
    elif sig == 'PUT_SWEEP':
        badge_cls  = 'flow-put'
        badge_text = '🔴 Put Sweep'
    elif sig == 'MIXED':
        badge_cls  = 'flow-mix'
        badge_text = '🟡 雙向異常'
    else:
        badge_cls  = 'flow-neu'
        badge_text = '⚪ 無異常'

    ratio_bar_c = min(int((cr or 0) / 2 * 100), 100)
    ratio_bar_p = min(int((pr or 0) / 2 * 100), 100)

    return f'''<div class="{badge_cls} flow-badge">{badge_text}</div>
<div class="flow-row"><span class="flow-lbl">C vol/OI</span><span class="flow-bar-wrap"><span class="flow-bar-c" style="width:{ratio_bar_c}%"></span></span><span class="flow-val">{cr:.2f}x</span></div>
<div class="flow-row"><span class="flow-lbl">P vol/OI</span><span class="flow-bar-wrap"><span class="flow-bar-p" style="width:{ratio_bar_p}%"></span></span><span class="flow-val">{pr:.2f}x</span></div>
<div class="flow-row" style="margin-top:3px"><span class="flow-lbl" style="color:#484f58">C{fmt_k(cv)}/P{fmt_k(pv)}</span><span class="flow-val" style="color:{"#3fb950" if (prem or 0)>=0 else "#f85149"}">{fmt_prem(prem)}</span></div>''' if cr is not None and pr is not None else f'<div class="{badge_cls} flow-badge">{badge_text}</div><div style="font-size:0.7em;color:#484f58">{note}</div>'


def ms_html(ticker, price):
    ms = MORNINGSTAR.get(ticker, {})
    if not ms:
        return '<div class="ms-fv">—</div>'
    fv    = ms['fv']
    stars = ms['stars']
    moat  = ms['moat']
    pfv   = price / fv
    disc  = round((1 - pfv) * 100)

    if disc > 0:
        pct_html = f'<span class="ms-disc">▼{disc}%折</span>'
    elif disc < 0:
        pct_html = f'<span class="ms-prem">▲{abs(disc)}%溢</span>'
    else:
        pct_html = '<span class="ms-neu">合理</span>'

    star_str = '★' * stars + '☆' * (5 - stars)

    moat_cls = {'Wide': 'ms-mw', 'Narrow': 'ms-mn', 'None': 'ms-m0'}.get(moat, 'ms-m0')
    moat_txt = {'Wide': 'Wide 護城河', 'Narrow': 'Narrow 護城河', 'None': 'No Moat'}.get(moat, '—')

    return f'''<div class="ms-fv">${fv} {pct_html}</div>
    <div class="ms-str">{star_str}</div>
    <span class="{moat_cls}">{moat_txt}</span>'''


def oi_html_single(d, pfx, label=''):
    """單一到期期別的 OI 格，pfx='w0'~'w3'"""
    cw    = d.get(f'{pfx}_call_wall')
    mp    = d.get(f'{pfx}_max_pain')
    pw    = d.get(f'{pfx}_put_wall')
    exp   = d.get(f'{pfx}_expiry', '') or ''
    is_mo = d.get(f'{pfx}_is_monthly', False)
    wk_label = d.get(f'{pfx}_label', label)
    exp_short = exp[5:] if exp else wk_label  # MM-DD

    def fmt(v):
        return f'${v:g}' if v is not None else '—'

    mo_badge = '<span style="font-size:0.65em;color:#e3b341;margin-left:3px;background:rgba(227,179,65,.12);border:1px solid rgba(227,179,65,.3);border-radius:3px;padding:0 3px">月結</span>' if is_mo else ''

    if cw is None and mp is None and pw is None:
        return f'<div style="color:#484f58;font-size:0.75em">{wk_label}{mo_badge}<br>N/A</div>'

    return f'''<div class="oi-lbl">{wk_label} {exp_short}{mo_badge}</div>
<div class="oi-call">▼ {fmt(cw)}</div>
<div class="oi-pain">⚡ {fmt(mp)}</div>
<div class="oi-put">▲ {fmt(pw)}</div>'''


def stock_row(d):
    if not d.get('ok'):
        return f'''<tr>
  <td><div class="tk"><span class="tb">{d["ticker"]}</span>{"<span class='m7'>Mag 7</span>" if d["ticker"] in MAG7 else ""}</div></td>
  <td colspan="17" style="color:#f85149;font-size:0.8em">⚠ 數據抓取失敗</td>
</tr>'''

    ticker  = d['ticker']
    price   = d['price']
    chg     = d['change']
    chg_pct = d['change_pct']
    vol     = d['vol_ratio']

    # Change styling
    if chg > 0:
        chg_cls = 'pos'
        chg_str = f'+{chg:.2f}'
        pct_str = f'+{chg_pct:.2f}%'
    elif chg < 0:
        chg_cls = 'neg'
        chg_str = f'{chg:.2f}'
        pct_str = f'{chg_pct:.2f}%'
    else:
        chg_cls = 'neu'
        chg_str = '—'
        pct_str = '—'

    # Volume
    if vol >= 1.5:
        vol_html = f'<span class="vh">{vol:.2f}x</span><span class="vb">放量</span>'
    else:
        vol_html = f'{vol:.2f}x'

    # MA
    ma20_h  = ma_html(price, d['ma20'],  'MA20')
    ma50_h  = ma_html(price, d['ma50'],  'MA50')
    ma200_h = ma_html(price, d['ma200'], 'MA200')

    # P/E
    ttm_h, fwd_h = pe_html(d['pe_ttm'], d['pe_fwd'])

    # IV
    iv_bar, iv_badge = iv_html(d['iv'], d['ivr'])

    # P/C
    pc_h = pc_html(d['pc_ratio'])

    # OI 四欄
    oi_w0_h = oi_html_single(d, 'w0')
    oi_w1_h = oi_html_single(d, 'w1')
    oi_w2_h = oi_html_single(d, 'w2')
    oi_w3_h = oi_html_single(d, 'w3')

    # 流量訊號
    flow_h = flow_html(d)

    # Morningstar
    ms_h = ms_html(ticker, price)

    # Market cap
    mc = fmt_cap(d.get('market_cap'))

    mag7_badge = f'<span class="m7">Mag 7</span>' if ticker in MAG7 else ''

    return f'''<tr>
  <td><div class="tk"><span class="tb">{ticker}</span>{mag7_badge}</div></td>
  <td class="pr">${price:.2f}</td>
  <td class="oi">{oi_w0_h}</td>
  <td class="oi">{oi_w1_h}</td>
  <td class="oi">{oi_w2_h}</td>
  <td class="oi">{oi_w3_h}</td>
  <td class="flow-td">{flow_h}</td>
  <td class="{chg_cls}">{chg_str}</td><td class="{chg_cls} pct">{pct_str}</td>
  <td>{vol_html}</td>
  <td>{ms_h}</td>
  <td>{ma20_h}</td><td>{ma50_h}</td><td>{ma200_h}</td>
  <td>{ttm_h}</td>
  <td>{fwd_h}</td>
  <td>{iv_bar}</td>
  <td>{iv_badge}</td>
  <td>{pc_h}</td>
  <td class="mc">{mc}</td>
</tr>'''


def summary_cards(stocks):
    ok = [s for s in stocks if s.get('ok')]
    up   = sum(1 for s in ok if s['change'] > 0)
    down = sum(1 for s in ok if s['change'] < 0)
    flat = len(stocks) - up - down
    best = max(ok, key=lambda s: s['change_pct'], default=None)
    best_str = f'{best["ticker"]} {best["change_pct"]:+.2f}%' if best else '—'
    iv_vals = [s['ivr'] for s in ok if s.get('ivr') is not None]
    avg_ivr = int(sum(iv_vals) / len(iv_vals)) if iv_vals else 0
    total = len(stocks)

    # 流量統計
    call_sweeps = [s['ticker'] for s in ok if s.get('flow_signal') == 'CALL_SWEEP']
    put_sweeps  = [s['ticker'] for s in ok if s.get('flow_signal') == 'PUT_SWEEP']
    mixed       = [s['ticker'] for s in ok if s.get('flow_signal') == 'MIXED']
    flow_str  = ' '.join(call_sweeps) if call_sweeps else '—'
    fput_str  = ' '.join(put_sweeps)  if put_sweeps  else '—'

    return f'''
  <div class="sc"><div class="l">上漲</div><div class="v gu">▲ {up}</div></div>
  <div class="sc"><div class="l">下跌</div><div class="v gd">▼ {down}</div></div>
  <div class="sc"><div class="l">數據不完整</div><div class="v gf">— {flat}</div></div>
  <div class="sc"><div class="l">市場情緒</div><div class="v gb">{"偏多 🟢" if up > down else "偏空 🔴" if down > up else "中性 ⚪"}</div></div>
  <div class="sc"><div class="l">最強漲幅</div><div class="v gu" style="font-size:1em">{best_str}</div></div>
  <div class="sc"><div class="l">整體 IV 水位</div><div class="v gmx">IVR≈{avg_ivr}%</div></div>
  <div class="sc"><div class="l">掃描股票</div><div class="v gbl">{total}</div></div>
  <div class="sc" style="min-width:160px"><div class="l">🟢 Call Sweep</div><div class="v" style="font-size:0.85em;color:#3fb950">{flow_str}</div></div>
  <div class="sc" style="min-width:160px"><div class="l">🔴 Put Sweep</div><div class="v" style="font-size:0.85em;color:#f85149">{fput_str}</div></div>'''


def vix_bar_pct(val):
    # 0-40 scale → 0-100%
    return min(int(val / 40 * 100), 100)


def generate_report(stocks, vix, spy_qqq, date_str):
    rows_html = '\n'.join(stock_row(s) for s in stocks)
    cards_html = summary_cards(stocks)

    vix_val = vix['value']
    vix_chg = vix['change']
    vix_pct = vix['pct']
    vix_sign = '▼' if vix_chg <= 0 else '▲'
    vix_cls  = 'vix-val' + (' style="color:#f85149"' if vix_val > 25 else '')

    spy = spy_qqq.get('SPY', {})
    qqq = spy_qqq.get('QQQ', {})

    spy_chg_cls = 'pos' if spy.get('change_pct', 0) >= 0 else 'neg'
    qqq_chg_cls = 'pos' if qqq.get('change_pct', 0) >= 0 else 'neg'

    # S5TW / S5FI breadth display
    def breadth_color(val):
        if val is None:
            return '#8b949e'
        if val >= 70:
            return '#3fb950'
        if val >= 40:
            return '#d29922'
        return '#f85149'

    s5tw_val = vix.get('s5tw')
    s5fi_val = vix.get('s5fi')
    s5tw_str = f'{s5tw_val:.1f}%' if s5tw_val is not None else 'N/A'
    s5fi_str = f'{s5fi_val:.1f}%' if s5fi_val is not None else 'N/A'
    s5tw_color = breadth_color(s5tw_val)
    s5fi_color = breadth_color(s5fi_val)

    # SPY / QQQ 四週 OI Grid HTML
    def etf_oi_grid_html(etf_data):
        WEEK_KEYS = ['w0', 'w1', 'w2', 'w3']
        cells = []
        for k in WEEK_KEYS:
            lbl    = etf_data.get(f'{k}_label', '')
            exp    = etf_data.get(f'{k}_expiry', '') or ''
            is_mo  = etf_data.get(f'{k}_is_monthly', False)
            cw     = etf_data.get(f'{k}_call_wall')
            mp     = etf_data.get(f'{k}_max_pain')
            pw     = etf_data.get(f'{k}_put_wall')
            exp_short = exp[5:] if exp else ''
            mo_badge  = '<span style="font-size:0.6em;color:#e3b341;background:rgba(227,179,65,.12);border:1px solid rgba(227,179,65,.3);border-radius:3px;padding:0 3px;margin-left:2px">月結</span>' if is_mo else ''
            fmt = lambda v: f'${v:g}' if v is not None else '—'
            cell = f'''<div style="background:#21262d;border-radius:6px;padding:7px 10px">
  <div class="oi-lbl">{lbl} {exp_short}{mo_badge}</div>
  <div class="oi-call" style="font-size:0.85em">▼ {fmt(cw)}</div>
  <div class="oi-pain" style="font-size:0.85em">⚡ {fmt(mp)}</div>
  <div class="oi-put" style="font-size:0.85em">▲ {fmt(pw)}</div>
</div>'''
            cells.append(cell)
        return '\n'.join(cells)

    spy_oi_grid = etf_oi_grid_html(spy)
    qqq_oi_grid = etf_oi_grid_html(qqq)

    spy_flow_html = flow_html(spy)
    qqq_flow_html = flow_html(qqq)

    now_tw = datetime.now().strftime('%Y-%m-%d %H:%M')
    weekday_map = {0:'一',1:'二',2:'三',3:'四',4:'五',5:'六',6:'日'}
    wd = weekday_map[datetime.now().weekday()]
    display_date = datetime.now().strftime(f'%Y年%m月%d日（週{wd}）')

    return f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>美股每日開盤掃描報告 - {date_str}</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','Noto Sans TC','PingFang TC',sans-serif;background:#0d1117;color:#c9d1d9;min-height:100vh;padding:16px;line-height:1.5;font-size:16px}}
.wrap{{max-width:1600px;margin:0 auto}}
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
td{{padding:13px 14px;vertical-align:middle}}
.tk{{display:flex;flex-direction:column;gap:3px}}
.tb{{background:#21262d;border:1px solid #30363d;border-radius:5px;padding:2px 8px;font-weight:700;font-size:0.86em;color:#79c0ff;font-family:monospace;letter-spacing:.5px;width:fit-content}}
.m7{{font-size:0.58em;background:rgba(121,192,255,.08);border:1px solid rgba(121,192,255,.15);color:#79c0ff;border-radius:3px;padding:1px 5px;width:fit-content}}
.cn{{color:#8b949e;font-size:0.8em;white-space:nowrap}}
.pr{{font-weight:700;color:#e6edf3;font-family:monospace;white-space:nowrap}}
.pos{{color:#3fb950}}.neg{{color:#f85149}}.neu{{color:#8b949e}}.pct{{font-weight:600}}
.vh{{color:#e3b341;font-weight:600}}
.vb{{display:inline-block;background:rgba(227,179,65,.12);border:1px solid rgba(227,179,65,.25);color:#e3b341;border-radius:4px;padding:1px 4px;font-size:0.72em;margin-left:3px;vertical-align:middle}}
.mau{{color:#3fb950;white-space:nowrap}}
.mad{{color:#f85149;white-space:nowrap}}
.man{{color:#484f58}}
.ma-price{{font-family:monospace;font-size:0.92em;font-weight:700}}
.ma-tag{{font-size:0.72em;margin-top:2px}}
.pe{{font-family:monospace;font-size:0.95em;font-weight:700;white-space:nowrap}}
.pe-w{{color:#e3b341}}.pe-l{{color:#f85149;font-size:0.85em}}
.fpe-dn{{color:#3fb950;font-family:monospace;font-size:0.92em;font-weight:700}}
.fpe-fl{{color:#8b949e;font-family:monospace;font-size:0.92em;font-weight:700}}
.pe-note{{font-size:0.7em;margin-top:2px}}
.iv{{font-family:monospace;font-size:0.84em;font-weight:700}}
.ivb{{display:inline-block;padding:2px 7px;border-radius:8px;font-size:0.72em;font-weight:600}}
.ivs{{background:rgba(63,185,80,.1);color:#3fb950;border:1px solid rgba(63,185,80,.2)}}
.ivn{{background:rgba(139,148,158,.1);color:#8b949e;border:1px solid rgba(139,148,158,.2)}}
.ive{{background:rgba(210,153,34,.12);color:#d29922;border:1px solid rgba(210,153,34,.25)}}
.ivh{{background:rgba(248,81,73,.12);color:#f85149;border:1px solid rgba(248,81,73,.25)}}
.ivr-bg{{background:#21262d;border-radius:3px;height:6px;margin-bottom:4px;overflow:hidden;width:78px}}
.ivr-f{{height:100%;border-radius:3px}}
.ivrl{{background:#3fb950}}.ivrm{{background:#d29922}}.ivrh{{background:#f85149}}
.ivr-lbl{{font-size:0.75em;color:#8b949e}}
.oi{{min-width:110px;max-width:130px}}
.oi-lbl{{font-size:0.7em;color:#484f58;margin-bottom:3px;letter-spacing:0.5px;text-transform:uppercase}}
.oi-call{{font-family:monospace;font-size:0.92em;color:#f85149;white-space:nowrap}}
.oi-pain{{font-family:monospace;font-size:0.92em;color:#d29922;white-space:nowrap;margin:3px 0}}
.oi-put{{font-family:monospace;font-size:0.92em;color:#3fb950;white-space:nowrap}}
.ms-fv{{font-family:monospace;font-size:0.95em;font-weight:700;color:#e6edf3}}
.ms-disc{{font-size:0.78em;color:#3fb950;font-weight:600}}
.ms-prem{{font-size:0.78em;color:#f85149;font-weight:600}}
.ms-neu{{font-size:0.78em;color:#8b949e;font-weight:600}}
.ms-str{{color:#e3b341;font-size:1em;letter-spacing:2px;margin:3px 0;line-height:1}}
.ms-mw{{display:inline-block;font-size:0.6em;padding:1px 5px;border-radius:3px;background:rgba(121,192,255,.1);color:#79c0ff;border:1px solid rgba(121,192,255,.2)}}
.ms-mn{{display:inline-block;font-size:0.6em;padding:1px 5px;border-radius:3px;background:rgba(210,153,34,.1);color:#d29922;border:1px solid rgba(210,153,34,.2)}}
.ms-m0{{display:inline-block;font-size:0.6em;padding:1px 5px;border-radius:3px;background:rgba(72,79,88,.1);color:#484f58;border:1px solid rgba(72,79,88,.2)}}
.pc-b{{color:#3fb950;font-weight:600;font-family:monospace}}
.pc-r{{color:#f85149;font-weight:600;font-family:monospace}}
.pc-n{{color:#8b949e;font-weight:600;font-family:monospace}}
.pc-l{{font-size:0.65em;color:#484f58;margin-top:2px}}
.mc{{color:#8b949e;font-size:0.8em;font-family:monospace;white-space:nowrap}}
.leg{{background:#161b22;border:1px solid #30363d;border-radius:10px;padding:10px 16px;margin-bottom:12px;display:flex;gap:14px;flex-wrap:wrap;align-items:center;font-size:0.73em;color:#8b949e}}
.leg strong{{color:#c9d1d9}}
.mkt-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}}
@media(max-width:780px){{.mkt-grid{{grid-template-columns:1fr}}}}
.mkc{{background:#161b22;border:1px solid #30363d;border-radius:12px;padding:18px 20px}}
.mkr{{display:flex;justify-content:space-between;align-items:center;font-size:0.78em;padding:4px 0;border-bottom:1px solid #21262d}}
.mkr:last-child{{border-bottom:none}}
.mkr .k{{color:#8b949e}}.mkr .v{{font-family:monospace;font-weight:600}}
.mkc-tk{{font-size:1.3em;font-weight:800;color:#e6edf3;font-family:monospace}}
.mkc-sub{{font-size:0.7em;color:#8b949e;margin-top:2px}}
.mkc-p{{font-size:1.25em;font-weight:700;font-family:monospace;color:#e6edf3;text-align:right}}
.mkc-hdr{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px}}
.ftr{{background:#161b22;border:1px solid #30363d;border-radius:10px;padding:12px 18px;text-align:center;color:#484f58;font-size:0.7em;line-height:2;margin-top:14px}}
.flow-td{{min-width:150px;max-width:170px;vertical-align:middle}}
.flow-badge{{display:inline-block;font-size:0.75em;font-weight:700;border-radius:6px;padding:3px 8px;margin-bottom:5px;white-space:nowrap}}
.flow-call{{background:rgba(63,185,80,.12);color:#3fb950;border:1px solid rgba(63,185,80,.3)}}
.flow-put{{background:rgba(248,81,73,.12);color:#f85149;border:1px solid rgba(248,81,73,.3)}}
.flow-mix{{background:rgba(210,153,34,.12);color:#d29922;border:1px solid rgba(210,153,34,.3)}}
.flow-neu{{background:rgba(139,148,158,.08);color:#8b949e;border:1px solid rgba(139,148,158,.2)}}
.flow-row{{display:flex;align-items:center;gap:5px;margin-bottom:2px}}
.flow-lbl{{font-size:0.65em;color:#8b949e;min-width:42px;white-space:nowrap}}
.flow-bar-wrap{{flex:1;background:#21262d;border-radius:2px;height:4px;overflow:hidden;min-width:30px}}
.flow-bar-c{{display:block;height:100%;background:#3fb950;border-radius:2px}}
.flow-bar-p{{display:block;height:100%;background:#f85149;border-radius:2px}}
.flow-val{{font-family:monospace;font-size:0.72em;color:#c9d1d9;white-space:nowrap}}
</style>
</head>
<body>
<div class="wrap">

<div class="hdr">
  <div>
    <h1>📊 美股每日開盤掃描報告</h1>
    <div class="sub">Magnificent 7 + TSM · MU · SNDK · NFLX · AMD · AVGO · COST　｜　真實數據 yfinance　｜　晨星估值 · OI · IV · P/C Ratio · 機構流量</div>
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
  <div style="display:flex;flex-direction:column;justify-content:center;gap:6px;min-width:160px;border-left:1px solid #30363d;padding-left:18px">
    <div>
      <div class="vix-lbl" style="margin-bottom:2px">站上20日均線 ($S5TW)</div>
      <div style="font-size:1.3em;font-weight:700;font-family:monospace;color:{s5tw_color}">{s5tw_str}</div>
    </div>
    <div>
      <div class="vix-lbl" style="margin-bottom:2px">站上50日均線 ($S5FI)</div>
      <div style="font-size:1.3em;font-weight:700;font-family:monospace;color:{s5fi_color}">{s5fi_str}</div>
    </div>
  </div>
  <div class="vix-gauge">
    <div class="vix-g-lbl"><span>極低 &lt;12</span><span>正常 12–20</span><span>警戒 20–30</span><span>恐慌 &gt;30</span></div>
    <div class="vix-bar"><div class="vix-dot" style="left:{vix_bar_pct(vix_val)}%"></div></div>
  </div>
  <div class="vix-ctx">
    <strong>{"低波動 ✓" if vix_val < 20 else "警戒 ⚠" if vix_val < 30 else "恐慌 🔥"}</strong>　VIX {vix_val}
    {"位於正常區間（12–20），期權保護成本偏低。" if vix_val < 20 else "進入警戒區間，市場波動加劇，留意風險。" if vix_val < 30 else "極度恐慌，考慮逢低布局或持有避險部位。"}
  </div>
</div>

<div class="sum">{cards_html}</div>

<div class="leg">
  <strong>期權欄位說明：</strong>
  <span><span class="ivb ivs">低 IVR</span> &lt;25%</span>
  <span><span class="ivb ivn">正常</span> 25–50%</span>
  <span><span class="ivb ive">偏高</span> 50–75%</span>
  <span><span class="ivb ivh">極高</span> &gt;75%</span>
  <span>｜</span>
  <span class="oi-call">▼ Call Wall</span> = 壓力
  <span class="oi-pain">⚡ Max Pain</span> = 最大痛苦點
  <span class="oi-put">▲ Put Wall</span> = 支撐　｜　<span style="color:#e3b341">月結</span> = 當月第三週五
  <span>｜</span>
  <strong style="color:#a371f7">機構流量：</strong>
  <span class="flow-badge flow-call">🟢 Call Sweep</span> vol/OI &gt;0.5 且 Call 主導
  <span class="flow-badge flow-put">🔴 Put Sweep</span> vol/OI &gt;0.5 且 Put 主導
  <span class="flow-badge flow-mix">🟡 雙向異常</span> 雙側同時放量　｜　vol/OI &gt;1 = 今日新建倉超過歷史累積
</div>

<div class="stitle">📋 個股全覽</div>
<div class="tbl-wrap">
<table>
<thead>
<tr>
  <th rowspan="2">代碼</th>
  <th rowspan="2">現價</th>
  <th colspan="4" class="th-grp" style="color:#58a6ff;min-width:440px">── OI 支撐壓力（本週 / 下週 / 下下週 / 下下下週）──</th>
  <th rowspan="2" class="th-grp" style="color:#a371f7;min-width:140px">── 機構流量 ──</th>
  <th rowspan="2">漲跌$</th>
  <th rowspan="2">漲跌%</th>
  <th rowspan="2">量比</th>
  <th rowspan="2" style="min-width:100px">晨星估值</th>
  <th rowspan="2">MA20</th>
  <th rowspan="2">MA50</th>
  <th rowspan="2">MA200</th>
  <th colspan="2" class="th-grp th-val">── 估 值 ──</th>
  <th colspan="3" class="th-grp th-opt">── 期 權 分 析 ──</th>
  <th rowspan="2">市值</th>
</tr>
<tr>
  <th style="color:#58a6ff;font-size:0.7em">本週 Wall</th>
  <th style="color:#58a6ff;font-size:0.7em">下週 Wall</th>
  <th style="color:#58a6ff;font-size:0.7em">下下週 Wall</th>
  <th style="color:#58a6ff;font-size:0.7em">下下下週 Wall</th>
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
        <div class="mkc-p">${spy.get("price","—")}</div>
        <div class="mkc-c {spy_chg_cls}" style="font-size:0.8em;text-align:right">{spy.get("change_pct",0):+.2f}%</div>
      </div>
    </div>
    <div class="mkr"><span class="k">MA50</span><span class="v {"pos" if spy.get("price",0)>spy.get("ma50",0) else "neg"}">${spy.get("ma50","—")} {"▲ 站上" if spy.get("price",0)>spy.get("ma50",0) else "▼ 跌破"}</span></div>
    <div class="mkr"><span class="k">MA200</span><span class="v {"pos" if spy.get("price",0)>spy.get("ma200",0) else "neg"}">${spy.get("ma200","—")} {"▲ 站上" if spy.get("price",0)>spy.get("ma200",0) else "▼ 跌破"}</span></div>
    <div style="margin-top:10px;border-top:1px solid #21262d;padding-top:10px">
      <div class="oi-lbl" style="margin-bottom:6px;letter-spacing:0.5px">期權 OI — 四週展望</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
        {spy_oi_grid}
      </div>
    </div>
    <div style="margin-top:10px;border-top:1px solid #21262d;padding-top:8px">
      <div class="oi-lbl" style="margin-bottom:6px;letter-spacing:0.5px;color:#a371f7">機構流量（四週加總）</div>
      {spy_flow_html}
    </div>
  </div>
  <div class="mkc">
    <div class="mkc-hdr">
      <div><div class="mkc-tk">QQQ</div><div class="mkc-sub">Nasdaq 100 ETF</div></div>
      <div>
        <div class="mkc-p">${qqq.get("price","—")}</div>
        <div class="mkc-c {qqq_chg_cls}" style="font-size:0.8em;text-align:right">{qqq.get("change_pct",0):+.2f}%</div>
      </div>
    </div>
    <div class="mkr"><span class="k">MA50</span><span class="v {"pos" if qqq.get("price",0)>qqq.get("ma50",0) else "neg"}">${qqq.get("ma50","—")} {"▲ 站上" if qqq.get("price",0)>qqq.get("ma50",0) else "▼ 跌破"}</span></div>
    <div class="mkr"><span class="k">MA200</span><span class="v {"pos" if qqq.get("price",0)>qqq.get("ma200",0) else "neg"}">${qqq.get("ma200","—")} {"▲ 站上" if qqq.get("price",0)>qqq.get("ma200",0) else "▼ 跌破"}</span></div>
    <div style="margin-top:10px;border-top:1px solid #21262d;padding-top:10px">
      <div class="oi-lbl" style="margin-bottom:6px;letter-spacing:0.5px">期權 OI — 四週展望</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
        {qqq_oi_grid}
      </div>
    </div>
    <div style="margin-top:10px;border-top:1px solid #21262d;padding-top:8px">
      <div class="oi-lbl" style="margin-bottom:6px;letter-spacing:0.5px;color:#a371f7">機構流量（四週加總）</div>
      {qqq_flow_html}
    </div>
  </div>
</div>

<div class="ftr">
  ⚠️ 本報告數據來自 Yahoo Finance（yfinance），OI / Max Pain 為真實期權鏈計算，晨星估值為靜態手動更新。僅供參考，不構成投資建議。<br>
  數據來源：Yahoo Finance · yfinance　｜　報告產生：{now_tw} 台灣時間<br>
  <span style="color:#21262d">本機版 v5.0 · python3 stock_scan.py · 四週 OI · S5TW/S5FI · 機構流量偵測</span>
</div>

</div>
</body>
</html>'''


# ═══════════════════════════════════════════════════════════════
#  主程式
# ═══════════════════════════════════════════════════════════════

def main():
    print('=' * 50)
    print('美股每日掃描報告 — 本機版 v3.0')
    print('=' * 50)

    date_str = datetime.now().strftime('%Y%m%d')
    # GitHub Pages 用 index.html；同時保留日期版本
    output_path = os.path.join(OUTPUT_DIR, 'index.html')
    dated_path  = os.path.join(OUTPUT_DIR, f'stock_report_{date_str}.html')

    # 確認輸出資料夾存在
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print('\n[1/4] 抓取 VIX...')
    vix = fetch_vix()
    print(f'  VIX: {vix["value"]} ({vix["change"]:+.2f})')

    print('\n[2/4] 抓取個股數據（共 14 支）...')
    stocks = []
    for ticker in TICKERS:
        stocks.append(fetch_stock(ticker))

    ok_count = sum(1 for s in stocks if s.get('ok'))
    print(f'  完成：{ok_count}/{len(TICKERS)} 支成功')

    print('\n[3/4] 抓取 SPY / QQQ...')
    spy_qqq = fetch_spy_qqq()

    print('\n[4/4] 生成 HTML 報告...')
    html = generate_report(stocks, vix, spy_qqq, date_str)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    with open(dated_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f'\n✅ 報告已存至：{output_path}')
    print(f'   日期版本：{dated_path}')
    print('=' * 50)


if __name__ == '__main__':
    main()
