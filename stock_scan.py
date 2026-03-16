#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
美股每日開盤掃描報告 Powered by 黑叔 — v7.1
HMA/短EMA/大EMA 最優參數支撐壓力 | 期權流五法共識 | Buy/Sell Zone（回測驗證係數）
使用 yfinance 抓取真實股價、期權 OI、IV、P/C Ratio、HMA/EMA 數據
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

TICKERS = ['NVDA', 'GOOGL', 'AAPL', 'TSM', 'MSFT', 'AMZN', 'META', 'TSLA',
           'AMD', 'MU', 'AVGO', 'SNDK', 'NFLX', 'LITE', 'COHR', 'JPM']

MAG7 = {'AAPL', 'MSFT', 'NVDA', 'AMZN', 'GOOGL', 'META', 'TSLA'}

# 每股建議觀察重點（顯示在股價欄下方）
STOCK_TIPS = {
    'TSLA':  ('gex',      '🔆 GEX',       '觀察做市商 Gamma 方向，Short γ 時波動放大'),
    'AAPL':  ('maxpain',  '⚡ Max Pain',   '價格磁吸力強，結算易回歸 Max Pain'),
    'AMD':   ('oi',       '📊 OI',         '關注 Put/Call Wall OI 集中點，決定支撐壓力'),
    'MSFT':  ('sellzone', '📤 Sell Zone',  '回測驗證 Sell Zone 為有效盤中出場點'),
    'GOOGL': ('wall',     '🧱 C/P Wall',   '關注 Call Wall 壓力與 Put Wall 支撐位'),
    'NFLX':  ('oi',       '📊 OI',         '流動性集中，OI Wall 為盤中關鍵翻轉點'),
    'AVGO':  ('sellzone', '📤 Sell Zone',  '多頭格局下 Sell Zone 為理想短線獲利出場'),
    'NVDA':  ('buyzone',  '📥 Buy Zone',   '回測勝率高，5 批分買策略有效性最強'),
    'TSM':   ('putwall',  '🛡 Put Wall',   'Put Wall OI 厚實，為短線重要下方支撐'),
    'META':  ('maxpain',  '⚡ Max Pain',   '結算收盤常回歸 Max Pain，週五磁吸明顯'),
    'AMZN':  ('ema',      '📐 EMA',        'EMA15/50 為趨勢判斷關鍵，注意均線交叉'),
    'MU':    ('iv',       '📈 IV/IVR',     'IV 波動大，IVR 高時賣方策略優先'),
    'SNDK':  ('oi',       '📊 OI',         'OI 集中點明確，可做為盤中方向參考'),
    'LITE':  ('ema',      '📐 EMA',        '趨勢跟蹤為主，EMA 排列決定多空方向'),
    'COHR':  ('ema',      '📐 EMA',        '中低流動性，EMA 為最可靠技術參考'),
    'JPM':   ('wall',     '🧱 C/P Wall',   '金融股期權 Wall 效果顯著，注意財報周風險'),
}

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
    'JPM':   {'fv': 220,  'stars': 3, 'moat': 'Wide'},
}

OUTPUT_DIR = os.environ.get('OUTPUT_DIR', os.path.expanduser('~/Desktop'))


# ═══════════════════════════════════════════════════════════════
#  HMA / EMA 最優參數表（52週 × 19股回測研究 2026-03-15）
#  (hma週期, 短ema週期, 大ema週期, hma支撐5d%, hma壓力5d%, 評級, 假穿越%)
# ═══════════════════════════════════════════════════════════════

DEFAULT_HMA_P, DEFAULT_EMA_P, DEFAULT_BIG_EMA_P = 18, 13, 57

EMA_HMA_CFG = {
    'SPY':  (24,  6, 57,  83.3, 87.5, '雙向強',   47.9),
    'QQQ':  (28,  6, 57,  89.5, 63.2, '支撐強',   54.8),
    'AAPL': (19, 14, 69,  76.5, 83.3, '雙向強',   28.8),
    'MSFT': (11, 17, 59,  78.8, 50.0, '支撐偏強', 59.9),
    'NVDA': (22,  8, 64,  76.0, 48.0, '支撐偏強', 54.0),
    'AMZN': (12, 19, 51,  79.3, 62.1, '雙向強',   50.7),
    'GOOGL':(14, 19, 66,  85.0, 42.1, '支撐偏強', 34.4),
    'META': (21, 16, 58,  70.6, 68.8, '雙向強',   41.2),
    'TSLA': (15,  5, 58,  53.3, 43.3, '雙向弱',   61.3),
    'AMD':  (19, 19, 53,  73.3, 56.2, '支撐偏強', 41.2),
    'TSM':  (11, 19, 59,  61.8, 48.6, '普通',     58.0),
    'MU':   (29, 13, 61,  72.2, 36.8, '支撐偏強', 56.4),
    'AVGO': (11, 19, 51,  65.7, 34.3, '壓力弱',   60.6),
    'SNDK': (12, 15, 51,  63.0, 25.9, '壓力弱',   58.4),
    'NFLX': (12, 11, 61,  57.7, 60.0, '壓力偏強', 44.2),
    'LITE': (28,  7, None,70.6, 23.5, '壓力弱',   42.6),
    'COHR': (25, 19, 58,  66.7, 38.1, '支撐偏強', 46.3),
    'JPM':  (28, 16, 53,  80.0, 62.5, '雙向強',   41.7),
}

# 多頭逆轉勝率（突破大EMA後5日上漲概率）
BULL_WIN = {
    'AVGO':71,'QQQ':71,'SPY':57,'NFLX':50,'WDC':50,'COHR':33,
    'AAPL':45,'NVDA':48,'META':52,'AMZN':55,'MSFT':40,'AMD':42,
    'MU':50,'TSM':45,'TSLA':38,'GOOGL':55,'JPM':52,'SNDK':40,'LITE':40,
}
# 空頭逆轉勝率（跌破大EMA後持續下跌概率）
BEAR_WIN = {
    'MSFT':50,'QQQ':40,'AAPL':38,'MU':36,'NFLX':31,'SPY':30,
    'NVDA':28,'META':32,'AMZN':28,'TSLA':35,'AMD':30,'AVGO':25,
    'WDC':22,'COHR':30,'TSM':28,'GOOGL':25,'JPM':30,'SNDK':25,'LITE':20,
}
SUPPORT_ONLY = {'AVGO','WDC','LITE','SNDK'}
WEAK_BOTH    = {'TSLA','TSM'}


def _hma(series, period):
    half  = max(int(period / 2), 2)
    sqrtn = max(int(period ** 0.5), 2)
    wma_h = series.ewm(span=half,   adjust=False).mean()
    wma_n = series.ewm(span=period, adjust=False).mean()
    return (2 * wma_h - wma_n).ewm(span=sqrtn, adjust=False).mean()


def _ema(series, period):
    return series.ewm(span=period, adjust=False).mean()


def calc_hma_ema_data(ticker, closes):
    """
    計算個股 HMA / 短EMA / 大EMA 當前值與角色，回傳 dict。
    closes: pd.Series 收盤價（至少 120 個交易日）
    """
    cfg = EMA_HMA_CFG.get(ticker.upper())
    if cfg:
        hma_p, ema_p, big_p, sup5d, res5d, rating, fake_pct = cfg
    else:
        hma_p, ema_p, big_p = DEFAULT_HMA_P, DEFAULT_EMA_P, DEFAULT_BIG_EMA_P
        sup5d, res5d, rating, fake_pct = 65.0, 47.0, '預設', 50.0

    big_p = big_p or DEFAULT_BIG_EMA_P
    bull_win = BULL_WIN.get(ticker.upper(), 48)
    bear_win = BEAR_WIN.get(ticker.upper(), 30)

    try:
        n = len(closes)
        if n < max(hma_p, ema_p, big_p) + 5:
            return None

        S         = float(closes.iloc[-1])
        prev      = float(closes.iloc[-2])
        hma_s     = _hma(closes, hma_p)
        ema_s     = _ema(closes, ema_p)
        big_s     = _ema(closes, big_p)

        hma_val   = float(hma_s.iloc[-1])
        ema_val   = float(ema_s.iloc[-1])
        big_val   = float(big_s.iloc[-1])
        prev_hma  = float(hma_s.iloc[-2])
        prev_ema  = float(ema_s.iloc[-2])
        prev_big  = float(big_s.iloc[-2])

        # 穿越偵測
        hma_cross_up   = prev < prev_hma and S > hma_val
        hma_cross_down = prev > prev_hma and S < hma_val
        ema_cross_up   = prev < prev_ema and S > ema_val
        ema_cross_down = prev > prev_ema and S < ema_val
        big_cross_up   = prev < prev_big and S > big_val
        big_cross_down = prev > prev_big and S < big_val

        above_hma = S > hma_val
        above_ema = S > ema_val
        above_big = S > big_val

        # 動態角色：在上=支撐，在下=壓力
        def _role(above, cross_up, cross_down):
            if cross_up:   return 'just_up'
            if cross_down: return 'just_down'
            return 'support' if above else 'resist'

        hma_role = _role(above_hma, hma_cross_up, hma_cross_down)
        ema_role = _role(above_ema, ema_cross_up, ema_cross_down)
        big_role = _role(above_big, big_cross_up, big_cross_down)

        # 盤整五條件評分
        atr_raw = (closes.diff().abs().rolling(14).mean()).iloc[-1]
        atr = max(atr_raw if not math.isnan(atr_raw) else S*0.02, S*0.015)
        c1 = abs(ema_val - hma_val) / atr < 0.8
        c2 = abs(float(hma_s.iloc[-1]) - float(hma_s.iloc[-6])) / atr < 0.3 if n >= 7 else False
        c3 = abs(S - big_val) / atr < 2.0
        c4 = False
        ema200_val = None
        above_ema200 = None
        if n >= 200:
            ema200_val = round(float(_ema(closes, 200).iloc[-1]), 2)
            above_ema200 = S > ema200_val
            c4 = abs(S - ema200_val) / atr < 3.0
        c5 = abs(ema_val - big_val) / atr < 1.0
        consol = sum([c1, c2, c3, c4, c5])
        is_consol = consol >= 3

        # 中期趨勢
        both_above = above_hma and above_ema and above_big
        both_below = not above_hma and not above_ema and not above_big
        if big_cross_up:
            mid = f'📈剛突破大EMA({bull_win}%上漲)'
        elif big_cross_down:
            mid = f'📉剛跌破大EMA({bear_win}%下跌)'
        elif is_consol:
            mid = f'⚪盤整({consol}/5) 結束後70%向上'
        elif both_above:
            mid = f'📈中期上行({bull_win}%)'
        elif both_below:
            mid = f'📉中期下行({bear_win}%)'
        else:
            mid = '🟡方向分歧，等兩線同步'

        return {
            'hma_p': hma_p, 'ema_p': ema_p, 'big_p': big_p,
            'hma_val': round(hma_val, 2),
            'ema_val': round(ema_val, 2),
            'big_val': round(big_val, 2),
            'ema200_val': ema200_val,
            'above_ema200': above_ema200,
            'hma_role': hma_role, 'ema_role': ema_role, 'big_role': big_role,
            'sup5d': sup5d, 'res5d': res5d,
            'rating': rating, 'fake_pct': fake_pct,
            'bull_win': bull_win, 'bear_win': bear_win,
            'consol': consol, 'is_consol': is_consol,
            'mid': mid,
            'support_only': ticker.upper() in SUPPORT_ONLY,
            'weak_both': ticker.upper() in WEAK_BOTH,
            'hma_cross_up': hma_cross_up, 'hma_cross_down': hma_cross_down,
            'ema_cross_up': ema_cross_up, 'ema_cross_down': ema_cross_down,
            'big_cross_up': big_cross_up, 'big_cross_down': big_cross_down,
        }
    except Exception:
        return None


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
#  EMA 計算 + 訊號引擎（SPY/QQQ 大盤用）
# ═══════════════════════════════════════════════════════════════

def _compute_emas(closes):
    """計算 EMA15/50/100 及前一日值，供 SPY/QQQ 大盤分析用"""
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
        # ── HMA / EMA 最優技術面計算 ──────────────────────────
        d['hma_ema'] = calc_hma_ema_data(ticker, closes)

        WEEK_KEYS = ['w0','w1','w2','w3']
        for k in WEEK_KEYS:
            d.update({f'{k}_expiry':None, f'{k}_put_wall':None, f'{k}_call_wall':None,
                      f'{k}_max_pain':None, f'{k}_gex':None,
                      f'{k}_is_monthly':False, f'{k}_label':'',
                      f'{k}_sell_hi':None, f'{k}_sell_lo':None,
                      f'{k}_buy_hi':None,  f'{k}_buy_lo':None,
                      f'{k}_settle_lo':None, f'{k}_settle_hi':None,
                      f'{k}_consensus':None})
        d.update({'iv':None,'ivr':None,'pc_ratio':None})

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

                    # ── Buy/Sell Zone v4.1（回測修正：排除CallWall + 硬下限）─
                    try:
                        mp = res.get('max_pain') or price
                        atm_idx = (cdf['strike'] - price).abs().argsort()[:3]
                        atm_c   = cdf.iloc[atm_idx]
                        atm_p   = pdf.iloc[(pdf['strike'] - price).abs().argsort()[:3]]
                        c_mid   = pd.to_numeric(atm_c['lastPrice'], errors='coerce').mean()
                        p_mid   = pd.to_numeric(atm_p['lastPrice'], errors='coerce').mean()
                        straddle = (c_mid + p_mid) * 0.85 if (c_mid > 0 and p_mid > 0) else price * 0.04
                        cs = straddle * 0.55  # conservative straddle
                        # GEX 最大阻力/支撐（排除 Call Wall）
                        try:
                            gex_map = {}
                            for _, row in cdf.iterrows():
                                k2 = float(row.get('strike', 0))
                                gex_map[k2] = gex_map.get(k2, 0) + float(row.get('openInterest', 0) or 0)
                            for _, row in pdf.iterrows():
                                k2 = float(row.get('strike', 0))
                                gex_map[k2] = gex_map.get(k2, 0) - float(row.get('openInterest', 0) or 0)
                            resist_2 = max(gex_map, key=gex_map.get) if gex_map else None
                            support_2 = min(gex_map, key=gex_map.get) if gex_map else None
                        except Exception:
                            resist_2 = None; support_2 = None
                        # Sell Zone — 排除CallWall，硬下限 S×1.005
                        sell_lo = round(mp * 1.005, 2) if mp else round(price * 1.005, 2)
                        cands_s = [v for v in [resist_2, price + cs] if v is not None]
                        sell_hi = round(min(cands_s), 2) if cands_s else round(price * 1.03, 2)
                        sell_hi = max(sell_hi, round(price * 1.005, 2))
                        if sell_hi <= sell_lo: sell_hi = round(sell_lo * 1.015, 2)
                        # Buy Zone
                        buy_hi = round(mp * 0.995, 2) if mp else round(price * 0.995, 2)
                        cands_b = [v for v in [support_2, price - cs] if v is not None]
                        buy_lo  = round(max(cands_b), 2) if cands_b else round(price * 0.97, 2)
                        buy_lo  = min(buy_lo, round(price * 0.995, 2))
                        if buy_lo >= buy_hi: buy_lo = round(buy_hi * 0.985, 2)
                        d[f'{k}_sell_lo'] = sell_lo
                        d[f'{k}_sell_hi'] = sell_hi
                        d[f'{k}_buy_lo']  = buy_lo
                        d[f'{k}_buy_hi']  = buy_hi
                        # 週五結算磁吸區 = Max Pain ± 1%
                        d[f'{k}_settle_lo'] = round(mp * 0.99, 2)
                        d[f'{k}_settle_hi'] = round(mp * 1.01, 2)
                        d[f'{k}_max_pain']  = res.get('max_pain')
                        # 五法共識 v4.1（回測最優：平均誤差1.15%）= (MaxPain + S+straddle + CallWall) / 3
                        cw_val = res.get('call_wall')
                        con_vals = [v for v in [mp, price + straddle, cw_val] if v is not None]
                        d[f'{k}_consensus'] = round(sum(con_vals) / len(con_vals), 2) if con_vals else None
                    except Exception as ez:
                        print(f'    Zone計算錯誤[{i}]: {ez}')
                    mo = '★月結' if wk['is_monthly'] else ''
                    print(f'    {wk["label"]}{mo} ({d[f"{k}_expiry"]})  '
                          f'Put牆:{d[f"{k}_put_wall"]}  Pain:{d[f"{k}_max_pain"]}  Call牆:{d[f"{k}_call_wall"]}')
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
            # ── HMA/EMA 最優參數計算 ──────────────────────────
            entry['hma_ema'] = calc_hma_ema_data(sym, closes)
            entry['ticker']  = sym

            for k in WEEK_KEYS:
                entry.update({f'{k}_expiry':None, f'{k}_put_wall':None,
                               f'{k}_call_wall':None, f'{k}_max_pain':None,
                               f'{k}_is_monthly':False, f'{k}_label':'',
                               f'{k}_sell_hi':None, f'{k}_sell_lo':None,
                               f'{k}_buy_hi':None,  f'{k}_buy_lo':None})
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
                    res,cdf,pdf=_fetch_etf(wk['date'])
                    entry[f'{k}_expiry']=res.get('expiry')
                    entry[f'{k}_put_wall']=res.get('put_wall')
                    entry[f'{k}_call_wall']=res.get('call_wall')
                    entry[f'{k}_max_pain']=res.get('max_pain')
                    mo='★月結' if wk['is_monthly'] else ''
                    print(f'    {sym} {wk["label"]}{mo} ({entry[f"{k}_expiry"]}) Put:{entry[f"{k}_put_wall"]} Pain:{entry[f"{k}_max_pain"]} Call:{entry[f"{k}_call_wall"]}')
                    # Buy/Sell Zone v4.1（排除CallWall + 硬下限）
                    try:
                        mp = res.get('max_pain') or price
                        atm_c = cdf.iloc[(cdf['strike']-price).abs().argsort()[:3]]
                        atm_p = pdf.iloc[(pdf['strike']-price).abs().argsort()[:3]]
                        c_mid = pd.to_numeric(atm_c['lastPrice'],errors='coerce').mean()
                        p_mid = pd.to_numeric(atm_p['lastPrice'],errors='coerce').mean()
                        st = (c_mid+p_mid)*0.85 if (c_mid>0 and p_mid>0) else price*0.03
                        cs = st * 0.55
                        try:
                            gm = {}
                            for _, row in cdf.iterrows():
                                k2=float(row.get('strike',0)); gm[k2]=gm.get(k2,0)+float(row.get('openInterest',0) or 0)
                            for _, row in pdf.iterrows():
                                k2=float(row.get('strike',0)); gm[k2]=gm.get(k2,0)-float(row.get('openInterest',0) or 0)
                            r2=max(gm,key=gm.get) if gm else None
                            s2=min(gm,key=gm.get) if gm else None
                        except Exception: r2=None; s2=None
                        sell_lo=round(mp*1.005,2) if mp else round(price*1.005,2)
                        cands_s=[v for v in [r2, price+cs] if v is not None]
                        sell_hi=round(min(cands_s),2) if cands_s else round(price*1.03,2)
                        sell_hi=max(sell_hi, round(price*1.005,2))
                        if sell_hi<=sell_lo: sell_hi=round(sell_lo*1.015,2)
                        buy_hi=round(mp*0.995,2) if mp else round(price*0.995,2)
                        cands_b=[v for v in [s2, price-cs] if v is not None]
                        buy_lo=round(max(cands_b),2) if cands_b else round(price*0.97,2)
                        buy_lo=min(buy_lo, round(price*0.995,2))
                        if buy_lo>=buy_hi: buy_lo=round(buy_hi*0.985,2)
                        entry[f'{k}_sell_lo']=sell_lo; entry[f'{k}_sell_hi']=sell_hi
                        entry[f'{k}_buy_lo']=buy_lo;   entry[f'{k}_buy_hi']=buy_hi
                        entry[f'{k}_settle_lo']=round(mp*0.99,2)
                        entry[f'{k}_settle_hi']=round(mp*1.01,2)
                        # 五法共識 v4.1 = (MaxPain + price+straddle + CallWall) / 3
                        cw_etf = res.get('call_wall')
                        con_etf = [v for v in [mp, price+st, cw_etf] if v is not None]
                        entry[f'{k}_consensus'] = round(sum(con_etf)/len(con_etf), 2) if con_etf else None
                    except: pass
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


def hma_ema_cell_html(d):
    """
    個股 HMA/短EMA/大EMA/EMA200 技術面欄
    每條線：第一行 標籤+$價格（不換行）
            第二行 角色+勝率%+漲跌%+回測建議（nowrap）
    """
    he = d.get('hma_ema')
    if not he:
        return '<div style="color:#484f58;font-size:0.75em">HMA/EMA 計算中…</div>'

    price   = d.get('price', 0)
    ticker  = (d.get('ticker') or '').upper()

    ROLES = {
        'just_up':   ('#3fb950', '🔔↑', '剛突破→支撐'),
        'just_down': ('#f85149', '🔔↓', '剛跌破→壓力'),
        'support':   ('#3fb950', '●',   '支撐'),
        'resist':    ('#f85149', '●',   '壓力'),
    }

    # ── 回測研究建議邏輯 ──────────────────────────────────────
    # 三線同步在大EMA上方/下方
    above_big = he.get('big_role') in ('support', 'just_up')
    above_hma = he.get('hma_role') in ('support', 'just_up')
    above_ema = he.get('ema_role') in ('support', 'just_up')
    triple_above = above_big and above_hma and above_ema   # 三線都在大EMA上 → 強多
    triple_below = not above_big and not above_hma and not above_ema  # 三線同步跌破

    # 個股特殊規則（來自記憶庫）
    BUY_ON_DROP  = {'AVGO','COHR','TSM','JPM','MU','SNDK','AMD','WDC'}  # 跌破大EMA=買點
    BEAR_VALID   = {'NFLX','MSFT','META'}       # 空頭逆轉有效股
    NO_SELL_BELOW = {'AVGO','WDC','LITE','SNDK'} # 壓力線失效，只做支撐

    def _adv(role_key, line_type='hma'):
        """根據角色+線型+個股特性生成建議"""
        is_up   = role_key in ('support', 'just_up')
        is_down = role_key in ('resist',  'just_down')
        cross_u = role_key == 'just_up'
        cross_d = role_key == 'just_down'

        if line_type == 'hma':
            if cross_u and triple_above:
                return '三線多頭，10日勝率68%'
            if cross_u:
                return 'HMA突破，趨勢型信號'
            if cross_d and triple_below:
                if ticker in BUY_ON_DROP:
                    return '三線跌破，此股逆向買'
                if ticker in BEAR_VALID:
                    return '三線同步跌破，空頭確認'
                return '假跌破概率高，等3日確認'
            if cross_d:
                return 'HMA跌破，等3日確認方向'
            if is_up and triple_above:
                return '三線多頭，守HMA持多'
            if is_up:
                return '守HMA支撐，偏多操作'
            if is_down:
                return '反彈至此可減碼'

        elif line_type == 'ema':
            if cross_u:
                return '短EMA突破=短暫反彈'
            if cross_d:
                return 'EMA跌破，短線轉弱'
            if is_up:
                return 'EMA支撐中，動能偏多'
            if is_down:
                if triple_below:
                    return '三線空排，EMA壓制'
                return '短EMA壓力，等突破'

        elif line_type == 'big':
            if cross_u:
                if ticker in BUY_ON_DROP:
                    return f'突破大EMA，多頭逆轉'
                return '突破大EMA，趨勢轉多'
            if cross_d:
                if ticker in BUY_ON_DROP:
                    return f'{ticker}跌破=逆向做多'
                if ticker in BEAR_VALID:
                    return '跌破大EMA，持空'
                return '跌破大EMA，多反彈'
            if is_up:
                return '站上大EMA，趨勢確認'
            if is_down:
                return '大EMA壓制，反彈減碼'

        return ''

    def two_line(tag, val, role_key, win_pct, line_type='hma'):
        if val is None:
            return ''
        col, icon, role_txt = ROLES.get(role_key, ('#8b949e','●','觀察'))
        pct      = (price - val) / val * 100 if val else 0
        pct_col  = '#3fb950' if pct >= 0 else '#f85149'
        pct_sign = '+' if pct >= 0 else ''
        adv      = _adv(role_key, line_type)
        return (
            # 行1：標籤 + 價格緊貼（貼左）
            f'<div style="display:flex;align-items:baseline;gap:6px;'
            f'margin-top:6px;white-space:nowrap">'
            f'<span style="font-size:0.65em;color:#8b949e;min-width:54px">{tag}</span>'
            f'<span style="font-family:monospace;font-size:0.95em;font-weight:700;'
            f'color:{col}">${val:.2f}</span>'
            f'</div>'
            # 行2：角色+勝率+漲跌+建議（nowrap）
            f'<div style="display:flex;justify-content:space-between;align-items:center;'
            f'padding-bottom:5px;border-bottom:1px solid #21262d;white-space:nowrap">'
            f'<span style="font-size:0.67em;color:{col};font-weight:600">'
            f'{icon} {role_txt} {win_pct:.0f}% '
            f'<span style="color:{pct_col}">{pct_sign}{pct:.1f}%</span></span>'
            f'<span style="font-size:0.6em;color:#6e7681;margin-left:6px">{adv}</span>'
            f'</div>'
        )

    hma_win = he['sup5d'] if he['hma_role'] in ('support','just_up') else he['res5d']
    ema_win = he['sup5d']*0.95 if he['ema_role'] in ('support','just_up') else he['res5d']*1.02
    big_win = he['bull_win'] if he['big_role'] in ('support','just_up') else he['bear_win']

    hma_row = two_line(f'HMA{he["hma_p"]}',   he['hma_val'], he['hma_role'], hma_win, 'hma')
    ema_row = two_line(f'EMA{he["ema_p"]}',   he['ema_val'], he['ema_role'], ema_win, 'ema')
    big_row = two_line(f'大EMA{he["big_p"]}', he['big_val'], he['big_role'], big_win, 'big')

    # EMA200 行
    ema200_row = ''
    if he.get('ema200_val') is not None:
        v200       = he['ema200_val']
        ab         = he.get('above_ema200', True)
        col200     = '#3fb950' if ab else '#f85149'
        pct200     = (price - v200) / v200 * 100
        pct200_col = '#3fb950' if pct200 >= 0 else '#f85149'
        # EMA200 個股特殊建議
        if ab:
            if ticker in ('JPM','TSLA','AVGO','AAPL','COHR'):
                adv200 = '長期多頭，逆轉分高持多'
            else:
                adv200 = '長期多頭結構，持多'
        else:
            if ticker in BUY_ON_DROP:
                adv200 = f'{ticker}跌破EMA200=買點'
            elif ticker in ('NFLX',):
                adv200 = 'NFLX空頭最強，可持空'
            else:
                adv200 = '跌破長期均線，謹慎'
        status = '● 長期支撐' if ab else '● 長期壓力'
        ema200_row = (
            f'<div style="display:flex;align-items:baseline;gap:6px;'
            f'margin-top:6px;white-space:nowrap">'
            f'<span style="font-size:0.65em;color:#8b949e;min-width:54px">EMA200</span>'
            f'<span style="font-family:monospace;font-size:0.95em;font-weight:700;'
            f'color:{col200}">${v200:.2f}</span>'
            f'</div>'
            f'<div style="display:flex;justify-content:space-between;align-items:center;'
            f'padding-bottom:5px;border-bottom:1px solid #21262d;white-space:nowrap">'
            f'<span style="font-size:0.67em;color:{col200};font-weight:600">{status}'
            f'<span style="color:{pct200_col};margin-left:5px">{pct200:+.1f}%</span></span>'
            f'<span style="font-size:0.6em;color:#6e7681;margin-left:6px">{adv200}</span>'
            f'</div>'
        )

    # 警告
    warn = ''
    if he.get('support_only'):
        warn = ('<div style="font-size:0.62em;color:#e3b341;margin-bottom:3px;padding:2px 5px;'
                'background:rgba(227,179,65,.08);border-radius:3px">⚠ 壓力線失效，只看支撐</div>')
    elif he.get('weak_both'):
        warn = ('<div style="font-size:0.62em;color:#e3b341;margin-bottom:3px;padding:2px 5px;'
                'background:rgba(227,179,65,.08);border-radius:3px">⚠ 雙向弱，需3日確認</div>')

    # 評級行 — 評級 + 假穿越 同行顯示
    rating_col = '#3fb950' if '強' in he['rating'] else '#e3b341' if '偏' in he['rating'] else '#8b949e'
    header = (
        f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:2px;flex-wrap:wrap">'
        f'<span style="font-size:0.62em;color:{rating_col};font-weight:600">{he["rating"]}</span>'
        f'<span style="font-size:0.6em;color:#484f58;background:#1c2128;border-radius:3px;'
        f'padding:1px 5px">假穿越 {he["fake_pct"]:.0f}%</span>'
        f'</div>'
    )

    # ── 均線做多/做空評分（簡潔版）──────────────────────────
    above_big = he.get('big_role') in ('support', 'just_up')
    above_hma = he.get('hma_role') in ('support', 'just_up')
    above_ema = he.get('ema_role') in ('support', 'just_up')
    hma_above_big = he.get('hma_val', 0) > he.get('big_val', 0)
    ema_above_big = he.get('ema_val', 0) > he.get('big_val', 0)
    above_ema200  = he.get('above_ema200')
    big_cross_up  = he.get('big_cross_up', False)

    BEAR_BANNED_S = {'MU','SNDK','AVGO','TSM','JPM','AMD'}
    BEAR_VALID_S  = {'META','MSFT','NFLX'}

    # 做多分（自動可算6分，下影線+2需人工）
    ls = sum([
        1 if above_ema else 0,
        1 if (he.get('ema_val',0) > he.get('hma_val',0)) else 0,
        1 if (he.get('hma_cross_up') or (hma_above_big and ema_above_big)) else 0,
        1 if (above_big and hma_above_big) else 0,
        1 if (above_ema200 is True) else 0,
        1 if big_cross_up else 0,
    ])
    if ls >= 5:   long_sig, long_col = '🟢 高信心做多', '#3fb950'
    elif ls >= 3: long_sig, long_col = '🟡 標準倉位',   '#e3b341'
    else:         long_sig, long_col = '⬜ 等待',        '#484f58'

    # 做空分（自動可算5分）
    short_banned = ticker in BEAR_BANNED_S
    triple_below = not above_big and not hma_above_big and not ema_above_big
    ss = sum([
        1 if not above_big else 0,
        1 if not hma_above_big else 0,
        1 if not ema_above_big else 0,
        1 if ticker in BEAR_VALID_S else 0,
        1 if (above_ema200 is False) else 0,
    ])
    if short_banned:
        short_sig, short_col = f'🚫 {ticker}跌破=買點', '#e3b341'
    elif ss >= 4:   short_sig, short_col = '🔴 高信心做空', '#f85149'
    elif ss >= 2:   short_sig, short_col = '🟡 輕倉觀察',   '#e3b341'
    else:           short_sig, short_col = '⬜ 做空條件弱', '#484f58'

    score_html = (
        f'<div style="margin-top:6px;padding:5px 7px;background:#1c2128;border-radius:4px;'
        f'border-left:2px solid #30363d">'
        f'<div style="font-size:0.6em;color:#484f58;margin-bottom:3px;letter-spacing:0.5px">均線評分</div>'
        f'<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">'
        f'<span style="font-size:0.72em;font-weight:700;color:{long_col}">{long_sig} {ls}/6</span>'
        f'<span style="font-size:0.62em;color:#484f58">｜</span>'
        f'<span style="font-size:0.72em;font-weight:700;color:{short_col}">{short_sig} {ss}/5</span>'
        f'</div>'
        f'<div style="font-size:0.6em;color:#484f58;margin-top:2px">+下影線可+2做多分</div>'
        f'</div>'
    )

    # 中期趨勢
    mid_col = ('#3fb950' if '上行' in he['mid'] or '剛突破' in he['mid']
               else '#f85149' if '下行' in he['mid'] or '剛跌破' in he['mid']
               else '#e3b341')
    mid_html = (
        f'<div style="margin-top:6px;padding:4px 7px;background:#1c2128;border-radius:4px">'
        f'<span style="font-size:0.62em;color:#484f58">中期趨勢　</span>'
        f'<span style="font-size:0.75em;font-weight:700;color:{mid_col}">{he["mid"]}</span>'
        f'</div>'
    )

    return f'{warn}{header}{hma_row}{ema_row}{big_row}{ema200_row}{score_html}{mid_html}'


def etf_ema_block_html(etf_data):
    """SPY/QQQ 卡片：HMA/EMA 最優四層分析（復用 hma_ema_cell_html）"""
    # 直接沿用個股格的渲染邏輯，etf_data 結構相同
    return (
        f'<div class="etf-ema-blk">'
        f'<div class="oi-lbl" style="margin-bottom:7px;letter-spacing:0.6px">📐 HMA/EMA 均線分析</div>'
        f'{hma_ema_cell_html(etf_data)}'
        f'</div>'
    )


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
        fwd = f'<div class="fpe-dn" style="color:{col}">{pe_fwd}x</div><div class="pe-note pos">↑ 獲利成長</div>'
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


def consensus_zone_html(d, pfx='w0'):
    """第二格：五法共識 + Sell/Buy Zone + 週五結算磁吸區"""
    consensus = d.get(f'{pfx}_consensus')
    sh = d.get(f'{pfx}_sell_hi')
    sl = d.get(f'{pfx}_sell_lo')
    bh = d.get(f'{pfx}_buy_hi')
    bl = d.get(f'{pfx}_buy_lo')
    sth = d.get(f'{pfx}_settle_hi')
    stl = d.get(f'{pfx}_settle_lo')
    mp  = d.get(f'{pfx}_max_pain')
    if sh is None and bh is None:
        return '<div style="color:#484f58;font-size:0.72em">Zone：無數據</div>'
    fmt = lambda v: f'${v:g}' if v is not None else '—'
    price = d.get('price', 0)
    # 觸及提醒
    alert = ''
    if bh and price <= bh:
        alert = '<div style="font-size:0.65em;color:#f78166;background:rgba(247,129,102,.1);border:1px solid rgba(247,129,102,.3);border-radius:3px;padding:1px 5px;margin-bottom:3px">🔴 分批買入區</div>'
    elif sl and price >= sl:
        alert = '<div style="font-size:0.65em;color:#56d364;background:rgba(86,211,100,.1);border:1px solid rgba(86,211,100,.3);border-radius:3px;padding:1px 5px;margin-bottom:3px">🟢 分批賣出區</div>'
    return (f'{alert}'
            f'<div style="font-size:0.65em;color:#484f58;margin-bottom:2px">五法共識 {fmt(consensus)}</div>'
            f'<div style="font-family:monospace;font-size:0.82em;color:#56d364;font-weight:600">🟢 Sell {fmt(sl)}~{fmt(sh)}</div>'
            f'<div style="font-family:monospace;font-size:0.82em;color:#f78166;font-weight:600;margin-top:2px">🔴 Buy {fmt(bl)}~{fmt(bh)}</div>'
            f'<div style="font-family:monospace;font-size:0.75em;color:#e3b341;margin-top:3px">📌結算 {fmt(stl)}~{fmt(sth)}</div>')


def oi_html_single(d, pfx, label=''):
    """
    單欄期權流：
    ① OI 結構（Call Wall / Max Pain / Put Wall）
    ② Buy/Sell Zone + 結算磁吸
    ③ GEX 流入流出
    """
    cw    = d.get(f'{pfx}_call_wall')
    mp    = d.get(f'{pfx}_max_pain')
    pw    = d.get(f'{pfx}_put_wall')
    gex   = d.get(f'{pfx}_gex')
    sh    = d.get(f'{pfx}_sell_hi')
    sl    = d.get(f'{pfx}_sell_lo')
    bh    = d.get(f'{pfx}_buy_hi')
    bl    = d.get(f'{pfx}_buy_lo')
    sth   = d.get(f'{pfx}_settle_hi')
    stl   = d.get(f'{pfx}_settle_lo')
    con   = d.get(f'{pfx}_consensus')
    exp   = d.get(f'{pfx}_expiry','') or ''
    is_mo = d.get(f'{pfx}_is_monthly', False)
    lbl   = d.get(f'{pfx}_label', label)
    exps  = exp[5:] if exp else lbl
    price = d.get('price', 0)
    fmt   = lambda v: f'${v:g}' if v is not None else '—'
    fmtf  = lambda v: f'${v:.2f}' if v is not None else '—'
    mob   = ('<span style="font-size:0.6em;color:#e3b341;margin-left:3px;'
             'background:rgba(227,179,65,.12);border:1px solid rgba(227,179,65,.3);'
             'border-radius:3px;padding:0 3px">月結</span>' if is_mo else '')

    # ── ① OI 結構 ──────────────────────────────────────
    if cw is None and mp is None and pw is None:
        oi_block = f'<div style="color:#484f58;font-size:0.72em">{lbl}{mob} N/A</div>'
    else:
        oi_block = (
            f'<div class="oi-lbl">{lbl} {exps}{mob}</div>'
            f'<div class="oi-call">▼ Call {fmt(cw)}</div>'
            f'<div class="oi-pain">⚡ MaxPain {fmt(mp)}</div>'
            f'<div class="oi-put">▲ Put {fmt(pw)}</div>'
        )

    # ── ② Zone + 結算磁吸 ──────────────────────────────
    zone_block = ''
    if sh is not None:
        # 觸及提醒
        alert = ''
        if bh and price <= bh:
            alert = ('<div style="font-size:0.62em;color:#f78166;background:rgba(247,129,102,.1);'
                     'border:1px solid rgba(247,129,102,.3);border-radius:3px;'
                     'padding:1px 5px;margin-bottom:2px">🔴 買入區</div>')
        elif sl and price >= sl:
            alert = ('<div style="font-size:0.62em;color:#56d364;background:rgba(86,211,100,.1);'
                     'border:1px solid rgba(86,211,100,.3);border-radius:3px;'
                     'padding:1px 5px;margin-bottom:2px">🟢 賣出區</div>')
        zone_block = (
            f'<div style="margin-top:5px;padding-top:4px;border-top:1px dashed #30363d">'
            f'{alert}'
            f'<div style="font-size:0.82em;color:#58a6ff;font-weight:600;margin-bottom:2px">五法共識 {fmtf(con)}</div>'
            f'<div style="font-family:monospace;font-size:0.8em;color:#56d364;font-weight:600">🟢 {fmtf(sl)}~{fmtf(sh)}</div>'
            f'<div style="font-family:monospace;font-size:0.8em;color:#f78166;font-weight:600;margin-top:1px">🔴 {fmtf(bl)}~{fmtf(bh)}</div>'
            f'<div style="font-family:monospace;font-size:0.72em;color:#e3b341;margin-top:2px">📌 {fmtf(stl)}~{fmtf(sth)}</div>'
            f'</div>'
        )

    # ── ③ GEX ──────────────────────────────────────────
    gex_block = ''
    if gex is not None:
        col  = '#3fb950' if gex >= 0 else '#f85149'
        arr  = '▲' if gex >= 0 else '▼'
        lbl2 = 'Long γ 流入' if gex >= 0 else 'Short γ 流出'
        bp   = min(int(abs(gex)/5000*100), 100)
        gex_block = (
            f'<div style="margin-top:5px;padding-top:4px;border-top:1px dashed #30363d">'
            f'<div style="font-size:0.62em;color:#484f58;margin-bottom:2px">GEX</div>'
            f'<div style="font-family:monospace;font-size:0.82em;font-weight:700;color:{col}">'
            f'{arr} ${abs(gex):,.0f}M</div>'
            f'<div style="font-size:0.62em;color:{col};margin-bottom:3px">{lbl2}</div>'
            f'<div style="background:#21262d;border-radius:2px;height:4px;width:55px">'
            f'<div style="background:{col};height:4px;border-radius:2px;width:{bp}%"></div></div>'
            f'</div>'
        )

    return f'{oi_block}{zone_block}{gex_block}'


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


def zone_html(d, pfx):
    """渲染單一週期的 Buy Zone / Sell Zone"""
    sh = d.get(f'{pfx}_sell_hi')
    sl = d.get(f'{pfx}_sell_lo')
    bh = d.get(f'{pfx}_buy_hi')
    bl = d.get(f'{pfx}_buy_lo')
    if sh is None and bh is None:
        return ''
    fmt = lambda v: f'${v:g}' if v is not None else '—'
    return (f'<div class="zone-wrap">'
            f'<div class="zone-sell">🟢 Sell {fmt(sl)}~{fmt(sh)}</div>'
            f'<div class="zone-buy">🔴 Buy {fmt(bl)}~{fmt(bh)}</div>'
            f'</div>')


def stock_tip_html(ticker):
    """渲染個股建議觀察提示"""
    tip = STOCK_TIPS.get(ticker)
    if not tip:
        return ''
    _, icon_label, desc = tip
    return (f'<div class="stk-tip">'
            f'<span class="stk-tip-icon">{icon_label}</span>'
            f'<span class="stk-tip-desc">{desc}</span>'
            f'</div>')


def zone_alert_html(d):
    """比較現價與本週(w0) Buy/Sell Zone，觸及時顯示提醒"""
    price  = d.get('price')
    buy_lo = d.get('w0_buy_lo')
    buy_hi = d.get('w0_buy_hi')
    sell_lo = d.get('w0_sell_lo')
    sell_hi = d.get('w0_sell_hi')
    if price is None: return ''

    alerts = []
    # 股價 <= buy_hi（進入或低於 Buy Zone）
    if buy_hi is not None and price <= buy_hi:
        alerts.append(
            '<div class="zone-alert-buy">🔴 分批買入提醒</div>'
        )
    # 股價 >= sell_lo（進入或高於 Sell Zone）
    if sell_lo is not None and price >= sell_lo:
        alerts.append(
            '<div class="zone-alert-sell">🟢 分批賣出提醒</div>'
        )
    if not alerts: return ''
    return '<div class="zone-alert-wrap">' + ''.join(alerts) + '</div>'


def stock_row(d):
    if not d.get('ok'):
        mag = "<span class='m7'>Mag 7</span>" if d['ticker'] in MAG7 else ''
        return (f'<tr><td><div class="tk"><span class="tb">{d["ticker"]}</span>{mag}</div></td>'
                f'<td colspan="9" style="color:#f85149;font-size:0.8em">⚠ 數據抓取失敗</td></tr>')

    ticker  = d['ticker']
    price   = d['price']
    chg     = d['change']
    chg_pct = d['change_pct']

    if chg > 0:   cc, cs, ps2 = 'pos', f'+{chg:.2f}', f'+{chg_pct:.2f}%'
    elif chg < 0: cc, cs, ps2 = 'neg', f'{chg:.2f}',  f'{chg_pct:.2f}%'
    else:         cc, cs, ps2 = 'neu', '—', '—'

    # 期權流欄（每週一格，含 OI + Zone + GEX）
    oi_w = [oi_html_single(d, f'w{i}') for i in range(4)]

    ttm_h, fwd_h     = pe_html(d['pe_ttm'], d['pe_fwd'])
    iv_bar, iv_badge = iv_html(d['iv'], d['ivr'])
    pc_h  = pc_html(d['pc_ratio'])
    ms_h  = ms_html(ticker, price)
    mc    = fmt_cap(d.get('market_cap'))
    vol_h = vol_cell_html(d)
    hma_h = hma_ema_cell_html(d)
    tip_h = stock_tip_html(ticker)
    mb    = f'<span class="m7">Mag 7</span>' if ticker in MAG7 else ''

    return f'''<tr>
  <td><div class="tk"><span class="tb">{ticker}</span>{mb}</div></td>
  <td class="pr-td">
    <div class="pr">${price:.2f}</div>
    <div class="chg-row {cc}">{cs} <span class="pct">{ps2}</span></div>
    <div class="vol-inline">{vol_h}</div>
    {tip_h}
  </td>
  <td class="hma-td">{hma_h}</td>
  <td class="oi">{oi_w[0]}</td>
  <td class="oi">{oi_w[1]}</td>
  <td class="oi">{oi_w[2]}</td>
  <td class="oi">{oi_w[3]}</td>
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
            sh, sl = ed.get(f'{k}_sell_hi'), ed.get(f'{k}_sell_lo')
            bh, bl = ed.get(f'{k}_buy_hi'),  ed.get(f'{k}_buy_lo')
            exps  = exp[5:] if exp else ''
            mob   = ('<span style="font-size:0.6em;color:#e3b341;background:rgba(227,179,65,.12);'
                     'border:1px solid rgba(227,179,65,.3);border-radius:3px;padding:0 3px;margin-left:2px">月結</span>' if is_mo else '')
            fmt   = lambda v: f'${v:g}' if v is not None else '—'
            zone_block = ''
            if sh is not None:
                sth = ed.get(f'{k}_settle_hi')
                stl = ed.get(f'{k}_settle_lo')
                fmt2 = lambda v: f'${v:.2f}' if v is not None else '—'
                zone_block = (f'<div style="margin-top:4px;padding-top:4px;border-top:1px dashed #30363d">'
                              f'<div style="font-family:monospace;font-size:0.78em;color:#56d364;font-weight:600">🟢 Sell {fmt(sl)}~{fmt(sh)}</div>'
                              f'<div style="font-family:monospace;font-size:0.78em;color:#f78166;font-weight:600;margin-top:2px">🔴 Buy {fmt(bl)}~{fmt(bh)}</div>'
                              f'<div style="font-family:monospace;font-size:0.72em;color:#e3b341;margin-top:2px">📌結算 {fmt2(stl)}~{fmt2(sth)}</div>'
                              f'</div>')
            cells.append(f'<div style="background:#21262d;border-radius:6px;padding:7px 10px">'
                         f'<div class="oi-lbl">{lbl} {exps}{mob}</div>'
                         f'<div class="oi-call" style="font-size:0.85em">▼ {fmt(cw)}</div>'
                         f'<div class="oi-pain" style="font-size:0.85em">⚡ {fmt(mp)}</div>'
                         f'<div class="oi-put"  style="font-size:0.85em">▲ {fmt(pw)}</div>'
                         f'{zone_block}</div>')
        return '\n'.join(cells)

    now_tw = datetime.now().strftime('%Y-%m-%d %H:%M')
    wd = {0:'一',1:'二',2:'三',3:'四',4:'五',5:'六',6:'日'}[datetime.now().weekday()]
    display_date = datetime.now().strftime(f'%Y年%m月%d日（週{wd}）')

    return f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>黑叔美股掃描 {display_date} — v7.1</title>
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
.oi{{min-width:155px;max-width:200px;vertical-align:top}}
.pr-td{{min-width:120px;max-width:160px;vertical-align:top}}
.chg-row{{font-size:0.85em;margin-top:3px;white-space:nowrap}}
.vol-inline{{margin-top:6px;padding-top:5px;border-top:1px dashed #30363d}}
.gex-inline{{margin-top:6px;padding-top:5px;border-top:1px dashed #30363d}}
.ema-td{{min-width:175px;max-width:210px;vertical-align:top}}
.hma-td{{min-width:240px;max-width:290px;vertical-align:top;padding:10px 13px}}
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
/* ═══ Buy/Sell Zone ═══ */
.zone-wrap{{margin-top:5px;padding-top:4px;border-top:1px dashed #30363d}}
.zone-sell{{font-family:monospace;font-size:0.78em;color:#56d364;white-space:nowrap;font-weight:600}}
.zone-buy{{font-family:monospace;font-size:0.78em;color:#f78166;white-space:nowrap;font-weight:600;margin-top:2px}}
/* ═══ 個股觀察建議 ═══ */
.stk-tip{{margin-top:6px;padding:4px 7px;background:rgba(121,192,255,.06);border:1px solid rgba(121,192,255,.15);border-radius:5px;line-height:1.45}}
.stk-tip-icon{{font-size:0.7em;color:#79c0ff;font-weight:700;display:block}}
.stk-tip-desc{{font-size:0.62em;color:#8b949e;display:block;margin-top:1px;line-height:1.3}}
/* ═══ Buy/Sell Zone 觸及提醒 ═══ */
.zone-alert-wrap{{margin-top:5px}}
.zone-alert-buy{{font-size:0.72em;font-weight:700;color:#f78166;background:rgba(247,129,102,.1);border:1px solid rgba(247,129,102,.3);border-radius:4px;padding:3px 7px;margin-bottom:3px;line-height:1.4}}
.zone-alert-sell{{font-size:0.72em;font-weight:700;color:#56d364;background:rgba(86,211,100,.1);border:1px solid rgba(86,211,100,.3);border-radius:4px;padding:3px 7px;line-height:1.4}}
.ftr{{background:#161b22;border:1px solid #30363d;border-radius:10px;padding:12px 18px;text-align:center;color:#484f58;font-size:0.7em;line-height:2;margin-top:14px}}
</style>
</head>
<body>
<div class="wrap">

<div class="hdr">
  <div>
    <h1>📊 美股每日開盤掃描報告 Powered by 黑叔 — v7.1</h1>
  </div>
  <div class="hdr-r">
    <div class="dt">{display_date}</div>
    <div>頁面生成：{now_tw} 台灣時間　｜　本機生成</div>
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
  <span style="color:#56d364;font-weight:600">🟢 Sell Zone</span>=週高壓力區（回測82%蓋頂）　<span style="color:#f78166;font-weight:600">🔴 Buy Zone</span>=週低支撐區（回測79%蓋底）　<span style="color:#e3b341">📌結算</span>=Max Pain±1%磁吸區
  <span>｜ HMA/EMA：</span>
  <span style="color:#3fb950">🟢支撐</span>
  <span style="color:#f85149">🔴壓力</span>
  <span style="color:#3fb950">🔔↑剛突破</span>
  <span style="color:#f85149">🔔↓剛跌破</span>
</div>

<div class="stitle">📋 個股全覽</div>
<div class="tbl-wrap">
<table>
<thead>
<tr>
  <th rowspan="2">代碼</th>
  <th rowspan="2" style="min-width:120px">現價 / 漲跌 / 量比</th>
  <th rowspan="2" style="min-width:165px;color:#a371f7">HMA/EMA 技術面</th>
  <th colspan="4" class="th-grp" style="color:#58a6ff">── 期 權 流（本週 / 下週 / 下下週 / 下下三週）──</th>
  <th rowspan="2" style="min-width:100px">晨星估值</th>
  <th colspan="2" class="th-grp th-val">── 估 值 ──</th>
  <th colspan="3" class="th-grp th-opt">── 期 權 分 析 ──</th>
  <th rowspan="2">市值</th>
</tr>
<tr>
  <th style="color:#58a6ff;font-size:0.65em;min-width:155px">本週 OI · Zone · GEX</th>
  <th style="color:#58a6ff;font-size:0.65em;min-width:155px">下週 OI · Zone · GEX</th>
  <th style="color:#58a6ff;font-size:0.65em;min-width:155px">下下週 OI · Zone · GEX</th>
  <th style="color:#58a6ff;font-size:0.65em;min-width:155px">下下三週 OI · Zone · GEX</th>
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
  <span style="color:#21262d">v7.1 · 四週OI + HMA/EMA最優參數 + Buy/Sell Zone v4.1（排除CallWall+硬下限）+ 均線評分 + 週五結算磁吸區</span>
</div>

</div>
</body>
</html>'''


# ═══════════════════════════════════════════════════════════════
#  主程式
# ═══════════════════════════════════════════════════════════════

def main():
    print('=' * 50)
    print('美股每日掃描報告 Powered by 黑叔 - v7.1')
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
