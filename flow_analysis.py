#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Options Flow Analysis — Barchart Flow 法 v4.0
用 yfinance 期權鏈計算：
  - Premium Flow / Delta-Weighted Flow / GEX
  - 上方壓力帶（三層）
  - 五法共識結算區
  - 下方三層防線 + 崩盤情境
  - IV Skew 警報

用法：
    python3 flow_analysis.py --ticker NVDA
    python3 flow_analysis.py --ticker AAPL --expiry 2026-03-20
    python3 flow_analysis.py --ticker TSLA --expiry 2026-03-06 --rate 0.053

依賴：pip install yfinance pandas
"""

import argparse
import math
import sys
from datetime import datetime, timedelta
from typing import Optional

try:
    import yfinance as yf
    import pandas as pd
except ImportError:
    print("❌ 缺少依賴，請執行：pip install yfinance pandas")
    sys.exit(1)


# ═══════════════════════════════════════════════════════════
#  EMA / HMA 最優參數表（52週 × 19股 回測研究）
#  來源：2026-03-15 HMA×短EMA 支撐壓力完整研究
#
#  欄位說明：
#    hma      : 最優 HMA 週期
#    ema      : 最優短 EMA 週期
#    sup5d    : HMA 支撐 5日勝率（%）
#    res5d    : HMA 壓力 5日勝率（%）
#    rating   : 評級（雙向強/支撐偏強/壓力弱/雙向弱）
#    fake_pct : 假穿越率（%）
# ═══════════════════════════════════════════════════════════

# 預設值（19股平均：HMA18，EMA13）
DEFAULT_HMA = 18
DEFAULT_EMA = 13

EMA_HMA_CONFIG = {
    # Ticker : (hma, ema, sup5d, res5d, rating, fake_pct)
    'SPY':  (24, 6,  83.3, 87.5, '雙向強',   47.9),
    'AAPL': (19, 14, 76.5, 83.3, '雙向強',   28.8),
    'QQQ':  (28, 6,  89.5, 63.2, '支撐強',   54.8),
    'JPM':  (28, 16, 80.0, 62.5, '雙向強',   41.7),
    'AMZN': (12, 19, 79.3, 62.1, '雙向強',   50.7),
    'META': (21, 16, 70.6, 68.8, '雙向強',   41.2),
    'AMD':  (19, 19, 73.3, 56.2, '支撐偏強', 41.2),
    'MSFT': (11, 17, 78.8, 50.0, '支撐偏強', 59.9),
    'GOOGL':(14, 19, 85.0, 42.1, '支撐偏強', 34.4),
    'NVDA': (22, 8,  76.0, 48.0, '支撐偏強', 54.0),
    'NFLX': (12, 11, 57.7, 60.0, '壓力偏強', 44.2),
    'MU':   (29, 13, 72.2, 36.8, '支撐偏強', 56.4),
    'TSM':  (11, 19, 61.8, 48.6, '普通',     58.0),
    'COHR': (25, 19, 66.7, 38.1, '支撐偏強', 46.3),
    'AVGO': (11, 19, 65.7, 34.3, '壓力弱',   60.6),
    'WDC':  (10, 10, 74.2, 25.8, '壓力弱',   57.3),
    'LITE': (28, 7,  70.6, 23.5, '壓力弱',   42.6),
    'TSLA': (15, 5,  53.3, 43.3, '雙向弱',   61.3),
    'SNDK': (12, 15, 63.0, 25.9, '壓力弱',   58.4),
}

# 壓力完全失效（只做多方支撐）
SUPPORT_ONLY = {'AVGO', 'WDC', 'LITE', 'SNDK'}
# 雙向弱（假穿越>58%，需等3日確認）
WEAK_BOTH    = {'TSLA', 'TSM'}


def _ema(series, period):
    """計算 EMA（指數移動平均）"""
    return series.ewm(span=period, adjust=False).mean()


def _hma(series, period):
    """
    計算 HMA（赫爾移動平均）
    HMA(n) = WMA(2*WMA(n/2) - WMA(n), sqrt(n))
    用 EWM 近似 WMA
    """
    half   = max(int(period / 2), 2)
    sqrtn  = max(int(period ** 0.5), 2)
    wma_h  = series.ewm(span=half,   adjust=False).mean()
    wma_n  = series.ewm(span=period, adjust=False).mean()
    raw    = 2 * wma_h - wma_n
    return raw.ewm(span=sqrtn, adjust=False).mean()


def calc_ema_hma_zones(ticker: str, current_price: float):
    """
    根據最優 HMA / 短EMA / 大EMA 計算技術面支撐壓力區、
    穿越方向偵測、盤整評分、中期趨勢概率。

    大EMA 參數：回測研究各股最優大EMA（趨勢逆轉用）
    若未收錄則用 EMA57（19股中位數）
    """
    # ── 大EMA 最優參數（趨勢逆轉研究，3年數據）─────────────
    BIG_EMA_CONFIG = {
        'SPY': 57, 'QQQ': 57, 'AAPL': 69, 'MSFT': 59, 'GOOGL': 66,
        'NVDA': 64, 'META': 58, 'AMZN': 51, 'TSLA': 58, 'AMD':  53,
        'AVGO': 51, 'NFLX': 61, 'JPM':  53, 'MU':   61, 'WDC':  56,
        'COHR': 58, 'TSM':  59, 'SNDK': 51, 'LITE': None,
    }
    DEFAULT_BIG_EMA = 57

    # ── 多頭逆轉勝率（突破大EMA後5日上漲概率，回測數據）────
    BULL_REVERSAL_WIN = {
        'AVGO': 71, 'QQQ': 71, 'SPY': 57, 'NFLX': 50, 'WDC': 50,
        'COHR': 33, 'AAPL': 45, 'NVDA': 48, 'META': 52, 'AMZN': 55,
        'MSFT': 40, 'AMD':  42, 'MU':   50, 'TSM':  45, 'TSLA': 38,
        'GOOGL': 55, 'JPM': 52, 'SNDK': 40, 'LITE': 40,
    }
    # ── 空頭逆轉勝率（跌破大EMA後5日持續下跌概率，回測數據）─
    BEAR_REVERSAL_WIN = {
        'MSFT': 50, 'QQQ': 40, 'AAPL': 38, 'MU': 36, 'NFLX': 31,
        'SPY': 30,  'NVDA': 28, 'META': 32, 'AMZN': 28, 'TSLA': 35,
        'AMD': 30,  'AVGO': 25, 'WDC':  22, 'COHR': 30, 'TSM': 28,
        'GOOGL': 25, 'JPM': 30, 'SNDK': 25, 'LITE': 20,
    }

    cfg = EMA_HMA_CONFIG.get(ticker.upper())
    use_default = cfg is None
    if use_default:
        hma_p, ema_p = DEFAULT_HMA, DEFAULT_EMA
        sup5d, res5d, rating, fake_pct = 65.0, 47.0, '預設參數', 50.0
    else:
        hma_p, ema_p, sup5d, res5d, rating, fake_pct = cfg

    big_ema_p = BIG_EMA_CONFIG.get(ticker.upper(), DEFAULT_BIG_EMA) or DEFAULT_BIG_EMA
    bull_win  = BULL_REVERSAL_WIN.get(ticker.upper(), 48)
    bear_win  = BEAR_REVERSAL_WIN.get(ticker.upper(), 30)

    try:
        hist_data = yf.Ticker(ticker).history(period='6mo')
        need = max(hma_p, ema_p, big_ema_p) + 10
        if hist_data.empty or len(hist_data) < need:
            return None

        close      = hist_data['Close'].astype(float)
        atr_series = (hist_data['High'] - hist_data['Low']).astype(float).rolling(14).mean()

        hma_series     = _hma(close, hma_p)
        ema_series     = _ema(close, ema_p)
        big_ema_series = _ema(close, big_ema_p)

        hma_val     = float(hma_series.iloc[-1])
        ema_val     = float(ema_series.iloc[-1])
        big_ema_val = float(big_ema_series.iloc[-1])
        atr_val     = float(atr_series.iloc[-1]) if not atr_series.iloc[-1:].isna().all() else current_price * 0.02
        S = current_price

        # ── 穿越方向偵測（今日 vs 前日）─────────────────────
        prev_close   = float(close.iloc[-2])
        prev_hma     = float(hma_series.iloc[-2])
        prev_ema     = float(ema_series.iloc[-2])
        prev_big_ema = float(big_ema_series.iloc[-2])

        # 由上而下穿越（從上跌破）→ 從支撐線變壓力線
        hma_cross_down = (prev_close > prev_hma) and (S < hma_val)
        # 由下而上穿越（從下突破）→ 從壓力線變支撐線
        hma_cross_up   = (prev_close < prev_hma) and (S > hma_val)
        ema_cross_down = (prev_close > prev_ema) and (S < ema_val)
        ema_cross_up   = (prev_close < prev_ema) and (S > ema_val)
        big_cross_down = (prev_close > prev_big_ema) and (S < big_ema_val)
        big_cross_up   = (prev_close < prev_big_ema) and (S > big_ema_val)

        # ── 當前股價與各線位置 ────────────────────────────────
        above_hma     = S > hma_val
        above_ema     = S > ema_val
        above_big_ema = S > big_ema_val
        hma_above_big = hma_val > big_ema_val
        ema_above_big = ema_val > big_ema_val

        # ── 動態角色判斷：支撐線 or 壓力線 ──────────────────
        # 股價在上方 → 該線是支撐線；股價在下方 → 該線是壓力線
        # 加入穿越事件強調
        def _line_role(above, cross_up, cross_down, line_name, val, win_rate_sup, win_rate_res):
            if cross_up:
                return f'🟢↑ {line_name} 剛突破 → 由壓力轉支撐 (${val:.2f})  勝率{win_rate_sup:.0f}%'
            elif cross_down:
                return f'🔴↓ {line_name} 剛跌破 → 由支撐轉壓力 (${val:.2f})  勝率{win_rate_res:.0f}%'
            elif above:
                return f'🟢  {line_name} = ${val:.2f} 作為支撐  勝率{win_rate_sup:.0f}%'
            else:
                return f'🔴  {line_name} = ${val:.2f} 作為壓力  勝率{win_rate_res:.0f}%'

        hma_role_str = _line_role(above_hma, hma_cross_up, hma_cross_down,
                                   f'HMA{hma_p}', hma_val, sup5d, res5d)
        ema_role_str = _line_role(above_ema, ema_cross_up, ema_cross_down,
                                   f'EMA{ema_p}', ema_val, sup5d * 0.95, res5d * 1.02)
        # 大EMA 用多/空頭逆轉概率
        big_ema_role = _line_role(above_big_ema, big_cross_up, big_cross_down,
                                   f'大EMA{big_ema_p}', big_ema_val, bull_win, bear_win)

        # ── 盤整評分（五條件法）─────────────────────────────
        # ATR 保護：若 ATR 太小則用股價的 1.5%
        atr = max(atr_val, S * 0.015)
        cond1 = abs(ema_val - hma_val) / atr < 0.8     # 兩線靠攏
        cond2 = abs(float(hma_series.iloc[-1]) - float(hma_series.iloc[-6])) / atr < 0.3  # HMA走平
        cond3 = abs(S - big_ema_val) / atr < 2.0       # 股價在大EMA附近
        cond4 = abs(S - float(_ema(close, 200).iloc[-1])) / atr < 3.0 if len(close) >= 200 else False
        cond5 = abs(ema_val - big_ema_val) / atr < 1.0 # 短EMA在大EMA附近
        consol_score = sum([cond1, cond2, cond3, cond4, cond5])
        is_consolidating = consol_score >= 3

        # ── 中期趨勢判斷（HMA + 短EMA 與大EMA 的關係）───────
        #
        # 研究結論：
        #   多頭逆轉：短EMA + HMA 同步突破大EMA → 延續上漲（各股平均48~71%）
        #   空頭逆轉：大部分跌破大EMA後20日內反彈，持續下跌概率僅22~50%
        #   盤整期間：兩線在大EMA附近纏繞，方向不明，等突破確認

        both_above_big = above_hma and above_ema and above_big_ema
        both_below_big = (not above_hma) and (not above_ema) and (not above_big_ema)
        just_crossed_up_big   = big_cross_up   or (hma_above_big and ema_above_big and not above_big_ema)
        just_crossed_down_big = big_cross_down or (not hma_above_big and not ema_above_big and above_big_ema)

        # 趨勢概率與建議
        if both_above_big and hma_above_big and ema_above_big:
            mid_trend_label = '📈 中期趨勢上行'
            mid_prob = bull_win
            mid_desc = (f'HMA{hma_p} + EMA{ema_p} 均在大EMA{big_ema_p}上方，'
                        f'趨勢延續上行概率 {bull_win}%')
            mid_action = '持多為主，回測大EMA支撐可加碼'
        elif big_cross_up:
            mid_trend_label = '📈 中期趨勢剛轉多'
            mid_prob = bull_win
            mid_desc = (f'收盤剛突破大EMA{big_ema_p}，多頭逆轉訊號，'
                        f'5日上漲概率 {bull_win}%')
            mid_action = f'等短EMA{ema_p}也站上大EMA確認，確認後可進多'
        elif both_below_big and not hma_above_big and not ema_above_big:
            mid_trend_label = '📉 中期趨勢下行'
            mid_prob = bear_win
            mid_desc = (f'HMA{hma_p} + EMA{ema_p} 均在大EMA{big_ema_p}下方，'
                        f'持續下跌概率 {bear_win}%（注意：多數個股跌破後仍會反彈）')
            mid_action = '偏空觀望，反彈至大EMA附近減碼，勿追空'
        elif big_cross_down:
            mid_trend_label = '📉 中期趨勢剛轉空'
            mid_prob = bear_win
            mid_desc = (f'收盤剛跌破大EMA{big_ema_p}，空頭逆轉訊號，'
                        f'5日持續下跌概率 {bear_win}%（大多數股票跌破後反彈）')
            mid_action = '謹慎，不建議直接做空，等3日確認是否真跌'
        elif is_consolidating:
            mid_trend_label = '⚪ 中期趨勢盤整'
            mid_prob = 70  # 盤整結束後70%向上突破（回測結論）
            mid_desc = (f'兩線纏繞大EMA{big_ema_p}附近（盤整分數 {consol_score}/5），'
                        f'盤整結束後70%概率向上突破')
            mid_action = '縮倉等待突破，盤整結束後第一根確認K棒再進場'
        else:
            # 混合狀態：一條線在上、一條線在下
            mid_trend_label = '🟡 中期趨勢分歧'
            mid_prob = 50
            mid_desc = (f'HMA/EMA 與大EMA{big_ema_p} 位置分歧，方向不明確')
            mid_action = '等 HMA 與短EMA 同步突破或跌破大EMA再操作'

        # ── HMA/EMA 緩衝帶（±0.5%）──────────────────────────
        buf_h   = hma_val * 0.005
        buf_e   = ema_val * 0.005
        buf_big = big_ema_val * 0.005

        # 支撐帶 / 壓力帶（根據股價位置動態切換）
        hma_zone = (round(hma_val - buf_h, 2), round(hma_val + buf_h, 2))
        ema_zone = (round(ema_val - buf_e, 2), round(ema_val + buf_e, 2))
        big_zone = (round(big_ema_val - buf_big, 2), round(big_ema_val + buf_big, 2))

        # 高假穿越率警告
        need_confirm = fake_pct >= 58.0
        confirm_note = '⚠ 假穿越率高，需等3日收盤確認' if need_confirm else '訊號相對乾淨'

        return {
            # 參數
            'hma_period':      hma_p,
            'ema_period':      ema_p,
            'big_ema_period':  big_ema_p,
            'sup5d':           sup5d,
            'res5d':           res5d,
            'rating':          rating,
            'fake_pct':        fake_pct,
            'use_default':     use_default,
            'support_only':    ticker.upper() in SUPPORT_ONLY,
            'weak_both':       ticker.upper() in WEAK_BOTH,
            'confirm_note':    confirm_note,
            # 線值
            'hma_val':         round(hma_val, 2),
            'ema_val':         round(ema_val, 2),
            'big_ema_val':     round(big_ema_val, 2),
            'atr_val':         round(atr_val, 2),
            # 緩衝帶
            'hma_zone':        hma_zone,
            'ema_zone':        ema_zone,
            'big_ema_zone':    big_zone,
            # 位置布林
            'above_hma':       above_hma,
            'above_ema':       above_ema,
            'above_big_ema':   above_big_ema,
            'hma_above_big':   hma_above_big,
            'ema_above_big':   ema_above_big,
            # 穿越事件
            'hma_cross_up':    hma_cross_up,
            'hma_cross_down':  hma_cross_down,
            'ema_cross_up':    ema_cross_up,
            'ema_cross_down':  ema_cross_down,
            'big_cross_up':    big_cross_up,
            'big_cross_down':  big_cross_down,
            # 動態角色字串
            'hma_role_str':    hma_role_str,
            'ema_role_str':    ema_role_str,
            'big_ema_role':    big_ema_role,
            # 盤整
            'consol_score':    consol_score,
            'is_consolidating': is_consolidating,
            'consol_conds':    [cond1, cond2, cond3, cond4, cond5],
            # 中期趨勢
            'mid_trend_label': mid_trend_label,
            'mid_prob':        mid_prob,
            'mid_desc':        mid_desc,
            'mid_action':      mid_action,
            # 概率參考
            'bull_win':        bull_win,
            'bear_win':        bear_win,
        }

    except Exception as e:
        return None


# ═══════════════════════════════════════════════════════════
#  Black-Scholes 工具函數
# ═══════════════════════════════════════════════════════════

def norm_cdf(x):
    return 0.5 * (1.0 + math.erf(x / math.sqrt(2.0)))

def bs_price(S, K, T, r, sigma, opt='call'):
    if T <= 0 or sigma <= 0:
        return max(0.0, S - K) if opt == 'call' else max(0.0, K - S)
    d1 = (math.log(S / K) + (r + 0.5 * sigma**2) * T) / (sigma * math.sqrt(T))
    d2 = d1 - sigma * math.sqrt(T)
    if opt == 'call':
        return S * norm_cdf(d1) - K * math.exp(-r * T) * norm_cdf(d2)
    else:
        return K * math.exp(-r * T) * norm_cdf(-d2) - S * norm_cdf(-d1)

def bs_delta(S, K, T, r, sigma, opt='call'):
    if T <= 0 or sigma <= 0:
        return (1.0 if S > K else 0.0) if opt == 'call' else (-1.0 if K > S else 0.0)
    d1 = (math.log(S / K) + (r + 0.5 * sigma**2) * T) / (sigma * math.sqrt(T))
    return norm_cdf(d1) if opt == 'call' else norm_cdf(d1) - 1.0

def bs_gamma(S, K, T, r, sigma):
    if T <= 0 or sigma <= 0:
        return 0.0
    try:
        d1 = (math.log(S / K) + (r + 0.5 * sigma**2) * T) / (sigma * math.sqrt(T))
        pdf = math.exp(-0.5 * d1**2) / math.sqrt(2.0 * math.pi)
        return pdf / (S * sigma * math.sqrt(T))
    except Exception:
        return 0.0


# ═══════════════════════════════════════════════════════════
#  Max Pain
# ═══════════════════════════════════════════════════════════

def calc_max_pain(calls_df, puts_df, current_price):
    try:
        lo, hi = current_price * 0.80, current_price * 1.20
        c = calls_df[(calls_df['strike'] >= lo) & (calls_df['strike'] <= hi)].copy()
        p = puts_df[(puts_df['strike']  >= lo) & (puts_df['strike']  <= hi)].copy()
        all_strikes = sorted(set(c['strike'].tolist() + p['strike'].tolist()))
        if not all_strikes:
            return None
        calls_map = dict(zip(c['strike'], c['openInterest'].fillna(0)))
        puts_map  = dict(zip(p['strike'], p['openInterest'].fillna(0)))
        best, min_pain = all_strikes[len(all_strikes) // 2], float('inf')
        for s in all_strikes:
            pain = (sum(max(0, s - k) * oi for k, oi in calls_map.items()) +
                    sum(max(0, k - s) * oi for k, oi in puts_map.items()))
            if pain < min_pain:
                min_pain, best = pain, s
        return best
    except Exception:
        return None


# ═══════════════════════════════════════════════════════════
#  IV Skew（25-delta）
# ═══════════════════════════════════════════════════════════

def calc_skew(call_rows, put_rows):
    try:
        call_25 = min(call_rows, key=lambda r: abs(abs(r['delta']) - 0.25))
        put_25  = min(put_rows,  key=lambda r: abs(abs(r['delta']) - 0.25))
        skew = put_25['iv'] - call_25['iv']
        return round(skew, 4), put_25['iv'], call_25['iv']
    except Exception:
        return None, None, None


# ═══════════════════════════════════════════════════════════
#  格式化工具
# ═══════════════════════════════════════════════════════════

def fmt(v):
    return f"${v:g}" if v is not None else '—'

def fmt_m(v):
    if v is None: return '—'
    return f"${abs(v)/1e6:.1f}M"

def section(title):
    print(f"\n{'='*66}")
    print(f"  {title}")
    print(f"{'='*66}")


def _print_hma_section(ticker: str, S: float, hma_data: dict):
    """輸出 HMA/EMA 技術面支撐壓力區塊（v3：動態角色 + 盤整 + 中期趨勢概率）"""
    section(f"📐 HMA/EMA 技術面支撐壓力 — {ticker}")

    if hma_data is None:
        print("  ⚠  無法計算 HMA/EMA（歷史數據不足）")
        return

    h = hma_data
    src_tag  = '（使用預設 HMA18/EMA13）' if h['use_default'] else f'（{ticker} 最優參數）'
    warn_txt = ''
    if h['support_only']:
        warn_txt = '⚠ 壓力線失效—只看支撐方向操作'
    elif h['weak_both']:
        warn_txt = '⚠ 雙向弱—需等3日收盤確認再進場'

    # ── 盤整條件細節 ──────────────────────────────────────
    conds = h['consol_conds']
    cond_symbols = ['✅' if c else '❌' for c in conds]
    cond_labels  = ['兩線靠攏', 'HMA走平', '近大EMA', '近EMA200', '短EMA近大EMA']
    cond_line    = '  '.join(f'{cond_symbols[i]}{cond_labels[i]}' for i in range(5))
    consol_bar   = '█' * h['consol_score'] + '░' * (5 - h['consol_score'])

    # ── 穿越事件 banner ───────────────────────────────────
    cross_events = []
    if h['hma_cross_up']:   cross_events.append(f'🔔 HMA{h["hma_period"]} 剛突破（壓力→支撐）')
    if h['hma_cross_down']: cross_events.append(f'🔔 HMA{h["hma_period"]} 剛跌破（支撐→壓力）')
    if h['ema_cross_up']:   cross_events.append(f'🔔 EMA{h["ema_period"]} 剛突破（壓力→支撐）')
    if h['ema_cross_down']: cross_events.append(f'🔔 EMA{h["ema_period"]} 剛跌破（支撐→壓力）')
    if h['big_cross_up']:   cross_events.append(f'🚨 大EMA{h["big_ema_period"]} 剛突破！趨勢逆轉訊號')
    if h['big_cross_down']: cross_events.append(f'🚨 大EMA{h["big_ema_period"]} 剛跌破！趨勢逆轉訊號')

    # ── 格式化輸出 ────────────────────────────────────────
    W = 57  # 框內寬度

    def row(txt, pad=W):
        # 補足空格到固定寬度（考慮中文字2格、emoji 2格）
        import unicodedata
        vis = 0
        for c in txt:
            if unicodedata.east_asian_width(c) in ('W', 'F'):
                vis += 2
            elif c in ('🟢','🔴','🟡','⚪','📈','📉','✅','❌','⚠','🔔','🚨','█','░','║','╠','╣','╔','╗','╚','╝'):
                vis += 2
            else:
                vis += 1
        padding = max(0, pad - vis)
        return f"  ║  {txt}{' ' * padding} ║"

    def divider(c='╠', r='╣'):
        return f"  {c}{'═' * (W + 4)}{r}"

    print(f"\n  參數來源  : HMA{h['hma_period']} / EMA{h['ema_period']} / 大EMA{h['big_ema_period']}  {src_tag}")
    print(f"  評級      : {h['rating']}　假穿越率：{h['fake_pct']:.0f}%　{h['confirm_note']}")
    if warn_txt:
        print(f"  ⚠ 警告    : {warn_txt}")
    if cross_events:
        print()
        for ev in cross_events:
            print(f"  {ev}")

    print(f"\n  ╔{'═' * (W + 4)}╗")
    print(f"  ║{'【均線當前值】':^{W + 2}} ║")
    print(divider())
    print(row(f"現價               : ${S:<10.2f}"))
    print(row(f"HMA{h['hma_period']:<3} 當前值      : ${h['hma_val']:<10.2f}  ATR={h['atr_val']:.2f}"))
    print(row(f"EMA{h['ema_period']:<3} 當前值      : ${h['ema_val']:<10.2f}"))
    print(row(f"大EMA{h['big_ema_period']:<3} 當前值    : ${h['big_ema_val']:<10.2f}"))

    print(divider())
    print(f"  ║{'【動態支撐 / 壓力角色】':^{W + 0}} ║")
    print(divider())
    # HMA 角色（±0.5% 緩衝帶）
    z = h['hma_zone']
    role_icon = '🟢' if h['above_hma'] else '🔴'
    role_word = '支撐帶' if h['above_hma'] else '壓力帶'
    sup_res_pct = h['sup5d'] if h['above_hma'] else h['res5d']
    print(row(f"{role_icon} HMA{h['hma_period']} {role_word} : ${z[0]:.2f}～${z[1]:.2f}  勝率{sup_res_pct:.0f}%"))
    if h['hma_cross_up']:
        print(row(f"   ↑ 剛突破：壓力線 → 支撐線轉換中"))
    elif h['hma_cross_down']:
        print(row(f"   ↓ 剛跌破：支撐線 → 壓力線轉換中"))

    # EMA 角色
    ez = h['ema_zone']
    role_icon_e = '🟢' if h['above_ema'] else '🔴'
    role_word_e = '支撐帶' if h['above_ema'] else '壓力帶'
    sup_res_e   = h['sup5d'] * 0.95 if h['above_ema'] else h['res5d'] * 1.02
    print(row(f"{role_icon_e} EMA{h['ema_period']} {role_word_e} : ${ez[0]:.2f}～${ez[1]:.2f}  勝率{sup_res_e:.0f}%"))
    if h['ema_cross_up']:
        print(row(f"   ↑ 剛突破：壓力線 → 支撐線轉換中"))
    elif h['ema_cross_down']:
        print(row(f"   ↓ 剛跌破：支撐線 → 壓力線轉換中"))

    # 大EMA 角色
    bz = h['big_ema_zone']
    role_icon_b = '🟢' if h['above_big_ema'] else '🔴'
    role_word_b = '支撐帶' if h['above_big_ema'] else '壓力帶'
    prob_b = h['bull_win'] if h['above_big_ema'] else h['bear_win']
    print(row(f"{role_icon_b} 大EMA{h['big_ema_period']} {role_word_b}: ${bz[0]:.2f}～${bz[1]:.2f}  勝率{prob_b:.0f}%"))
    if h['big_cross_up']:
        print(row(f"   ↑ 剛突破大EMA：多頭逆轉！勝率{h['bull_win']}%"))
    elif h['big_cross_down']:
        print(row(f"   ↓ 剛跌破大EMA：空頭警報，勝率{h['bear_win']}%"))

    # ── 盤整評分 ──────────────────────────────────────────
    print(divider())
    consol_title = '【⚠ 盤整區間警示】' if h['is_consolidating'] else '【盤整評分】'
    print(f"  ║  {consol_title:<{W - 1}} ║")
    print(divider())
    score_str = f"盤整分數：{h['consol_score']}/5  [{consol_bar}]"
    if h['is_consolidating']:
        print(row(f"{score_str}  ← 確認盤整！"))
        print(row(f"建議：縮倉等突破，70%概率向上突破"))
    else:
        print(row(f"{score_str}"))
    print(row(f"{cond_line}"))

    # ── 中期趨勢（最下方）────────────────────────────────
    print(divider())
    print(f"  ║  {'【中期趨勢判斷】':<{W - 1}} ║")
    print(divider())
    print(row(f"{h['mid_trend_label']}  概率 {h['mid_prob']}%"))
    # 折行顯示長說明
    desc = h['mid_desc']
    chunk = 26  # 每行約26中文字
    while desc:
        print(row(f"  {desc[:chunk]}"))
        desc = desc[chunk:]
    print(row(f"操作：{h['mid_action']}"))
    print(f"  ╚{'═' * (W + 4)}╝\n")


# ═══════════════════════════════════════════════════════════
#  主分析
# ═══════════════════════════════════════════════════════════

def analyze(ticker, target_expiry=None, rate=0.053):
    section(f"Options Flow Analysis — {ticker}  v4.0")

    t = yf.Ticker(ticker)
    hist = t.history(period='5d')
    if hist.empty:
        print(f"❌ 無法取得 {ticker} 股價數據"); return

    S    = float(hist['Close'].iloc[-1])
    prev = float(hist['Close'].iloc[-2]) if len(hist) > 1 else S
    chg_pct = (S - prev) / prev * 100
    print(f"  現價：${S:.2f}  ({chg_pct:+.2f}%)")

    # ── HMA / EMA 技術面計算 ──────────────────────────────
    print(f"  計算 HMA/EMA 技術面支撐壓力中…", end='', flush=True)
    hma_data = calc_ema_hma_zones(ticker, S)
    print(" ✓" if hma_data else " 略過")

    available = t.options
    if not available:
        print(f"❌ {ticker} 無可用期權數據"); return

    if target_expiry:
        target_dt = datetime.strptime(target_expiry, '%Y-%m-%d')
    else:
        today = datetime.now()
        days = (4 - today.weekday()) % 7 or 7
        target_dt = today + timedelta(days=days)

    best_expiry = min(available,
                      key=lambda e: abs((datetime.strptime(e, '%Y-%m-%d') - target_dt).days))
    expiry_dt = datetime.strptime(best_expiry, '%Y-%m-%d')
    T = max((expiry_dt - datetime.now()).days, 1) / 252
    print(f"  到期日：{best_expiry}（T={T*252:.0f} 交易日）")

    chain = t.option_chain(best_expiry)
    calls = chain.calls.copy()
    puts  = chain.puts.copy()
    for df in [calls, puts]:
        df['openInterest']      = pd.to_numeric(df['openInterest'],      errors='coerce').fillna(0)
        df['volume']            = pd.to_numeric(df['volume'],            errors='coerce').fillna(0)
        df['impliedVolatility'] = pd.to_numeric(df['impliedVolatility'], errors='coerce').fillna(0.5)

    lo, hi = S * 0.70, S * 1.30
    calls = calls[(calls['strike'] >= lo) & (calls['strike'] <= hi)].copy()
    puts  = puts[ (puts['strike']  >= lo) & (puts['strike']  <= hi)].copy()

    # ── BS 計算 ────────────────────────────────────────────
    def enrich(df, opt):
        rows = []
        for _, row in df.iterrows():
            K  = float(row['strike'])
            oi = float(row['openInterest'])
            v  = float(row['volume'])
            iv = float(row['impliedVolatility']) or 0.5
            mid   = bs_price(S, K, T, rate, iv, opt)
            delta = bs_delta(S, K, T, rate, iv, opt)
            gamma = bs_gamma(S, K, T, rate, iv)
            rows.append({
                'strike':     K,  'oi': oi, 'vol': v, 'iv': iv,
                'mid':        mid, 'delta': delta, 'gamma': gamma,
                'prem_flow':  v * mid * 100,
                'delta_flow': v * delta * 100,
                'gex':        oi * gamma * 100 * S * S,
                'vol_oi':     v / oi if oi > 0 else 0.0,
                'opt':        opt,
            })
        return rows

    call_rows = enrich(calls, 'call')
    put_rows  = enrich(puts,  'put')

    # ── Flow 彙總 ──────────────────────────────────────────
    total_call_prem  = sum(r['prem_flow']  for r in call_rows)
    total_put_prem   = sum(r['prem_flow']  for r in put_rows)
    total_call_delta = sum(r['delta_flow'] for r in call_rows)
    total_put_delta  = sum(r['delta_flow'] for r in put_rows)
    net_prem  = total_call_prem - total_put_prem
    net_delta = total_call_delta + total_put_delta
    call_pct  = (total_call_prem / (total_call_prem + total_put_prem) * 100
                 if (total_call_prem + total_put_prem) > 0 else 50)

    # ── GEX map ────────────────────────────────────────────
    gex_map = {}
    for r in call_rows:
        gex_map[r['strike']] = gex_map.get(r['strike'], 0.0) + r['gex']
    for r in put_rows:
        gex_map[r['strike']] = gex_map.get(r['strike'], 0.0) - r['gex']
    gex_asc  = sorted(gex_map.items(), key=lambda x:  x[1])
    gex_desc = sorted(gex_map.items(), key=lambda x: -x[1])

    # ── 聰明錢 ─────────────────────────────────────────────
    smart_calls = [r for r in call_rows if r['vol_oi'] > 0.3 and r['vol'] > 10_000]
    smart_puts  = [r for r in put_rows  if r['vol_oi'] > 0.3 and r['vol'] >  5_000]

    def weighted_target(rows, key):
        w = sum(abs(r[key]) for r in rows)
        return sum(r['strike'] * abs(r[key]) for r in rows) / w if w > 0 else None

    c_prem_tgt  = weighted_target(smart_calls, 'prem_flow')
    c_delta_tgt = weighted_target(smart_calls, 'delta_flow')
    p_prem_tgt  = weighted_target(smart_puts,  'prem_flow')
    p_delta_tgt = weighted_target(smart_puts,  'delta_flow')

    # ── ATM Straddle ───────────────────────────────────────
    atm_call = min(call_rows, key=lambda r: abs(r['strike'] - S), default=None)
    atm_put  = min(put_rows,  key=lambda r: abs(r['strike'] - S), default=None)
    straddle   = (atm_call['mid'] + atm_put['mid']) * 0.85 if atm_call and atm_put else 0
    straddle2s = straddle * 2

    # ── Max Pain ──────────────────────────────────────────
    max_pain = calc_max_pain(calls, puts, S)

    # ── IV Skew ───────────────────────────────────────────
    skew, put25_iv, call25_iv = calc_skew(call_rows, put_rows)

    # ── Call Wall / Put Wall ───────────────────────────────
    calls_above = [r for r in call_rows if S * 0.98 <= r['strike'] <= S * 1.30]
    puts_below  = [r for r in put_rows  if S * 0.70 <= r['strike'] <= S * 1.02]
    call_wall = max(calls_above, key=lambda r: r['oi'], default=None)
    put_wall  = max(puts_below,  key=lambda r: r['oi'], default=None)

    # ── 上方壓力三層 ──────────────────────────────────────
    resist_1     = call_wall['strike'] if call_wall else None
    resist_1_oi  = call_wall['oi']     if call_wall else 0
    resist_2     = gex_desc[0][0]      if gex_desc  else None
    resist_2_gex = gex_desc[0][1]      if gex_desc  else 0
    resist_3     = round(max_pain * 1.02, 1) if max_pain else None

    # ── 五法共識 ──────────────────────────────────────────
    five_vals = [v for v in [max_pain, resist_2, S + straddle, c_prem_tgt, c_delta_tgt]
                 if v is not None]
    consensus_hi = sum(five_vals) / len(five_vals) if five_vals else S
    confidence = ('🟢 高信心' if len(five_vals) >= 4 else
                  '🟡 中等信心' if len(five_vals) >= 3 else '🔴 低信心')

    # ── 下方三層防線 ──────────────────────────────────────
    support_1     = put_wall['strike'] if put_wall else None
    support_1_oi  = put_wall['oi']     if put_wall else 0
    support_2     = gex_asc[0][0]      if gex_asc  else None
    support_2_gex = gex_asc[0][1]      if gex_asc  else 0
    support_3     = round(max_pain * 0.97, 1) if max_pain else None
    crash_limit   = round(S - straddle2s, 1)

    # ── 實際操作區間（Sell Zone / Buy Zone）──────────────
    # 邏輯：做市商對沖帶來的真實買賣壓力帶，比五法均值更保守可靠
    #
    # Sell Zone（逢高賣出）：
    #   下限 = Max Pain 上方微幅（+0.5%），做市商開始賣股的起點
    #   上限 = GEX 最大阻力 與 保守Straddle(×0.55) 取小值，超過就是追高
    #
    # Buy Zone（逢低買入）：
    #   上限 = Max Pain 下方微幅（-0.5%），做市商開始買股的起點
    #   下限 = GEX 最大支撐 與 保守Straddle(×0.55) 取大值，跌破即止損

    conservative_straddle = straddle * 0.55  # 實際移動通常只有理論的 50~60%

    # Sell Zone
    sell_lo = round(max_pain * 1.005, 1) if max_pain else round(S * 1.005, 1)
    sell_hi_candidates = [v for v in [resist_2, S + conservative_straddle, resist_1] if v is not None]
    sell_hi = round(min(sell_hi_candidates), 1) if sell_hi_candidates else round(S * 1.03, 1)
    if sell_hi <= sell_lo:
        sell_hi = round(resist_1, 1) if resist_1 else round(sell_lo * 1.02, 1)

    # Buy Zone
    buy_hi = round(max_pain * 0.995, 1) if max_pain else round(S * 0.995, 1)
    buy_lo_candidates = [v for v in [support_2, S - conservative_straddle, support_1] if v is not None]
    buy_lo = round(max(buy_lo_candidates), 1) if buy_lo_candidates else round(S * 0.97, 1)
    if buy_lo >= buy_hi:
        buy_lo = round(support_1, 1) if support_1 else round(buy_hi * 0.98, 1)

    # ── 突破後延伸情境（基於 90筆回測：straddle×1.3 蓋頂 82% / straddle×1.5 蓋頂 87%）──
    # 取代舊版 OTM：舊版在非交易時段 vol=0 退化成純 OI 排序，無預測意義
    # 新版用 straddle 係數 + 次級 Wall 提供有統計支撐的延伸目標

    # 上行突破延伸（站上 Sell Zone 後）
    up_ext_80  = round(S + straddle * 1.3, 1)   # 80% 週高蓋頂線（P75 straddle乘數）
    up_ext_90  = round(S + straddle * 1.5, 1)   # 90% 週高蓋頂線（P90 straddle乘數）
    # 下行突破延伸（跌破 Buy Zone 後）
    dn_ext_80  = round(S - straddle * 1.3, 1)   # 80% 週低蓋底線
    dn_ext_90  = round(S - straddle * 1.5, 1)   # 90% 週低蓋底線

    # 次級 Call Wall / Put Wall（OI 第二大，突破第一層後的下一磁吸）
    calls_above_sorted = sorted(
        [r for r in call_rows if r['strike'] > S * 0.98 and r['oi'] > 0],
        key=lambda r: -r['oi'])
    puts_below_sorted  = sorted(
        [r for r in put_rows  if r['strike'] < S * 1.02 and r['oi'] > 0],
        key=lambda r: -r['oi'])
    call_wall2 = calls_above_sorted[1]['strike'] if len(calls_above_sorted) >= 2 else None
    put_wall2  = puts_below_sorted[1]['strike']  if len(puts_below_sorted)  >= 2 else None

    # 突破觸發門檻仍使用 sell_hi / buy_lo
    up_trigger = sell_hi
    dn_trigger = buy_lo

    # ════════════════════════════════════════════════════════
    #  輸出
    # ════════════════════════════════════════════════════════

    section("CALL 資金流")
    print(f"  {'Strike':>8}  {'Vol':>8}  {'MidPx':>7}  {'PremFlow$M':>11}  {'Delta':>7}  {'Vol/OI':>7}  聰明錢")
    print(f"  {'-'*70}")
    for r in sorted(call_rows, key=lambda x: -x['prem_flow'])[:12]:
        flag = '🔥' if r['vol_oi'] > 0.5 and r['vol'] > 10_000 else ''
        print(f"  ${r['strike']:>7.1f}  {r['vol']:>8,.0f}  ${r['mid']:>6.2f}  "
              f"${r['prem_flow']/1e6:>9.2f}M  {r['delta']:>7.3f}  {r['vol_oi']:>7.2f}x  {flag}")

    section("PUT 資金流")
    print(f"  {'Strike':>8}  {'Vol':>8}  {'MidPx':>7}  {'PremFlow$M':>11}  {'Delta':>7}  {'Vol/OI':>7}  聰明錢")
    print(f"  {'-'*70}")
    for r in sorted(put_rows, key=lambda x: -x['prem_flow'])[:12]:
        flag = '🔥' if r['vol_oi'] > 0.5 and r['vol'] > 5_000 else ''
        print(f"  ${r['strike']:>7.1f}  {r['vol']:>8,.0f}  ${r['mid']:>6.2f}  "
              f"${r['prem_flow']/1e6:>9.2f}M  {r['delta']:>7.3f}  {r['vol_oi']:>7.2f}x  {flag}")

    section("聰明錢大單排行（Vol/OI > 0.3）")
    all_smart = [('CALL', r) for r in smart_calls] + [('PUT', r) for r in smart_puts]
    all_smart.sort(key=lambda x: -x[1]['prem_flow'])
    print(f"  {'方向':>5}  {'Strike':>8}  {'Vol/OI':>7}  {'PremFlow$M':>11}  {'DeltaFlow':>11}")
    print(f"  {'-'*55}")
    for side, r in all_smart[:10]:
        emoji = '📈' if side == 'CALL' else '📉'
        print(f"  {emoji}{side:>4}  ${r['strike']:>7.1f}  {r['vol_oi']:>7.2f}x  "
              f"${r['prem_flow']/1e6:>9.2f}M  ${r['delta_flow']/1e6:>+9.2f}M")

    section("資金流彙總")
    flow_dir  = '📈 Call 主導（偏多）' if net_prem  > 0 else '📉 Put 主導（偏空）'
    delta_dir = '多頭' if net_delta > 0 else '空頭'
    print(f"  Call Premium Flow  ：${total_call_prem/1e6:>8.2f}M")
    print(f"  Put  Premium Flow  ：${total_put_prem/1e6:>8.2f}M")
    print(f"  淨 Flow            ：${net_prem/1e6:>+8.2f}M　{flow_dir}")
    print(f"  Call 佔比          ：{call_pct:.1f}%")
    print(f"  淨 Delta 暴露      ：{net_delta/1e6:>+8.2f}M 股當量（{delta_dir}主導）")

    section("IV Skew 警報（25-delta）")
    if skew is not None:
        skew_status = ('✅ 正常　無特別尾部風險'     if skew < 0.03 else
                       '⚠️  偏高　市場開始定價下行風險' if skew < 0.07 else
                       '🚨 極端　機構大量買 Put，注意崩盤')
        print(f"  25-delta Put IV    ：{put25_iv*100:.1f}%")
        print(f"  25-delta Call IV   ：{call25_iv*100:.1f}%")
        print(f"  Skew（Put－Call）  ：{skew*100:+.2f}%")
        print(f"  狀態               ：{skew_status}")
    else:
        print("  Skew：無法計算")

    # ── 最終價位預測大表 ──────────────────────────────────
    section(f"完整價位預測　｜　{confidence}（{len(five_vals)}/5 法有效）")
    skew_str = f"{skew*100:+.2f}%" if skew is not None else 'N/A'
    tail_str = '尾部風險偏高 ⚠️' if skew and skew > 0.05 else '無異常'

    print(f"""
  現價 ${S:.2f}　到期 {best_expiry}

  ╔═══════════════════════════════════════════════════════╗
  ║               【上方壓力帶（三層）】                  ║
  ╠═══════════════════════════════════════════════════════╣
  ║  第一層 Call Wall  : {fmt(resist_1):<8} OI={resist_1_oi:>8,.0f}          ║
  ║  第二層 GEX 阻力   : {fmt(resist_2):<8} 做市商賣壓 {fmt_m(resist_2_gex):>8}      ║
  ║  第三層 Max Pain↑  : {fmt(resist_3):<8} (Max Pain×1.02)            ║
  ╠═══════════════════════════════════════════════════════╣
  ║               【五法共識結算區】                      ║
  ╠═══════════════════════════════════════════════════════╣
  ║  Max Pain          : {fmt(max_pain):<10}                          ║
  ║  GEX 最大阻力      : {fmt(resist_2):<10}                          ║
  ║  Straddle 上限85%  : ${S+straddle:<10.2f}                          ║
  ║  聰明錢 Premium加權: {fmt(c_prem_tgt):<10}                          ║
  ║  聰明錢 Delta 加權 : {fmt(c_delta_tgt):<10}                          ║
  ║  ─────────────────────────────────────────────────  ║
  ║  ★ 五法均值共識   : ${consensus_hi:.2f}  (理論錨點，偏樂觀)           ║
  ╠═══════════════════════════════════════════════════════╣
  ║          【實際操作區間（回測驗證）】                 ║
  ╠═══════════════════════════════════════════════════════╣
  ║  🔴 逢高賣出 Sell Zone : ${sell_lo:<6.1f} ～ ${sell_hi:<6.1f}                    ║
  ║     依據：Max Pain上緣 ＋ GEX阻力 ＋ Straddle×0.3~1.3 ║
  ║  🟢 逢低買入 Buy  Zone : ${buy_lo:<6.1f} ～ ${buy_hi:<6.1f}                    ║
  ║     依據：Max Pain下緣 ＋ GEX支撐 ＋ Straddle×0.3~1.3 ║
  ║  📌 週五結算磁吸區     : ${round(max_pain*0.99,1):<6.1f} ～ ${round(max_pain*1.01,1):<6.1f}  (Max Pain±1%)      ║
  ╠═══════════════════════════════════════════════════════╣
  ║          【突破後延伸情境（Zone 被穿越時參考）】       ║
  ╠═══════════════════════════════════════════════════════╣
  ║  ↑ 站上 ${up_trigger:<6.1f}（Sell Zone上緣）觸發：         ║
  ║    延伸目標80% : ${up_ext_80:<8.1f} (Straddle×1.3, P75)           ║
  ║    延伸目標90% : ${up_ext_90:<8.1f} (Straddle×1.5, P90)           ║
  ║    次級Call Wall: {fmt(call_wall2):<8} (OI第二大壓力)             ║
  ║  ↓ 跌破 ${dn_trigger:<6.1f}（Buy  Zone下緣）觸發：         ║
  ║    延伸目標80% : ${dn_ext_80:<8.1f} (Straddle×1.3, P75)           ║
  ║    延伸目標90% : ${dn_ext_90:<8.1f} (Straddle×1.5, P90)           ║
  ║    次級Put Wall : {fmt(put_wall2):<8} (OI第二大支撐)             ║
  ╠═══════════════════════════════════════════════════════╣
  ║               【下方防線（三層）】                    ║
  ╠═══════════════════════════════════════════════════════╣
  ║  第一防線 Put Wall : {fmt(support_1):<8} OI={support_1_oi:>8,.0f}          ║
  ║  第二防線 GEX 支撐 : {fmt(support_2):<8} 做市商買盤 {fmt_m(support_2_gex):>8}      ║
  ║  第三防線 MaxPain↓ : {fmt(support_3):<8} (Max Pain×0.97)            ║
  ╠═══════════════════════════════════════════════════════╣
  ║               【崩盤情境（2σ）】                      ║
  ╠═══════════════════════════════════════════════════════╣
  ║  ATM Straddle 1σ   : ±${straddle:.2f}  (68% 機率區間)           ║
  ║  2σ 壓力區間       : ${S-straddle2s:.1f} ～ ${S+straddle2s:.1f}  (95% 機率)           ║
  ║  崩盤觸發下限      : ${crash_limit:.1f}  (突破即加速下殺)             ║
  ║  最終護城河        : {fmt(support_2):<8} (GEX最大負值，自動托底)    ║
  ╠═══════════════════════════════════════════════════════╣
  ║  資金流  : {flow_dir:<20}                      ║
  ║  Delta   : {net_delta/1e6:>+8.1f}M 股當量                          ║
  ║  IV Skew : {skew_str:<8}  {tail_str:<20}              ║
  ╚═══════════════════════════════════════════════════════╝
""")

    # ── HMA / EMA 技術面支撐壓力輸出 ─────────────────────
    _print_hma_section(ticker, S, hma_data)

    print("  ⚠  不構成投資建議　｜　數據來源 Yahoo Finance（yfinance）\n")

    result = {
        'ticker': ticker, 'price': S, 'expiry': best_expiry,
        'net_flow_m': round(net_prem/1e6, 2),
        'net_delta_m': round(net_delta/1e6, 2),
        'resist_1_call_wall': resist_1,
        'resist_2_gex': resist_2,
        'resist_3_mp_upper': resist_3,
        'consensus_hi': round(consensus_hi, 2),
        'sell_zone': (sell_lo, sell_hi),
        'buy_zone': (buy_lo, buy_hi),
        'settlement_zone': (round(max_pain*0.99, 1), round(max_pain*1.01, 1)),
        'up_ext_80': up_ext_80,
        'up_ext_90': up_ext_90,
        'dn_ext_80': dn_ext_80,
        'dn_ext_90': dn_ext_90,
        'call_wall2': call_wall2,
        'put_wall2': put_wall2,
        'up_trigger': up_trigger,
        'dn_trigger': dn_trigger,
        'support_1_put_wall': support_1,
        'support_2_gex': support_2,
        'support_3_mp_lower': support_3,
        'crash_limit_2sigma': crash_limit,
        'max_pain': max_pain,
        'straddle': round(straddle, 2),
        'iv_skew': round(skew, 4) if skew else None,
    }
    if hma_data:
        result['hma_zone']       = hma_data['hma_zone']
        result['ema_zone']       = hma_data['ema_zone']
        result['big_ema_zone']   = hma_data['big_ema_zone']
        result['hma_val']        = hma_data['hma_val']
        result['ema_val']        = hma_data['ema_val']
        result['big_ema_val']    = hma_data['big_ema_val']
        result['mid_trend']      = hma_data['mid_trend_label']
        result['mid_prob']       = hma_data['mid_prob']
        result['is_consolidating'] = hma_data['is_consolidating']
    return result


# ═══════════════════════════════════════════════════════════
#  CLI 入口
# ═══════════════════════════════════════════════════════════

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Options Flow Analysis v4.0')
    parser.add_argument('--ticker', '-t', default='NVDA')
    parser.add_argument('--expiry', '-e', default=None)
    parser.add_argument('--rate',   '-r', default=0.053, type=float)
    args = parser.parse_args()
    analyze(args.ticker, args.expiry, args.rate)
