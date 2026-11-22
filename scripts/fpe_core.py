"""数据读取、区间处理、频率识别与指标计算。

这里的 Python 口径与 fpe_excel.py 写入的 Excel 公式逐项一致，
tests/ 会用真实 Excel / LibreOffice 重算来核对两边结果。
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

CALENDAR_DAYS = 365
FREQ_FACTORS = {'daily': 252, 'weekly': 52, 'monthly': 12, 'quarterly': 4}
FREQ_LABELS = {'daily': '日频', 'weekly': '周频', 'monthly': '月频', 'quarterly': '季频'}

_DATE_PAT = re.compile(r'(?i)(date|日期)')
_DATE_EXCLUDE = re.compile(r'(到期|成立|申购|赎回|分红|除息|公告)')
_UNIT_PAT = re.compile(r'(?i)(单位净值|unit.?nav|^nav$|^净值$)')
_ACCUM_PAT = re.compile(r'(?i)(累计|accum|cumulative)')
_NAV_PAT = re.compile(r'(?i)(净值|nav)')
_RETURN_PAT = re.compile(r'(?i)(收益率|增长率|return|涨跌)')
_BENCH_PAT = re.compile(r'(?i)(收盘|close|点位|指数|净值|nav|价格|price|value)')
_CN_NUM = '零〇一二三四五六七八九十百千'


class SkipFund(Exception):
    """该文件无法评估（数据不足、识别不到列等），批量时跳过并汇报原因。"""


# ---------------------------------------------------------------- 读取与识别
def _first_match(cols, include, exclude=None):
    for c in cols:
        cs = str(c).strip()
        if include.search(cs) and not (exclude and exclude.search(cs)):
            return c
    return None


def find_columns(df: pd.DataFrame, use_accum: bool = False):
    """返回 (日期列, 净值列, 是否用了累计净值)。

    use_accum=True 时优先找累计净值列（同时含“累计/accum”与“净值/nav”，排除收益率类），
    找不到则回退到单位净值列。
    """
    cols = df.columns.tolist()
    date_col = _first_match(cols, _DATE_PAT, _DATE_EXCLUDE)

    nav_col, used_accum = None, False
    if use_accum:
        for c in cols:
            cs = str(c).strip()
            if _ACCUM_PAT.search(cs) and _NAV_PAT.search(cs) and not _RETURN_PAT.search(cs):
                nav_col, used_accum = c, True
                break
    if nav_col is None:
        for c in cols:
            cs = str(c).strip()
            if _UNIT_PAT.search(cs) and not _ACCUM_PAT.search(cs):
                nav_col = c
                break

    if date_col is None or nav_col is None:
        raise SkipFund(f'无法识别日期列或净值列，可用列: {[str(c) for c in cols]}')
    return date_col, nav_col, used_accum


def _clean_series(df, date_col, value_col, value_name):
    out = pd.DataFrame({
        'date': pd.to_datetime(df[date_col], errors='coerce').dt.normalize(),
        value_name: pd.to_numeric(df[value_col], errors='coerce'),
    })
    before = len(out)
    out = out.dropna()
    out = out[out[value_name] > 0]
    out = out.drop_duplicates('date', keep='last').sort_values('date').reset_index(drop=True)
    return out, before - len(out)


def load_nav(path: str, use_accum: bool = False):
    """读取净值文件，返回 (DataFrame[date, nav], 备注列表)。"""
    raw = pd.read_excel(path)
    date_col, nav_col, used_accum = find_columns(raw, use_accum)
    df, dropped = _clean_series(raw, date_col, nav_col, 'nav')
    notes = []
    if use_accum and not used_accum:
        notes.append('未找到累计净值列，已使用单位净值')
    if dropped:
        notes.append(f'剔除 {dropped} 行空值/无效数据')
    if len(df) < 2:
        raise SkipFund('有效数据不足 2 行')
    return df, notes


def load_benchmark(path: str) -> pd.DataFrame:
    """读取基准文件，返回 DataFrame[date, bench]。"""
    raw = pd.read_excel(path)
    cols = raw.columns.tolist()
    date_col = _first_match(cols, _DATE_PAT, _DATE_EXCLUDE)
    if date_col is None:
        raise ValueError(f'基准文件无法识别日期列，可用列: {[str(c) for c in cols]}')
    value_col = _first_match([c for c in cols if c != date_col], _BENCH_PAT, _RETURN_PAT)
    if value_col is None:  # 退而求其次：第一个数值列
        for c in cols:
            if c != date_col and pd.to_numeric(raw[c], errors='coerce').notna().mean() > 0.8:
                value_col = c
                break
    if value_col is None:
        raise ValueError(f'基准文件无法识别点位/净值列，可用列: {[str(c) for c in cols]}')
    df, _ = _clean_series(raw, date_col, value_col, 'bench')
    if len(df) < 2:
        raise ValueError('基准文件有效数据不足 2 行')
    return df


def extract_prod_name(path: str):
    """产品名：从第一个中文字符到“X号”（X 可为阿拉伯或中文数字），其次“X期”，
    再退到去掉“原始数据/数据/净值”后缀的文件名。文件名含“绩效评估”返回 None（输出文件）。"""
    base = os.path.splitext(os.path.basename(path))[0]
    if '绩效评估' in base:
        return None
    for suffix in ('号', '期'):
        m = re.search(rf'[一-鿿].*?(?:\d+|[{_CN_NUM}]+){suffix}', base)
        if m:
            return m.group(0)
    name = re.sub(r'(原始)?(净值)?数据$|净值$', '', base).strip(' _-')
    return name or base


# ---------------------------------------------------------------- 区间与频率
def slice_interval(df: pd.DataFrame, start=None, end=None):
    """按 [start, end] 截取区间，并往前多取一个净值日作为收益率分母。

    返回 (截取后的 df, 用户区间起点在 df 中的行号)。行号为 1 表示第 0 行是扩展行。
    """
    display_idx = 0
    if start:
        in_range = df.index[df['date'] >= pd.to_datetime(start)]
        if len(in_range) == 0:
            raise SkipFund('区间内有效数据不足 2 行')
        first = in_range[0]
        if first > 0:
            df = df.iloc[first - 1:]
            display_idx = 1
        else:
            df = df.iloc[first:]
    if end:
        df = df[df['date'] <= pd.to_datetime(end)]
    df = df.reset_index(drop=True)
    if len(df) < 2:
        raise SkipFund('区间内有效数据不足 2 行')
    return df, display_idx


def detect_freq(dates: pd.Series) -> str:
    """按相邻净值日间隔（自然日）的中位数判断频率。"""
    gap = pd.Series(dates).diff().dt.days.dropna().median()
    if gap <= 4:
        return 'daily'
    if gap <= 10:
        return 'weekly'
    if gap <= 45:
        return 'monthly'
    return 'quarterly'


# ---------------------------------------------------------------- 指标
def annualize(growth: float, days: int, linear: bool) -> float:
    """growth 为期末/期初（如 1.05）。复利: g^(365/d)-1；线性: (g-1)*365/d。"""
    days = max(days, 1)
    if linear:
        return (growth - 1) * CALENDAR_DAYS / days
    return growth ** (CALENDAR_DAYS / days) - 1


def _safe_div(a, b):
    if b is None or b == 0 or not np.isfinite(b):
        return None
    return a / b


def drawdown_info(dates: pd.Series, nav: np.ndarray) -> dict:
    """最大回撤及其峰值日、谷值日、修复日、修复天数（谷值日→修复日的自然日）。

    口径与 Excel 公式一致：谷值取回撤最小的第一行；峰值取该处历史最高净值第一次出现的日期；
    修复日为谷值之后第一个净值 >= 峰值净值的日期。
    """
    cummax = np.maximum.accumulate(nav)
    dd = nav / cummax - 1
    mdd = float(dd.min())
    info = {'max_drawdown': mdd, 'peak_nav': None, 'peak_date': None, 'trough_date': None,
            'recovery_date': None, 'recovery_days': None, 'status': None}
    if mdd == 0:
        info['status'] = '无回撤'
        return info
    trough = int(np.argmin(dd))
    peak_nav = float(cummax[trough])
    peak = int(np.flatnonzero(nav == peak_nav)[0])
    after = np.flatnonzero((np.arange(len(nav)) > trough) & (nav >= peak_nav))
    dates = pd.Series(dates).reset_index(drop=True)
    info.update(peak_nav=peak_nav, peak_date=dates[peak], trough_date=dates[trough])
    if len(after):
        rec = int(after[0])
        info.update(recovery_date=dates[rec], recovery_days=int((dates[rec] - dates[trough]).days))
    else:
        info['status'] = '未修复'
    return info


def compute_metrics(df: pd.DataFrame, factor: int, rf: float = 0.0, linear: bool = False) -> dict:
    nav = df['nav'].to_numpy(dtype=float)
    dates = df['date']
    rets = nav[1:] / nav[:-1] - 1
    days = int((dates.iloc[-1] - dates.iloc[0]).days)

    ann = annualize(nav[-1] / nav[0], days, linear)
    vol = float(np.std(rets, ddof=1) * np.sqrt(factor)) if len(rets) >= 2 else None
    downside = float(np.sqrt(np.mean(np.minimum(rets, 0) ** 2)) * np.sqrt(factor))
    dd = drawdown_info(dates, nav)

    return {
        'days': days,
        'interval_return': nav[-1] / nav[0] - 1,
        'annual_return': ann,
        'annual_vol': vol,
        'sharpe': _safe_div(ann - rf, vol),
        'calmar': _safe_div(ann, abs(dd['max_drawdown'])),
        'sortino': _safe_div(ann - rf, downside),
        **dd,
    }


def align_benchmark(df: pd.DataFrame, bench: pd.DataFrame) -> pd.DataFrame:
    """只保留产品与基准都有数据的日期。"""
    return df.merge(bench, on='date', how='inner').reset_index(drop=True)


def excess_curve(aligned: pd.DataFrame, mode: str) -> np.ndarray:
    p = aligned['nav'].to_numpy(float) / aligned['nav'].iloc[0]
    b = aligned['bench'].to_numpy(float) / aligned['bench'].iloc[0]
    return p / b if mode == 'geometric' else 1 + p - b


def compute_excess(aligned: pd.DataFrame, factor: int, mode: str = 'geometric',
                   linear: bool = False) -> dict:
    """基准与超额指标（在共同日期上计算）。

    几何: 超额净值 = 产品归一 / 基准归一，超额期收益 = 超额净值环比；
    算术: 超额净值 = 1 + 产品归一 - 基准归一，超额期收益 = 产品期收益 - 基准期收益，
          超额年化 = 产品同期年化 - 基准年化。
    """
    nav = aligned['nav'].to_numpy(float)
    bench = aligned['bench'].to_numpy(float)
    days = int((aligned['date'].iloc[-1] - aligned['date'].iloc[0]).days)
    ex = excess_curve(aligned, mode)

    prod_ann = annualize(nav[-1] / nav[0], days, linear)
    bench_ann = annualize(bench[-1] / bench[0], days, linear)
    if mode == 'geometric':
        ex_rets = ex[1:] / ex[:-1] - 1
        ex_ann = annualize(ex[-1], days, linear)
    else:
        ex_rets = (nav[1:] / nav[:-1] - 1) - (bench[1:] / bench[:-1] - 1)
        ex_ann = prod_ann - bench_ann
    te = float(np.std(ex_rets, ddof=1) * np.sqrt(factor)) if len(ex_rets) >= 2 else None
    ex_dd = float((ex / np.maximum.accumulate(ex) - 1).min())

    return {
        'common_count': len(aligned),
        'bench_interval_return': bench[-1] / bench[0] - 1,
        'bench_annual_return': bench_ann,
        'prod_annual_return': prod_ann,
        'excess_interval_return': ex[-1] - 1,
        'excess_annual_return': ex_ann,
        'tracking_error': te,
        'information_ratio': _safe_div(ex_ann, te),
        'excess_max_drawdown': ex_dd,
    }


# ---------------------------------------------------------------- 单只产品
@dataclass
class Options:
    start: str | None = None
    end: str | None = None
    use_accum: bool = False
    linear: bool = False
    freq: str = 'auto'
    rf: float = 0.0
    excess: str = 'geometric'


@dataclass
class FundResult:
    src: str
    name: str
    df: pd.DataFrame
    display_idx: int
    freq: str
    freq_auto: bool
    factor: int
    metrics: dict
    notes: list = field(default_factory=list)
    aligned: pd.DataFrame | None = None
    excess: dict | None = None

    @property
    def interval_label(self) -> str:
        d = self.df['date']
        return f"{d.iloc[self.display_idx]:%Y/%m/%d}-{d.iloc[-1]:%Y/%m/%d}"

    @property
    def freq_label(self) -> str:
        return FREQ_LABELS[self.freq] + ('（自动识别）' if self.freq_auto else '（手动指定）')


def evaluate(path: str, opts: Options, bench: pd.DataFrame | None = None) -> FundResult:
    name = extract_prod_name(path)
    if name is None:
        raise SkipFund('已是绩效评估输出文件')
    df, notes = load_nav(path, opts.use_accum)
    df, display_idx = slice_interval(df, opts.start, opts.end)

    freq_auto = opts.freq == 'auto'
    freq = detect_freq(df['date']) if freq_auto else opts.freq
    factor = FREQ_FACTORS[freq]
    metrics = compute_metrics(df, factor, opts.rf, opts.linear)
    if metrics['days'] < CALENDAR_DAYS:
        notes.append('区间不足一年，年化指标仅供参考')

    aligned = excess = None
    if bench is not None:
        a = align_benchmark(df, bench)
        if len(a) < 2:
            notes.append('与基准的共同日期不足 2 个，未计算超额')
        else:
            aligned = a
            excess = compute_excess(a, factor, opts.excess, opts.linear)
            if len(a) < len(df):
                notes.append(f'与基准共同日期 {len(a)}/{len(df)} 个，超额指标只用共同日期')

    return FundResult(path, name, df, display_idx, freq, freq_auto, factor, metrics,
                      notes, aligned, excess)
