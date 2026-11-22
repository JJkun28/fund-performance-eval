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
_ACCUM_PAT = re.compile(r'(?i)(累计|accum|cumulative)')
_NAV_PAT = re.compile(r'(?i)(净值|nav)')
# 名字里带这些词的“净值”列不是我们要的单位/累计净值（复权净值、净值增长率、净值日期等）
_NAV_EXCLUDE = re.compile(r'(?i)(复权|adj|收益|增长|return|涨跌|日期|date|规模|总额|份额)')
_RETURN_PAT = re.compile(r'(?i)(收益率|增长率|return|涨跌)')
_UNIT_EXACT = {'单位净值', '净值', '基金净值', '产品净值', 'nav', 'unitnav'}
# 基准取值列：先找明确的收盘/点位列，再找净值/指数/价格类；代码、名称这类标识列一律排除
_BENCH_PRIMARY = re.compile(r'(?i)(收盘|close|点位)')
_BENCH_SECONDARY = re.compile(r'(?i)(净值|nav|指数|价格|price|value)')
_BENCH_EXCLUDE = re.compile(r'(?i)(代码|code|名称|简称|name|ticker|symbol|证券|^id$|编号)')
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


def _norm(name) -> str:
    """列名归一：小写、去空白和括号里的单位，如“净值(元)”→“净值”、“Unit NAV”→“unitnav”。"""
    s = re.sub(r'[（(][^）)]*[）)]', '', str(name))
    return re.sub(r'[\s_\-.]', '', s).lower()


def _col_list(cols):
    return [str(c) for c in cols]


def _pick(cols, name, what):
    """按用户指定的列名取列（忽略首尾空白）。"""
    for c in cols:
        if str(c).strip() == str(name).strip():
            return c
    raise SkipFund(f'找不到指定的{what} "{name}"，可用列: {_col_list(cols)}')


def find_columns(df: pd.DataFrame, use_accum: bool = False, date_col=None, nav_col=None):
    """返回 (日期列, 净值列, 是否用了累计净值)。

    date_col / nav_col 为用户手动指定的列名，优先使用。
    自动识别时：use_accum=True 优先找累计净值列（含“累计/accum”和“净值/nav”），找不到回退单位净值。
    单位净值按优先级取：列名正好是“单位净值/净值/NAV”等 > 含“单位净值/unit nav” > 其他含“净值/nav”的列；
    复权净值、净值增长率、净值日期之类一律排除。
    """
    cols = df.columns.tolist()
    date_col = _pick(cols, date_col, '日期列') if date_col else _first_match(cols, _DATE_PAT, _DATE_EXCLUDE)

    if nav_col:
        nav_col = _pick(cols, nav_col, '净值列')
        used_accum = bool(_ACCUM_PAT.search(str(nav_col)))
    else:
        nav_col, used_accum = None, False
        usable = [c for c in cols if c != date_col and _NAV_PAT.search(str(c))
                  and not _NAV_EXCLUDE.search(str(c))]
        if use_accum:
            accum = [c for c in usable if _ACCUM_PAT.search(str(c))]
            if accum:
                nav_col, used_accum = accum[0], True
        if nav_col is None:
            unit = [c for c in usable if not _ACCUM_PAT.search(str(c))]

            def rank(c):
                n = _norm(c)
                return 0 if n in _UNIT_EXACT else 1 if ('单位净值' in n or 'unitnav' in n) else 2
            if unit:
                nav_col = min(unit, key=rank)  # min 在并列时保留靠前的列

    if date_col is None or nav_col is None:
        raise SkipFund(f'无法识别日期列或净值列，可用列: {_col_list(cols)}。'
                       '可用 --date-col / --nav-col 手动指定')
    return date_col, nav_col, used_accum


def parse_dates(s: pd.Series) -> pd.Series:
    """解析日期列：支持日期/文本日期、20240105 这类整数日期、Excel 日期序列号。"""
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    num = pd.to_numeric(s, errors='coerce')
    valid = num.notna()
    if valid.mean() > 0.8:
        v = num[valid]
        if v.between(19000101, 21001231).all():
            return pd.to_datetime(num.astype('Int64').astype(str), format='%Y%m%d', errors='coerce')
        if v.between(1, 80000).all():
            return pd.to_datetime(num, unit='D', origin='1899-12-30', errors='coerce')
    try:
        return pd.to_datetime(s, errors='coerce', format='mixed')
    except (TypeError, ValueError):
        return pd.to_datetime(s, errors='coerce')


def read_table(path: str, value_pat=None, date_name=None, value_name=None) -> pd.DataFrame:
    """读取 Excel，自动跳过表头上方的标题行：在前 10 行里找第一行同时有日期列名和数值列名
    （且不在同一个单元格）的行当表头；找不到就按第一行当表头。"""
    raw = pd.read_excel(path, header=None)
    for i in range(min(10, len(raw))):
        cells = [(j, str(v).strip()) for j, v in enumerate(raw.iloc[i]) if pd.notna(v)]
        if date_name:
            d_idx = {j for j, c in cells if c == str(date_name).strip()}
        else:
            d_idx = {j for j, c in cells if _DATE_PAT.search(c) and not _DATE_EXCLUDE.search(c)}
        if value_name:
            v_idx = {j for j, c in cells if c == str(value_name).strip()}
        elif value_pat is not None:
            v_idx = {j for j, c in cells if value_pat.search(c)}
        else:
            v_idx = {j for j, _ in cells}
        if d_idx and v_idx - d_idx:
            header = [str(v).strip() if pd.notna(v) else f'未命名{j}' for j, v in enumerate(raw.iloc[i])]
            df = raw.iloc[i + 1:].reset_index(drop=True)
            df.columns = header
            return df
    return pd.read_excel(path)


def _clean_series(df, date_col, value_col, value_name):
    out = pd.DataFrame({
        'date': parse_dates(df[date_col]).dt.normalize(),
        value_name: pd.to_numeric(df[value_col], errors='coerce'),
    })
    before = len(out)
    out = out.dropna()
    out = out[out[value_name] > 0]
    out = out.drop_duplicates('date', keep='last').sort_values('date').reset_index(drop=True)
    return out, before - len(out)


def load_nav(path: str, use_accum: bool = False, date_col=None, nav_col=None):
    """读取净值文件，返回 (DataFrame[date, nav], 备注列表, 实际使用的净值列名)。"""
    raw = read_table(path, _NAV_PAT, date_col, nav_col)
    date_col, nav_col, used_accum = find_columns(raw, use_accum, date_col, nav_col)
    df, dropped = _clean_series(raw, date_col, nav_col, 'nav')
    notes = []
    if use_accum and not used_accum:
        notes.append('未找到累计净值列，已使用单位净值')
    if dropped:
        notes.append(f'剔除 {dropped} 行空值/无效数据')
    if len(df) < 2:
        raise SkipFund(f'有效数据不足 2 行（日期列“{date_col}”、净值列“{nav_col}”，'
                       '请检查日期和净值格式）')
    return df, notes, str(nav_col)


def load_benchmark(path: str) -> pd.DataFrame:
    """读取基准文件，返回 DataFrame[date, bench]；实际使用的列名记在 df.attrs['value_col']。

    取值列只在“大部分是数字、且不是常数”的列里选（指数代码 000300 也是数字，但每行都一样），
    排除代码、名称等标识列；优先收盘/点位，其次净值/指数/价格，最后才取第一个可用数值列。
    """
    raw = read_table(path)
    cols = raw.columns.tolist()
    date_col = _first_match(cols, _DATE_PAT, _DATE_EXCLUDE)
    if date_col is None:
        raise ValueError(f'基准文件无法识别日期列，可用列: {_col_list(cols)}')

    def usable(c):
        if c == date_col or _BENCH_EXCLUDE.search(str(c)) or _RETURN_PAT.search(str(c)):
            return False
        v = pd.to_numeric(raw[c], errors='coerce')
        return v.notna().mean() > 0.8 and v.nunique() > 1

    candidates = [c for c in cols if usable(c)]
    value_col = (next((c for c in candidates if _BENCH_PRIMARY.search(str(c))), None)
                 or next((c for c in candidates if _BENCH_SECONDARY.search(str(c))), None)
                 or (candidates[0] if candidates else None))
    if value_col is None:
        raise ValueError(f'基准文件无法识别点位/净值列，可用列: {_col_list(cols)}')
    df, _ = _clean_series(raw, date_col, value_col, 'bench')
    if len(df) < 2:
        raise ValueError('基准文件有效数据不足 2 行')
    df.attrs['value_col'] = str(value_col)
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
    date_col: str | None = None
    nav_col: str | None = None


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
    nav_col: str = ''
    bench_freq: str | None = None

    @property
    def bench_factor(self) -> int | None:
        return FREQ_FACTORS[self.bench_freq] if self.bench_freq else None

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
    df, notes, nav_col = load_nav(path, opts.use_accum, opts.date_col, opts.nav_col)
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
            # 共同日期的频率可能比产品低（如日频产品 vs 周频基准），超额指标按共同日期自己的频率年化
            bench_freq = detect_freq(a['date'])
            excess = compute_excess(a, FREQ_FACTORS[bench_freq], opts.excess, opts.linear)
            if len(a) < len(df):
                notes.append(f'与基准共同日期 {len(a)}/{len(df)} 个，超额指标只用共同日期')
            if bench_freq != freq:
                notes.append(f'共同日期为{FREQ_LABELS[bench_freq]}，超额指标按{FREQ_LABELS[bench_freq]}年化')

    return FundResult(path, name, df, display_idx, freq, freq_auto, factor, metrics,
                      notes, aligned, excess, nav_col, bench_freq if aligned is not None else None)
