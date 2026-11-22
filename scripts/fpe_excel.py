"""把评估结果写成 Excel：数据区 + 全公式指标表 + 原生折线图。

只用 Excel 2007 起就有的函数（STDEV、SUMSQ、INDEX/MATCH、IFERROR 等），
不用 STDEV.S / MINIFS 这类需要 _xlfn 前缀的新函数，避免打开后出现 #NAME?。
"""
from __future__ import annotations

import math
import re

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, quote_sheetname

from fpe_core import FundResult, Options

# ---------------------------------------------------------------- 样式
BLUE = '2F5496'
F_HDR = Font(name='Arial', bold=True, color='FFFFFF', size=10)
F_DATA = Font(name='Arial', size=10)
F_LABEL = Font(name='Arial', bold=True, size=10)
F_VALUE = Font(name='Arial', bold=True, size=10, color=BLUE)
FILL_HDR = PatternFill('solid', fgColor=BLUE)
FILL_IND_HDR = PatternFill('solid', fgColor='D6E4F0')
FILL_EVEN = PatternFill('solid', fgColor='F2F2F2')
FILL_ODD = PatternFill('solid', fgColor='FFFFFF')
FILL_INPUT = PatternFill('solid', fgColor='FFF2CC')
_side = Side('thin', color='D9E2F3')
B_THIN = Border(left=_side, right=_side, top=_side, bottom=_side)
CENTER = Alignment(horizontal='center', vertical='center')
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

PCT, RATIO, DATE, NAV, DAYS, TEXT = '0.00%', '0.0000', 'yyyy-mm-dd', '0.0000', '0', '@'

# 主 sheet 指标表行号（H 列标签，I 列数值）
R = {k: i for i, k in enumerate([
    'interval', 'freq', 'factor', 'rf', 'method', 'interval_return', 'annual_return',
    'annual_vol', 'max_drawdown', 'sharpe', 'calmar', 'sortino', 'peak_nav', 'peak_date',
    'trough_date', 'recovery_date', 'recovery_days', 'notes'], start=2)}
# 基准 sheet 指标表行号（M 列标签，N 列数值）
BR = {k: i for i, k in enumerate([
    'interval', 'common_count', 'mode', 'bench_interval_return', 'bench_annual_return',
    'prod_annual_return', 'excess_interval_return', 'excess_annual_return',
    'tracking_error', 'information_ratio', 'excess_max_drawdown'], start=2)}


def _hdr(ws, row, col, text, fill=FILL_HDR, font=F_HDR):
    c = ws.cell(row=row, column=col, value=text)
    c.font, c.fill, c.alignment, c.border = font, fill, CENTER, B_THIN
    return c


def _data(ws, row, col, value, fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font, c.alignment, c.border = F_DATA, CENTER, B_THIN
    if fmt:
        c.number_format = fmt
    return c


def _ind(ws, row, lcol, label, value, fmt, input_cell=False):
    fill = FILL_EVEN if row % 2 == 0 else FILL_ODD
    lc = ws.cell(row=row, column=lcol, value=label)
    lc.font, lc.alignment, lc.border, lc.fill = F_LABEL, LEFT, B_THIN, fill
    vc = ws.cell(row=row, column=lcol + 1, value=value)
    vc.font, vc.alignment, vc.border, vc.number_format = F_VALUE, CENTER, B_THIN, fmt
    vc.fill = FILL_INPUT if input_cell else fill
    return vc


def _ann(growth: str, days: str, linear: bool) -> str:
    if linear:
        return f'(({growth})-1)*365/({days})'
    return f'({growth})^(365/({days}))-1'


def _line_chart(ws, title, cols, last_row, anchor, y_min=None, y_fmt='0.000'):
    ch = LineChart()
    ch.title = title
    ch.varyColors = False
    ch.height, ch.width = 10, 20
    ch.y_axis.number_format = y_fmt
    if y_min is not None:  # 不从 0 起，避免曲线被压扁
        ch.y_axis.scaling.min = y_min
    ch.x_axis.number_format = 'yyyy-mm'
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    for col in cols:
        ch.add_data(Reference(ws, min_col=col, min_row=1, max_row=last_row), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last_row))
    for s in ch.series:
        s.smooth = False
        s.marker.symbol = 'none'
        s.graphicalProperties.line.width = 15000
    if len(cols) == 1:
        ch.legend = None
    else:
        ch.legend.position = 'b'
        ch.legend.overlay = False  # 图例单独占位，不盖住曲线和坐标轴
    ws.add_chart(ch, anchor)


def _axis_floor(*series) -> float:
    lo = min(float(min(s)) for s in series)
    return max(0.0, math.floor(lo * 0.95 * 10) / 10)


def _widths(ws, widths: dict):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def safe_title(name: str, used: set, suffix: str = '') -> str:
    clean = re.sub(r"[\[\]:*?/\\']", '_', name)
    room = 31 - len(suffix)
    title, i = clean[:room] + suffix, 2
    while title in used:
        tag = f'({i})'
        title, i = clean[:room - len(tag)] + tag + suffix, i + 1
    used.add(title)
    return title


# ---------------------------------------------------------------- 主 sheet
def write_fund_sheet(ws, r: FundResult, opts: Options):
    """A 日期 | B 净值 | C 期收益率 | D 历史最高 | E 回撤 | F 下行收益 | G 修复辅助 ;  H-I 指标表"""
    df = r.df
    L = len(df) + 1
    for ci, h in enumerate(['日期', '净值', '期收益率', '历史最高净值', '回撤', '下行收益', '修复辅助'], 1):
        _hdr(ws, 1, ci, h)

    I = lambda k: f'$I${R[k]}'  # noqa: E731
    for i, row in enumerate(df.itertuples(index=False)):
        x = i + 2
        _data(ws, x, 1, row.date.to_pydatetime(), DATE)
        _data(ws, x, 2, float(row.nav), NAV)
        _data(ws, x, 3, None if i == 0 else f'=B{x}/B{x-1}-1', PCT)
        _data(ws, x, 4, f'=MAX(B$2:B{x})', NAV)
        _data(ws, x, 5, f'=B{x}/D{x}-1', PCT)
        _data(ws, x, 6, None if i == 0 else f'=MIN(C{x},0)', PCT)
        _data(ws, x, 7, f'=IF(AND(ISNUMBER({I("trough_date")}),A{x}>{I("trough_date")},'
                        f'B{x}>={I("peak_nav")}),A{x},"")', DATE)

    rng = lambda c, first=2: f'{c}{first}:{c}{L}'  # noqa: E731
    i_ = lambda k: f'I{R[k]}'  # noqa: E731
    na = lambda expr: f'=IFERROR({expr},"N/A")'  # noqa: E731
    nodd = lambda expr: f'=IF({i_("max_drawdown")}=0,"无回撤",{expr})'  # noqa: E731

    _hdr(ws, 1, 8, '指标', FILL_IND_HDR, Font(name='Arial', bold=True, size=10, color=BLUE))
    _hdr(ws, 1, 9, '数值', FILL_IND_HDR, Font(name='Arial', bold=True, size=10, color=BLUE))
    rows = [
        ('interval', '区间', r.interval_label, TEXT),
        ('freq', '数据频率', r.freq_label, TEXT),
        ('factor', '年化系数（可改）', r.factor, '0'),
        ('rf', '无风险利率（可改）', opts.rf, PCT),
        ('method', '年化口径', '线性' if opts.linear else '复利', TEXT),
        ('interval_return', '区间收益率', f'=B{L}/B2-1', PCT),
        ('annual_return', '年化收益率', '=' + _ann(f'B{L}/B2', f'A{L}-A2', opts.linear), PCT),
        ('annual_vol', '年化波动率', na(f'STDEV({rng("C", 3)})*SQRT({i_("factor")})'), PCT),
        ('max_drawdown', '最大回撤', f'=MIN({rng("E")})', PCT),
        ('sharpe', '夏普比率', na(f'({i_("annual_return")}-{i_("rf")})/{i_("annual_vol")}'), RATIO),
        ('calmar', '卡玛比率', na(f'{i_("annual_return")}/ABS({i_("max_drawdown")})'), RATIO),
        ('sortino', '索提诺比率', na(f'({i_("annual_return")}-{i_("rf")})/(SQRT(SUMSQ({rng("F", 3)})'
                                    f'/COUNT({rng("C", 3)}))*SQRT({i_("factor")}))'), RATIO),
        ('peak_nav', '回撤峰值净值', nodd(f'INDEX({rng("D")},MATCH({i_("max_drawdown")},{rng("E")},0))'), NAV),
        ('peak_date', '回撤峰值日', nodd(f'INDEX({rng("A")},MATCH({i_("peak_nav")},{rng("B")},0))'), DATE),
        ('trough_date', '回撤谷值日', nodd(f'INDEX({rng("A")},MATCH({i_("max_drawdown")},{rng("E")},0))'), DATE),
        ('recovery_date', '回撤修复日', nodd(f'IF(COUNT({rng("G")})=0,"未修复",MIN({rng("G")}))'), DATE),
        ('recovery_days', '修复天数（谷值→修复）',
         f'=IF(ISNUMBER({i_("recovery_date")}),{i_("recovery_date")}-{i_("trough_date")},'
         f'{i_("recovery_date")})', DAYS),
        ('notes', '备注', '；'.join(r.notes) or '—', TEXT),
    ]
    for key, label, val, fmt in rows:
        _ind(ws, R[key], 8, label, val, fmt, input_cell=key in ('factor', 'rf'))

    _line_chart(ws, f'{r.name} 净值走势', [2], L, 'K2', y_min=_axis_floor(df['nav']))
    _widths(ws, {'A': 12, 'B': 10, 'C': 10, 'D': 12, 'E': 10, 'F': 10, 'G': 11, 'H': 24, 'I': 22})
    ws.freeze_panes = 'A2'


# ---------------------------------------------------------------- 基准 sheet
def write_bench_sheet(ws, r: FundResult, opts: Options, main_title: str):
    """A 日期 | B 产品净值 | C 基准 | D 产品归一 | E 基准归一 | F 超额净值 | G 产品期收益 |
    H 基准期收益 | I 超额期收益 | J 超额历史最高 | K 超额回撤 ;  M-N 指标表"""
    a = r.aligned
    M = len(a) + 1
    geo = opts.excess == 'geometric'
    heads = ['日期', '产品净值', '基准', '产品(归一)', '基准(归一)', '超额净值',
             '产品期收益', '基准期收益', '超额期收益', '超额历史最高', '超额回撤']
    for ci, h in enumerate(heads, 1):
        _hdr(ws, 1, ci, h)
    for i, row in enumerate(a.itertuples(index=False)):
        x = i + 2
        _data(ws, x, 1, row.date.to_pydatetime(), DATE)
        _data(ws, x, 2, float(row.nav), NAV)
        _data(ws, x, 3, float(row.bench), NAV)
        _data(ws, x, 4, f'=B{x}/B$2', NAV)
        _data(ws, x, 5, f'=C{x}/C$2', NAV)
        _data(ws, x, 6, f'=D{x}/E{x}' if geo else f'=1+D{x}-E{x}', NAV)
        _data(ws, x, 7, None if i == 0 else f'=B{x}/B{x-1}-1', PCT)
        _data(ws, x, 8, None if i == 0 else f'=C{x}/C{x-1}-1', PCT)
        _data(ws, x, 9, None if i == 0 else (f'=F{x}/F{x-1}-1' if geo else f'=G{x}-H{x}'), PCT)
        _data(ws, x, 10, f'=MAX(F$2:F{x})', NAV)
        _data(ws, x, 11, f'=F{x}/J{x}-1', PCT)

    n_ = lambda k: f'N{BR[k]}'  # noqa: E731
    days = f'A{M}-A2'
    factor = f'{quote_sheetname(main_title)}!$I${R["factor"]}'
    ex_ann = ('=' + _ann(f'F{M}', days, opts.linear)) if geo else \
        f'={n_("prod_annual_return")}-{n_("bench_annual_return")}'
    d = a['date']
    _hdr(ws, 1, 13, '基准与超额', FILL_IND_HDR, Font(name='Arial', bold=True, size=10, color=BLUE))
    _hdr(ws, 1, 14, '数值', FILL_IND_HDR, Font(name='Arial', bold=True, size=10, color=BLUE))
    rows = [
        ('interval', '共同区间', f'{d.iloc[0]:%Y/%m/%d}-{d.iloc[-1]:%Y/%m/%d}', TEXT),
        ('common_count', '共同日期数', f'=COUNT(A2:A{M})', '0'),
        ('mode', '超额口径', '几何（产品÷基准）' if geo else '算术（产品−基准）', TEXT),
        ('bench_interval_return', '基准区间收益', f'=C{M}/C2-1', PCT),
        ('bench_annual_return', '基准年化收益', '=' + _ann(f'C{M}/C2', days, opts.linear), PCT),
        ('prod_annual_return', '产品同期年化', '=' + _ann(f'B{M}/B2', days, opts.linear), PCT),
        ('excess_interval_return', '超额区间收益', f'=F{M}-1', PCT),
        ('excess_annual_return', '超额年化收益', ex_ann, PCT),
        ('tracking_error', '跟踪误差', f'=IFERROR(STDEV(I3:I{M})*SQRT({factor}),"N/A")', PCT),
        ('information_ratio', '信息比率',
         f'=IFERROR({n_("excess_annual_return")}/{n_("tracking_error")},"N/A")', RATIO),
        ('excess_max_drawdown', '超额最大回撤', f'=MIN(K2:K{M})', PCT),
    ]
    for key, label, val, fmt in rows:
        _ind(ws, BR[key], 13, label, val, fmt)

    p = a['nav'] / a['nav'].iloc[0]
    b = a['bench'] / a['bench'].iloc[0]
    ex = p / b if geo else 1 + p - b
    _line_chart(ws, f'{r.name} 产品 vs 基准', [4, 5, 6], M, 'P2', y_min=_axis_floor(p, b, ex))
    _widths(ws, {get_column_letter(c): 11 for c in range(1, 12)})
    _widths(ws, {'M': 16, 'N': 20})
    ws.freeze_panes = 'A2'


# ---------------------------------------------------------------- 工作簿
def _add_fund(wb, r: FundResult, opts: Options, used: set, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = safe_title(r.name, used)
    write_fund_sheet(ws, r, opts)
    bench_title = None
    if r.aligned is not None:
        wb_ = wb.create_sheet()
        wb_.title = bench_title = safe_title(r.name, used, '-基准')
        write_bench_sheet(wb_, r, opts, ws.title)
    return ws.title, bench_title


def build_fund_workbook(r: FundResult, opts: Options) -> Workbook:
    wb = Workbook()
    _add_fund(wb, r, opts, set(), first=True)
    return wb


SUMMARY_MAIN = [
    ('区间', 'interval', TEXT), ('数据频率', 'freq', TEXT),
    ('区间收益率', 'interval_return', PCT), ('年化收益率', 'annual_return', PCT),
    ('年化波动率', 'annual_vol', PCT), ('最大回撤', 'max_drawdown', PCT),
    ('夏普比率', 'sharpe', RATIO), ('卡玛比率', 'calmar', RATIO), ('索提诺比率', 'sortino', RATIO),
    ('回撤峰值日', 'peak_date', DATE), ('回撤谷值日', 'trough_date', DATE),
    ('回撤修复日', 'recovery_date', DATE), ('修复天数', 'recovery_days', DAYS),
]
SUMMARY_BENCH = [
    ('基准区间收益', 'bench_interval_return', PCT), ('基准年化收益', 'bench_annual_return', PCT),
    ('超额区间收益', 'excess_interval_return', PCT), ('超额年化收益', 'excess_annual_return', PCT),
    ('跟踪误差', 'tracking_error', PCT), ('信息比率', 'information_ratio', RATIO),
    ('超额最大回撤', 'excess_max_drawdown', PCT),
]


def build_summary_workbook(results: list[FundResult], opts: Options) -> Workbook:
    """汇总 sheet 在最前，后面是每只产品的明细 sheet；汇总单元格全部引用明细 sheet。"""
    wb = Workbook()
    ws = wb.active
    ws.title = '汇总'
    used = {'汇总'}
    has_bench = any(r.aligned is not None for r in results)
    cols = SUMMARY_MAIN + (SUMMARY_BENCH if has_bench else []) + [('备注', 'notes', TEXT)]

    _hdr(ws, 1, 1, '产品名称')
    for ci, (h, _, _) in enumerate(cols, 2):
        _hdr(ws, 1, ci, h)

    for j, r in enumerate(results):
        main, bench = _add_fund(wb, r, opts, used)
        x = j + 2
        fill = PatternFill('solid', fgColor='EDF2F9') if j % 2 == 0 else FILL_ODD
        c = _data(ws, x, 1, r.name)
        c.fill = fill
        for ci, (_, key, fmt) in enumerate(cols, 2):
            if key in R:
                v = f'={quote_sheetname(main)}!I{R[key]}'
            elif bench:
                v = f'={quote_sheetname(bench)}!N{BR[key]}'
            else:
                v = '—'
            c = _data(ws, x, ci, v, 'General' if fmt == TEXT else fmt)
            c.fill = fill

    _widths(ws, {'A': 20, 'B': 23, 'C': 14})
    for ci in range(4, len(cols) + 2):
        ws.column_dimensions[get_column_letter(ci)].width = 13
    ws.column_dimensions[get_column_letter(len(cols) + 1)].width = 40
    ws.freeze_panes = 'B2'
    return wb
