"""把生成的 Excel 交给真实 Excel / LibreOffice 重算，核对公式结果与 Python 口径逐项一致，
并确认整本工作簿没有 #NAME? / #VALUE! 之类的错误值。没有重算引擎时自动跳过。"""
import datetime as dt

import pytest
from openpyxl import load_workbook

from fpe_core import Options, evaluate, load_benchmark
from fpe_excel import BR, R, SUMMARY_BENCH, SUMMARY_MAIN, build_fund_workbook, build_summary_workbook
from recalc import available_engine, recalc

pytestmark = pytest.mark.skipif(available_engine() is None, reason='没有 Excel 或 LibreOffice 可用于重算')

ERRORS = ('#NAME?', '#VALUE!', '#DIV/0!', '#REF!', '#N/A', '#NUM!')
NUM_KEYS = {'interval_return', 'annual_return', 'annual_vol', 'max_drawdown', 'sharpe', 'calmar',
            'sortino', 'peak_nav', 'recovery_days'}
DATE_KEYS = {'peak_date', 'trough_date', 'recovery_date'}


def expected(m: dict, key: str):
    v = m.get(key)
    if key in ('peak_nav', 'peak_date', 'trough_date', 'recovery_date', 'recovery_days') and v is None:
        return m['status']  # 无回撤 / 未修复
    return 'N/A' if v is None else v


def check(actual, exp, where):
    if isinstance(exp, str):
        assert actual == exp, where
    elif isinstance(exp, dt.datetime):  # pd.Timestamp 也是 datetime 子类
        assert actual == dt.datetime(exp.year, exp.month, exp.day), where
    else:
        assert actual == pytest.approx(float(exp), rel=1e-9, abs=1e-12), where


def no_error_cells(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for v in row:
                assert not (isinstance(v, str) and v in ERRORS), f'{ws.title} 出现 {v}'


SCENARIOS = {
    'daily_compound': (dict(freq='B', periods=260, seed=1), Options(), False),
    'weekly_linear_rf': (dict(freq='W-FRI', periods=120, seed=2), Options(linear=True, rf=0.02), False),
    'interval_slice': (dict(freq='B', periods=300, seed=3),
                       Options(start='2024-03-01', end='2024-09-30'), False),
    'no_drawdown': (dict(freq='B', periods=40, shape='up'), Options(), False),
    'unrecovered': (dict(freq='B', periods=120, seed=4, shape='crash_end'), Options(), False),
    'bench_geometric': (dict(freq='B', periods=250, seed=5), Options(), True),
    'bench_arith_weekly': (dict(freq='W-FRI', periods=60, seed=6),
                           Options(excess='arithmetic', linear=True), True),
}


@pytest.mark.parametrize('name', SCENARIOS)
def test_fund_workbook_formulas_match_python(name, make_file, bench_file, tmp_path):
    kw, opts, use_bench = SCENARIOS[name]
    bench = load_benchmark(bench_file()) if use_bench else None
    r = evaluate(make_file(**kw), opts, bench)
    src = tmp_path / f'{name}.xlsx'
    build_fund_workbook(r, opts).save(src)
    wb = load_workbook(recalc(str(src), str(tmp_path / 'calc')), data_only=True)
    no_error_cells(wb)

    ws = wb.worksheets[0]
    for key in NUM_KEYS | DATE_KEYS:
        check(ws[f'I{R[key]}'].value, expected(r.metrics, key), f'{name}:{key}')
    if use_bench:
        assert r.excess is not None
        wsb = wb.worksheets[1]
        for key, row in BR.items():
            if key in r.excess:
                check(wsb[f'N{row}'].value, expected(r.excess, key), f'{name}:{key}')
        assert wsb[f'N{BR["common_count"]}'].value == r.excess['common_count']


def test_summary_references_fund_sheets(make_file, bench_file, tmp_path):
    bench = load_benchmark(bench_file())
    opts = Options()
    rs = [evaluate(make_file(name=f'产品{i}号.xlsx', seed=i, freq='B', periods=200), opts, bench)
          for i in (1, 2)]
    src = tmp_path / 'summary.xlsx'
    build_summary_workbook(rs, opts).save(src)
    wb = load_workbook(recalc(str(src), str(tmp_path / 'calc')), data_only=True)
    no_error_cells(wb)
    ws = wb['汇总']
    headers = [c.value for c in ws[1]]
    for j, r in enumerate(rs):
        x = j + 2
        assert ws.cell(x, 1).value == r.name
        for label, key, _ in SUMMARY_MAIN:
            if key in NUM_KEYS | DATE_KEYS:
                check(ws.cell(x, headers.index(label) + 1).value, expected(r.metrics, key), f'汇总:{key}')
        for label, key, _ in SUMMARY_BENCH:
            check(ws.cell(x, headers.index(label) + 1).value, expected(r.excess, key), f'汇总:{key}')
