import math

import numpy as np
import pandas as pd
import pytest

from fpe_core import (FundResult, Options, SkipFund, align_benchmark, compute_excess,  # noqa: F401
                      compute_metrics, detect_freq, evaluate, extract_prod_name, find_columns,
                      slice_interval)


def _df(dates, nav):
    return pd.DataFrame({'date': pd.to_datetime(dates), 'nav': nav})


KNOWN = _df(['2024-01-01', '2024-01-08', '2024-01-15', '2024-01-22', '2024-01-29'],
            [1.0, 1.1, 0.99, 1.045, 1.12])


def test_known_answer_weekly():
    m = compute_metrics(KNOWN, factor=52, rf=0.0)
    rets = np.array([0.1, -0.1, 1.045 / 0.99 - 1, 1.12 / 1.045 - 1])
    ann = 1.12 ** (365 / 28) - 1
    vol = math.sqrt(((rets - rets.mean()) ** 2).sum() / 3) * math.sqrt(52)
    down = math.sqrt(0.1 ** 2 / 4) * math.sqrt(52)

    assert m['interval_return'] == pytest.approx(0.12)
    assert m['annual_return'] == pytest.approx(ann)
    assert m['annual_vol'] == pytest.approx(vol)
    assert m['sharpe'] == pytest.approx(ann / vol)
    assert m['max_drawdown'] == pytest.approx(-0.1)
    assert m['calmar'] == pytest.approx(ann / 0.1)
    assert m['sortino'] == pytest.approx(ann / down)
    assert m['peak_date'] == pd.Timestamp('2024-01-08')
    assert m['trough_date'] == pd.Timestamp('2024-01-15')
    assert m['recovery_date'] == pd.Timestamp('2024-01-29')
    assert m['recovery_days'] == 14


def test_rf_and_linear():
    m = compute_metrics(KNOWN, factor=52, rf=0.02, linear=True)
    ann = 0.12 * 365 / 28
    assert m['annual_return'] == pytest.approx(ann)
    assert m['sharpe'] == pytest.approx((ann - 0.02) / m['annual_vol'])


def test_no_drawdown_and_unrecovered():
    up = _df(pd.date_range('2024-01-01', periods=10, freq='B'), np.cumprod(np.full(10, 1.01)))
    m = compute_metrics(up, 252)
    assert m['max_drawdown'] == 0 and m['status'] == '无回撤'
    assert m['calmar'] is None and m['sortino'] is None

    down = _df(pd.date_range('2024-01-01', periods=4, freq='B'), [1.0, 1.2, 1.1, 1.0])
    m = compute_metrics(down, 252)
    assert m['status'] == '未修复' and m['recovery_date'] is None
    assert m['peak_date'] == pd.Timestamp('2024-01-02')


@pytest.mark.parametrize('freq,expect', [('B', 'daily'), ('W-FRI', 'weekly'),
                                         ('ME', 'monthly'), ('QE', 'quarterly')])
def test_detect_freq(freq, expect):
    assert detect_freq(pd.Series(pd.date_range('2023-01-01', periods=30, freq=freq))) == expect


def test_slice_interval_takes_one_extra_row():
    df = _df(pd.date_range('2024-01-01', periods=10, freq='D'), np.linspace(1, 1.1, 10))
    out, idx = slice_interval(df, '2024-01-04', '2024-01-07')
    assert out['date'].iloc[0] == pd.Timestamp('2024-01-03') and idx == 1
    assert out['date'].iloc[-1] == pd.Timestamp('2024-01-07')
    out, idx = slice_interval(df, '2023-12-01', None)
    assert idx == 0 and len(out) == 10
    with pytest.raises(SkipFund):
        slice_interval(df, '2025-01-01', None)


@pytest.mark.parametrize('fname,expect', [
    ('大鹏9号原始数据.xlsx', '大鹏9号'),
    ('测试一号原始数据.xlsx', '测试一号'),
    ('某某信托-稳健3期数据.xlsx', '某某信托-稳健3期'),
    ('宁静成长净值数据.xlsx', '宁静成长'),
    ('大鹏9号绩效评估.xlsx', None),
])
def test_extract_prod_name(fname, expect):
    assert extract_prod_name(fname) == expect


def test_find_columns_accum_fallback():
    df = pd.DataFrame(columns=['成立日期', '净值日期', '单位净值', '累计收益率'])
    assert find_columns(df, use_accum=True) == ('净值日期', '单位净值', False)
    df = pd.DataFrame(columns=['Date', 'Unit NAV', 'Accumulated NAV'])
    assert find_columns(df, use_accum=True) == ('Date', 'Accumulated NAV', True)
    with pytest.raises(SkipFund):
        find_columns(pd.DataFrame(columns=['a', 'b']))


def test_excess_geometric_and_arithmetic():
    a = pd.DataFrame({'date': pd.to_datetime(['2024-01-01', '2024-01-08', '2024-01-15']),
                      'nav': [1.0, 1.1, 1.21], 'bench': [100.0, 100.0, 110.0]})
    g = compute_excess(a, 52, 'geometric')
    assert g['excess_interval_return'] == pytest.approx(0.1)
    assert g['tracking_error'] == pytest.approx(np.std([0.1, 0.0], ddof=1) * math.sqrt(52))
    ar = compute_excess(a, 52, 'arithmetic')
    assert ar['excess_interval_return'] == pytest.approx(0.11)
    assert ar['excess_annual_return'] == pytest.approx(ar['prod_annual_return'] - ar['bench_annual_return'])


def test_align_keeps_common_dates_only():
    p = _df(['2024-01-05', '2024-01-12', '2024-01-19'], [1, 1.01, 1.02])
    b = pd.DataFrame({'date': pd.to_datetime(['2024-01-05', '2024-01-08', '2024-01-19']),
                      'bench': [1, 2, 3]})
    assert list(align_benchmark(p, b)['date'].dt.day) == [5, 19]


def test_evaluate_product_metrics_unaffected_by_benchmark(make_file, bench_file):
    from fpe_core import load_benchmark
    f = make_file(freq='W-FRI', periods=80)
    plain = evaluate(f, Options())
    with_b = evaluate(f, Options(), load_benchmark(bench_file()))
    assert plain.freq == 'weekly' and plain.factor == 52
    assert plain.metrics == with_b.metrics
    assert with_b.excess is not None
