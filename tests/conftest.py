import os
import sys

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'scripts'))


def make_nav(freq: str = 'B', periods: int = 300, seed: int = 0, start='2024-01-02',
             drift=0.0004, vol=0.008, shape=None):
    """生成一条合成净值序列。shape: None 随机游走 / 'up' 单调上涨 / 'crash_end' 尾部大跌不修复。"""
    dates = pd.date_range(start, periods=periods, freq=freq)
    rng = np.random.default_rng(seed)
    if shape == 'up':
        rets = np.full(periods, 0.001)
    else:
        rets = rng.normal(drift, vol, periods)
        if shape == 'crash_end':
            rets[-20:] = -0.01
    nav = np.round(np.cumprod(1 + rets), 4)
    return pd.DataFrame({'净值日期': dates, '单位净值': nav,
                         '累计净值': np.round(nav + 0.25, 4)})


def write_xlsx(df: pd.DataFrame, path) -> str:
    df.to_excel(path, index=False)
    return str(path)


@pytest.fixture
def make_file(tmp_path):
    def _make(name='测试1号原始数据.xlsx', **kw):
        return write_xlsx(make_nav(**kw), tmp_path / name)
    return _make


@pytest.fixture
def bench_file(tmp_path):
    def _make(freq='B', periods=320, seed=99, start='2023-12-20'):
        dates = pd.date_range(start, periods=periods, freq=freq)
        rng = np.random.default_rng(seed)
        close = np.round(3500 * np.cumprod(1 + rng.normal(0.0002, 0.01, periods)), 2)
        d = tmp_path / 'bench'
        d.mkdir(exist_ok=True)
        return write_xlsx(pd.DataFrame({'日期': dates, '收盘价': close}), d / '基准_沪深300.xlsx')
    return _make
