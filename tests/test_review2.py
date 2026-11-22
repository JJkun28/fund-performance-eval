"""第二轮审阅发现的问题，每条一个回归测试。"""
import os

import numpy as np
import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

import fund_eval
from fpe_core import Options, evaluate, load_benchmark
from fund_eval import EXIT_LOCKED, EXIT_OK, SUMMARY_NAME, main

DAILY = pd.bdate_range('2024-01-01', periods=120)


def _nav(tmp_path, name, dates=DAILY, seed=0):
    rng = np.random.default_rng(seed)
    nav = np.round(np.cumprod(1 + rng.normal(0.0005, 0.01, len(dates))), 4)
    pd.DataFrame({'净值日期': dates, '单位净值': nav}).to_excel(tmp_path / name, index=False)
    return str(tmp_path / name)


def test_benchmark_skips_index_code_column(tmp_path, capsys):
    """表头为 日期/指数代码/收盘点位 时，用收盘点位，而不是每行都一样的指数代码。"""
    d = pd.date_range('2024-01-01', periods=5, freq='W-FRI')
    pd.DataFrame({'日期': d, '指数代码': [300] * 5, '收盘点位': [100, 102, 104, 106, 110]}) \
        .to_excel(tmp_path / 'b.xlsx', index=False)
    b = load_benchmark(str(tmp_path / 'b.xlsx'))
    assert b.attrs['value_col'] == '收盘点位'
    assert b['bench'].iloc[-1] / b['bench'].iloc[0] - 1 == pytest.approx(0.10)


@pytest.mark.parametrize('cols,expect', [
    (['日期', '证券简称', '指数', '涨跌幅'], '指数'),
    (['日期', 'code', 'close'], 'close'),
    (['日期', '代码', '数值'], '数值'),  # 没有关键词时取第一个可用数值列，代码列仍被排除
])
def test_benchmark_value_column_priority(tmp_path, cols, expect):
    d = pd.date_range('2024-01-01', periods=4, freq='D')
    data = {cols[0]: d}
    for c in cols[1:]:
        data[c] = ['沪深300'] * 4 if c == '证券简称' else [1, 2, 3, 5] if c != '涨跌幅' else [0.1] * 4
    if 'code' in cols or '代码' in cols:
        data[cols[1]] = [300, 300, 300, 300]
    pd.DataFrame(data).to_excel(tmp_path / 'b.xlsx', index=False)
    assert load_benchmark(str(tmp_path / 'b.xlsx')).attrs['value_col'] == expect


def test_product_named_like_summary_does_not_collide(tmp_path, capsys):
    _nav(tmp_path, '基金汇总.xlsx', seed=1)
    _nav(tmp_path, '甲1号.xlsx', seed=2)
    assert main([str(tmp_path)]) == EXIT_OK
    out = sorted(os.listdir(tmp_path / 'result'))
    assert out == ['基金汇总(2)绩效评估.xlsx', SUMMARY_NAME, '甲1号绩效评估.xlsx']
    assert load_workbook(tmp_path / 'result' / SUMMARY_NAME)['汇总'].max_row == 3  # 两只产品都在


def test_case_only_difference_does_not_collide(tmp_path):
    _nav(tmp_path, 'Alpha原始数据.xlsx', seed=1)
    _nav(tmp_path, 'alpha数据.xlsx', seed=2)
    assert main([str(tmp_path)]) == EXIT_OK
    files = [f for f in os.listdir(tmp_path / 'result') if f != SUMMARY_NAME]
    assert len(files) == 2 and len({f.casefold() for f in files}) == 2
    wb = load_workbook(tmp_path / 'result' / SUMMARY_NAME)
    assert len({n.casefold() for n in wb.sheetnames}) == len(wb.sheetnames) == 3


def test_tracking_error_uses_common_date_frequency(tmp_path):
    """日频产品对比周频基准：共同日期是周频，跟踪误差按 √52 年化。"""
    f = _nav(tmp_path, '甲1号.xlsx')
    weekly = pd.date_range('2024-01-05', periods=24, freq='W-FRI')
    pd.DataFrame({'日期': weekly, '收盘': np.linspace(3000, 3300, 24)}).to_excel(tmp_path / 'b.xlsx', index=False)
    r = evaluate(f, Options(), load_benchmark(str(tmp_path / 'b.xlsx')))
    assert r.freq == 'daily' and r.bench_freq == 'weekly' and r.bench_factor == 52
    a = r.aligned
    p, b = a['nav'] / a['nav'].iloc[0], a['bench'] / a['bench'].iloc[0]
    ex = (p / b).to_numpy()
    expect = np.std(ex[1:] / ex[:-1] - 1, ddof=1) * np.sqrt(52)
    assert r.excess['tracking_error'] == pytest.approx(expect)
    assert any('按周频年化' in n for n in r.notes)


def test_partial_write_is_reported(tmp_path, capsys, monkeypatch):
    """写到一半遇到文件被打开：退出码 3，并列出已经写入的文件。"""
    _nav(tmp_path, '甲1号.xlsx', seed=1)
    _nav(tmp_path, '乙2号.xlsx', seed=2)
    real_save, calls = Workbook.save, []

    def flaky_save(self, path):
        calls.append(path)
        if len(calls) == 2:
            raise PermissionError(13, 'locked', path)
        return real_save(self, path)

    monkeypatch.setattr(Workbook, 'save', flaky_save)
    assert main([str(tmp_path)]) == EXIT_LOCKED
    out = capsys.readouterr().out
    assert '只写入了部分文件' in out
    assert f'已写入: {calls[0]}' in out and f'被占用: {calls[1]}' in out
    assert '没有写入任何文件' not in out


def test_benchmark_column_is_printed(tmp_path, capsys):
    f = _nav(tmp_path, '甲1号.xlsx')
    pd.DataFrame({'日期': DAILY, '指数代码': [300] * len(DAILY), '收盘点位': np.linspace(1, 2, len(DAILY))}) \
        .to_excel(tmp_path / 'bench.xlsx', index=False)
    assert main([f, '--benchmark', str(tmp_path / 'bench.xlsx')]) == EXIT_OK
    assert '基准列: “收盘点位”' in capsys.readouterr().out


def test_module_exposes_summary_name():
    assert fund_eval.SUMMARY_NAME == '基金汇总绩效评估.xlsx'


def _mixed_batch(tmp_path):
    """甲是标准表头；乙的列名自动识别不了；丙有标题行、表头也标准。"""
    d = pd.date_range('2024-01-05', periods=30, freq='W-FRI')
    nav = np.round(np.linspace(1, 1.1, 30), 4)
    pd.DataFrame({'净值日期': d, '单位净值': nav}).to_excel(tmp_path / '甲1号.xlsx', index=False)
    pd.DataFrame({'时点': d, '价值A': nav * 1.02}).to_excel(tmp_path / '乙2号.xlsx', index=False)
    wb = Workbook()
    wb.active.append(['丙3号 净值表'])
    wb.active.append(['净值日期', '单位净值'])
    for x, v in zip(d, nav * 0.98):
        wb.active.append([x.to_pydatetime(), float(v)])
    wb.save(tmp_path / '丙3号.xlsx')


def test_batch_rerun_with_manual_columns_keeps_other_products(tmp_path, capsys):
    """按 SKILL.md 补跑：带上乙的列名重跑整个文件夹，甲、丙不能因此被跳过，汇总表三只都在。"""
    _mixed_batch(tmp_path)
    assert main([str(tmp_path)]) == EXIT_OK
    assert '跳过: 乙2号.xlsx' in capsys.readouterr().out

    assert main([str(tmp_path), '--date-col', '时点', '--nav-col', '价值A', '--overwrite']) == EXIT_OK
    out = capsys.readouterr().out
    assert '跳过' not in out
    assert '没有指定的日期列“时点”，已自动识别' in out and '净值列“价值A”' in out

    ws = load_workbook(tmp_path / 'result' / SUMMARY_NAME)['汇总']
    assert sorted(r[0].value for r in ws.iter_rows(min_row=2)) == ['丙3号', '乙2号', '甲1号']


def test_single_file_manual_column_still_strict(tmp_path, capsys):
    _mixed_batch(tmp_path)
    from fund_eval import EXIT_ERROR
    assert main([str(tmp_path / '甲1号.xlsx'), '--nav-col', '价值A']) == EXIT_ERROR
    assert '找不到指定的净值列' in capsys.readouterr().out
