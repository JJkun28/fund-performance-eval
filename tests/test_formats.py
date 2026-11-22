"""真实数据里常见的格式问题（来自审阅实测），每条一个回归测试。"""
import os
import shutil
import subprocess
import sys

import pandas as pd
import pytest
from openpyxl import Workbook

from fpe_core import Options, SkipFund, evaluate, find_columns, load_benchmark, load_nav
from fund_eval import EXIT_ERROR, EXIT_OK, main, parse_rf

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATES = pd.date_range('2024-01-01', periods=6, freq='W-FRI')
UNIT = [1.0000, 1.0020, 0.9980, 1.0050, 1.0070, 1.0089]


def _write_rows(path, rows):
    wb = Workbook()
    for r in rows:
        wb.active.append(r)
    wb.save(path)
    return str(path)


def test_adjusted_nav_column_is_not_used(tmp_path):
    """复权单位净值排在单位净值前面时，仍然用单位净值。"""
    df = pd.DataFrame({'净值日期': DATES, '复权单位净值': [u * 1.1 for u in UNIT], '单位净值': UNIT})
    df.to_excel(tmp_path / '甲1号.xlsx', index=False)
    nav, _, col = load_nav(str(tmp_path / '甲1号.xlsx'))
    assert col == '单位净值'
    assert nav['nav'].iloc[-1] == pytest.approx(1.0089)


@pytest.mark.parametrize('name', ['净值(元)', '基金净值', '单位净值（元）', 'NAV', '产品净值'])
def test_common_nav_column_names(name):
    df = pd.DataFrame(columns=['日期', name, '净值增长率', '复权净值'])
    assert find_columns(df)[1] == name


def test_unit_nav_exact_name_beats_fuzzy():
    df = pd.DataFrame(columns=['日期', '母基金单位净值', '单位净值'])
    assert find_columns(df)[1] == '单位净值'


def test_integer_dates(tmp_path):
    ints = [int(d.strftime('%Y%m%d')) for d in DATES]
    pd.DataFrame({'净值日期': ints, '单位净值': UNIT}).to_excel(tmp_path / 'a.xlsx', index=False)
    nav, _, _ = load_nav(str(tmp_path / 'a.xlsx'))
    assert list(nav['date']) == list(DATES)


def test_excel_serial_dates(tmp_path):
    serial = [(d - pd.Timestamp('1899-12-30')).days for d in DATES]
    pd.DataFrame({'日期': serial, '单位净值': UNIT}).to_excel(tmp_path / 'a.xlsx', index=False)
    nav, _, _ = load_nav(str(tmp_path / 'a.xlsx'))
    assert list(nav['date']) == list(DATES)


def test_title_rows_above_header(tmp_path):
    rows = [['某某1号私募证券投资基金 单位净值表（截至日期 2024-02-09）'],  # 同一格里有“日期”和“净值”
            ['数据来源：托管行', None],
            ['净值日期', '单位净值', '累计净值']]
    rows += [[d.to_pydatetime(), u, u + 0.1] for d, u in zip(DATES, UNIT)]
    path = _write_rows(tmp_path / '某某1号.xlsx', rows)
    r = evaluate(path, Options())
    assert r.nav_col == '单位净值' and len(r.df) == 6


def test_benchmark_with_title_row_and_int_dates(tmp_path):
    rows = [['沪深300 指数行情'], ['日期', '收盘']]
    rows += [[int(d.strftime('%Y%m%d')), 3500 + i] for i, d in enumerate(DATES)]
    b = load_benchmark(_write_rows(tmp_path / 'b.xlsx', rows))
    assert list(b['date']) == list(DATES) and b['bench'].iloc[-1] == 3505


def test_manual_columns_via_cli(tmp_path, capsys):
    pd.DataFrame({'时点': DATES, '价值A': UNIT}).to_excel(tmp_path / '乙2号.xlsx', index=False)
    f = str(tmp_path / '乙2号.xlsx')
    assert main([f]) == EXIT_ERROR
    assert '可用列' in capsys.readouterr().out
    assert main([f, '--date-col', '时点', '--nav-col', '价值A']) == EXIT_OK
    assert main([f, '--nav-col', '不存在', '--overwrite']) == EXIT_ERROR
    assert '找不到指定的净值列' in capsys.readouterr().out


def test_manual_column_with_adjusted_nav(tmp_path):
    df = pd.DataFrame({'净值日期': DATES, '复权单位净值': [u * 1.1 for u in UNIT], '单位净值': UNIT})
    df.to_excel(tmp_path / '甲1号.xlsx', index=False)
    r = evaluate(str(tmp_path / '甲1号.xlsx'), Options(nav_col='复权单位净值'))
    assert r.df['nav'].iloc[-1] == pytest.approx(1.0089 * 1.1)


def test_stdout_is_utf8_without_env(tmp_path):
    """不设 PYTHONIOENCODING 时输出也是 UTF-8（Windows 默认 GBK 会让 agent 读到乱码）。"""
    pd.DataFrame({'净值日期': DATES, '单位净值': UNIT}).to_excel(tmp_path / '丙3号.xlsx', index=False)
    env = {k: v for k, v in os.environ.items() if k not in ('PYTHONIOENCODING', 'PYTHONUTF8')}
    p = subprocess.run([sys.executable, os.path.join(ROOT, 'scripts', 'fund_eval.py'),
                        str(tmp_path / '丙3号.xlsx')], capture_output=True, env=env, timeout=120)
    assert p.returncode == 0
    out = p.stdout.decode('utf-8')
    assert '已完成' in out and '丙3号' in out


def test_rf_warning_when_too_high(capsys):
    assert parse_rf('0.5') == 0.5
    assert '偏高' in capsys.readouterr().out
    assert parse_rf('2%') == pytest.approx(0.02)
    assert parse_rf('0.02') == 0.02
    assert '偏高' not in capsys.readouterr().out


def test_readme_example_command_runs(tmp_path):
    """README 里的示例命令在新克隆的仓库上直接能跑（不能因为提交了结果文件而退出码 2）。"""
    assert not os.path.exists(os.path.join(ROOT, 'examples', 'data', 'result'))
    shutil.copytree(os.path.join(ROOT, 'examples'), tmp_path / 'examples')
    ex = tmp_path / 'examples'
    assert main([str(ex / 'data'), '--benchmark', str(ex / 'benchmark' / '示例基准指数.xlsx')]) == EXIT_OK


def test_unreadable_columns_still_skip_cleanly():
    with pytest.raises(SkipFund, match='--date-col'):
        find_columns(pd.DataFrame(columns=['a', 'b']))
