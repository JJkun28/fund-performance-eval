import os

from openpyxl import load_workbook

from fund_eval import EXIT_ERROR, EXIT_EXISTS, EXIT_LOCKED, EXIT_OK, SUMMARY_NAME, main


def test_single_file_then_exists_then_overwrite(make_file, tmp_path, capsys):
    f = make_file('大鹏9号原始数据.xlsx', freq='W-FRI', periods=60)
    out = tmp_path / '大鹏9号绩效评估.xlsx'
    assert main([f]) == EXIT_OK and out.exists()
    mtime = out.stat().st_mtime_ns

    assert main([f]) == EXIT_EXISTS
    assert '已存在:' in capsys.readouterr().out
    assert out.stat().st_mtime_ns == mtime  # 没被改动

    assert main([f, '--overwrite']) == EXIT_OK
    assert '已覆盖:' in capsys.readouterr().out


def test_locked_output_blocks_even_with_overwrite(make_file, tmp_path, capsys):
    f = make_file('大鹏9号原始数据.xlsx')
    assert main([f]) == EXIT_OK
    (tmp_path / '~$大鹏9号绩效评估.xlsx').write_bytes(b'')  # 模拟 Excel 打开时的锁文件
    assert main([f, '--overwrite']) == EXIT_LOCKED
    assert '被占用:' in capsys.readouterr().out


def test_batch_prechecks_all_outputs_before_writing(make_file, tmp_path):
    make_file('甲1号.xlsx', seed=1)
    make_file('乙2号.xlsx', seed=2)
    res = tmp_path / 'result'
    res.mkdir()
    (res / f'~${SUMMARY_NAME}').write_bytes(b'')  # 只有汇总表被占用
    assert main([str(tmp_path)]) == EXIT_LOCKED
    assert sorted(os.listdir(res)) == [f'~${SUMMARY_NAME}']  # 一个产品文件都没写


def test_batch_skips_bad_files_and_ignores_benchmark(make_file, bench_file, tmp_path, capsys):
    make_file('甲1号.xlsx', seed=1)
    make_file('乙2号.xlsx', seed=2)
    (tmp_path / '坏文件.xlsx').write_bytes(b'not an xlsx')
    b = bench_file()
    os.replace(b, tmp_path / '基准.xlsx')  # 基准放进批量文件夹
    assert main([str(tmp_path), '--benchmark', str(tmp_path / '基准.xlsx'), '--freq', 'daily']) == EXIT_OK
    out = capsys.readouterr().out
    assert '跳过: 坏文件.xlsx' in out and '基准.xlsx --' not in out

    wb = load_workbook(tmp_path / 'result' / SUMMARY_NAME)
    assert wb.sheetnames == ['汇总', '乙2号', '乙2号-基准', '甲1号', '甲1号-基准']  # 按文件名排序
    assert wb['甲1号']['I3'].value == '日频（手动指定）'


def test_missing_path_and_bad_benchmark(make_file, tmp_path):
    assert main([str(tmp_path / 'nope.xlsx')]) == EXIT_ERROR
    f = make_file()
    assert main([f, '--benchmark', str(tmp_path / 'nope.xlsx')]) == EXIT_ERROR


def test_rf_parsing():
    from fund_eval import parse_rf
    assert parse_rf('2%') == 0.02
    assert parse_rf('0.015') == 0.015
    assert parse_rf('2') == 0.02
    assert parse_rf('') == 0.0
