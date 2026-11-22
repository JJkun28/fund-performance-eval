"""基金绩效评估命令行入口（单文件 / 文件夹批量自动判断）。

用法:
  python fund_eval.py <源 .xlsx 或文件夹> [--start YYYY-MM-DD] [--end YYYY-MM-DD]
         [--use-accum] [--linear] [--freq auto|daily|weekly|monthly|quarterly]
         [--rf 0.02|2%] [--benchmark 基准.xlsx] [--excess geometric|arithmetic]
         [--date-col 列名] [--nav-col 列名] [--output-dir 目录] [--overwrite]

输出固定为 UTF-8。退出码:
  0  完成（个别文件跳过也算完成，原因会打印出来）
  1  出错：路径不存在、没有可评估的文件、基准文件读不了
  2  输出文件已存在且没加 --overwrite：什么都没写，先问用户是否覆盖
  3  输出文件正被 Excel 等程序打开：什么都没写，请用户关闭后重跑
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fpe_core import FREQ_FACTORS, Options, SkipFund, evaluate, load_benchmark  # noqa: E402
from fpe_excel import build_fund_workbook, build_summary_workbook  # noqa: E402

EXIT_OK, EXIT_ERROR, EXIT_EXISTS, EXIT_LOCKED = 0, 1, 2, 3
SUMMARY_NAME = '基金汇总绩效评估.xlsx'


def parse_rf(text: str) -> float:
    t = str(text).strip()
    if not t:
        return 0.0
    if t.endswith('%'):
        t = str(float(t[:-1]) / 100)
    v = float(t)
    if v >= 1:  # “2” 理解为 2%
        print(f'提示: 无风险利率 {t} 按百分比理解为 {v / 100:.2%}')
        v /= 100
    if v > 0.2:
        print(f'提示: 无风险利率为 {v:.2%}，明显偏高，请确认（2% 应写成 2% 或 0.02）')
    return v


def is_locked(path: str) -> bool:
    """文件是否正被 Excel / WPS / LibreOffice 打开。"""
    folder, name = os.path.split(path)
    for lock in (f'~${name}', f'.~lock.{name}#'):
        if os.path.exists(os.path.join(folder, lock)):
            return True
    if not os.path.exists(path):
        return False
    try:
        with open(path, 'r+b'):
            return False
    except PermissionError:
        return True


def collect_sources(src: str):
    """返回 (源文件列表, 默认输出目录, 是否批量)。"""
    if os.path.isdir(src):
        files = sorted(os.path.join(src, f) for f in os.listdir(src)
                       if f.lower().endswith('.xlsx') and not f.startswith(('~$', '.~lock'))
                       and '绩效评估' not in f)
        return files, os.path.join(src, 'result'), True
    return [src], os.path.dirname(os.path.abspath(src)), False


def build_parser():
    p = argparse.ArgumentParser(description='基金绩效评估')
    p.add_argument('source', help='源 .xlsx 文件，或包含多个 .xlsx 的文件夹')
    p.add_argument('--start', default='', help='开始日期 YYYY-MM-DD（含当天）')
    p.add_argument('--end', default='', help='结束日期 YYYY-MM-DD（含当天）')
    p.add_argument('--use-accum', action='store_true', help='用累计净值（找不到回退单位净值）')
    p.add_argument('--linear', action='store_true', help='年化收益率用线性口径（默认复利）')
    p.add_argument('--freq', default='auto', choices=['auto', *FREQ_FACTORS],
                   help='数据频率，决定年化系数 252/52/12/4，默认自动识别')
    p.add_argument('--rf', default='0', help='年化无风险利率，如 0.02 或 2%%，默认 0')
    p.add_argument('--benchmark', default='', help='基准 .xlsx（日期 + 点位/净值）')
    p.add_argument('--excess', default='geometric', choices=['geometric', 'arithmetic'],
                   help='超额口径，默认几何')
    p.add_argument('--date-col', default='', help='手动指定日期列名；批量时只用于有这一列的文件，其余自动识别')
    p.add_argument('--nav-col', default='', help='手动指定净值列名；批量时只用于有这一列的文件，其余自动识别')
    p.add_argument('--output-dir', default='', help='输出目录；默认单文件=源文件目录，批量=源文件夹\\result')
    p.add_argument('--overwrite', action='store_true', help='允许覆盖已存在的输出文件')
    return p


def main(argv=None) -> int:
    # 固定 UTF-8 输出：Windows 下默认是 GBK，agent 读到的会是乱码
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding='utf-8', errors='replace')
        except (AttributeError, ValueError):
            pass
    args = build_parser().parse_args(argv)

    if not os.path.exists(args.source):
        print(f'错误: 路径不存在: {args.source}')
        return EXIT_ERROR
    opts = Options(start=args.start or None, end=args.end or None, use_accum=args.use_accum,
                   linear=args.linear, freq=args.freq, rf=parse_rf(args.rf), excess=args.excess,
                   date_col=args.date_col or None, nav_col=args.nav_col or None)

    bench = None
    if args.benchmark:
        try:
            bench = load_benchmark(args.benchmark)
        except Exception as e:  # noqa: BLE001
            print(f'错误: 基准文件读取失败: {args.benchmark} ({e})')
            return EXIT_ERROR
        print(f'基准列: “{bench.attrs["value_col"]}”（{os.path.basename(args.benchmark)}）')

    files, out_dir, batch = collect_sources(args.source)
    opts.strict_cols = not batch  # 批量时指定列只用于有这一列的文件
    if args.benchmark:  # 基准文件放在批量文件夹里时不当作产品
        files = [f for f in files
                 if os.path.normcase(os.path.abspath(f)) != os.path.normcase(os.path.abspath(args.benchmark))]
    out_dir = os.path.abspath(args.output_dir or out_dir)
    if not files:
        print(f'错误: {args.source} 中没有可评估的 .xlsx 文件')
        return EXIT_ERROR

    # 1) 先在内存里算完全部产品
    # 输出文件名按不区分大小写去重（Windows/macOS 文件系统如此），批量时预留汇总表的文件名
    taken = {SUMMARY_NAME.casefold()} if batch else set()
    results, skipped = [], []
    for f in files:
        try:
            r = evaluate(f, opts, bench)
        except SkipFund as e:
            skipped.append((f, str(e)))
            continue
        except Exception as e:  # noqa: BLE001
            skipped.append((f, f'读取或计算出错: {e}'))
            continue
        base, i = r.name, 2
        while f'{r.name}绩效评估.xlsx'.casefold() in taken:
            r.name, i = f'{base}({i})', i + 1
        if r.name != base:
            r.notes.append(f'与其他输出文件重名，已改名为“{r.name}”')
        taken.add(f'{r.name}绩效评估.xlsx'.casefold())
        results.append(r)

    if not results:
        print('错误: 没有可评估的产品')
        for f, why in skipped:
            print(f'跳过: {os.path.basename(f)} -- {why}')
        return EXIT_ERROR

    # 2) 写之前统一检查全部输出，有问题就一个文件都不写
    planned = [(r, os.path.join(out_dir, f'{r.name}绩效评估.xlsx')) for r in results]
    outputs = [p for _, p in planned] + ([os.path.join(out_dir, SUMMARY_NAME)] if batch else [])
    locked = [p for p in outputs if is_locked(p)]
    if locked:
        print('输出文件正被打开，无法写入。请关闭以下文件后重新运行（本次没有写入任何文件）:')
        for p in locked:
            print(f'被占用: {p}')
        return EXIT_LOCKED
    existing = [p for p in outputs if os.path.exists(p)]
    if existing and not args.overwrite:
        print('以下输出文件已存在（本次没有写入任何文件）。确认覆盖请加 --overwrite 重新运行:')
        for p in existing:
            print(f'已存在: {p}')
        return EXIT_EXISTS

    # 3) 写入
    os.makedirs(out_dir, exist_ok=True)
    jobs = [(build_fund_workbook(r, opts), p) for r, p in planned]
    if batch:
        jobs.append((build_summary_workbook(results, opts), os.path.join(out_dir, SUMMARY_NAME)))
    written = []
    for wb, p in jobs:
        try:
            wb.save(p)
        except PermissionError:
            # 写入前的检查和保存之间有时间差，这时前面的文件已经落盘，必须如实说明
            print('写入中途遇到文件被打开，本次只写入了部分文件。')
            print(f'被占用: {p}')
            for w in written:
                print(f'已写入: {w}')
            if not written:
                print('（本次没有写入任何文件）')
            print('请关闭被占用的文件后，加 --overwrite 重新运行整批。')
            return EXIT_LOCKED
        written.append(p)
        print(('已覆盖: ' if p in existing else '已完成: ') + p)

    for r in results:
        m = r.metrics
        tag = f'  {r.name}: 区间收益 {m["interval_return"]:.2%}，年化 {m["annual_return"]:.2%}，' \
              f'最大回撤 {m["max_drawdown"]:.2%}，{r.freq_label}，净值列“{r.nav_col}”'
        if r.notes:
            tag += f'；备注: {"；".join(r.notes)}'
        print(tag)
    for f, why in skipped:
        print(f'跳过: {os.path.basename(f)} -- {why}')
    print(f'共评估 {len(results)}/{len(files)} 个产品，输出目录: {out_dir}')
    return EXIT_OK


if __name__ == '__main__':
    sys.exit(main())
