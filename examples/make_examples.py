"""生成示例数据（合成数据，非真实产品）：两只周频产品 + 一个日频基准指数。

python examples/make_examples.py
python scripts/fund_eval.py examples/data --benchmark examples/benchmark/示例基准指数.xlsx
"""
import os

import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))


def walk(dates, drift, vol, seed):
    rng = np.random.default_rng(seed)
    return np.round(np.cumprod(1 + rng.normal(drift, vol, len(dates))), 4)


def main():
    os.makedirs(os.path.join(HERE, 'data'), exist_ok=True)
    os.makedirs(os.path.join(HERE, 'benchmark'), exist_ok=True)
    weekly = pd.bdate_range('2023-01-06', '2025-06-27', freq='W-FRI')
    for name, drift, vol, seed in [('示例成长1号', 0.0030, 0.022, 9), ('示例稳健2号', 0.0012, 0.008, 3)]:
        nav = walk(weekly, drift, vol, seed)
        pd.DataFrame({'净值日期': weekly, '单位净值': nav, '累计净值': np.round(nav + 0.12, 4)}) \
            .to_excel(os.path.join(HERE, 'data', f'{name}原始数据.xlsx'), index=False)
    daily = pd.bdate_range('2022-12-30', '2025-06-30')
    pd.DataFrame({'日期': daily, '收盘点位': np.round(3800 * walk(daily, 0.0002, 0.011, 8), 2)}) \
        .to_excel(os.path.join(HERE, 'benchmark', '示例基准指数.xlsx'), index=False)
    print('示例数据已生成:', HERE)


if __name__ == '__main__':
    main()
