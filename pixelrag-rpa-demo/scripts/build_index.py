#!/usr/bin/env python3
"""构建案例索引

用法:
    python scripts/build_index.py --cases-dir ./cases --index-dir ./index
"""

import argparse
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.index_builder import CaseIndex


def main():
    parser = argparse.ArgumentParser(description="构建 RPA 故障案例索引")
    parser.add_argument("--cases-dir", default="./cases", help="案例库目录")
    parser.add_argument("--index-dir", default="./index", help="索引输出目录")
    parser.add_argument(
        "--mode",
        choices=["auto", "real", "mock"],
        default="auto",
        help="向量化模式（默认 auto，优先 real，失败回退 mock）",
    )
    args = parser.parse_args()

    idx = CaseIndex(
        cases_dir=args.cases_dir,
        index_dir=args.index_dir,
        embed_mode=args.mode,
    )
    idx.build()
    print("\n✅ 索引构建完成！")


if __name__ == "__main__":
    main()
