#!/usr/bin/env python3
"""搜索相似案例

用法:
    python scripts/search.py --image ./error.png --index-dir ./index --top-k 3
"""

import argparse
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.index_builder import CaseIndex
from src.case_manager import CaseManager


def main():
    parser = argparse.ArgumentParser(description="搜索相似 RPA 故障案例")
    parser.add_argument("--image", required=True, help="异常截图路径")
    parser.add_argument("--cases-dir", default="./cases", help="案例库目录")
    parser.add_argument("--index-dir", default="./index", help="索引目录")
    parser.add_argument("--top-k", type=int, default=3, help="返回前 N 个结果")
    args = parser.parse_args()

    idx = CaseIndex(cases_dir=args.cases_dir, index_dir=args.index_dir)
    cm = CaseManager(cases_dir=args.cases_dir)

    results = idx.search(args.image, top_k=args.top_k)

    print(f"\n搜索结果（Top-{args.top_k}）:")
    print("-" * 60)
    for r in results:
        case = cm.get_case(r["case_id"])
        desc = case.get("description", "") if case else ""
        print(f"  #{r['rank']}  {r['case_id']}  相似度: {r['score']:.4f}")
        if desc:
            print(f"         {desc}")
    print("-" * 60)

    if results:
        best = results[0]
        print(f"\n最匹配: {best['case_id']} (相似度: {best['score']:.4f})")


if __name__ == "__main__":
    main()
