#!/usr/bin/env python3
"""执行修复流程

用法:
    python scripts/execute.py --case-id case_001 --cases-dir ./cases
    python scripts/execute.py --case-id case_001 --dry-run
"""

import argparse
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.executor import FixExecutor


def main():
    parser = argparse.ArgumentParser(description="执行 RPA 修复流程")
    parser.add_argument("--case-id", required=True, help="案例 ID")
    parser.add_argument("--cases-dir", default="./cases", help="案例库目录")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="演示模式（不实际操作鼠标键盘）",
    )
    args = parser.parse_args()

    executor = FixExecutor(cases_dir=args.cases_dir, dry_run=args.dry_run)
    success = executor.execute_case(args.case_id)

    if success:
        print("\n✅ 修复执行成功")
    else:
        print("\n❌ 修复执行失败")
        sys.exit(1)


if __name__ == "__main__":
    main()
