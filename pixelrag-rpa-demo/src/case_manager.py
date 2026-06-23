"""案例管理模块：读取/新增/更新案例"""

import json
import shutil
from pathlib import Path


class CaseManager:
    """案例库管理器"""

    def __init__(self, cases_dir: str):
        self.cases_dir = Path(cases_dir)
        self.cases_dir.mkdir(parents=True, exist_ok=True)

    def list_cases(self) -> list[str]:
        """列出所有案例 ID"""
        return sorted([
            d.name for d in self.cases_dir.iterdir()
            if d.is_dir() and (d / "meta.json").exists()
        ])

    def get_case(self, case_id: str) -> dict | None:
        """获取案例详情"""
        case_dir = self.cases_dir / case_id
        meta_path = case_dir / "meta.json"
        if not meta_path.exists():
            return None

        with open(meta_path, "r", encoding="utf-8") as f:
            meta = json.load(f)

        steps_path = case_dir / "steps.json"
        if steps_path.exists():
            with open(steps_path, "r", encoding="utf-8") as f:
                meta["steps"] = json.load(f)

        return meta

    def add_case(
        self,
        case_id: str,
        before_image: str,
        after_image: str,
        steps: list[dict],
        description: str = "",
        error_type: str = "",
    ) -> str:
        """新增案例

        Args:
            case_id: 案例 ID
            before_image: 异常截图路径
            after_image: 修复后截图路径
            steps: 操作步骤列表
            description: 描述
            error_type: 错误类型

        Returns:
            案例目录路径
        """
        case_dir = self.cases_dir / case_id
        case_dir.mkdir(parents=True, exist_ok=True)

        # 复制截图
        shutil.copy(before_image, case_dir / "before.png")
        shutil.copy(after_image, case_dir / "after.png")

        # 保存步骤
        steps_data = {
            "case_id": case_id,
            "description": description,
            "error_type": error_type,
            "steps": steps,
        }
        with open(case_dir / "steps.json", "w", encoding="utf-8") as f:
            json.dump(steps_data, f, indent=2, ensure_ascii=False)

        # 保存元数据
        meta = {
            "case_id": case_id,
            "description": description,
            "error_type": error_type,
            "created_at": "",  # 可以加时间
            "used_count": 0,
            "success_count": 0,
        }
        with open(case_dir / "meta.json", "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2, ensure_ascii=False)

        print(f"案例已添加: {case_dir}")
        return str(case_dir)

    def get_steps(self, case_id: str) -> list[dict] | None:
        """获取案例的操作步骤"""
        steps_path = self.cases_dir / case_id / "steps.json"
        if not steps_path.exists():
            return None
        with open(steps_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("steps", [])

    def increment_usage(self, case_id: str, success: bool = True):
        """更新案例使用计数"""
        meta_path = self.cases_dir / case_id / "meta.json"
        if not meta_path.exists():
            return
        with open(meta_path, "r", encoding="utf-8") as f:
            meta = json.load(f)
        meta["used_count"] = meta.get("used_count", 0) + 1
        if success:
            meta["success_count"] = meta.get("success_count", 0) + 1
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2, ensure_ascii=False)
