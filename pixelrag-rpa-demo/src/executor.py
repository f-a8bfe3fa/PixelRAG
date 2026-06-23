"""修复执行模块：根据 steps.json 执行像素级操作

支持的操作类型：
- click: 点击指定坐标
- type: 输入文本
- wait: 等待
- hotkey: 快捷键
- double_click: 双击
- right_click: 右键点击
"""

import json
import time
from pathlib import Path


class FixExecutor:
    """修复流程执行器"""

    def __init__(self, cases_dir: str, dry_run: bool = False):
        """
        Args:
            cases_dir: 案例库目录
            dry_run: 演示模式，不实际执行操作（无 GUI 环境时用）
        """
        self.cases_dir = Path(cases_dir)
        self.dry_run = dry_run
        self._pyautogui = None

    def _get_pyautogui(self):
        if self._pyautogui is None:
            import pyautogui
            pyautogui.FAILSAFE = True
            self._pyautogui = pyautogui
        return self._pyautogui

    def execute_case(self, case_id: str) -> bool:
        """执行指定案例的修复流程"""
        steps_path = self.cases_dir / case_id / "steps.json"
        if not steps_path.exists():
            print(f"未找到步骤文件: {steps_path}")
            return False

        with open(steps_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        steps = data.get("steps", [])
        description = data.get("description", case_id)

        print(f"开始执行修复: {description}")
        print(f"共 {len(steps)} 步")

        for i, step in enumerate(steps, 1):
            action = step.get("action", "")
            desc = step.get("description", f"步骤 {i}")

            print(f"  [{i}/{len(steps)}] {desc}")

            try:
                if action == "click":
                    self._do_click(step)
                elif action == "double_click":
                    self._do_double_click(step)
                elif action == "right_click":
                    self._do_right_click(step)
                elif action == "type":
                    self._do_type(step)
                elif action == "hotkey":
                    self._do_hotkey(step)
                elif action == "wait":
                    self._do_wait(step)
                else:
                    print(f"    ⚠️  未知操作类型: {action}，跳过")
            except Exception as e:
                print(f"    ❌ 执行失败: {e}")
                return False

        print("✅ 修复流程执行完成")
        return True

    def _do_click(self, step: dict):
        x, y = step.get("x", 0), step.get("y", 0)
        button = step.get("button", "left")
        if self.dry_run:
            print(f"    → [DRY] 点击 ({x}, {y}), 按钮={button}")
        else:
            pg = self._get_pyautogui()
            pg.click(x, y, button=button)

    def _do_double_click(self, step: dict):
        x, y = step.get("x", 0), step.get("y", 0)
        if self.dry_run:
            print(f"    → [DRY] 双击 ({x}, {y})")
        else:
            pg = self._get_pyautogui()
            pg.doubleClick(x, y)

    def _do_right_click(self, step: dict):
        x, y = step.get("x", 0), step.get("y", 0)
        if self.dry_run:
            print(f"    → [DRY] 右键点击 ({x}, {y})")
        else:
            pg = self._get_pyautogui()
            pg.rightClick(x, y)

    def _do_type(self, step: dict):
        text = step.get("text", "")
        x = step.get("x")
        y = step.get("y")
        if self.dry_run:
            if x and y:
                print(f"    → [DRY] 点击 ({x}, {y}) 后输入: {text}")
            else:
                print(f"    → [DRY] 输入文本: {text}")
        else:
            pg = self._get_pyautogui()
            if x and y:
                pg.click(x, y)
                time.sleep(0.2)
            pg.write(text)

    def _do_hotkey(self, step: dict):
        keys = step.get("keys", [])
        if self.dry_run:
            print(f"    → [DRY] 快捷键: {'+'.join(keys)}")
        else:
            pg = self._get_pyautogui()
            pg.hotkey(*keys)

    def _do_wait(self, step: dict):
        ms = step.get("ms", 1000)
        if self.dry_run:
            print(f"    → [DRY] 等待 {ms}ms")
        else:
            time.sleep(ms / 1000)
