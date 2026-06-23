# pip install playwright httpx openpyxl python-docx && playwright install chromium
# 任务：自动登录V3
# 录制时间：2026-04-07 15:25:35
# 由 OpenClaw RPA Recorder（headed 真实录制）生成 — 可脱离 OpenClaw 独立运行

import asyncio
import os
import urllib.parse
from pathlib import Path

import httpx
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

CONFIG = {
    "output_dir":    Path.home() / "Desktop",
    "headless":      False,
    "timeout":       60_000,
    "slow_mo":       300,
    # 导航后等待 SPA 内容渲染的额外时间（重型 SPA 如 Yahoo Finance 需要 1-2 秒）
    "spa_settle_ms": 1_500,
    # extract_text 等待目标元素出现的超时（毫秒）
    "content_wait":  15_000,
    # httpx 调用外部 API 的超时（秒）
    "api_timeout":   60.0,
    # 已保存的登录 Cookie 路径（由 #rpa-login 生成；留空则不注入）
    "cookies_path":  '/Users/jiangzhiwei/.openclaw/rpa/sessions/saucedemo.com/cookies.json',}

_UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)

def _format_extract_section(field_label: str, lines: list[str]) -> str:
    """Format extracted DOM text: show field name, separator line, then body."""
    name = (field_label or "").strip() or "extract"
    title = f"【字段：{name}】"
    if not lines:
        body = "(no text matched)\n"
    elif len(lines) == 1:
        body = lines[0].strip() + "\n"
    else:
        parts = [f"{i + 1}. {s.strip()}" for i, s in enumerate(lines)]
        body = "\n\n".join(parts) + "\n"
    sep = "─" * 32
    return f"{title}\n{sep}\n{body}\n"


_EXTRACT_JS = '([s,n])=>{const r=document.querySelector("main")||document.querySelector(\'[role="main"]\');const bare=/^[a-zA-Z][a-zA-Z0-9-]*$/.test(s)&&s.indexOf("#")<0&&s.indexOf(".")<0&&s.indexOf("[")<0&&s.indexOf(" ")<0;const sc=bare&&r?r:document;return Array.from(sc.querySelectorAll(s)).slice(0,n).map(e=>(e.textContent||"").replace(/\\s+/g," ").trim()).filter(Boolean)}'


async def _wait_for_content(page, selector: str) -> None:
    """等待 selector 对应的元素出现在 DOM 中（容错：超时也继续）。"""
    try:
        await page.wait_for_selector(selector, timeout=CONFIG["content_wait"])
    except Exception:
        pass  # 元素未出现也继续，evaluate 会返回空列表


async def _scroll_window(page, dy: int) -> None:
    """窗口滚动：导航后若再用 evaluate(scrollBy)，易因执行上下文销毁报错；用 mouse.wheel 并在滚动前等待页面稳定。"""
    try:
        await page.wait_for_load_state("domcontentloaded", timeout=10_000)
    except Exception:
        pass
    vp = page.viewport_size
    if vp:
        await page.mouse.move(vp["width"] // 2, vp["height"] // 2)
    else:
        await page.mouse.move(720, 450)
    await page.mouse.wheel(0, float(dy))


async def run():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=CONFIG["headless"],
            args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
            slow_mo=CONFIG["slow_mo"],
        )
        context = await browser.new_context(
            user_agent=_UA,
            viewport={"width": 1440, "height": 900},
            locale="en-US",
            timezone_id="America/New_York",
            extra_http_headers={"Accept-Language": "en-US,en;q=0.9"},
        )
        await context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        # 若配置了 cookies_path，注入 Cookie 模拟已登录态（由 rpa_manager login-start 生成）
        import json as _json
        _cp = CONFIG.get("cookies_path", "")
        if _cp and Path(_cp).exists():
            await context.add_cookies(_json.loads(Path(_cp).read_text()))
        page = await context.new_page()
        page.set_default_timeout(CONFIG["timeout"])

        try:
            # ── 步骤 1：访问商品列表页
            try:
                await page.goto('https://www.saucedemo.com/inventory.html', wait_until="domcontentloaded")
                await page.wait_for_timeout(CONFIG["spa_settle_ms"])
            except Exception:
                await page.screenshot(path="step_1_error.png")
                raise

            # ── 步骤 2：按价格从高到低排序
            try:
                await page.locator('#header_container div div span select').first.select_option('hilo')
                await page.wait_for_load_state("domcontentloaded")
                await page.wait_for_timeout(800)
            except Exception:
                await page.screenshot(path="step_2_error.png")
                raise

        except PlaywrightTimeout as e:
            await page.screenshot(path="error_timeout.png")
            raise RuntimeError(f"超时：{e}") from e
        except Exception:
            await page.screenshot(path="error_unexpected.png")
            raise
        finally:
            await browser.close()


if __name__ == "__main__":
    asyncio.run(run())
