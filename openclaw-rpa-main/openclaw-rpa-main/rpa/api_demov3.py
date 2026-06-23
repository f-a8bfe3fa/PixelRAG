# pip install playwright httpx && playwright install chromium
# 任务：api_demoV3
# 录制时间：2026-04-06 11:30:18
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
    # --- API 密钥（从环境变量读取；个人使用时也可直接在此填写）
    "ALPHAVANTAGE_API_KEY": os.environ.get("ALPHAVANTAGE_API_KEY", ""),
}

_missing = [v for v in ['ALPHAVANTAGE_API_KEY'] if not CONFIG.get(v)]
if _missing:
    print("\n❌  以下 API 密钥未配置，脚本无法运行：")
    for _v in _missing:
        print(f"   {_v} 为空  →  请在终端执行：  export {_v}='你的密钥'")
        print(f"   （或直接在脚本的 CONFIG 字典里填写 {_v} 的值）")
    raise SystemExit(1)

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
        page = await context.new_page()
        page.set_default_timeout(CONFIG["timeout"])

        try:
            # ── 步骤 1：拉取 NVDA 股票的日线数据
            try:
                _params = {'function': 'TIME_SERIES_DAILY', 'symbol': 'NVDA', 'apikey': CONFIG["ALPHAVANTAGE_API_KEY"]}
                _api_url = 'https://www.alphavantage.co/query' + "?" + urllib.parse.urlencode(_params)
                async with httpx.AsyncClient(timeout=CONFIG["api_timeout"]) as _hc:
                    _r = await _hc.get(_api_url)
                    _r.raise_for_status()
                (CONFIG["output_dir"] / 'nvda_time_series_daily.json').write_text(_r.text, encoding="utf-8")
                print("API 响应已写入", CONFIG["output_dir"] / 'nvda_time_series_daily.json')
            except Exception:
                await page.screenshot(path="step_1_error.png")
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
