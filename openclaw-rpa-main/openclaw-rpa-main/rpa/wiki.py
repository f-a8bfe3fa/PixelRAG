# pip install playwright && playwright install chromium
# 任务：wiki
# 录制时间：2026-03-31 12:06:23
# 由 OpenClaw RPA Recorder（headed 真实录制）生成 — 可脱离 OpenClaw 独立运行

import asyncio
from pathlib import Path
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
}

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


_EXTRACT_JS = (
    "([s,n])=>{return Array.from(document.querySelectorAll(s))"
    ".slice(0,n).map(e=>(e.textContent||'')"
    ".replace(/\\s+/g,' ').trim()).filter(Boolean)}"
)


async def _wait_for_content(page, selector: str) -> None:
    """等待 selector 对应的元素出现在 DOM 中（容错：超时也继续）。"""
    try:
        await page.wait_for_selector(selector, timeout=CONFIG["content_wait"])
    except Exception:
        pass  # 元素未出现也继续，evaluate 会返回空列表


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
            # ── 步骤 1：Visit Wikipedia
            try:
                await page.goto('https://www.wikipedia.org', wait_until="domcontentloaded")
                await page.wait_for_timeout(CONFIG["spa_settle_ms"])
            except Exception:
                await page.screenshot(path="step_1_error.png")
                raise

            # ── 步骤 2：Type Google into the search bar
            try:
                await page.locator('#searchInput').first.fill('Google')
            except Exception:
                await page.screenshot(path="step_2_error.png")
                raise

            # ── 步骤 3：Click search button
            try:
                await page.locator('#search-form fieldset button').first.click()
                await page.wait_for_load_state("domcontentloaded")
                await page.wait_for_timeout(800)
            except Exception:
                await page.screenshot(path="step_3_error.png")
                raise

            # ── 步骤 4：Extract the title
            try:
                _sel = '#firstHeading'
                _lim = 9999
                await _wait_for_content(page, _sel)
                _texts = await page.evaluate(_EXTRACT_JS, [_sel, _lim])
                _out = CONFIG["output_dir"] / 'wikipedia.txt'
                _field = 'Title'
                _block = _format_extract_section(_field, _texts)
                _out.write_text(_block, encoding="utf-8")
                print(f"已提取 {len(_texts)} 条，写入 {_out}（本文件首次写入）")
            except Exception:
                await page.screenshot(path="step_4_error.png")
                raise

            # ── 步骤 5：Extract the first 3 paragraphs of content
            try:
                _sel = '#mw-content-text p'
                _lim = 3
                await _wait_for_content(page, _sel)
                _texts = await page.evaluate(_EXTRACT_JS, [_sel, _lim])
                _out = CONFIG["output_dir"] / 'wikipedia.txt'
                _field = 'Content'
                _block = _format_extract_section(_field, _texts)
                with _out.open("a", encoding="utf-8") as _f:
                    _f.write(_block)
                print(f"已提取 {len(_texts)} 条，追加写入 {_out}")
            except Exception:
                await page.screenshot(path="step_5_error.png")
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
