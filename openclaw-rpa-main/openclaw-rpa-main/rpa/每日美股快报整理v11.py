# pip install playwright && playwright install chromium
# 任务：每日美股快报整理V11
# 录制时间：2026-03-28 17:27:24
# 由 OpenClaw RPA Recorder（headed 真实录制）生成 — 可脱离 OpenClaw 独立运行

import asyncio
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

CONFIG = {
    "output_dir": Path.home() / "Desktop",
    "headless":   False,
    "timeout":    60_000,
    "slow_mo":    300,
}

_UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)


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
            # ── 步骤 1：打开雅虎财经香港
            try:
                await page.goto('https://hk.finance.yahoo.com', wait_until="domcontentloaded")
            except Exception:
                await page.screenshot(path="step_1_error.png")
                raise

            # ── 步骤 2：在搜索框输入 NVDA
            try:
                await page.locator('#ybar-sbq').first.fill('NVDA')
            except Exception:
                await page.screenshot(path="step_2_error.png")
                raise

            # ── 步骤 3：按回车执行搜索
            try:
                await page.keyboard.press('Enter')
                await page.wait_for_load_state("domcontentloaded")
            except Exception:
                await page.screenshot(path="step_3_error.png")
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
