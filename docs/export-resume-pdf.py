"""Export resume HTML → A4 PDF using Playwright page.pdf() with temp file fallback"""
import asyncio, os, tempfile
from playwright.async_api import async_playwright

BASE = r"C:\Users\64107\Desktop\charlotte-portfolio\docs"
FILES = [
    ("resume-en.html", "resume-en.pdf"),
    ("resume-cn.html", "resume-cn.pdf"),
]

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        for html_name, pdf_name in FILES:
            html_path = os.path.join(BASE, html_name)
            pdf_path = os.path.join(BASE, pdf_name)
            page = await browser.new_page(viewport={"width": 1280, "height": 720})
            await page.goto(f"file:///{html_path}", wait_until="networkidle")
            await page.wait_for_timeout(800)
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            tmp.close()
            await page.pdf(path=tmp.name, format="A4", print_background=True, margin={"top": "0", "bottom": "0", "left": "0", "right": "0"})
            os.replace(tmp.name, pdf_path)
            print(f"OK {html_name} -> {pdf_name}  ({round(os.path.getsize(pdf_path)/1024)}K)")
            await page.close()
        await browser.close()

asyncio.run(main())
