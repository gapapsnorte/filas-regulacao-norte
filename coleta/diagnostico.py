"""Diagnóstico: abre o Mapa Social e lista o que a página carregou (uso: python coleta/diagnostico.py)."""
import asyncio, os
from playwright.async_api import async_playwright
URL="https://paineis-ext.mpdft.mp.br/extensions/mapasauderegulamentacao/mapasauderegulamentacao.html"
async def main():
    async with async_playwright() as p:
        b=await p.chromium.launch(channel=os.environ.get("CANAL") or None, headless=True); pg=await b.new_page()
        reqs=[]
        pg.on("response", lambda r: reqs.append((r.status, r.url[:120])))
        r=await pg.goto(URL, wait_until="networkidle", timeout=120000)
        await pg.wait_for_timeout(5000)
        print("HTTP", r.status, pg.url, await pg.title(), "require:", await pg.evaluate("typeof require"))
        print((await pg.inner_text("body"))[:300])
        for s,u in reqs: print(s,u)
        await b.close()
asyncio.run(main())
