import asyncio
from playwright.async_api import async_playwright

async def test():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        await page.set_content("""
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                body {
                    margin: 0;
                    width: 1920px;
                    height: 1080px;
                    background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%);
                    color: white;
                    font-family: 'Inter', system-ui, -apple-system, sans-serif;
                    padding: 80px;
                    box-sizing: border-box;
                    display: flex;
                    flex-direction: column;
                    justify-content: space-between;
                }
                .badge {
                    background: rgba(99, 102, 241, 0.2);
                    border: 1px solid #818cf8;
                    color: #a5b4fc;
                    padding: 12px 24px;
                    border-radius: 30px;
                    font-size: 18px;
                    font-weight: 700;
                    letter-spacing: 2px;
                    width: fit-content;
                }
                .title {
                    font-size: 64px;
                    font-weight: 800;
                    background: linear-gradient(90deg, #ffffff, #c7d2fe);
                    -webkit-background-clip: text;
                    -webkit-text-fill-color: transparent;
                    margin-top: 30px;
                    line-height: 1.2;
                }
                .cards {
                    display: grid;
                    grid-template-columns: repeat(3, 1fr);
                    gap: 30px;
                }
                .card {
                    background: rgba(255, 255, 255, 0.05);
                    border: 1px solid rgba(255, 255, 255, 0.1);
                    backdrop-filter: blur(10px);
                    border-radius: 24px;
                    padding: 40px;
                }
                .card-title {
                    font-size: 28px;
                    font-weight: 700;
                    color: #38bdf8;
                    margin-bottom: 16px;
                }
                .card-text {
                    font-size: 18px;
                    color: #94a3b8;
                    line-height: 1.6;
                }
            </style>
        </head>
        <body>
            <div>
                <div class="badge">GAMMA STYLE INFOGRAPHIC</div>
                <div class="title">Yangi O'zbekiston Taraqqiyot Strategiyasi</div>
            </div>
            <div class="cards">
                <div class="card">
                    <div class="card-title">01. Inson Qadri</div>
                    <div class="card-text">Inson - jamiyat - davlat tamoyili ustuvor etib belgilandi va ijtimoiy davlat maqomi mustahkamlandi.</div>
                </div>
                <div class="card">
                    <div class="card-title">02. Erkin Bozor</div>
                    <div class="card-text">Valyuta konvertatsiyasi, soliq imtiyozlari va xususiy mulk daxlsizligi qonuniy ta'minlandi.</div>
                </div>
                <div class="card">
                    <div class="card-title">03. Ochiq Diplomatiya</div>
                    <div class="card-text">Markaziy Osiyoda yaxshi qo'shnichilik hamda ko'p vektorli va pragmatik tashqi siyosat.</div>
                </div>
            </div>
        </body>
        </html>
        """)
        await page.screenshot(path=r"C:\Users\user\Desktop\SlideTranslate_AI\gamma_style_preview.png")
        await browser.close()
        print("PLAYWRIGHT_RENDER_SUCCESS")

asyncio.run(test())
