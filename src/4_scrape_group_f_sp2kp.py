"""
Scraper Tabulasi Harga SP2KP (Kemendag)
Sumber: https://sp2kp.kemendag.go.id/statistik/tabulasi-harga

Cara kerja:
1. Buka halaman SP2KP pakai Playwright (headless browser), biar proses
   Trusted Authentication ke Tableau Server jalan otomatis (dapat cookie + ticket).
2. Cari iframe Tableau yang ke-embed di halaman itu.
3. Klik tombol download bawaan Tableau -> pilih Crosstab -> pilih format CSV.
4. Tangkep file yang di-download, simpan ke path tujuan.

Requirement:
    pip install playwright --break-system-packages
    playwright install chromium
"""

import asyncio
import sys
from pathlib import Path
from playwright.async_api import async_playwright

SP2KP_URL = "https://sp2kp.kemendag.go.id/statistik/tabulasi-harga"
OUTPUT_PATH = Path("../data/raw/groupF/9_ekonomi_harga_pangan_harian.csv")

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
)


async def scrape_tabulasi_harga(output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent=USER_AGENT,
            accept_downloads=True,
        )
        page = await context.new_page()

        await page.goto(SP2KP_URL, wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_timeout(
            10000
        )  # kasih waktu Tableau bootstrap + trusted auth selesai

        # cari iframe Tableau yang ke-embed di halaman (bukan main frame)
        target_frame = None
        for f in page.frames:
            if "analitik.kemendag.go.id/views/" in f.url:
                target_frame = f
                break

        if target_frame is None:
            await browser.close()
            raise RuntimeError("Frame Tableau tidak ditemukan di halaman SP2KP")

        # klik tombol download di toolbar Tableau
        await target_frame.locator("#download").click()
        await page.wait_for_timeout(1500)

        # pilih opsi "Crosstab" di menu yang muncul
        await target_frame.locator("text=Crosstab").first.click()
        await page.wait_for_timeout(1500)

        # pilih format CSV di dialog "Download Crosstab"
        csv_radio = target_frame.locator("text=CSV").first
        if await csv_radio.count() > 0:
            await csv_radio.click()

        # klik tombol Download final, tangkep file yang ter-download
        async with page.expect_download(timeout=20000) as download_info:
            await target_frame.locator('button:has-text("Download")').last.click()

        download = await download_info.value
        await download.save_as(str(output_path))

        await browser.close()

    return output_path


def verify_output(path: Path) -> None:
    import pandas as pd

    df = pd.read_csv(path, sep="\t", encoding="utf-16")
    print(f"Shape: {df.shape}")
    print(f"10 kolom tanggal terakhir: {df.columns.tolist()[-10:]}")


def main() -> int:
    try:
        saved_path = asyncio.run(scrape_tabulasi_harga(OUTPUT_PATH))
    except Exception as e:
        print(f"Gagal scraping: {e}", file=sys.stderr)
        return 1

    print(f"Berhasil disimpan ke: {saved_path}")

    try:
        verify_output(saved_path)
    except Exception as e:
        print(f"File tersimpan tapi gagal diverifikasi: {e}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
