import asyncio
import threading
import re
from datetime import datetime
import pandas as pd
from playwright.async_api import async_playwright

log = []
download_path = {}

SOURCE_URL = "https://tradingeconomics.com/indonesia/manufacturing-pmi"
OUTPUT_CSV = "data/raw/groupG/5_ekonomi_tren_pmi_bulanan.csv"


async def run():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
            accept_downloads=True,
        )
        page = await context.new_page()

        await page.goto(SOURCE_URL, wait_until="domcontentloaded", timeout=60000)
        await page.wait_for_timeout(8000)
        log.append("halaman ke-load")

        # klik "3Y" dan verifikasi beneran aktif
        try:
            btn_3y = page.locator('text="3Y"').first
            await btn_3y.click(timeout=5000)
            await page.wait_for_timeout(2000)

            is_active = await btn_3y.evaluate("""
                (el) => {
                    const style = window.getComputedStyle(el);
                    return style.backgroundColor !== 'rgba(0, 0, 0, 0)' && style.backgroundColor !== 'transparent';
                }
            """)
            log.append(f"Tombol 3Y aktif? {is_active}")
            await page.wait_for_timeout(2000)
        except Exception as e:
            log.append(f"klik 3Y gagal: {e}")

        # klik ikon titik tiga (menu export chart) -> class-nya "auxExportingBtn"
        try:
            menu_candidates = [
                "button.auxExportingBtn",
                '[class*="auxExportingBtn"]',
                'button:has(svg):near(:text("Compare"))',
                '[class*="dots"]',
                '[class*="menu-icon"]',
                "button >> nth=-1",
            ]
            clicked = False
            for sel in menu_candidates:
                try:
                    el = page.locator(sel).last
                    if await el.count() > 0:
                        await el.click(timeout=3000)
                        clicked = True
                        log.append(f"klik menu pakai selector: {sel}")
                        break
                except Exception:
                    continue
            if not clicked:
                log.append("Gagal klik menu titik-tiga lewat semua selector")
        except Exception as e:
            log.append(f"error klik menu: {e}")

        await page.wait_for_timeout(1000)

        # klik "SVG Image" dan tangkep download-nya
        try:
            async with page.expect_download(timeout=15000) as download_info:
                svg_option = page.locator('text="SVG Image"').first
                await svg_option.click()

            download = await download_info.value
            save_path = "pmi_chart_3y.svg"
            await download.save_as(save_path)
            download_path["path"] = save_path
            log.append(f"SVG tersimpan: {save_path}")
        except Exception as e:
            log.append(f"gagal download SVG: {e}")
            await page.screenshot(path="debug_pmi_3y.png")

        await browser.close()


def run_in_thread(coro_func):
    def runner():
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(coro_func())
        except Exception as e:
            log.append(f"ERROR FATAL: {e}")
        finally:
            loop.close()

    t = threading.Thread(target=runner)
    t.start()
    t.join()


def parse_pmi_svg(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()

    # cari x posisi label sumbu-Y otomatis (bukan hardcode, karena chart 3Y bisa beda layout)
    # skip elemen hidden (helper ukur teks doang, visibility="hidden", gak punya y=)
    ylabel_group = re.search(
        r'<g class="highcharts-axis-labels highcharts-yaxis-labels[^"]*"[^>]*>(.*?)</g>',
        content,
        re.DOTALL,
    )
    label_x = None
    if ylabel_group:
        m = re.search(r'<text x="([\d.]+)"[^>]*\sy="[\d.]+"', ylabel_group.group(1))
        if m:
            label_x = m.group(1)

    if label_x is None:
        # fallback ke nilai lama kalau auto-detect gagal
        label_x = "573"

    label_pattern = re.compile(
        rf'<text x="{label_x}"[^>]*\sy="([\d.]+)"[^>]*><tspan[^>]*>(\d+)</tspan></text>'
    )
    label_matches = label_pattern.findall(content)
    y_pixels = [float(y) for y, v in label_matches]
    y_values = [float(v) for y, v in label_matches]

    n = len(y_pixels)
    mean_x = sum(y_pixels) / n
    mean_y = sum(y_values) / n
    m = sum((x - mean_x) * (y - mean_y) for x, y in zip(y_pixels, y_values)) / sum(
        (x - mean_x) ** 2 for x in y_pixels
    )
    b = mean_y - m * mean_x

    transform_match = re.search(
        r'class="highcharts-series highcharts-series-0[^"]*"[^>]*transform="translate\(([\d.-]+),\s*([\d.-]+)\)',
        content,
    )
    offset_x = float(transform_match.group(1)) if transform_match else 0
    offset_y = float(transform_match.group(2)) if transform_match else 0

    bar_pattern = re.compile(
        r'<rect x="([\d.-]+)" y="([\d.-]+)" width="([\d.-]+)" height="([\d.-]+)" fill="[^"]*" opacity="1" class="highcharts-point">'
    )
    bars = bar_pattern.findall(content)
    bars = [(float(x), float(y), float(w), float(h)) for x, y, w, h in bars]
    bars.sort(key=lambda b: b[0])

    values = []
    for x, y, w, h in bars:
        y_abs = y + offset_y
        val = m * y_abs + b
        x_center = x + w / 2 + offset_x
        values.append((x_center, round(val, 2)))

    xlabel_pattern = re.compile(
        r'<g class="highcharts-axis-labels highcharts-xaxis-labels"[^>]*>(.*?)</g>',
        re.DOTALL,
    )
    xlabel_block = xlabel_pattern.search(content).group(1)
    xlabel_items = re.findall(r'<text x="([\d.]+)"[^>]*>([^<]+)</text>', xlabel_block)
    xlabel_items = [(float(x), txt) for x, txt in xlabel_items]

    bulan_list = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
    ]

    bar_x_centers = [v[0] for v in values]
    anchor_idx_to_label = {}
    for x, txt in xlabel_items:
        idx = min(range(len(bar_x_centers)), key=lambda i: abs(bar_x_centers[i] - x))
        anchor_idx_to_label[idx] = txt

    year_anchor_idx = None
    year_anchor_val = None
    for idx, txt in anchor_idx_to_label.items():
        if re.fullmatch(r"\d{4}", txt):
            year_anchor_idx = idx
            year_anchor_val = int(txt)
            break

    if year_anchor_idx is None:
        raise ValueError(
            "Gak ketemu anchor tahun, gak bisa rekonstruksi bulan otomatis"
        )

    results = []
    for i, (x_center, val) in enumerate(values):
        month_offset = i - year_anchor_idx
        month_num = (0 + month_offset) % 12
        year = year_anchor_val + (0 + month_offset) // 12
        bulan_nama = bulan_list[month_num]
        results.append((f"{bulan_nama} {year}", val))

    return results


if __name__ == "__main__":
    run_in_thread(run)

    print("=== LOG ===")
    for l in log:
        print(l)
    print(download_path)

    hasil = parse_pmi_svg(download_path["path"])
    for bulan, val in hasil:
        print(f"{bulan}: {val}")

    # convert ke CSV
    df = pd.DataFrame(hasil, columns=["periode", "nilai_pmi"])
    df["sumber_url"] = SOURCE_URL
    df["tanggal_scrape"] = datetime.now().strftime("%Y-%m-%d")

    import os

    os.makedirs("data/raw/groupG", exist_ok=True)
    df.to_csv(OUTPUT_CSV, index=False)
    print(f"\nCSV tersimpan di: {OUTPUT_CSV} ({len(df)} baris)")
