import argparse
import json
import re
from pathlib import Path

import pandas as pd
import requests
from playwright.sync_api import sync_playwright


def read_group_a(excel_path: Path) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name="Data Lengkap")
    return df[df["Grup"].astype(str).str.strip().str.upper().eq("A")].copy()


def norm_name(text: str) -> str:
    text = str(text).lower()
    text = re.sub(r"^usecase_ekonomi\.", "", text)
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return re.sub(r"_+", "_", text).strip("_") or "page"


def inspect_raw_html(url: str, timeout: int = 30) -> dict:
    out = {
        "status_code": None,
        "final_url": None,
        "content_type": None,
        "html_len": 0,
        "table_tag_count": 0,
        "download_word_count": 0,
        "iframe_count": 0,
        "script_count": 0,
        "notice_direktori": False,
        "title": "",
        "error": "",
    }

    try:
        r = requests.get(
            url,
            timeout=timeout,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                )
            },
        )
        out["status_code"] = r.status_code
        out["final_url"] = r.url
        out["content_type"] = r.headers.get("content-type", "")
        html = r.text or ""
        out["html_len"] = len(html)
        out["table_tag_count"] = len(re.findall(r"<table\b", html, flags=re.I))
        out["download_word_count"] = len(re.findall(r"(download|unduh|excel|xlsx|xls|csv)", html, flags=re.I))
        out["iframe_count"] = len(re.findall(r"<iframe\b", html, flags=re.I))
        out["script_count"] = len(re.findall(r"<script\b", html, flags=re.I))
        out["notice_direktori"] = "Tabel Publikasi Indikator Ekonomi" in html or "direktori.web.bps.go.id" in html

        m = re.search(r"<title[^>]*>(.*?)</title>", html, flags=re.I | re.S)
        if m:
            out["title"] = re.sub(r"\s+", " ", m.group(1)).strip()

    except Exception as e:
        out["error"] = f"{type(e).__name__}: {e}"

    return out


def inspect_rendered_page(url: str, out_dir: Path, stem: str, timeout_ms: int = 60000, headed: bool = False) -> dict:
    result = {
        "render_final_url": "",
        "render_title": "",
        "render_html_len": 0,
        "render_table_tag_count": 0,
        "render_iframe_count": 0,
        "render_download_candidates": 0,
        "render_buttons": 0,
        "render_links": 0,
        "network_interesting_count": 0,
        "screenshot": "",
        "render_html": "",
        "network_json": "",
        "error_render": "",
    }

    out_dir.mkdir(parents=True, exist_ok=True)

    interesting_urls = []

    def on_request(req):
        u = req.url
        if re.search(r"(api|ajax|data|download|export|metabase|dataset|statistic|table|dynamictable|statictable)", u, re.I):
            interesting_urls.append({"method": req.method, "url": u, "resource_type": req.resource_type})

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=not headed)
            context = browser.new_context(
                viewport={"width": 1366, "height": 900},
                locale="id-ID",
                timezone_id="Asia/Jakarta",
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
            )
            page = context.new_page()
            page.on("request", on_request)

            page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            try:
                page.wait_for_load_state("networkidle", timeout=min(timeout_ms, 15000))
            except Exception:
                pass
            page.wait_for_timeout(5000)

            html = page.content()
            result["render_final_url"] = page.url
            result["render_title"] = page.title()
            result["render_html_len"] = len(html)
            result["render_table_tag_count"] = len(re.findall(r"<table\b", html, flags=re.I))
            result["render_iframe_count"] = len(re.findall(r"<iframe\b", html, flags=re.I))

            links = page.locator("a").evaluate_all(
                """els => els.map(a => ({
                    text: (a.innerText || a.textContent || '').trim(),
                    href: a.href || a.getAttribute('href') || '',
                    title: a.getAttribute('title') || '',
                    aria: a.getAttribute('aria-label') || ''
                }))"""
            )
            buttons = page.locator("button").evaluate_all(
                """els => els.map(b => ({
                    text: (b.innerText || b.textContent || '').trim(),
                    title: b.getAttribute('title') || '',
                    aria: b.getAttribute('aria-label') || ''
                }))"""
            )

            download_re = re.compile(r"(unduh|download|excel|xlsx|xls|csv|export)", re.I)
            candidates = [
                x for x in links + buttons
                if download_re.search(" ".join(str(v) for v in x.values()))
            ]

            screenshot_path = out_dir / f"{stem}_render.png"
            html_path = out_dir / f"{stem}_render.html"
            links_path = out_dir / f"{stem}_links_buttons.json"
            network_path = out_dir / f"{stem}_network_interesting.json"

            page.screenshot(path=str(screenshot_path), full_page=True)
            html_path.write_text(html, encoding="utf-8")
            links_path.write_text(json.dumps({"links": links, "buttons": buttons, "download_candidates": candidates}, ensure_ascii=False, indent=2), encoding="utf-8")
            network_path.write_text(json.dumps(interesting_urls, ensure_ascii=False, indent=2), encoding="utf-8")

            result["render_links"] = len(links)
            result["render_buttons"] = len(buttons)
            result["render_download_candidates"] = len(candidates)
            result["network_interesting_count"] = len(interesting_urls)
            result["screenshot"] = str(screenshot_path)
            result["render_html"] = str(html_path)
            result["network_json"] = str(network_path)

            browser.close()

    except Exception as e:
        result["error_render"] = f"{type(e).__name__}: {e}"

    return result


def main():
    parser = argparse.ArgumentParser(description="Cek kondisi halaman BPS: HTML statis atau JS/render/direktori/metabase.")
    parser.add_argument("--excel", required=True)
    parser.add_argument("--output-dir", default="data/raw/groupA_debug")
    parser.add_argument("--limit", type=int, default=3, help="Jumlah baris awal grup A yang dicek. Pakai 0 untuk semua.")
    parser.add_argument("--timeout", type=int, default=60)
    parser.add_argument("--headed", action="store_true")
    args = parser.parse_args()

    excel_path = Path(args.excel)
    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    group_a = read_group_a(excel_path)
    if args.limit and args.limit > 0:
        group_a = group_a.head(args.limit)

    rows = []

    for _, row in group_a.iterrows():
        no = int(row["No"])
        nama = str(row["Nama Ramah"])
        tabel_db = str(row["Tabel DB"])
        url = str(row["sumber URL/Dokumen"])
        stem = f"{no}_{norm_name(tabel_db)}"

        print(f"\n[{stem}] {nama}")
        print(url)

        raw = inspect_raw_html(url, timeout=args.timeout)
        print(f"  RAW: status={raw['status_code']} table={raw['table_tag_count']} iframe={raw['iframe_count']} notice_direktori={raw['notice_direktori']}")

        rendered = inspect_rendered_page(url, out_dir, stem, timeout_ms=args.timeout * 1000, headed=args.headed)
        print(
            "  RENDER:",
            f"table={rendered['render_table_tag_count']}",
            f"iframe={rendered['render_iframe_count']}",
            f"download_candidates={rendered['render_download_candidates']}",
            f"network_interesting={rendered['network_interesting_count']}",
        )

        rows.append({
            "no": no,
            "nama_ramah": nama,
            "tabel_db": tabel_db,
            "url": url,
            **raw,
            **rendered,
        })

        pd.DataFrame(rows).to_csv(out_dir / "_page_condition_report.csv", index=False, encoding="utf-8-sig")

    print("\nReport:", out_dir / "_page_condition_report.csv")


if __name__ == "__main__":
    main()
