"""
Sumber:
https://www.bps.go.id/id/statistics-table/2/MTk1NiMy/-seri-2010--2-
-pdb-triwulanan-atas-dasar-harga-konstan-menurut-pengeluaran--milyar-rupiah-.html

Pemakaian:
    python a_1_bps_all.py

API key di .env sebagai BPS_API_KEY=...
"""

import argparse
import base64
import csv
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any

import requests

try:
    from dotenv import load_dotenv

    load_dotenv()
except Exception:
    pass

BPS_API_BASE = "https://webapi.bps.go.id/v1/api/list"

DEFAULT_URL = (
    "https://www.bps.go.id/id/statistics-table/2/MTk1NiMy/"
    "-seri-2010--2--pdb-triwulanan-atas-dasar-harga-konstan-menurut-pengeluaran--milyar-rupiah-.html"
)

ENTITAS_PREFIX_RE = re.compile(r"^((?:[A-Za-z0-9]+\.)+)\s*(.+)$")


def extract_var_id_from_bps_url(url: str) -> str:
    m = re.search(r"/statistics-table/2/([^/?#]+)", str(url))
    if not m:
        raise ValueError(
            f"Tidak menemukan segmen /statistics-table/2/<kode> pada URL: {url}"
        )

    token = m.group(1).split(".")[0].strip()
    padded = token + ("=" * (-len(token) % 4))

    try:
        decoded = base64.urlsafe_b64decode(padded.encode("ascii")).decode("utf-8")
    except Exception as e:
        raise ValueError(f"Gagal decode kode BPS '{token}' dari URL: {url}") from e

    m_id = re.match(r"(\d+)", decoded)
    if not m_id:
        raise ValueError(f"Hasil decode kode BPS tidak berisi var_id: {decoded}")
    return m_id.group(1)


def bps_get(params: dict[str, Any], api_key: str, timeout: int = 40) -> dict:
    req_params = {**params, "key": api_key}
    resp = requests.get(BPS_API_BASE, params=req_params, timeout=timeout)
    resp.raise_for_status()

    try:
        data = resp.json()
    except Exception as e:
        raise RuntimeError(
            f"Respons WebAPI BPS bukan JSON. Awal respons: {resp.text[:200]!r}"
        ) from e

    if data.get("status") != "OK":
        raise RuntimeError(f"WebAPI BPS status bukan OK: {data}")

    return data


def parse_th_rows(th_resp: dict) -> list[dict]:
    data = th_resp.get("data")
    if isinstance(data, list) and len(data) > 1 and isinstance(data[1], list):
        return data[1]
    if isinstance(data, dict) and isinstance(data.get("data"), list):
        return data["data"]
    return []


def normalize_records(raw) -> list[dict]:
    if raw is None:
        raw = []
    if isinstance(raw, dict):
        raw = [raw]
    if not isinstance(raw, list):
        return []

    rows = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        val = item.get("val") if item.get("val") is not None else item.get("id")
        label = item.get("label")
        rows.append({"id": str(val), "label": str(label or "")})
    return rows


def get_var_meta(data_obj: dict) -> dict:
    var_raw = data_obj.get("var") or {}
    if isinstance(var_raw, list):
        item = var_raw[0] if var_raw else {}
    elif isinstance(var_raw, dict):
        item = var_raw
    else:
        item = {}
    return {
        "var_id": str(item.get("val") or item.get("id") or ""),
        "variabel": str(item.get("label") or ""),
        "unit": str(item.get("unit") or ""),
    }


def fetch_th_list(
    var_id: str, api_key: str, domain: str, timeout: int
) -> dict[str, str]:
    result = {}
    page = 1

    while True:
        th_resp = bps_get(
            {
                "model": "th",
                "lang": "ind",
                "domain": domain,
                "var": var_id,
                "page": page,
            },
            api_key,
            timeout,
        )

        th_rows = parse_th_rows(th_resp)
        if not th_rows:
            break

        for r in th_rows:
            if r.get("th") is not None and r.get("th_id") is not None:
                result[str(r.get("th"))] = str(r.get("th_id"))

        data = th_resp.get("data")
        paging = (
            data[0]
            if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict)
            else {}
        )

        total_pages = int(
            paging.get("pages")
            or paging.get("total_pages")
            or paging.get("page_total")
            or page
        )

        if page >= total_pages:
            break

        page += 1

    return result


def parse_entitas_label(label: str) -> tuple[bool, str]:
    text = label.strip()
    m = ENTITAS_PREFIX_RE.match(text)
    if not m:
        return True, text

    prefix, rest = m.groups()
    first_token = prefix.split(".")[0]
    is_parent = first_token.isdigit()
    return is_parent, rest.strip()


def build_entitas_map(vervar: list[dict]) -> dict[str, tuple[str, str]]:
    ordered = sorted(vervar, key=lambda v: int(v["id"]) if v["id"].isdigit() else 0)

    mapping: dict[str, tuple[str, str]] = {}
    current_parent = ""
    for v in ordered:
        is_parent, text = parse_entitas_label(v["label"])
        if is_parent:
            current_parent = text
            mapping[v["id"]] = (text, "")
        else:
            mapping[v["id"]] = (current_parent, text)
    return mapping


def fetch_year_all_rows(
    var_id: str,
    th_id: str,
    api_key: str,
    domain: str,
    timeout: int,
) -> list[dict]:
    data_obj = bps_get(
        {"model": "data", "lang": "ind", "domain": domain, "var": var_id, "th": th_id},
        api_key,
        timeout,
    )
    datacontent = data_obj.get("datacontent") or {}
    if not isinstance(datacontent, dict):
        datacontent = {}

    meta = get_var_meta(data_obj)
    var_id_resp = meta["var_id"] or var_id

    vervar = normalize_records(data_obj.get("vervar"))
    turvar = normalize_records(data_obj.get("turvar")) or [{"id": "0", "label": ""}]
    tahun = normalize_records(data_obj.get("tahun")) or [{"id": "", "label": ""}]
    turtahun = normalize_records(data_obj.get("turtahun")) or [{"id": "0", "label": ""}]

    if not vervar:
        raise RuntimeError("Tidak menemukan daftar entitas (vervar) di respons BPS.")

    entitas_map = build_entitas_map(vervar)
    tv = turvar[0]
    th = tahun[0]

    rows: list[dict] = []
    vervar_ordered = sorted(
        vervar, key=lambda v: int(v["id"]) if v["id"].isdigit() else 0
    )
    for v in vervar_ordered:
        entitas_parent, entitas_child = entitas_map.get(v["id"], (v["label"], ""))
        for turth in turtahun:
            key = f"{v['id']}{var_id_resp}{tv['id']}{th['id']}{turth['id']}"
            raw_val = datacontent.get(key)
            if raw_val in (None, "-", ""):
                continue
            nilai_str = str(raw_val).replace(",", "").strip()
            try:
                float(nilai_str)
            except ValueError:
                continue
            rows.append(
                {
                    "entitas_parent": entitas_parent,
                    "entitas_child": entitas_child,
                    "periode": turth["label"],
                    "nilai": nilai_str,
                }
            )

    return rows


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Scrape WebAPI BPS untuk SEMUA entitas/kategori PDB Triwulanan "
            "Atas Dasar Harga Konstan menurut Pengeluaran (bukan cuma total PDB)."
        )
    )
    parser.add_argument(
        "--url", default=DEFAULT_URL, help="URL sumber BPS statistics-table"
    )
    parser.add_argument(
        "--years",
        default="2010,2011,2012,2013,2014,2015,2016,2017,2018,2019,2020,2021,2022,2023,2024,2025,2026",
        help="Daftar tahun target, pisah koma",
    )
    parser.add_argument("--domain", default="0000", help="Domain BPS. Nasional = 0000")
    parser.add_argument(
        "--api-key",
        default=os.getenv("BPS_API_KEY", ""),
        help="BPS API key. Bisa juga lewat .env BPS_API_KEY",
    )
    parser.add_argument("--timeout", type=int, default=40)
    parser.add_argument(
        "--kode-sumber-prefix",
        default="s20102ptadhkmp",
        help=(
            "Prefix kode_sumber (tanpa tahun). Default mengikuti pola tabel "
            "'Seri 2010, tabel 2, PDB Triwulanan ADHK Menurut Pengeluaran'."
        ),
    )
    parser.add_argument(
        "--kategori",
        default="nan",
        help="Nilai konstan untuk kolom 'kategori' (tabel ini tidak punya dimensi kategori terpisah).",
    )
    parser.add_argument(
        "--output",
        default="data/raw/groupA/pdb_nasional_tahunan.csv",
    )
    args = parser.parse_args()

    if not args.api_key:
        raise RuntimeError(
            "BPS_API_KEY belum ada. Isi .env dengan BPS_API_KEY=... atau pakai --api-key ..."
        )

    target_years = sorted({int(y.strip()) for y in args.years.split(",") if y.strip()})
    if not target_years:
        raise RuntimeError("--years kosong / tidak valid.")

    var_id = extract_var_id_from_bps_url(args.url)
    print(f"var_id dari URL: {var_id}")
    print(f"Tahun target : {target_years}")

    th_map = fetch_th_list(var_id, args.api_key, args.domain, args.timeout)

    all_rows: list[dict] = []
    for year in target_years:
        th_id = th_map.get(str(year))
        if not th_id:
            print(
                f"  [WARNING] Tahun {year} tidak tersedia di BPS (belum dirilis / di luar cakupan). Dilewati."
            )
            continue
        print(f"  Ambil tahun {year} (th_id={th_id}) ...")
        year_rows = fetch_year_all_rows(
            var_id, th_id, args.api_key, args.domain, args.timeout
        )
        print(f"    -> {len(year_rows)} baris ditemukan.")
        for r in year_rows:
            r["tahun"] = year
            all_rows.append(r)

    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    kode_sumber_by_year = {
        year: f"{args.kode_sumber_prefix}{year}" for year in target_years
    }

    final_rows = []
    for idx, r in enumerate(all_rows, start=1):
        final_rows.append(
            {
                "id": idx,
                "kode_sumber": kode_sumber_by_year[r["tahun"]],
                "entitas_parent": r["entitas_parent"],
                "entitas_child": r["entitas_child"],
                "kategori": args.kategori,
                "tahun": r["tahun"],
                "periode": r["periode"],
                "nilai": r["nilai"],
                "created_at": created_at,
            }
        )

    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "id",
                "kode_sumber",
                "entitas_parent",
                "entitas_child",
                "kategori",
                "tahun",
                "periode",
                "nilai",
                "created_at",
            ],
        )
        writer.writeheader()
        writer.writerows(final_rows)

    print(f"\nSelesai. {len(final_rows)} baris ditulis ke: {out_path.resolve()}")


if __name__ == "__main__":
    main()
