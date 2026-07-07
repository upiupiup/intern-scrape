"""
Sumber:
https://www.bps.go.id/id/statistics-table/2/MTk1NiMy/-seri-2010--2-
-pdb-triwulanan-atas-dasar-harga-konstan-menurut-pengeluaran--milyar-rupiah-.html

Tabel ini berisi "PDB Triwulanan Atas Dasar Harga Konstan menurut Pengeluaran"
(Milyar Rupiah). Baris terakhir tabel adalah total "PRODUK DOMESTIK BRUTO".
growth_pct dihitung sebagai pertumbuhan y-o-y per kuartal dari baris total PDB
tsb, yaitu (PDB_tahun_ini_Qx / PDB_tahun_lalu_Qx - 1) * 100.

Karena growth_pct butuh nilai tahun sebelumnya (t-1), script otomatis
mengambil satu tahun tambahan di belakang tahun paling awal yang diminta.
Contoh: --years 2024,2025,2026 -> otomatis juga ambil 2023 dari BPS supaya
growth 2024 bisa dihitung.

Pemakaian:
    python 2_scrape_ekonomi_pertumbuhan_ekonomi_kuartal.py \
        --api-key XXXXXXXX \
        --years 2024,2025,2026 \
        --output data/raw/groupA/1_ekonomi_pertumbuhan_ekonomi_kuartal.csv

API key bisa juga taruh di .env sebagai BPS_API_KEY=...
"""

import argparse
import base64
import csv
import os
import re
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

QUARTER_PATTERN = {
    "i": 1,
    "ii": 2,
    "iii": 3,
    "iv": 4,
}


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


def find_total_pdb_vervar(vervar: list[dict]) -> dict | None:
    """Cari baris total PDB. Biasanya berlabel 'PRODUK DOMESTIK BRUTO'."""
    for v in vervar:
        if "produk domestik bruto" in v["label"].lower():
            return v
    if vervar:
        return max(vervar, key=lambda v: int(v["id"]) if v["id"].isdigit() else -1)
    return None


def parse_quarter_label(label: str) -> int | None:
    m = re.match(r"triwulan\s+(i{1,3}v?|iv)\b", label.strip().lower())
    if not m:
        return None
    return QUARTER_PATTERN.get(m.group(1))


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

        # Cek info pagination dari response BPS
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


def fetch_year_quarter_values(
    var_id: str,
    th_id: str,
    api_key: str,
    domain: str,
    timeout: int,
) -> dict[int, float]:
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

    total_ver = find_total_pdb_vervar(vervar)
    if total_ver is None:
        raise RuntimeError(
            "Tidak menemukan baris total PRODUK DOMESTIK BRUTO di respons BPS."
        )

    th = tahun[0]
    tv = turvar[0]

    result: dict[int, float] = {}
    for turth in turtahun:
        q = parse_quarter_label(turth["label"])
        if q is None:
            continue
        key = f"{total_ver['id']}{var_id_resp}{tv['id']}{th['id']}{turth['id']}"
        raw_val = datacontent.get(key)
        if raw_val in (None, "-", ""):
            continue
        try:
            result[q] = float(str(raw_val).replace(",", ""))
        except ValueError:
            continue

    return result


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Scrape BPS WebAPI untuk ekonomi_pertumbuhan_ekonomi_kuartal (growth_pct y-o-y per kuartal)."
    )
    parser.add_argument(
        "--url", default=DEFAULT_URL, help="URL sumber BPS statistics-table"
    )
    parser.add_argument(
        "--years", default="2024,2025,2026", help="Daftar tahun target, pisah koma"
    )
    parser.add_argument("--domain", default="0000", help="Domain BPS. Nasional = 0000")
    parser.add_argument(
        "--api-key",
        default=os.getenv("BPS_API_KEY", ""),
        help="BPS API key. Bisa juga lewat .env BPS_API_KEY",
    )
    parser.add_argument("--timeout", type=int, default=40)
    parser.add_argument(
        "--output",
        default="data/raw/groupA/usecase_ekonomi.ekonomi_pertumbuhan_ekonomi_kuartal.csv",
    )
    args = parser.parse_args()

    if not args.api_key:
        raise RuntimeError(
            "BPS_API_KEY belum ada. Isi .env dengan BPS_API_KEY=... atau pakai --api-key ..."
        )

    target_years = sorted({int(y.strip()) for y in args.years.split(",") if y.strip()})
    if not target_years:
        raise RuntimeError("--years kosong / tidak valid.")

    # butuh 1 tahun tambahan sebelum tahun paling awal, buat basis growth y-o-y
    fetch_years = sorted(set(target_years) | {target_years[0] - 1})

    var_id = extract_var_id_from_bps_url(args.url)
    print(f"var_id dari URL: {var_id}")
    print(f"Tahun target : {target_years}")
    print(
        f"Tahun diambil: {fetch_years} (tambahan {target_years[0] - 1} buat basis growth {target_years[0]})"
    )

    th_map = fetch_th_list(var_id, args.api_key, args.domain, args.timeout)

    year_quarter_values: dict[int, dict[int, float]] = {}
    for year in fetch_years:
        th_id = th_map.get(str(year))
        if not th_id:
            print(
                f"  [WARNING] Tahun {year} tidak tersedia di BPS (belum dirilis / di luar cakupan). Dilewati."
            )
            continue
        print(f"  Ambil tahun {year} (th_id={th_id}) ...")
        year_quarter_values[year] = fetch_year_quarter_values(
            var_id, th_id, args.api_key, args.domain, args.timeout
        )

    rows = []
    row_id = 1
    for year in target_years:
        cur = year_quarter_values.get(year, {})
        prev = year_quarter_values.get(year - 1, {})
        for q in (1, 2, 3, 4):
            if q not in cur:
                continue
            growth_pct = ""
            if q in prev and prev[q]:
                growth_pct = round((cur[q] / prev[q] - 1) * 100, 2)
            rows.append(
                {
                    "id": row_id,
                    "tahun": year,
                    "kuartal": q,
                    "growth_pct": growth_pct,
                    "data_source": args.url,
                }
            )
            row_id += 1

    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f, fieldnames=["id", "tahun", "kuartal", "growth_pct", "data_source"]
        )
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nSelesai. {len(rows)} baris ditulis ke: {out_path.resolve()}")


if __name__ == "__main__":
    main()
