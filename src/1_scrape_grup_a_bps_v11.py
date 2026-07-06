import argparse
import base64
import csv
import json
import os
import re
import time
from pathlib import Path
from typing import Any

import pandas as pd
import requests

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

BPS_API_BASE = "https://webapi.bps.go.id/v1/api/list"


def slugify(text: str) -> str:
    text = str(text).strip().lower()
    text = re.sub(r"^usecase_ekonomi\.", "", text)
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text or "tabel"


def output_stem(no: int, table_db: str) -> str:
    return f"{int(no)}_{slugify(table_db)}"


def normalize_text(text: str) -> str:
    text = str(text or "").lower()
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"\b(19|20)\d{2}\b", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


STOPWORDS = {
    "dan", "yang", "menurut", "dengan", "atas", "dasar", "harga",
    "persen", "milyar", "rupiah", "orang", "kunjungan", "perjalanan",
    "jumlah", "seri", "tabel", "statistik", "provinsi", "bulan",
    "triwulanan", "tahunan", "tahun", "per", "di", "ke", "dari",
}


def token_set(text: str) -> set[str]:
    return {
        w for w in normalize_text(text).split()
        if len(w) >= 3 and w not in STOPWORDS and not w.isdigit()
    }


def title_score(expected_title: str, data_obj: dict) -> tuple[float, str]:
    var_raw = data_obj.get("var") or {}
    if isinstance(var_raw, list):
        var_item = var_raw[0] if var_raw else {}
    elif isinstance(var_raw, dict):
        var_item = var_raw
    else:
        var_item = {}

    var_label = str(var_item.get("label") or var_item.get("title") or var_item.get("var") or "")
    unit = str(var_item.get("unit") or "")

    expected_tokens = token_set(expected_title)
    actual_tokens = token_set(var_label + " " + unit)

    if not expected_tokens or not actual_tokens:
        return 0.0, var_label

    inter = expected_tokens & actual_tokens
    union = expected_tokens | actual_tokens
    jaccard = len(inter) / max(len(union), 1)
    coverage = len(inter) / max(len(expected_tokens), 1)
    score = round((0.35 * jaccard) + (0.65 * coverage), 4)
    return score, var_label


def extract_var_id_from_bps_url(url: str) -> str:
    """
    URL BPS statistics-table biasanya punya segmen base64 seperti:
    /statistics-table/2/MTk1NiMy/... -> base64 decode = 1956#2 -> var_id = 1956
    """
    m = re.search(r"/statistics-table/2/([^/?#]+)", str(url))
    if not m:
        raise ValueError(f"Tidak menemukan segmen /statistics-table/2/<kode> pada URL: {url}")

    token = m.group(1).split(".")[0]
    token = token.strip()
    padded = token + ("=" * (-len(token) % 4))

    try:
        decoded = base64.urlsafe_b64decode(padded.encode("ascii")).decode("utf-8")
    except Exception as e:
        raise ValueError(f"Gagal decode kode BPS '{token}' dari URL: {url}") from e

    m_id = re.match(r"(\d+)", decoded)
    if not m_id:
        raise ValueError(f"Hasil decode kode BPS tidak berisi var_id: {decoded}")
    return m_id.group(1)


def find_target_year(text: str) -> str | None:
    years = re.findall(r"\b(?:19|20)\d{2}\b", str(text or ""))
    return years[-1] if years else None


def bps_get(params: dict[str, Any], api_key: str, timeout: int = 40) -> dict:
    req_params = {**params, "key": api_key}
    resp = requests.get(BPS_API_BASE, params=req_params, timeout=timeout)
    resp.raise_for_status()

    try:
        data = resp.json()
    except Exception as e:
        raise RuntimeError(f"Respons WebAPI BPS bukan JSON. Awal respons: {resp.text[:200]!r}") from e

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


def pick_th_id(th_rows: list[dict], target_year: str | None) -> tuple[str | None, str | None, str]:
    if not th_rows:
        return None, None, "Tidak ada daftar th; request data tanpa parameter th."

    if target_year:
        for r in th_rows:
            if str(r.get("th")) == str(target_year):
                return str(r.get("th_id")), str(r.get("th")), "match_target_year"

    # BPS biasanya mengurutkan terbaru dulu, seperti 2024, 2023, dst.
    r = th_rows[0]
    return str(r.get("th_id")), str(r.get("th")), "fallback_latest_year"


def fetch_bps_data_by_var(
    var_id: str,
    api_key: str,
    domain: str = "0000",
    target_year: str | None = None,
    timeout: int = 40,
) -> tuple[dict, dict]:
    th_resp = bps_get({"model": "th", "lang": "ind", "domain": domain, "var": var_id}, api_key, timeout)
    th_rows = parse_th_rows(th_resp)
    th_id, th_year, th_pick_reason = pick_th_id(th_rows, target_year)

    params: dict[str, Any] = {"model": "data", "lang": "ind", "domain": domain, "var": var_id}
    if th_id:
        params["th"] = th_id

    data_obj = bps_get(params, api_key, timeout)

    info = {
        "var_id": var_id,
        "target_year_from_title": target_year or "",
        "selected_th_id": th_id or "",
        "selected_th_year": th_year or "",
        "selected_th_reason": th_pick_reason,
        "available_years": th_rows,
    }
    return data_obj, info


def normalize_records(raw, default_id="0", default_label="Tidak ada") -> list[dict]:
    if raw is None:
        raw = []
    if isinstance(raw, dict):
        raw = [raw]
    if not isinstance(raw, list) or len(raw) == 0:
        return [{"id": str(default_id), "label": str(default_label), "raw": {}}]

    rows = []
    for item in raw:
        if not isinstance(item, dict):
            rows.append({"id": str(item), "label": str(item), "raw": {"value": item}})
            continue

        val = (
            item.get("val")
            if item.get("val") is not None
            else item.get("id")
            if item.get("id") is not None
            else item.get("th_id")
            if item.get("th_id") is not None
            else item.get("turth_id")
            if item.get("turth_id") is not None
            else item.get("vervar_id")
            if item.get("vervar_id") is not None
            else item.get("turvar_id")
        )

        label = (
            item.get("label")
            if item.get("label") is not None
            else item.get("th")
            if item.get("th") is not None
            else item.get("name")
            if item.get("name") is not None
            else item.get("nama")
        )

        if val is None:
            val = label if label is not None else default_id
        if label is None:
            label = val

        rows.append({"id": str(val), "label": str(label), "raw": item})

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
        "variabel": str(item.get("label") or item.get("title") or item.get("var") or ""),
        "unit": str(item.get("unit") or item.get("satuan") or ""),
        "subject": str(item.get("subj") or item.get("subject") or ""),
        "note": str(item.get("note") or ""),
    }


def flatten_bps_data(data_obj: dict) -> pd.DataFrame:
    datacontent = data_obj.get("datacontent")
    if not isinstance(datacontent, dict):
        raise ValueError("data_obj tidak punya datacontent dict")

    var = get_var_meta(data_obj)
    var_id = var["var_id"]

    vervar = normalize_records(data_obj.get("vervar"))
    turvar = normalize_records(data_obj.get("turvar"))
    tahun = normalize_records(data_obj.get("tahun"))
    turtahun = normalize_records(data_obj.get("turtahun"))

    labelvervar = data_obj.get("labelvervar") or "vervar"
    labelturvar = data_obj.get("labelturvar") or "turvar"
    labeltahun = data_obj.get("labeltahun") or "tahun"
    labelturtahun = data_obj.get("labelturtahun") or "turtahun"

    rows = []

    for ver in vervar:
        for tur in turvar:
            for th in tahun:
                for turth in turtahun:
                    key = f"{ver['id']}{var_id}{tur['id']}{th['id']}{turth['id']}"
                    if key not in datacontent:
                        continue

                    rows.append({
                        "labelvervar": labelvervar,
                        "vervar_id": ver["id"],
                        "vervar": ver["label"],
                        "var_id": var_id,
                        "variabel": var["variabel"],
                        "unit": var["unit"],
                        "subject": var["subject"],
                        "labelturvar": labelturvar,
                        "turvar_id": tur["id"],
                        "turvar": tur["label"],
                        "labeltahun": labeltahun,
                        "tahun_id": th["id"],
                        "tahun": th["label"],
                        "labelturtahun": labelturtahun,
                        "turth_id": turth["id"],
                        "turth": turth["label"],
                        "nilai": datacontent[key],
                        "datacontent_key": key,
                    })

    if not rows and datacontent:
        # Fallback kalau kombinasi key BPS berubah/ambigu.
        for key, val in datacontent.items():
            rows.append({
                "labelvervar": labelvervar,
                "vervar_id": None,
                "vervar": None,
                "var_id": var_id,
                "variabel": var["variabel"],
                "unit": var["unit"],
                "subject": var["subject"],
                "labelturvar": labelturvar,
                "turvar_id": None,
                "turvar": None,
                "labeltahun": labeltahun,
                "tahun_id": None,
                "tahun": None,
                "labelturtahun": labelturtahun,
                "turth_id": None,
                "turth": None,
                "nilai": val,
                "datacontent_key": key,
            })

    return pd.DataFrame(rows)



def _is_default_dimension(records: list[dict]) -> bool:
    """True kalau dimensi cuma placeholder Tidak ada/0."""
    if not records:
        return True
    if len(records) != 1:
        return False
    r = records[0]
    return str(r.get("id", "")) in ("0", "", "None") and str(r.get("label", "")).lower() in ("tidak ada", "", "none", "nan")


def _format_bps_value(value: Any) -> str:
    """Format angka supaya mirip CSV download BPS: 468712.0 -> 468712, missing -> -."""
    if value is None:
        return "-"
    if isinstance(value, float) and pd.isna(value):
        return "-"
    text = str(value).strip()
    if text == "" or text.lower() in ("nan", "none", "null"):
        return "-"
    if text == "-":
        return "-"
    try:
        num = float(text.replace(",", ""))
        if num.is_integer():
            return str(int(num))
        return (f"{num:.10f}").rstrip("0").rstrip(".")
    except Exception:
        return text


def _label_with_unit(label: str, unit: str) -> str:
    label = str(label or "").strip()
    unit = str(unit or "").strip()
    if unit and f"({unit})" not in label:
        return f"{label} ({unit})"
    return label


def make_bps_download_rows(data_obj: dict, selected_year: str | None = None) -> list[list[str]]:
    """
    Buat baris CSV yang lebih menyerupai file hasil download BPS:

    row 1: label dimensi baris, mis. "38 Provinsi" / "PDB Pengeluaran (Seri 2010)"
    row 2: nama variabel + satuan
    row 3: tahun
    row 4: header periode/kolom, mis. Februari, Agustus, Tahunan
    row 5+: data, dengan missing value diisi "-"
    """
    datacontent = data_obj.get("datacontent") or {}
    if not isinstance(datacontent, dict):
        datacontent = {}

    meta = get_var_meta(data_obj)
    var_id = meta["var_id"]

    vervar = normalize_records(data_obj.get("vervar"))
    turvar = normalize_records(data_obj.get("turvar"))
    tahun = normalize_records(data_obj.get("tahun"))
    turtahun = normalize_records(data_obj.get("turtahun"))

    labelvervar = str(data_obj.get("labelvervar") or "vervar")
    indicator_title = _label_with_unit(meta["variabel"], meta["unit"])

    # Tahun yang ditampilkan di baris judul.
    year_label = selected_year or ""
    if not year_label and tahun:
        year_label = str(tahun[0].get("label") or tahun[0].get("id") or "")

    # Pilih dimensi kolom mengikuti pola umum download BPS.
    # Biasanya: baris = vervar, kolom = turtahun/periode. Jika ada turvar yang benar-benar
    # bermakna, kolom dibuat kombinasi turvar + turtahun.
    use_turvar = not _is_default_dimension(turvar)
    use_turtahun = not _is_default_dimension(turtahun)

    col_combos: list[tuple[dict, dict]] = []
    if use_turvar and use_turtahun:
        col_combos = [(tv, tt) for tv in turvar for tt in turtahun]
        col_labels = [f"{tv['label']} - {tt['label']}" for tv, tt in col_combos]
    elif use_turvar:
        default_turtahun = turtahun[0] if turtahun else {"id": "0", "label": ""}
        col_combos = [(tv, default_turtahun) for tv in turvar]
        col_labels = [str(tv["label"]) for tv, _ in col_combos]
    elif use_turtahun:
        default_turvar = turvar[0] if turvar else {"id": "0", "label": ""}
        col_combos = [(default_turvar, tt) for tt in turtahun]
        col_labels = [str(tt["label"]) for _, tt in col_combos]
    else:
        default_turvar = turvar[0] if turvar else {"id": "0", "label": ""}
        default_turtahun = turtahun[0] if turtahun else {"id": "0", "label": ""}
        col_combos = [(default_turvar, default_turtahun)]
        col_labels = [year_label or "Nilai"]

    # Tahun untuk key datacontent. Kalau API dipanggil dengan th tertentu, biasanya hanya ada 1 tahun.
    th = tahun[0] if tahun else {"id": "", "label": year_label}

    n_cols = 1 + len(col_labels)
    rows: list[list[str]] = []
    rows.append([labelvervar] + [""] * (n_cols - 1))
    rows.append([""] + [indicator_title] + [""] * max(0, n_cols - 2))
    rows.append([""] + [str(year_label)] + [""] * max(0, n_cols - 2))
    rows.append([""] + col_labels)

    for ver in vervar:
        row = [str(ver["label"])]
        for tv, tt in col_combos:
            key = f"{ver['id']}{var_id}{tv['id']}{th['id']}{tt['id']}"
            row.append(_format_bps_value(datacontent.get(key, "-")))
        rows.append(row)

    return rows


def write_bps_download_csv(path: Path, rows: list[list[str]]) -> None:
    """Tulis CSV dengan BOM UTF-8 supaya aman dibuka di Excel seperti download BPS."""
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(rows)

def read_group_rows(excel_path: Path, group: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, sheet_name="Data Lengkap")
    required = ["No", "Tabel DB", "Nama Ramah", "sumber URL/Dokumen", "Grup"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise RuntimeError(f"Kolom wajib tidak ketemu: {missing}")
    return df[df["Grup"].astype(str).str.strip().str.upper().eq(group.upper())].copy()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Scrape Grup A BPS lewat WebAPI BPS dan simpan output utama dalam bentuk CSV seperti download BPS."
    )
    parser.add_argument("--excel", required=True, help="Path Excel skema")
    parser.add_argument("--output-dir", default="data/raw/groupA")
    parser.add_argument("--group", default="A")
    parser.add_argument("--judul-col", default="Judul Resmi Tabel")
    parser.add_argument("--domain", default="0000", help="Domain BPS. Nasional = 0000")
    parser.add_argument("--api-key", default=os.getenv("BPS_API_KEY", ""), help="BPS API key. Bisa juga lewat .env BPS_API_KEY")
    parser.add_argument("--timeout", type=int, default=40)
    parser.add_argument("--sleep", type=float, default=0.5)
    parser.add_argument("--min-score-warning", type=float, default=0.35)
    parser.add_argument("--only", default="", help="Opsional, contoh: --only 1,12,30")
    args = parser.parse_args()

    if not args.api_key:
        raise RuntimeError("BPS_API_KEY belum ada. Isi .env dengan BPS_API_KEY=... atau pakai --api-key ...")

    excel_path = Path(args.excel)
    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    rows = read_group_rows(excel_path, args.group)
    if args.only.strip():
        allowed = {int(x.strip()) for x in args.only.split(",") if x.strip()}
        rows = rows[rows["No"].astype(int).isin(allowed)].copy()

    print(f"Total baris grup {args.group}: {len(rows)}")
    print(f"Output dir: {out_dir.resolve()}")
    print(f"Metode: BPS WebAPI | domain={args.domain}")

    manifest = []

    for _, row in rows.iterrows():
        no = int(row["No"])
        nama = str(row["Nama Ramah"])
        table_db = str(row["Tabel DB"])
        url = str(row["sumber URL/Dokumen"]).strip()
        judul = str(row.get(args.judul_col, "") or "").strip()
        stem = output_stem(no, table_db)

        print(f"\n[{stem}] {nama}")
        print(f"  Judul: {judul}")
        print(f"  URL  : {url}")

        item = {
            "no": no,
            "nama_ramah": nama,
            "tabel_db": table_db,
            "url": url,
            "judul_resmi_tabel": judul,
            "status": "GAGAL",
            "rows": "",
            "cols": "",
            "var_id_from_url": "",
            "var_id_bps": "",
            "variabel_bps": "",
            "target_year_from_title": "",
            "selected_th_id": "",
            "selected_th_year": "",
            "selected_th_reason": "",
            "title_score": "",
            "csv_file": "",
            "json_file": "",
            "warning": "",
            "error": "",
        }

        try:
            var_id = extract_var_id_from_bps_url(url)
            target_year = find_target_year(judul)
            data_obj, api_info = fetch_bps_data_by_var(
                var_id=var_id,
                api_key=args.api_key,
                domain=args.domain,
                target_year=target_year,
                timeout=args.timeout,
            )

            df_out = flatten_bps_data(data_obj)
            meta = get_var_meta(data_obj)
            score, label = title_score(judul, data_obj)

            item["var_id_from_url"] = var_id
            item["var_id_bps"] = meta["var_id"]
            item["variabel_bps"] = meta["variabel"]
            item["target_year_from_title"] = api_info.get("target_year_from_title", "")
            item["selected_th_id"] = api_info.get("selected_th_id", "")
            item["selected_th_year"] = api_info.get("selected_th_year", "")
            item["selected_th_reason"] = api_info.get("selected_th_reason", "")
            item["title_score"] = score

            if score < args.min_score_warning and judul:
                item["warning"] = f"Judul agak beda dari metadata BPS. score={score}, bps_label={label}"

            if target_year and api_info.get("selected_th_year") and str(target_year) != str(api_info.get("selected_th_year")):
                year_warning = f"Target year {target_year} tidak tersedia; pakai {api_info.get('selected_th_year')}"
                item["warning"] = (item["warning"] + " | " + year_warning).strip(" |")

            df_out.insert(0, "source_no", no)
            df_out.insert(1, "source_nama_ramah", nama)
            df_out.insert(2, "source_tabel_db", table_db)
            df_out.insert(3, "source_url", url)
            df_out.insert(4, "source_judul_resmi_tabel", judul)
            df_out.insert(5, "source_title_score", score)
            df_out.insert(6, "source_selected_th_id", item["selected_th_id"])
            df_out.insert(7, "source_selected_th_year", item["selected_th_year"])

            # Output utama dibuat mirip tabel BPS (wide/display).
            # Output long tetap disimpan untuk kebutuhan database/ETL.
            csv_path = out_dir / f"{stem}.csv"
            long_csv_path = out_dir / f"{stem}_long.csv"
            json_path = out_dir / f"{stem}.json"

            bps_rows = make_bps_download_rows(data_obj, selected_year=item["selected_th_year"] or target_year)
            write_bps_download_csv(csv_path, bps_rows)
            df_out.to_csv(long_csv_path, index=False, encoding="utf-8-sig")
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "source": "bps_webapi",
                        "source_no": no,
                        "source_nama_ramah": nama,
                        "source_tabel_db": table_db,
                        "source_url": url,
                        "source_judul_resmi_tabel": judul,
                        "api_info": api_info,
                        "title_score": score,
                        "chosen_bps_label": label,
                        "meta": meta,
                        "data": data_obj,
                    },
                    f,
                    ensure_ascii=False,
                    indent=2,
                )

            item["status"] = "OK"
            item["rows"] = max(0, len(bps_rows) - 4)
            item["cols"] = len(bps_rows[0]) if bps_rows else 0
            item["csv_file"] = str(csv_path)
            item["json_file"] = str(json_path)

            msg = (
                f"  OK -> {csv_path.name} | tabel_bps={item["rows"]} rows x {item["cols"]} cols "
                f"| long={long_csv_path.name} ({len(df_out)} rows x {df_out.shape[1]} cols) "
                f"| var={meta['var_id']} | th={item['selected_th_id']} ({item['selected_th_year']})"
            )
            if item["warning"]:
                msg += " | WARNING"
            print(msg)

        except Exception as e:
            item["error"] = f"{type(e).__name__}: {e}"
            print(f"  GAGAL -> {item['error']}")

        manifest.append(item)
        pd.DataFrame(manifest).to_csv(out_dir / "_manifest_grup_A_webapi.csv", index=False, encoding="utf-8-sig")
        time.sleep(args.sleep)

    print(f"\nManifest tersimpan: {out_dir / '_manifest_grup_A_webapi.csv'}")


if __name__ == "__main__":
    main()
