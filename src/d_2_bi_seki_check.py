"""
d_2_bi_seki_check.py
=====================
Cek ketersediaan data "Indeks Harga Konsumen Umum - Perubahan Tahun ke Tahun
(YoY)" di dalam arsip ZIP bulanan SEKI (hasil download_seki.py), dari bulan
TERBARU mundur ke bulan TERLAMA.

Untuk tiap bulan, script ini:
  1. Buka file TABEL8_1.xls di dalam zip (kalau ada) -> cek dulu di situ,
     baik di sheet '8.1' (format baru) maupun sheet lain di file yang sama.
  2. Kalau TIDAK ketemu di TABEL8_1.xls, script akan fallback: buka SEMUA
     file .xls/.xlsx lain di dalam zip itu dan cari baris yang cocok
     (kemungkinan indeksnya dipindah ke nomor tabel lain di tahun2 lama).
  3. Kalau ketemu di mana pun, dicatat: nama file, sheet, nomor baris, teks
     label baris, dan nilai terakhir yang ada di baris itu (buat verifikasi
     manual cepat).
  4. Kalau 6 bulan berturut-turut (default, bisa diganti) SAMA SEKALI tidak
     ketemu, pencarian dihentikan otomatis.

Cara pakai:
    python d_2_bi_seki_check.py
    python d_2_bi_seki_check.py --root data/raw/groupD/seki_downloads
    python d_2_bi_seki_check.py --stop-after-missing 6
    python d_2_bi_seki_check.py --stop-after-missing 0   # jangan pernah stop

Hasil:
  - Dicetak ke layar per bulan.
  - Disimpan sebagai CSV: <root>/_report_inflasi_yoy.csv
"""

import argparse
import csv
import glob
import os
import re
import sys
import zipfile
from io import BytesIO

import pandas as pd

# Konfigurasi

BULAN_ID = [
    "JANUARI",
    "FEBRUARI",
    "MARET",
    "APRIL",
    "MEI",
    "JUNI",
    "JULI",
    "AGUSTUS",
    "SEPTEMBER",
    "OKTOBER",
    "NOVEMBER",
    "DESEMBER",
]
BULAN_INDEX = {b: i + 1 for i, b in enumerate(BULAN_ID)}

MONTH_ABBR_EN = {
    "jan",
    "feb",
    "mar",
    "apr",
    "may",
    "jun",
    "jul",
    "aug",
    "sep",
    "oct",
    "nov",
    "dec",
}
MONTH_ABBR_ID = {
    "jan",
    "feb",
    "mar",
    "apr",
    "mei",
    "jun",
    "jul",
    "agu",
    "aug",
    "sep",
    "okt",
    "oct",
    "nov",
    "des",
    "dec",
}

# kata kunci baris yang menandakan "Perubahan YoY / Tahun ke Tahun"
YOY_PATTERNS = [
    r"tahun\s*ke\s*tahun",
    r"\byoy\b",
    r"year\s*on\s*year",
    r"annual\s*inflation",
    r"inflasi\s*tahunan",
]
YOY_RE = re.compile("|".join(YOY_PATTERNS), re.IGNORECASE)

# kata kunci konteks yang menandakan ini bagian "umum" (headline), bukan
# per kelompok barang
UMUM_RE = re.compile(r"\bumum\b|headline|indeks harga konsumen|\bihk\b", re.IGNORECASE)

CONTEXT_WINDOW = 8  # berapa baris ke atas dicek untuk konteks 'umum'

FNAME_RE = re.compile(r"SEKI_([A-Z]+)_(\d{4})\.zip$", re.IGNORECASE)

# sheet dengan nama seperti ini biasanya arsip histori lama (statis, tidak
# berubah antar terbitan) -> prioritas paling rendah
ARCHIVE_SHEET_RE = re.compile(r"^\s*th\.?\s*\d{4}\s*-\s*\d{4}\s*$", re.IGNORECASE)

YEAR_RE = re.compile(r"(19|20)\d{2}")


# Helper: cari file zip dan urutkan dari terbaru


def discover_zips(root):
    zips = glob.glob(os.path.join(root, "**", "SEKI_*.zip"), recursive=True)
    items = []
    for z in zips:
        m = FNAME_RE.search(os.path.basename(z))
        if not m:
            continue
        bulan_raw, tahun = m.group(1).upper(), int(m.group(2))
        if bulan_raw not in BULAN_INDEX:
            continue
        items.append((tahun, BULAN_INDEX[bulan_raw], bulan_raw, z))
    items.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return items


# Helper: cek satu sheet apakah ada baris YoY yang relevan
def find_data_columns(df, header_search_rows=15):
    """Kembalikan set kolom yang punya header bulan (Jan/Feb/... atau Januari/
    Februari/...) di beberapa baris pertama -> ini kolom data asli, BUKAN
    kolom nomor urut/catatan kaki di ujung kanan sheet."""
    n_rows, n_cols = df.shape
    r_limit = min(n_rows, header_search_rows)
    data_cols = set()
    for j in range(n_cols):
        for i in range(r_limit):
            v = df.iat[i, j]
            if isinstance(v, str):
                s = v.strip().lower()
                if s in MONTH_ABBR_EN or s in MONTH_ABBR_ID:
                    data_cols.add(j)
                    break
    return data_cols


def scan_dataframe_for_yoy(df, data_cols=None):
    """Return list of dict kandidat: {row, label, confirmed_umum, last_value, last_col}."""
    candidates = []
    n_rows, n_cols = df.shape
    # batasi jumlah kolom yang discan teksnya biar cepat (label biasanya di kolom awal)
    label_cols = min(n_cols, 6)

    if data_cols is None:
        data_cols = find_data_columns(df)
    sorted_data_cols = sorted(data_cols, reverse=True)  # dari kanan (terbaru) ke kiri

    for i in range(n_rows):
        row_text_cells = []
        for j in range(label_cols):
            v = df.iat[i, j]
            if isinstance(v, str):
                row_text_cells.append(v)
        row_text = " | ".join(row_text_cells)
        if not row_text or not YOY_RE.search(row_text):
            continue

        # cek konteks 'umum' di baris ini atau beberapa baris di atasnya
        confirmed = bool(UMUM_RE.search(row_text))
        if not confirmed:
            start = max(0, i - CONTEXT_WINDOW)
            for k in range(start, i):
                ctx_cells = [
                    df.iat[k, j]
                    for j in range(label_cols)
                    if isinstance(df.iat[k, j], str)
                ]
                ctx_text = " | ".join(ctx_cells)
                if ctx_text and UMUM_RE.search(ctx_text):
                    confirmed = True
                    break

        # cari nilai numerik terakhir (paling kanan) HANYA di kolom data asli
        # (kalau tidak ketemu kolom data sama sekali, fallback ke cara lama)
        last_value, last_col = None, None
        search_cols = (
            sorted_data_cols
            if sorted_data_cols
            else list(range(n_cols - 1, label_cols - 1, -1))
        )
        for j in search_cols:
            v = df.iat[i, j]
            if isinstance(v, (int, float)) and pd.notna(v):
                last_value = v
                last_col = j
                break

        candidates.append(
            {
                "row": i,
                "label": row_text[:120],
                "confirmed_umum": confirmed,
                "last_value": last_value,
                "last_col": last_col,
            }
        )

    return candidates


def detect_year_near(df, row, col, radius=12):
    """Cari angka tahun (1900-2099) di sekitar (row,col) -> dipakai buat menilai
    apakah suatu kandidat baris merepresentasikan data 'terkini' atau arsip lama."""
    n_rows, n_cols = df.shape
    r0, r1 = max(0, row - radius), min(n_rows, row + 1)
    c0, c1 = max(0, col - radius), min(n_cols, col + 1)
    years = []
    for i in range(r0, r1):
        for j in range(c0, c1):
            v = df.iat[i, j]
            if isinstance(v, (int, float)) and pd.notna(v) and 1900 <= v <= 2099:
                years.append(int(v))
            elif isinstance(v, str):
                m = YEAR_RE.search(v)
                if m:
                    years.append(int(m.group(0)))
    return max(years) if years else None


def best_candidate(candidates):
    """Pilih kandidat terbaik dalam SATU sheet: prioritas confirmed_umum=True
    dan punya last_value. (Pemilihan ANTAR sheet dilakukan di scan_workbook_bytes.)"""
    if not candidates:
        return None
    with_value = [c for c in candidates if c["last_value"] is not None]
    pool = with_value if with_value else candidates
    confirmed = [c for c in pool if c["confirmed_umum"]]
    chosen_pool = confirmed if confirmed else pool
    return chosen_pool[0]


# Helper: buka xls/xlsx dari bytes dan scan semua sheet


def scan_workbook_bytes(data_bytes, filename, expected_year=None):
    """Return (found:bool, detail:dict|None) hasil scan semua sheet di workbook ini.

    Kalau ketemu kandidat valid di lebih dari satu sheet (misal ada sheet arsip
    lama '1966-1978' dan sheet aktif '8.1'), pilih yang tahun datanya PALING
    DEKAT dengan expected_year (tahun terbitan file zip-nya) -- supaya tidak
    salah ambil dari sheet arsip statis yang nilainya sama terus di semua file.
    """
    try:
        xl = pd.ExcelFile(BytesIO(data_bytes))
    except Exception as e:
        return False, {"error": f"gagal buka {filename}: {e}"}

    sheet_names = xl.sheet_names
    all_hits = []  # list of dict per sheet yang confirmed_umum

    for sheet in sheet_names:
        try:
            df = pd.read_excel(xl, sheet_name=sheet, header=None)
        except Exception:
            continue
        data_cols = find_data_columns(df)
        candidates = scan_dataframe_for_yoy(df, data_cols=data_cols)
        best = best_candidate(candidates)
        if best is None or not best["confirmed_umum"]:
            continue

        year_found = None
        if best["last_col"] is not None:
            year_found = detect_year_near(df, best["row"], best["last_col"])
        is_archive_name = bool(ARCHIVE_SHEET_RE.match(str(sheet)))

        all_hits.append(
            {
                "file": filename,
                "sheet": sheet,
                "row": best["row"],
                "label": best["label"],
                "last_value": best["last_value"],
                "year_found": year_found,
                "is_archive_name": is_archive_name,
            }
        )

    if not all_hits:
        return False, None

    def score(hit):
        # skor lebih kecil = lebih bagus
        year_diff = 999
        if expected_year is not None and hit["year_found"] is not None:
            year_diff = abs(hit["year_found"] - expected_year)
        elif hit["year_found"] is None:
            year_diff = 500  # gak ketauan tahunnya, taruh di bawah yg ketauan
        archive_penalty = 1000 if hit["is_archive_name"] else 0
        return archive_penalty + year_diff

    all_hits.sort(key=score)
    chosen = all_hits[0]
    return True, {
        "file": chosen["file"],
        "sheet": chosen["sheet"],
        "row": chosen["row"],
        "label": chosen["label"],
        "last_value": chosen["last_value"],
    }


# Main check per zip


def check_zip(zip_path, expected_year=None):
    """Return dict hasil pengecekan satu file zip bulanan."""
    try:
        zf = zipfile.ZipFile(zip_path)
    except Exception as e:
        return {"status": "zip_error", "detail": str(e)}

    names = zf.namelist()
    xls_like = [n for n in names if n.lower().endswith((".xls", ".xlsx"))]

    # 1) prioritas: TABEL8_1.xls (atau .xlsx)
    table8_1 = [
        n for n in xls_like if re.search(r"tabel8[_\.]?1\.xlsx?$", n, re.IGNORECASE)
    ]
    for n in table8_1:
        data = zf.read(n)
        found, detail = scan_workbook_bytes(data, n, expected_year=expected_year)
        if found:
            detail["source"] = "TABEL8_1"
            return {"status": "found", **detail}

    # 2) fallback: semua file xls/xlsx lain di zip (indeks mungkin ada di tabel lain)
    others = [n for n in xls_like if n not in table8_1]
    for n in others:
        data = zf.read(n)
        found, detail = scan_workbook_bytes(data, n, expected_year=expected_year)
        if found:
            detail["source"] = "TABEL_LAIN"
            return {"status": "found_elsewhere", **detail}

    return {
        "status": "not_found",
        "detail": f"{len(xls_like)} file xls/xlsx dicek, tidak ada baris YoY umum",
    }


# Main


def main():
    parser = argparse.ArgumentParser(
        description="Cek ketersediaan indeks inflasi umum YoY di arsip SEKI"
    )
    parser.add_argument(
        "--root", default=os.path.join("data", "raw", "groupD", "seki_downloads")
    )
    parser.add_argument(
        "--stop-after-missing",
        type=int,
        default=6,
        help="Stop kalau N bulan berturut-turut tidak ketemu sama sekali. 0 = jangan stop.",
    )
    parser.add_argument(
        "--report",
        default=None,
        help="Path CSV output (default: <root>/_report_inflasi_yoy.csv)",
    )
    args = parser.parse_args()

    report_path = args.report or os.path.join(args.root, "_report_inflasi_yoy.csv")

    items = discover_zips(args.root)
    if not items:
        print(f"Tidak ada file SEKI_*.zip ditemukan di {args.root}")
        return 1

    print(f"Ditemukan {len(items)} file zip. Mulai cek dari terbaru ke terlama...\n")

    rows_out = []
    consecutive_missing = 0

    for tahun, bulan_idx, bulan_raw, zip_path in items:
        label = f"{bulan_raw} {tahun}"
        result = check_zip(zip_path, expected_year=tahun)
        status = result["status"]

        if status == "found":
            print(
                f"[{label}] OK - ketemu di TABEL8_1 (sheet={result['sheet']}, baris={result['row']}, "
                f"nilai_terakhir={result['last_value']})"
            )
            consecutive_missing = 0
        elif status == "found_elsewhere":
            print(
                f"[{label}] KETEMU TAPI BUKAN DI TABEL8_1 -> file={result['file']}, sheet={result['sheet']}, "
                f"baris={result['row']}, label='{result['label']}', nilai_terakhir={result['last_value']}"
            )
            consecutive_missing = 0
        elif status == "zip_error":
            print(f"[{label}] ERROR buka zip: {result['detail']}")
            consecutive_missing += 1
        else:
            print(f"[{label}] TIDAK KETEMU - {result.get('detail', '')}")
            consecutive_missing += 1

        rows_out.append(
            {
                "tahun": tahun,
                "bulan": bulan_raw,
                "status": status,
                "file": result.get("file", ""),
                "sheet": result.get("sheet", ""),
                "row": result.get("row", ""),
                "label": result.get("label", ""),
                "last_value": result.get("last_value", ""),
                "source": result.get("source", ""),
                "detail": result.get("detail", ""),
            }
        )

        if args.stop_after_missing and consecutive_missing >= args.stop_after_missing:
            print(
                f"\nBerhenti: {consecutive_missing} bulan berturut-turut tidak ketemu (berhenti di {label})."
            )
            break

    os.makedirs(os.path.dirname(report_path) or ".", exist_ok=True)
    with open(report_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows_out[0].keys()))
        writer.writeheader()
        writer.writerows(rows_out)

    n_found = sum(1 for r in rows_out if r["status"] in ("found", "found_elsewhere"))
    n_missing = sum(1 for r in rows_out if r["status"] == "not_found")
    print(
        f"\nSelesai. Ketemu: {n_found}, tidak ketemu: {n_missing}, total dicek: {len(rows_out)}"
    )
    print(f"Laporan lengkap: {report_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
