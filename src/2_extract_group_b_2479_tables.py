"""
Extract tabel Group B dari PDF BPS Press Release 2479/2497 ke CSV rapi.

Target dari skema:
- usecase_ekonomi.ekonomi_tpt_by_education_level
- usecase_ekonomi.ekonomi_rata_rata_upah_buruh_per_provinsi

Contoh run:
python src/extract_groupB_2479_tables.py \
  --pdf "data/raw/groupB/2497.pdf" \
  --schema "data/raw/aufi_Copy of Skema Data Ekonomi.xlsm" \
  --output-dir "data/raw/groupB"

Kalau nama PDF yang benar adalah 2479.pdf, ganti argumen --pdf saja.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Callable

import pandas as pd
import pdfplumber


TARGET_TABLES = {
    "usecase_ekonomi.ekonomi_tpt_by_education_level": "extract_tpt_by_education_level",
    "usecase_ekonomi.ekonomi_rata_rata_upah_buruh_per_provinsi": "extract_wage_by_province",
}

FALLBACK_SCHEMA_NO = {
    "usecase_ekonomi.ekonomi_tpt_by_education_level": "37",
    "usecase_ekonomi.ekonomi_rata_rata_upah_buruh_per_provinsi": "38",
}


def clean_cell(value) -> str:
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\n", " ")
    text = text.replace("–", "-").replace("−", "-")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_id_number(value):
    text = clean_cell(value)
    if text == "" or text.lower() in {"nan", "none"}:
        return pd.NA

    text = re.sub(r"[^0-9,\.\-]", "", text)
    if text in {"", "-"}:
        return pd.NA

    if "," in text:
        text = text.replace(".", "").replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return pd.NA

    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+", text):
        try:
            return int(text.replace(".", ""))
        except ValueError:
            return pd.NA

    try:
        number = float(text)
        return int(number) if number.is_integer() else number
    except ValueError:
        return pd.NA


def normalize_schema_no(value) -> str:
    text = clean_cell(value)
    if not text:
        return ""
    try:
        return str(int(float(text)))
    except ValueError:
        return re.sub(r"\D+", "", text) or text


def table_db_to_file_stem(table_db: str) -> str:
    return table_db.split(".")[-1].strip()


def read_schema_meta(schema_path: Path) -> dict[str, dict[str, str]]:
    if not schema_path.exists():
        raise FileNotFoundError(f"File skema tidak ditemukan: {schema_path}")

    sheets = pd.read_excel(schema_path, sheet_name=None, dtype=str, engine="openpyxl")

    candidate_df = None
    if "Data Lengkap" in sheets:
        candidate_df = sheets["Data Lengkap"]
    else:
        for _, df in sheets.items():
            cols = {str(c).strip().lower(): c for c in df.columns}
            if "no" in cols and "tabel db" in cols:
                candidate_df = df
                break

    if candidate_df is None:
        raise ValueError("Tidak menemukan sheet dengan kolom 'No' dan 'Tabel DB'.")

    df = candidate_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    if "No" not in df.columns or "Tabel DB" not in df.columns:
        raise ValueError("Kolom wajib 'No' dan/atau 'Tabel DB' tidak ada di skema.")

    meta: dict[str, dict[str, str]] = {}
    for _, row in df.iterrows():
        table_db = clean_cell(row.get("Tabel DB"))
        if not table_db:
            continue
        meta[table_db] = {
            "no": normalize_schema_no(row.get("No")),
            "table_db": table_db,
            "file_stem": table_db_to_file_stem(table_db),
        }

    missing = [t for t in TARGET_TABLES if t not in meta]
    if missing:
        raise ValueError(f"Tabel target tidak ditemukan di skema: {missing}")

    return {t: meta[t] for t in TARGET_TABLES}


def fallback_schema_meta() -> dict[str, dict[str, str]]:
    return {
        table_db: {
            "no": FALLBACK_SCHEMA_NO[table_db],
            "table_db": table_db,
            "file_stem": table_db_to_file_stem(table_db),
        }
        for table_db in TARGET_TABLES
    }


def find_page(
    pdf: pdfplumber.PDF,
    must_contain: list[str],
    must_not_contain: list[str] | None = None,
) -> int:
    must_not_contain = must_not_contain or []
    for idx, page in enumerate(pdf.pages):
        text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
        text_low = text.lower()
        if all(k.lower() in text_low for k in must_contain) and not any(
            k.lower() in text_low for k in must_not_contain
        ):
            return idx
    raise ValueError(f"Tidak menemukan halaman dengan keyword: {must_contain}")


def first_table_on_page(pdf: pdfplumber.PDF, page_idx: int) -> list[list[str]]:
    page = pdf.pages[page_idx]
    tables = page.extract_tables()
    if not tables:
        raise ValueError(f"Tidak ada tabel terdeteksi di halaman PDF {page_idx + 1}.")
    return max(tables, key=lambda t: len(t) * max((len(r) for r in t), default=0))


def extract_tpt_by_education_level(pdf_path: Path) -> pd.DataFrame:
    with pdfplumber.open(pdf_path) as pdf:
        page_idx = find_page(
            pdf,
            must_contain=[
                "Tingkat Pengangguran Terbuka Menurut Karakteristik",
                "TPT Menurut Pendidikan Tertinggi yang Ditamatkan",
            ],
        )
        table = first_table_on_page(pdf, page_idx)

    rows = [[clean_cell(c) for c in row] for row in table]

    captured: list[list[str]] = []
    is_capture = False
    for row in rows:
        first = clean_cell(row[0]) if row else ""
        first_low = first.lower()

        if "tpt menurut pendidikan" in first_low:
            is_capture = True
            continue

        if not is_capture:
            continue

        if first_low.startswith("tpt menurut") and "pendidikan" not in first_low:
            break

        if not first or first.startswith("("):
            continue
        if len(row) < 7:
            continue
        if all(clean_cell(x) == "" for x in row[1:7]):
            continue

        captured.append(row[:7])

    if not captured:
        raise ValueError("Bagian TPT menurut pendidikan tidak berhasil diekstrak.")

    df = pd.DataFrame(
        captured,
        columns=[
            "pendidikan_tertinggi",
            "agustus_2023",
            "februari_2024",
            "agustus_2024",
            "februari_2025",
            "agustus_2025",
            "perubahan_ags_2024_ags_2025_persen_poin",
        ],
    )
    df["pendidikan_tertinggi"] = (
        df["pendidikan_tertinggi"]
        .map(clean_cell)
        .str.replace(r"^[-\s]+", "", regex=True)
    )

    numeric_cols = [c for c in df.columns if c != "pendidikan_tertinggi"]
    for col in numeric_cols:
        df[col] = df[col].map(parse_id_number)

    return df


def extract_wage_by_province(pdf_path: Path) -> pd.DataFrame:
    with pdfplumber.open(pdf_path) as pdf:
        page_idx = find_page(
            pdf,
            must_contain=["Rata-Rata Upah Buruh Menurut Provinsi", "Agustus 2023"],
            must_not_contain=["Jenis Kelamin"],
        )
        table = first_table_on_page(pdf, page_idx)

    rows = [[clean_cell(c) for c in row] for row in table]
    data_rows: list[list[str]] = []

    for row in rows:
        if len(row) < 6:
            continue
        first = clean_cell(row[0])
        if not first or first.lower() == "provinsi" or first.startswith("("):
            continue
        values = [parse_id_number(x) for x in row[1:6]]
        if sum(pd.notna(v) for v in values) < 5:
            continue
        data_rows.append([first, *values])

    if not data_rows:
        raise ValueError("Lampiran 6 upah buruh per provinsi tidak berhasil diekstrak.")

    df = pd.DataFrame(
        data_rows,
        columns=[
            "provinsi",
            "agustus_2023",
            "februari_2024",
            "agustus_2024",
            "februari_2025",
            "agustus_2025",
        ],
    )

    return df


def save_output(df: pd.DataFrame, meta: dict[str, str], output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    filename = f"{meta['no']}_{meta['file_stem']}.csv"
    path = output_dir / filename
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract tabel BPS 2479/2497 PDF ke CSV sesuai skema DB."
    )
    parser.add_argument(
        "--pdf",
        default="data/raw/groupB/2497.pdf",
        help="Path PDF BPS, misal data/raw/groupB/2479.pdf",
    )
    parser.add_argument(
        "--schema",
        default="data/raw/aufi_Copy of Skema Data Ekonomi.xlsm",
        help="Path file skema .xlsm",
    )
    parser.add_argument(
        "--output-dir", default="data/raw/groupB", help="Folder output CSV"
    )
    parser.add_argument(
        "--no-schema",
        action="store_true",
        help="Pakai nomor fallback 37/38 kalau file skema tidak tersedia.",
    )
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    schema_path = Path(args.schema)
    output_dir = Path(args.output_dir)

    if not pdf_path.exists():
        raise FileNotFoundError(
            f"PDF tidak ditemukan: {pdf_path}\n"
            "Cek lagi nama file-nya. Untuk BRS ini biasanya ID-nya 2479, bukan 2497."
        )

    if args.no_schema:
        schema_meta = fallback_schema_meta()
    else:
        schema_meta = read_schema_meta(schema_path)

    extractors: dict[str, Callable[[Path], pd.DataFrame]] = {
        "usecase_ekonomi.ekonomi_tpt_by_education_level": extract_tpt_by_education_level,
        "usecase_ekonomi.ekonomi_rata_rata_upah_buruh_per_provinsi": extract_wage_by_province,
    }

    print(f"PDF        : {pdf_path}")
    print(f"Output dir : {output_dir}")
    print()

    for table_db, extractor in extractors.items():
        df = extractor(pdf_path)
        out_path = save_output(df, schema_meta[table_db], output_dir)
        print(f"OK  {out_path}  ({len(df)} baris x {len(df.columns)} kolom)")


if __name__ == "__main__":
    main()
