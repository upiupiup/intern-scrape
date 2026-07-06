import re
import os
from datetime import datetime
import requests
import pandas as pd

API_URL = "https://satudata.kemnaker.go.id/api/v1/infographics/detail_data"
INFOGRAFIK_ID = "104"  # ganti sesuai id infografik periode terbaru
SOURCE_URL = f"https://satudata.kemnaker.go.id/infografik/{INFOGRAFIK_ID}"
OUTPUT_DIR = "data/raw/groupH"
OUTPUT_CSV = f"{OUTPUT_DIR}/31_ekonomi_pekerja_penuh_dan_tidak_penuh.csv"


def to_float(s):
    return float(s.replace(".", "").replace(",", "."))


def scrape_pekerja_penuh_tidak_penuh():
    resp = requests.post(API_URL, json={"id": INFOGRAFIK_ID}, timeout=15)
    resp.raise_for_status()
    payload = resp.json()

    if payload.get("status") != 200:
        raise ValueError(f"API gagal: {payload.get('message')}")

    data = payload["data"]
    judul = data["judul"]
    content = data["content"]

    # periode dari judul, misal "Profil Ketenagakerjaan Umum Agustus 2025"
    m_periode = re.search(r"([A-Za-z]+)\s+(\d{4})$", judul.strip())
    periode = f"{m_periode.group(1)} {m_periode.group(2)}" if m_periode else None

    # PUK + komposisi AK/BAK
    m_puk = re.search(
        r"sebanyak ([\d.,]+) juta orang; terdiri dari ([\d.,]+) persen Angkatan Kerja \(AK\) dan ([\d.,]+) persen Bukan Angkatan Kerja \(BAK\)",
        content,
    )
    puk_juta = to_float(m_puk.group(1))
    ak_persen = to_float(m_puk.group(2))
    bak_persen = to_float(m_puk.group(3))

    # Bekerja Penuh + Sementara Tidak Bekerja
    m_penuh = re.search(
        r"pekerja penuh .*? yaitu sebanyak ([\d.,]+) juta orang atau ([\d.,]+) persen",
        content,
    )
    bekerja_penuh_juta = to_float(m_penuh.group(1))
    bekerja_penuh_persen = to_float(m_penuh.group(2))

    # TPT nasional
    m_tpt = re.search(
        r"Tingkat Pengangguran Terbuka \(TPT\) secara nasional yaitu ([\d.,]+) persen",
        content,
    )
    tpt_persen = to_float(m_tpt.group(1))

    # derivasi: Bekerja Tidak Penuh = Bekerja - Bekerja Penuh
    # Bekerja = Angkatan Kerja x (1 - TPT%)
    ak_juta = puk_juta * ak_persen / 100
    bekerja_juta = ak_juta * (1 - tpt_persen / 100)
    bekerja_tidak_penuh_juta = round(bekerja_juta - bekerja_penuh_juta, 2)
    bekerja_tidak_penuh_persen = round(100 - bekerja_penuh_persen, 2)

    row = {
        "periode": periode,
        "puk_juta_orang": puk_juta,
        "ak_persen": ak_persen,
        "bak_persen": bak_persen,
        "bekerja_penuh_juta_orang": bekerja_penuh_juta,
        "bekerja_penuh_persen": bekerja_penuh_persen,
        "bekerja_tidak_penuh_juta_orang": bekerja_tidak_penuh_juta,
        "bekerja_tidak_penuh_persen": bekerja_tidak_penuh_persen,
        "bekerja_tidak_penuh_metode": "derivasi (bekerja - bekerja_penuh)",
        "tpt_persen": tpt_persen,
        "sumber_url": SOURCE_URL,
        "tanggal_scrape": datetime.now().strftime("%Y-%m-%d"),
    }
    return row


if __name__ == "__main__":
    row = scrape_pekerja_penuh_tidak_penuh()
    for k, v in row.items():
        print(f"{k}: {v}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    df = pd.DataFrame([row])
    df.to_csv(OUTPUT_CSV, index=False)
    print(f"\nCSV tersimpan di: {OUTPUT_CSV}")
