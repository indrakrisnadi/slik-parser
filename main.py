import re
import fitz
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse, JSONResponse
from openpyxl import Workbook
from io import BytesIO

# ==============================
# INIT FASTAPI
# ==============================
app = FastAPI()


import re
import pdfplumber
import pandas as pd


# =====================================================
# EXTRACT TEXT PDF
# =====================================================
def extract_text_from_pdf(file_path):
    text = ""

    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += "\n" + t

    return text


# =====================================================
# AMBIL NAMA DEBITUR (DARI DATA POKOK DEBITUR)
# =====================================================
def extract_nama_debitur(text):
    match = re.search(
        r"Nama Sesuai Identitas\s*\n?\s*([A-Z\s'.-]+)",
        text,
        re.IGNORECASE
    )

    if match:
        return match.group(1).strip()

    return "Tidak Diketahui"


# =====================================================
# NORMALISASI STATUS FASILITAS
# =====================================================
def normalize_kondisi(text):
    if not text:
        return None

    t = text.lower()

    if "aktif" in t:
        return "Fasilitas Aktif"

    if "hapus" in t:
        return "Dihapusbukukan"

    return None   # selain ini diabaikan


# =====================================================
# PARSE SLIK
# =====================================================
def parse_slik(full_text):
    facilities = []

    nama_debitur = extract_nama_debitur(full_text)

    # split blok fasilitas (lebih stabil)
    blocks = re.split(r"\n\d{3}\s-\sPT", full_text)

    for block in blocks:

        if "Tbk" not in block:
            continue

        block = "PT " + block

        # -----------------------------
        # pelapor
        # -----------------------------
        pelapor_match = re.search(r"(PT.*?Tbk)", block)
        pelapor = pelapor_match.group(1).strip() if pelapor_match else ""

        # -----------------------------
        # baki debet
        # -----------------------------
        baki_match = re.search(r"Rp\s?[\d\.,]+", block)
        baki = baki_match.group().strip() if baki_match else ""

        # -----------------------------
        # kualitas
        # -----------------------------
        kualitas_match = re.search(r"Kualitas\s([1-5]\s-\s.*)", block)
        kualitas = kualitas_match.group(1).strip() if kualitas_match else ""

        # -----------------------------
        # tunggakan
        # -----------------------------
        tunggakan_match = re.search(r"Jumlah Hari Tunggakan\s(\d+)", block)
        tunggakan = tunggakan_match.group(1) if tunggakan_match else "0"

        # -----------------------------
        # jenis kredit
        # -----------------------------
        jenis_match = re.search(r"Jenis Kredit/Pembiayaan\s(.+)", block)
        if jenis_match:
            jenis = re.sub(r"Nilai Proyek.*", "", jenis_match.group(1)).strip()
        else:
            jenis = ""

        # -----------------------------
        # penggunaan
        # -----------------------------
        penggunaan_match = re.search(
            r"Jenis Penggunaan\s(Konsumsi|Modal Kerja|Investasi)",
            block
        )
        penggunaan = penggunaan_match.group(1) if penggunaan_match else ""

        # -----------------------------
        # restruktur
        # -----------------------------
        freq_match = re.search(r"Frekuensi Restrukturisasi\s(\d+)", block)
        frekuensi = freq_match.group(1) if freq_match else ""

        tgl_match = re.search(
            r"Tanggal Restrukturisasi Akhir\s(\d{1,2}\s\w+\s\d{4})",
            block
        )
        tanggal_restruktur = tgl_match.group(1) if tgl_match else ""

        # -----------------------------
        # kondisi (FILTER PENTING)
        # -----------------------------
        kondisi_match = re.search(r"Kondisi\s(.+)", block)
        kondisi_raw = kondisi_match.group(1).strip() if kondisi_match else ""
        kondisi = normalize_kondisi(kondisi_raw)

        if not kondisi:
            continue

        # -----------------------------
        # bunga
        # -----------------------------
        bunga_match = re.search(r"Suku Bunga/Imbalan\s([\d\,\.]+%)", block)
        bunga = bunga_match.group(1) if bunga_match else ""

        # -----------------------------
        # save
        # -----------------------------
        facilities.append({
            "Nama Debitur": nama_debitur,
            "Pelapor": pelapor,
            "Baki Debet": baki,
            "Kualitas": kualitas,
            "Jumlah Hari Tunggakan": tunggakan,
            "Jenis Kredit": jenis,
            "Jenis Penggunaan": penggunaan,
            "Frekuensi Restrukturisasi": frekuensi,
            "Tanggal Restrukturisasi Akhir": tanggal_restruktur,
            "Kondisi": kondisi,
            "Suku Bunga": bunga
        })

    return facilities


# =====================================================
# EXPORT EXCEL
# =====================================================
def export_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)


# =====================================================
# MAIN
# =====================================================
if __name__ == "__main__":

    pdf_file = "slik.pdf"          # ganti path
    output_excel = "hasil_slik.xlsx"

    print("Reading PDF...")
    text = extract_text_from_pdf(pdf_file)

    print("Parsing SLIK...")
    data = parse_slik(text)

    print("Exporting Excel...")
    export_to_excel(data, output_excel)

    print("DONE ✔")
    print("Total fasilitas:", len(data))
