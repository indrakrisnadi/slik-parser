import re
import pdfplumber
from fastapi import FastAPI, UploadFile, File
import shutil
import os
from openpyxl import Workbook
from fastapi.responses import FileResponse

app = FastAPI()


# ==============================
# FUNCTION: PARSE SLIK
# ==============================
def parse_slik(text):
    facilities = []

    # ========================
    # Ambil Nama Debitur
    # ========================
    nama_match = re.search(r"\n([A-Z\s]+)\s+LAKI-LAKI|PEREMPUAN", text)
    nama_debitur = nama_match.group(1).strip() if nama_match else "Tidak Diketahui"

    # ========================
    # Split per fasilitas
    # ========================
    pattern = r"\n\d{3}\s-\sPT\s.*?Tbk.*?(?=\n\d{3}\s-\sPT|\Z)"
    matches = re.finditer(pattern, text, re.DOTALL)

    for match in matches:
        block = match.group()

        # ========================
        # Pelapor
        # ========================
        pelapor_match = re.search(r"\d{3}\s-\s(PT.*?Tbk)", block)
        pelapor = pelapor_match.group(1).strip() if pelapor_match else ""

        # ========================
        # Baki Debet
        # ========================
        baki_match = re.search(r"Rp\s[\d\.,]+", block)
        baki = baki_match.group().strip() if baki_match else ""

        # ========================
        # Kualitas
        # ========================
        kualitas_match = re.search(r"Kualitas\s([1-5]\s-\s.*)", block)
        kualitas = kualitas_match.group(1).strip() if kualitas_match else ""

        # ========================
        # Jumlah Hari Tunggakan
        # ========================
        tunggakan_match = re.search(r"Jumlah Hari Tunggakan\s(\d+)", block)
        tunggakan = tunggakan_match.group(1).strip() if tunggakan_match else "0"

        # ========================
        # Jenis Kredit (bersihkan Nilai Proyek)
        # ========================
        jenis_match = re.search(r"Jenis Kredit/Pembiayaan\s(.+)", block)
        if jenis_match:
            jenis = jenis_match.group(1).strip()
            jenis = re.sub(r"Nilai Proyek.*", "", jenis).strip()
        else:
            jenis = ""

        # ========================
        # Jenis Penggunaan
        # ========================
        penggunaan_match = re.search(
            r"Jenis Penggunaan\s(Konsumsi|Modal Kerja|Investasi)",
            block
        )
        penggunaan = penggunaan_match.group(1).strip() if penggunaan_match else ""

        # ========================
        # Frekuensi Restrukturisasi
        # ========================
        freq_match = re.search(r"Frekuensi Restrukturisasi\s(\d+)", block)
        frekuensi = freq_match.group(1).strip() if freq_match else ""

        # ========================
        # Tanggal Restrukturisasi Akhir
        # ========================
        tgl_match = re.search(
            r"Tanggal Restrukturisasi Akhir\s(\d{1,2}\s\w+\s\d{4})",
            block
        )
        tanggal_restruktur = tgl_match.group(1).strip() if tgl_match else ""

        # ========================
        # Kondisi (FILTER HANYA 2)
        # ========================
        kondisi_match = re.search(r"Kondisi\s(.+)", block)
        raw_kondisi = kondisi_match.group(1).strip() if kondisi_match else ""

        if "Fasilitas Aktif" in raw_kondisi:
            kondisi = "Fasilitas Aktif"
        elif "Dihapusbukukan" in raw_kondisi:
            kondisi = "Dihapusbukukan"
        else:
            # Skip fasilitas jika Sudah Lunas atau lainnya
            continue

        # ========================
        # Suku Bunga
        # ========================
        bunga_match = re.search(r"Suku Bunga/Imbalan\s([\d\,\.]+%)", block)
        bunga = bunga_match.group(1).strip() if bunga_match else ""

        # ========================
        # Append Data
        # ========================
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


# ==============================
# UPLOAD ENDPOINT
# ==============================
@app.post("/upload/")
async def upload_file(file: UploadFile = File(...)):

    file_location = f"temp_{file.filename}"

    with open(file_location, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    full_text = ""

    with pdfplumber.open(file_location) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += "\n" + text

    os.remove(file_location)

    parsed_data = parse_slik(full_text)

    # ==========================
    # CREATE EXCEL
    # ==========================
    wb = Workbook()
    ws = wb.active
    ws.title = "SLIK Result"

    headers = [
        "Nama Debitur",
        "Pelapor",
        "Baki Debet",
        "Kualitas",
        "Jumlah Hari Tunggakan",
        "Jenis Kredit",
        "Jenis Penggunaan",
        "Frekuensi Restrukturisasi",
        "Tanggal Restrukturisasi Akhir",
        "Kondisi",
        "Suku Bunga"
    ]

    ws.append(headers)

    for item in parsed_data:
        ws.append([
            item["Nama Debitur"],
            item["Pelapor"],
            item["Baki Debet"],
            item["Kualitas"],
            item["Jumlah Hari Tunggakan"],
            item["Jenis Kredit"],
            item["Jenis Penggunaan"],
            item["Frekuensi Restrukturisasi"],
            item["Tanggal Restrukturisasi Akhir"],
            item["Kondisi"],
            item["Suku Bunga"]
        ])

    output_file = "hasil_slik.xlsx"
    wb.save(output_file)

    return FileResponse(
        path=output_file,
        filename="hasil_slik.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
