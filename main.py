import re
import pdfplumber
import pandas as pd
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from io import BytesIO

# =====================================================
# INIT FASTAPI
# =====================================================
app = FastAPI(title="SLIK Parser API")


# =====================================================
# EXTRACT TEXT PDF (ALL PAGES)
# =====================================================
def extract_text_from_pdf(file_bytes):
    text = ""

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += "\n" + t

    return text


# =====================================================
# AMBIL NAMA DEBITUR (DATA POKOK DEBITUR)
# =====================================================
def extract_nama_debitur(text):
    match = re.search(
        r"Nama Sesuai Identitas\s+([A-Z][A-Z\s'.-]+?)\s+NIK",
        text,
        re.IGNORECASE
    )
    return match.group(1).strip() if match else "Tidak Diketahui"


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

    return None


# =====================================================
# PARSE SLIK
# =====================================================
def parse_slik(full_text):
    facilities = []
    nama_debitur = extract_nama_debitur(full_text)

    # split blok fasilitas berdasarkan kode pelapor
    blocks = re.split(r"\n(?=\d{3,6}\s-\sPT)", full_text)

    for block in blocks:

        if "PT" not in block:
            continue

        # =====================
        # PELAPOR
        # =====================
        pelapor_match = re.search(r"(PT.*?Tbk|PT.*?Finance)", block)
        pelapor = pelapor_match.group(1).strip() if pelapor_match else ""

        # =====================
        # BAKI DEBET
        # =====================
        baki_match = re.search(r"Rp\s?[\d\.,]+", block)
        baki = baki_match.group() if baki_match else ""

        # =====================
        # KUALITAS
        # =====================
        kualitas_match = re.search(r"Kualitas\s([1-5]\s-\s.*)", block)
        kualitas = kualitas_match.group(1) if kualitas_match else ""

        # =====================
        # TUNGGAKAN
        # =====================
        tunggakan_match = re.search(r"Jumlah Hari Tunggakan\s(\d+)", block)
        tunggakan = tunggakan_match.group(1) if tunggakan_match else "0"

        # =====================
        # JENIS KREDIT
        # =====================
        jenis_match = re.search(r"Jenis Kredit/Pembiayaan\s(.+)", block)
        jenis = jenis_match.group(1).strip() if jenis_match else ""

        # =====================
        # PENGGUNAAN
        # =====================
        penggunaan_match = re.search(
            r"Jenis Penggunaan\s(Konsumsi|Modal Kerja|Investasi)",
            block
        )
        penggunaan = penggunaan_match.group(1) if penggunaan_match else ""

        # =====================
        # RESTRUKTUR
        # =====================
        freq_match = re.search(r"Frekuensi Restrukturisasi\s(\d+)", block)
        frekuensi = freq_match.group(1) if freq_match else ""

        tgl_match = re.search(
            r"Tanggal Restrukturisasi Akhir\s(\d{1,2}\s\w+\s\d{4})",
            block
        )
        tanggal_restruktur = tgl_match.group(1) if tgl_match else ""

        # =====================
        # KONDISI (FILTER)
        # =====================
        kondisi_match = re.search(r"Kondisi\s(.+)", block)
        kondisi_raw = kondisi_match.group(1).strip() if kondisi_match else ""
        kondisi = normalize_kondisi(kondisi_raw)

        if not kondisi:
            continue

        # =====================
        # BUNGA
        # =====================
        bunga_match = re.search(r"Suku Bunga/Imbalan\s([\d\,\.]+%)", block)
        bunga = bunga_match.group(1) if bunga_match else ""

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

    return facilities, nama_debitur


# =====================================================
# EXPORT EXCEL
# =====================================================
def generate_excel(data):
    df = pd.DataFrame(data)

    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return output


# =====================================================
# API ENDPOINT
# =====================================================
@app.post("/parse-slik")
async def parse_slik_api(file: UploadFile = File(...)):

    file_bytes = await file.read()

    full_text = extract_text_from_pdf(file_bytes)

    data, nama_debitur = parse_slik(full_text)

    if not data:
        return {"message": "Tidak ada fasilitas aktif / hapusbuku ditemukan"}

    excel_file = generate_excel(data)

    filename = f"SLIK_{nama_debitur.replace(' ', '_')}.xlsx"

    return StreamingResponse(
        excel_file,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )
