import re
import pdfplumber
import pandas as pd
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse, JSONResponse
from io import BytesIO

# ==============================
# INIT FASTAPI
# ==============================
app = FastAPI()


# =====================================================
# EXTRACT TEXT PDF (FROM BYTES)
# =====================================================
def extract_text_from_pdf_bytes(file_bytes):
    text = ""

    with pdfplumber.open(BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += "\n" + t

    return text


# =====================================================
# AMBIL NAMA DEBITUR
# =====================================================
def extract_nama_debitur(text):
    match = re.search(
        r"Nama Sesuai Identitas\s*\n?\s*([A-Z\s'.-]+)",
        text,
        re.IGNORECASE
    )
    return match.group(1).strip() if match else "Tidak Diketahui"


# =====================================================
# NORMALISASI STATUS
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

    blocks = re.split(r"\n\d{3}\s-\sPT", full_text)

    for block in blocks:

        if "Tbk" not in block:
            continue

        block = "PT " + block

        pelapor_match = re.search(r"(PT.*?Tbk)", block)
        pelapor = pelapor_match.group(1).strip() if pelapor_match else ""

        baki_match = re.search(r"Rp\s?[\d\.,]+", block)
        baki = baki_match.group().strip() if baki_match else ""

        kualitas_match = re.search(r"Kualitas\s([1-5]\s-\s.*)", block)
        kualitas = kualitas_match.group(1).strip() if kualitas_match else ""

        tunggakan_match = re.search(r"Jumlah Hari Tunggakan\s(\d+)", block)
        tunggakan = tunggakan_match.group(1) if tunggakan_match else "0"

        jenis_match = re.search(r"Jenis Kredit/Pembiayaan\s(.+)", block)
        jenis = re.sub(r"Nilai Proyek.*", "", jenis_match.group(1)).strip() if jenis_match else ""

        penggunaan_match = re.search(
            r"Jenis Penggunaan\s(Konsumsi|Modal Kerja|Investasi)",
            block
        )
        penggunaan = penggunaan_match.group(1) if penggunaan_match else ""

        freq_match = re.search(r"Frekuensi Restrukturisasi\s(\d+)", block)
        frekuensi = freq_match.group(1) if freq_match else ""

        tgl_match = re.search(
            r"Tanggal Restrukturisasi Akhir\s(\d{1,2}\s\w+\s\d{4})",
            block
        )
        tanggal_restruktur = tgl_match.group(1) if tgl_match else ""

        kondisi_match = re.search(r"Kondisi\s(.+)", block)
        kondisi_raw = kondisi_match.group(1).strip() if kondisi_match else ""
        kondisi = normalize_kondisi(kondisi_raw)

        if not kondisi:
            continue

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

    return facilities


# =====================================================
# ROOT TEST
# =====================================================
@app.get("/")
def root():
    return {"status": "SLIK parser running"}


# =====================================================
# UPLOAD PDF → DOWNLOAD EXCEL
# =====================================================
@app.post("/upload/")
async def upload_file(file: UploadFile = File(...)):
    try:
        content = await file.read()

        text = extract_text_from_pdf_bytes(content)
        data = parse_slik(text)

        df = pd.DataFrame(data)

        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=hasil_slik.xlsx"}
        )

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
