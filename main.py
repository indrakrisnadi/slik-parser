import re
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from openpyxl import Workbook
from io import BytesIO
import fitz  # PyMuPDF

app = FastAPI()


# ==============================
# FUNCTION: PARSE SLIK
# ==============================
def parse_slik(text):
    facilities = []

nama_match = re.search(
    r"(?<=\n)([A-Z][A-Z\s]{3,})\s+(?:LAKI-LAKI|PEREMPUAN)",
    text
)

nama_debitur = (
    nama_match.group(1).strip()
    if nama_match and nama_match.group(1)
    else "Tidak Diketahui"
)


    pattern = r"\n\d{3}\s-\sPT\s.*?Tbk.*?(?=\n\d{3}\s-\sPT|\Z)"
    matches = re.finditer(pattern, text, re.DOTALL)

    for match in matches:
        block = match.group()

        pelapor_match = re.search(r"\d{3}\s-\s(PT.*?Tbk)", block)
        pelapor = pelapor_match.group(1).strip() if pelapor_match else ""

        baki_match = re.search(r"Rp\s[\d\.,]+", block)
        baki = baki_match.group().strip() if baki_match else ""

        kualitas_match = re.search(r"Kualitas\s([1-5]\s-\s.*)", block)
        kualitas = kualitas_match.group(1).strip() if kualitas_match else ""

        tunggakan_match = re.search(r"Jumlah Hari Tunggakan\s(\d+)", block)
        tunggakan = tunggakan_match.group(1).strip() if tunggakan_match else "0"

        jenis_match = re.search(r"Jenis Kredit/Pembiayaan\s(.+)", block)
        if jenis_match:
            jenis = re.sub(r"Nilai Proyek.*", "", jenis_match.group(1)).strip()
        else:
            jenis = ""

        penggunaan_match = re.search(
            r"Jenis Penggunaan\s(Konsumsi|Modal Kerja|Investasi)",
            block
        )
        penggunaan = penggunaan_match.group(1).strip() if penggunaan_match else ""

        freq_match = re.search(r"Frekuensi Restrukturisasi\s(\d+)", block)
        frekuensi = freq_match.group(1).strip() if freq_match else ""

        tgl_match = re.search(
            r"Tanggal Restrukturisasi Akhir\s(\d{1,2}\s\w+\s\d{4})",
            block
        )
        tanggal_restruktur = tgl_match.group(1).strip() if tgl_match else ""

        kondisi_match = re.search(r"Kondisi\s(.+)", block)
        raw_kondisi = kondisi_match.group(1).strip() if kondisi_match else ""

        if "Fasilitas Aktif" in raw_kondisi:
            kondisi = "Fasilitas Aktif"
        elif "Dihapusbukukan" in raw_kondisi:
            kondisi = "Dihapusbukukan"
        else:
            continue

        bunga_match = re.search(r"Suku Bunga/Imbalan\s([\d\,\.]+%)", block)
        bunga = bunga_match.group(1).strip() if bunga_match else ""

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

    contents = await file.read()

    # ===== EXTRACT PDF TEXT =====
    full_text = ""
    doc = fitz.open(stream=contents, filetype="pdf")

    for page in doc:
        text = page.get_text()
        if text:
            full_text += "\n" + text

    doc.close()

    parsed_data = parse_slik(full_text)

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

    excel_stream = BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0)

    return StreamingResponse(
        excel_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=hasil_slik.xlsx"}
    )
