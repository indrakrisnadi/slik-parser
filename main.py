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


# ==============================
# FUNCTION PARSE SLIK
# ==============================
def parse_slik(text):
    facilities = []

    # =========================
    # NORMALIZE TEXT (SUPER PENTING)
    # =========================
    text = re.sub(r'\r', '', text)
    text = re.sub(r'\s+\n', '\n', text)
    text = re.sub(r'\n+', '\n', text)

    # =========================
    # NAMA DEBITUR (MULTI FORMAT)
    # =========================
    nama_match = re.search(
        r"\n([A-Z\s]+)\s+(?:LAKI-LAKI|PEREMPUAN)",
        text
    )

    if not nama_match:
        nama_match = re.search(r"Nama Debitur\s*:\s*(.+)", text)

    nama_debitur = (
        nama_match.group(1).strip()
        if nama_match and nama_match.group(1)
        else "Tidak Diketahui"
    )

    # =========================
    # BLOCK FASILITAS (ANTI KEPUTUS HALAMAN)
    # =========================
    pattern = r"Kredit/Pembiayaan.*?(?=Kredit/Pembiayaan|\Z)"
    matches = re.finditer(pattern, text, re.DOTALL | re.IGNORECASE)

    for match in matches:
        block = match.group()

        # =========================
        # PELAPOR (TIDAK DIUBAH SESUAI REQUEST)
        # =========================
        pelapor_match = re.search(r"\d{3}\s-\s(.+)", block)
        pelapor = (
            pelapor_match.group(1).split("\n")[0].strip()
            if pelapor_match else ""
        )

        # =========================
        # BAKI DEBET
        # =========================
        baki_match = re.search(r"Rp\s?[\d\.,]+", block)
        baki = baki_match.group().strip() if baki_match else ""

        # =========================
        # KUALITAS
        # =========================
        kualitas_match = re.search(r"Kualitas\s([1-5]\s-\s.*)", block)
        kualitas = kualitas_match.group(1).strip() if kualitas_match else ""

        # =========================
        # TUNGGAKAN
        # =========================
        tunggakan_match = re.search(r"Jumlah Hari Tunggakan\s(\d+)", block)
        tunggakan = tunggakan_match.group(1).strip() if tunggakan_match else "0"

        # =========================
        # JENIS KREDIT
        # =========================
        jenis_match = re.search(r"Jenis Kredit/Pembiayaan\s(.+)", block)
        jenis = jenis_match.group(1).strip() if jenis_match else ""

        # =========================
        # PENGGUNAAN
        # =========================
        penggunaan_match = re.search(
            r"Jenis Penggunaan\s(Konsumsi|Modal Kerja|Investasi)",
            block
        )
        penggunaan = penggunaan_match.group(1).strip() if penggunaan_match else ""

        # =========================
        # FREKUENSI RESTRUK
        # =========================
        freq_match = re.search(r"Frekuensi Restrukturisasi\s(\d+)", block)
        frekuensi = freq_match.group(1).strip() if freq_match else ""

        # =========================
        # TANGGAL RESTRUK AKHIR
        # =========================
        tgl_match = re.search(
            r"Tanggal Restrukturisasi Akhir\s(\d{1,2}\s\w+\s\d{4})",
            block
        )
        tanggal_restruktur = tgl_match.group(1).strip() if tgl_match else ""

        # =========================
        # KONDISI
        # =========================
        kondisi_match = re.search(r"Kondisi\s(.+)", block)
        raw_kondisi = kondisi_match.group(1).strip() if kondisi_match else ""

        if re.search(r"Aktif|Lancar|Berjalan", raw_kondisi, re.I):
            kondisi = "Fasilitas Aktif"
        elif re.search(r"Hapus|Write Off|Closed|Selesai", raw_kondisi, re.I):
            kondisi = "Dihapusbukukan"
        else:
            kondisi = raw_kondisi

        # =========================
        # SUKU BUNGA
        # =========================
        bunga_match = re.search(r"Suku Bunga/Imbalan\s([\d\,\.]+%)", block)
        bunga = bunga_match.group(1).strip() if bunga_match else ""

        # =========================
        # APPEND
        # =========================
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
# ROOT TEST
# ==============================
@app.get("/")
def root():
    return {"status": "SLIK parser running"}


# ==============================
# UPLOAD PDF → DOWNLOAD EXCEL
# ==============================
@app.post("/upload/")
async def upload_file(file: UploadFile = File(...)):
    try:
        content = await file.read()

        # ========= EXTRACT PDF TEXT (MERGE HALAMAN KUAT) =========
        doc = fitz.open(stream=content, filetype="pdf")

        text = ""
        for page in doc:
            text += "\n===PAGE_BREAK===\n"
            text += page.get_text()

        doc.close()

        print("===== PDF TEXT SAMPLE =====")
        print(text[:1000])
        print("===== END =====")

        # ========= PARSE =========
        data = parse_slik(text)

        # ========= CREATE EXCEL =========
        wb = Workbook()
        ws = wb.active
        ws.title = "SLIK Result"

        headers = list(data[0].keys()) if data else [
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

        for row in data:
            ws.append(list(row.values()))

        stream = BytesIO()
        wb.save(stream)
        stream.seek(0)

        return StreamingResponse(
            stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": "attachment; filename=hasil_slik.xlsx"
            }
        )

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
