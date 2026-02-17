import re
import fitz  # PyMuPDF
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse

app = FastAPI()


# ================================
# ROOT CHECK
# ================================
@app.get("/")
def root():
    return {"status": "API hidup"}


# ================================
# UPLOAD FILE ENDPOINT
# ================================
@app.post("/upload/")
async def upload_file(file: UploadFile = File(...)):
    try:
        content = await file.read()

        # BUKA PDF
        doc = fitz.open(stream=content, filetype="pdf")

        text = ""
        for page in doc:
            text += page.get_text()

        doc.close()

        # DEBUG
        print("===== TEXT PDF =====")
        print(text[:1000])
        print("===== END =====")

        data = parse_slik(text)

        return JSONResponse(content=data)

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": str(e)}
        )


# ================================
# PARSER SLIK
# ================================
def parse_slik(text):
    facilities = []

    # ========================
    # Ambil Nama Debitur
    # ========================
    nama_match = re.search(
        r"\n([A-Z\s]+)\s+(?:LAKI-LAKI|PEREMPUAN)",
        text
    )

    nama_debitur = (
        nama_match.group(1).strip()
        if nama_match and nama_match.group(1)
        else "Tidak Diketahui"
    )

    # ========================
    # Split per fasilitas
    # ========================
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
            jenis = jenis_match.group(1).strip()
            jenis = re.sub(r"Nilai Proyek.*", "", jenis).strip()
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
