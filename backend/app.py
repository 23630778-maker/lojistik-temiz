from flask import Flask, render_template, request, redirect, url_for, flash
from datetime import datetime
import os
import io
import json
import openpyxl
from openpyxl import Workbook, load_workbook
import shutil
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

app = Flask(__name__)
app.secret_key = "supersecretkey"

# -------------------------
# Dosya yolları
# -------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_LOCAL = os.path.join(BASE_DIR, "lojistik.xlsx")  # Lokal kopya
EXCEL_FILE_ONEDRIVE = os.path.join(BASE_DIR, "OneDrive_lojistik.xlsx")  # Yedek kopya

# -------------------------
# Google Drive ayarları
# -------------------------
EXCEL_FILE_DRIVE_ID = "1Rvg3nQkHsVjh9QicnU5ViYvzJm1EwO8T"  # Drive Excel ID
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# -------------------------
# Google Drive servis fonksiyonları
# -------------------------
def get_drive_service():
    """
    GOOGLE_CREDENTIALS_JSON env var'ı varsa onu,
    yoksa backend/service_account.json dosyasını kullanır.
    """
    credentials_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")

    if credentials_json:
        credentials_info = json.loads(credentials_json)
    else:
        json_path = os.path.join(BASE_DIR, "service_account.json")
        if not os.path.exists(json_path):
            raise Exception("GOOGLE_CREDENTIALS_JSON yok ve 'service_account.json' bulunamadı")
        with open(json_path, "r", encoding="utf-8") as f:
            credentials_info = json.load(f)

    # Private key format hatası varsa düzeltelim
    if "private_key" in credentials_info and "\\n" in credentials_info["private_key"]:
        credentials_info["private_key"] = credentials_info["private_key"].replace("\\n", "\n")

    creds = service_account.Credentials.from_service_account_info(credentials_info, scopes=SCOPES)
    service = build('drive', 'v3', credentials=creds)
    return service

def download_excel(service, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)
    wb = openpyxl.load_workbook(fh)
    return wb

def upload_excel(service, file_id, wb):
    fh = io.BytesIO()
    wb.save(fh)
    fh.seek(0)
    media = MediaIoBaseUpload(
        fh,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    service.files().update(fileId=file_id, media_body=media).execute()

# -------------------------
# Ana route
# -------------------------
@app.route("/", methods=["GET", "POST"])
def form():
    if request.method == "POST":
        try:
            # Form verilerini al
            tarih = request.form.get("tarih") or datetime.now().strftime("%Y-%m-%d")
            iscikissaat = request.form.get("iscikissaat") or "00:00"
            plaka = request.form.get("plaka")
            cikiskm = float(request.form.get("cikiskm") or 0)
            kumgirissaat = request.form.get("kumgirissaat") or "00:00"
            giriskm = float(request.form.get("giriskm") or 0)
            kumcikissaat = request.form.get("kumcikissaat") or "00:00"
            isletmegiriskm = float(request.form.get("isletmegiriskm") or 0)
            isletmegirissaat = request.form.get("isletmegirissaat") or "00:00"
            farkkm = giriskm - cikiskm
            uretici = request.form.get("uretici")
            ureticikm = float(request.form.get("ureticikm") or 0)
            tonaj = int(request.form.get("tonaj") or 0)

            # -------------------------
            # Lokal Excel kaydı
            # -------------------------
            if not os.path.exists(EXCEL_FILE_LOCAL):
                wb = Workbook()
                ws = wb.active
                ws.append([
                    "tarih","iscikissaat","plaka","cikiskm","kumgirissaat",
                    "giriskm","kumcikissaat","isletmegiriskm","isletmegirissaat",
                    "farkkm","uretici","ureticikm","tonaj"
                ])
                wb.save(EXCEL_FILE_LOCAL)

            wb = load_workbook(EXCEL_FILE_LOCAL)
            ws = wb.active
            ws.append([
                tarih, iscikissaat, plaka, cikiskm, kumgirissaat,
                giriskm, kumcikissaat, isletmegiriskm, isletmegirissaat,
                farkkm, uretici, ureticikm, tonaj
            ])
            wb.save(EXCEL_FILE_LOCAL)
            shutil.copy(EXCEL_FILE_LOCAL, EXCEL_FILE_ONEDRIVE)

            # -------------------------
            # Google Drive kaydı
            # -------------------------
            service = get_drive_service()
            wb_drive = download_excel(service, EXCEL_FILE_DRIVE_ID)
            ws_drive = wb_drive.active
            ws_drive.append([
                tarih, iscikissaat, plaka, cikiskm, kumgirissaat,
                giriskm, kumcikissaat, isletmegiriskm, isletmegirissaat,
                farkkm, uretici, ureticikm, tonaj
            ])
            upload_excel(service, EXCEL_FILE_DRIVE_ID, wb_drive)

            flash("Kayıt başarıyla eklendi, OneDrive ve Google Drive’a kaydedildi!", "success")
            return redirect(url_for("form"))

        except Exception as e:
            flash(f"Hata oluştu: {e}", "danger")
            return redirect(url_for("form"))

    return render_template("form.html")

# -------------------------
# Uygulama başlatma
# -------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
