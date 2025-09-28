from flask import Flask, render_template, request, redirect, url_for, flash
from datetime import datetime
import os
import json
from google.oauth2 import service_account
import gspread

app = Flask(__name__)
app.secret_key = "supersecretkey"

# Google Sheet ayarları
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(BASE_DIR, "credentials.json")  # Servis hesabı dosyası
SHEET_ID = "1Z6KU-IztPS9TxwjywDv3uidqyCRMPzqz"        # Senin Google Sheet ID

# Servis hesabıyla bağlantı
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

def get_sheet():
    creds = service_account.Credentials.from_service_account_file(JSON_PATH, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SHEET_ID).sheet1
    return sheet

@app.route("/", methods=["GET", "POST"])
def form():
    if request.method == "POST":
        try:
            # Form verilerini al
            tarih = request.form.get("tarih") or datetime.now().strftime("%Y-%m-%d")
            iscikissaat = request.form.get("iscikissaat") or "00:00"
            plaka = request.form.get("plaka") or "Bilinmiyor"
            cikiskm = request.form.get("cikiskm") or "0"
            kumgirissaat = request.form.get("kumgirissaat") or "00:00"
            giriskm = request.form.get("giriskm") or "0"
            kumcikissaat = request.form.get("kumcikissaat") or "00:00"
            isletmegiriskm = request.form.get("isletmegiriskm") or "0"
            isletmegirissaat = request.form.get("isletmegirissaat") or "00:00"
            farkkm = float(giriskm) - float(cikiskm)
            uretici = request.form.get("uretici") or "Bilinmiyor"
            ureticikm = request.form.get("ureticikm") or "0"
            tonaj = request.form.get("tonaj") or "0"

            # Google Sheet'e ekle
            sheet = get_sheet()
            
            # Eğer Sheet boşsa başlık ekle
            if len(sheet.get_all_values()) == 0:
                sheet.append_row([
                    "Tarih","İşlem Çıkış Saati","Plaka","Çıkış KM","Kümes Giriş Saati",
                    "Giriş KM","Kümes Çıkış Saati","İşletme Giriş KM","İşletme Giriş Saati",
                    "Fark KM","Üretici","Üretici KM","Tonaj"
                ])

            sheet.append_row([
                tarih, iscikissaat, plaka, cikiskm, kumgirissaat,
                giriskm, kumcikissaat, isletmegiriskm, isletmegirissaat,
                farkkm, uretici, ureticikm, tonaj
            ])

            flash("Kayıt başarıyla eklendi!", "success")
            return redirect(url_for("form"))

        except Exception as e:
            flash(f"Hata oluştu: {e}", "danger")
            return redirect(url_for("form"))

    return render_template("form.html")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
