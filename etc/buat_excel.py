"""
Membuat dokumen Excel studi kasus TGS 3 (Klasifikasi Email Spam).
Output: TGS_3_Kelompok_1_Klasifikasi_Email_Spam_NRP_Nama_Excel.xlsx

Sheet:
  1. Variabel          - definisi tiap kolom/fitur dan arti angkanya
  2. Dataset           - 240 baris data dummy (dari spam_dummy.csv)
  3. Studi Kasus NB    - perhitungan manual Naive Bayes dengan formula Excel hidup
  4. Hasil Evaluasi    - metrik 3 model + confusion matrix (dari hasil_eksperimen.json)
  5. Formula           - rumus matematika yang dipakai
"""

import json
import os

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.join(BASE, "spam_dummy.csv")
JSON = os.path.join(BASE, "hasil", "hasil_eksperimen.json")
OUT = os.path.join(BASE, "TGS_3_Kelompok_1_Klasifikasi_Email_Spam_NRP_Nama_Excel.xlsx")

# Contoh email uji (sama dengan contoh pada makalah & program Python)
CONTOH = {"kata_free": 6, "kata_menang": 4, "jumlah_link": 4, "huruf_kapital": 12, "jumlah_dolar": 3}
KETERANGAN_CONTOH = "Email promosi berhadiah: banyak kata FREE/WIN, link, huruf kapital, dan simbol $"

JUDUL = Font(bold=True, size=14, color="1F4E79")
TEBAL = Font(bold=True)
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF")
KUNING = PatternFill("solid", fgColor="FFF2CC")
HIJAU = PatternFill("solid", fgColor="E2EFDA")
BORDER = Border(*[Side(style="thin", color="999999")] * 4)


def header(ws, row, kolom, teks):
    c = ws.cell(row=row, column=kolom, value=teks)
    c.fill, c.font, c.border = HEADER_FILL, HEADER_FONT, BORDER
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def lebar(ws, lebar_list):
    for i, w in enumerate(lebar_list, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def sheet_variabel(wb, ds):
    ws = wb.create_sheet("Variabel")
    ws["A1"] = "DEFINISI VARIABEL (FITUR) DATASET"
    ws["A1"].font = JUDUL
    baris = [
        ["No", "Nama Kolom", "Arti", "Contoh nilai & maknanya"],
        [1, "kata_free", "Jumlah kata promosi 'free'/'gratis' dalam email", "6 = ada 6 kata promosi; 0 = tidak ada"],
        [2, "kata_menang", "Jumlah kata 'win'/'menang'/'hadiah'", "4 = ajakan menang hadiah 4 kali; 0 = tidak ada"],
        [3, "jumlah_link", "Jumlah tautan/URL dalam email", "4 = ada 4 link; 1 = normal (1 tautan kontak)"],
        [4, "huruf_kapital", "Panjang rangkaian HURUF KAPITAL terpanjang", "12 = kata 'GRATIS' (12 huruf kapital)"],
        [5, "jumlah_dolar", "Jumlah simbol '$' dalam email", "3 = ada 3 simbol uang; 0 = tidak ada"],
        [6, "label", "Kelas target (yang diprediksi)", "1 = SPAM, 0 = BUKAN SPAM (ham)"],
    ]
    for j, teks in enumerate(baris[0], start=1):
        header(ws, 3, j, teks)
    for i, b in enumerate(baris[1:], start=4):
        for j, v in enumerate(b, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.border = BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top")
    catatan = [
        "",
        "CATATAN PENTING:",
        "1. Dataset ini adalah data DUMMY/buatan (sintetis) dengan seed tetap, BUKAN data asli UCI Spambase.",
        "2. Struktur fiturnya terinspirasi UCI Spambase (frekuensi kata, huruf kapital, simbol $).",
        f"3. Jumlah data: {ds['jumlah']} email ({ds['bukan_spam']} bukan spam, {ds['spam']} spam).",
        "4. Semua fitur berupa bilangan bulat, sehingga cocok untuk model Gaussian Naive Bayes.",
    ]
    for i, t in enumerate(catatan, start=11):
        ws.cell(row=i, column=1, value=t).font = Font(bold=(i in (12,)))
    lebar(ws, [5, 16, 45, 48])


def sheet_dataset(wb, df):
    ws = wb.create_sheet("Dataset")
    ws["A1"] = "DATASET EMAIL SPAM (DATA DUMMY) - spam_dummy.csv"
    ws["A1"].font = JUDUL
    for j, kol in enumerate(df.columns, start=1):
        header(ws, 3, j, kol)
    for i, baris in enumerate(df.itertuples(index=False), start=4):
        for j, v in enumerate(baris, start=1):
            c = ws.cell(row=i, column=j, value=int(v))
            c.border = BORDER
            if j == 6:
                c.fill = KUNING if v == 1 else HIJAU
                c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A4"
    lebar(ws, [11, 13, 12, 14, 13, 8])
    ws.cell(row=len(df) + 5, column=1,
            value="label: 1 = SPAM (kuning), 0 = BUKAN SPAM (hijau). Boleh ditambah/diedit baris baru "
                  "lalu jalankan ulang program Python.").font = Font(italic=True)


def sheet_studi_kasus(wb, nb):
    ws = wb.create_sheet("Studi Kasus NB")
    fitur = nb["fitur"]
    prior, mean, std = nb["prior"], nb["mean"], nb["std"]

    ws["A1"] = "STUDI KASUS: PREDIKSI 1 EMAIL DENGAN NAIVE BAYES (dihitung langsung di Excel)"
    ws["A1"].font = JUDUL
    ws["A2"] = KETERANGAN_CONTOH
    ws["A2"].font = Font(italic=True)

    ws["A4"] = "1. PRIOR P(C) - peluang awal tiap kelas dari data latih"
    ws["A4"].font = TEBAL
    header(ws, 5, 1, "Kelas"); header(ws, 5, 2, "Prior")
    ws["A6"] = "0 = Bukan Spam"; ws["B6"] = round(prior[0], 4)
    ws["A7"] = "1 = Spam"; ws["B7"] = round(prior[1], 4)

    ws["A9"] = "2. PARAMETER NORMAL (mean & std) hasil belajar dari 180 data latih"
    ws["A9"].font = TEBAL
    for j, t in enumerate(["Fitur", "mean kelas 0", "std kelas 0", "mean kelas 1", "std kelas 1"], 1):
        header(ws, 10, j, t)
    for i, f in enumerate(fitur, start=11):          # baris 11-15
        ws.cell(row=i, column=1, value=f)
        ws.cell(row=i, column=2, value=mean[f][0])
        ws.cell(row=i, column=3, value=std[f][0])
        ws.cell(row=i, column=4, value=mean[f][1])
        ws.cell(row=i, column=5, value=std[f][1])

    ws["A17"] = "3. INPUT EMAIL YANG DIUJI (silakan diganti untuk menguji email lain)"
    ws["A17"].font = TEBAL
    header(ws, 18, 1, "Fitur"); header(ws, 18, 2, "Nilai")
    for i, f in enumerate(fitur, start=19):          # baris 19-23
        ws.cell(row=i, column=1, value=f)
        c = ws.cell(row=i, column=2, value=CONTOH[f])
        c.fill, c.font, c.border = KUNING, TEBAL, BORDER

    ws["A25"] = "4. LIKELIHOOD P(x|C) memakai fungsi kepadatan normal NORM.DIST(x, mean, std, FALSE)"
    ws["A25"].font = TEBAL
    header(ws, 26, 1, "Fitur"); header(ws, 26, 2, "P(x|Bukan Spam)"); header(ws, 26, 3, "P(x|Spam)")
    for i, f in enumerate(fitur, start=27):          # baris 27-31
        r_in, r_par = 19 + (i - 27), 11 + (i - 27)
        ws.cell(row=i, column=1, value=f)
        ws.cell(row=i, column=2, value=f"=NORM.DIST(B{r_in},B{r_par},C{r_par},FALSE)")
        ws.cell(row=i, column=3, value=f"=NORM.DIST(B{r_in},D{r_par},E{r_par},FALSE)")
    ws["A32"] = "Produk likelihood (asumsi: fitur saling independen)"
    ws["B32"] = "=PRODUCT(B27:B31)"
    ws["C32"] = "=PRODUCT(C27:C31)"

    ws["A34"] = "5. POSTERIOR TIDAK TERNORMALISASI = prior x produk likelihood"
    ws["A34"].font = TEBAL
    header(ws, 35, 1, "Kelas"); header(ws, 35, 2, "Prior x Likelihood")
    ws["A36"] = "0 = Bukan Spam"; ws["B36"] = "=B6*B32"
    ws["A37"] = "1 = Spam"; ws["B37"] = "=B7*C32"

    ws["A39"] = "6. POSTERIOR TERNORMALISASI P(C|X) - total kedua kelas = 100%"
    ws["A39"].font = TEBAL
    header(ws, 40, 1, "Kelas"); header(ws, 40, 2, "Probabilitas")
    ws["A41"] = "P(Bukan Spam | X)"; ws["B41"] = "=B36/($B$36+$B$37)"
    ws["A42"] = "P(Spam | X)"; ws["B42"] = "=B37/($B$36+$B$37)"
    ws["B41"].number_format = ws["B42"].number_format = "0.00%"

    ws["A44"] = "7. KEPUTUSAN (argmax: kelas dengan probabilitas terbesar)"
    ws["A44"].font = TEBAL
    ws["A45"] = "Hasil klasifikasi:"
    ws["B45"] = '=IF(B42>=B41,"1 - SPAM","0 - BUKAN SPAM")'
    ws["B45"].font = Font(bold=True, size=12, color="C00000")
    ws["A47"] = ("Cara membaca: makin besar probabilitas Spam, makin yakin model bahwa email itu spam. "
                 "Ganti nilai kuning di bagian 3 untuk menguji email lain; semua angka dihitung ulang otomatis.")
    ws["A47"].font = Font(italic=True)
    ws["A48"] = ("Hasil verifikasi untuk input di atas (dihitung program Python): posterior bukan spam "
                 "1,83E-40 vs spam 2,22E-07 -> P(Spam|X) hampir 100% -> keputusan SPAM.")
    ws["A48"].font = Font(italic=True)
    for row in ws.iter_rows(min_row=4, max_row=47, min_col=1, max_col=5):
        for cell in row:
            if cell.value is not None:
                cell.border = BORDER
    lebar(ws, [38, 20, 20, 18, 18])


def sheet_evaluasi(wb, payload):
    ws = wb.create_sheet("Hasil Evaluasi")
    ds, hasil = payload["dataset"], payload["hasil"]
    ws["A1"] = "HASIL EVALUASI 3 MODEL (eksperimen nyata: python klasifikasi_spam.py)"
    ws["A1"].font = JUDUL
    ws["A2"] = (f"Data: {ds['nama']} | {ds['jumlah']} email | pembagian "
                f"{int((1-ds['test_size'])*100)}% latih / {int(ds['test_size']*100)}% uji "
                f"({round(ds['jumlah']*ds['test_size'])} data uji) | random_state={ds['random_state']}")
    ws["A2"].font = Font(italic=True)

    for j, t in enumerate(["Model", "Accuracy", "Precision", "Recall", "F1-score"], 1):
        header(ws, 4, j, t)
    for i, (nama, m) in enumerate(hasil.items(), start=5):
        ws.cell(row=i, column=1, value=nama + (" (MODEL UTAMA)" if nama == "Naive Bayes" else ""))
        for j, k in enumerate(["accuracy", "precision", "recall", "f1"], start=2):
            c = ws.cell(row=i, column=j, value=m[k])
            c.number_format = "0.00%"
        if nama == payload["model_terbaik_f1"]:
            for j in range(1, 6):
                ws.cell(row=i, column=j).fill = HIJAU

    baris = 10
    ws.cell(row=baris, column=1, value="CONFUSION MATRIX tiap model (data uji)").font = TEBAL
    for nama, m in hasil.items():
        tn, fp = m["confusion_matrix"][0]
        fn, tp = m["confusion_matrix"][1]
        baris += 1
        ws.cell(row=baris, column=1, value=nama).font = TEBAL
        baris += 1
        for j, t in enumerate(["aktual \\ prediksi", "Bukan Spam", "Spam"], 1):
            header(ws, baris, j, t)
        baris += 1
        ws.cell(row=baris, column=1, value="Bukan Spam"); ws.cell(row=baris, column=2, value=tn)
        ws.cell(row=baris, column=3, value=fp)
        baris += 1
        ws.cell(row=baris, column=1, value="Spam"); ws.cell(row=baris, column=2, value=fn)
        ws.cell(row=baris, column=3, value=tp)
        baris += 2

    ws.cell(row=baris, column=1, value="Arti setiap sel (kelas positif = SPAM):").font = TEBAL
    for t in [
        "TP (True Positive)  : spam diprediksi spam - benar.",
        "TN (True Negative)  : email normal diprediksi normal - benar.",
        "FP (False Positive) : email NORMAL dikira spam - email penting masuk kotak spam (biaya: kehilangan email).",
        "FN (False Negative) : spam diprediksi normal - spam lolos ke kotak masuk (biaya: user menerima spam).",
        "Precision = TP/(TP+FP) | Recall = TP/(TP+FN) | F1 = 2PR/(P+R) | Accuracy = (TP+TN)/Semua",
    ]:
        baris += 1
        ws.cell(row=baris, column=1, value=t)
    baris += 2
    ws.cell(row=baris, column=1,
            value=("Kesimpulan: pada eksperimen ini Naive Bayes (model utama) meraih F1 tertinggi, "
                   "sehingga dipilih sebagai model utama; Logistic Regression menjadi pembanding, "
                   "Decision Tree dipakai untuk memvisualkan aturan keputusan.")
            ).font = Font(italic=True)
    lebar(ws, [32, 16, 16, 16, 16])


def sheet_formula(wb):
    ws = wb.create_sheet("Formula")
    ws["A1"] = "FORMULA MATEMATIKA YANG DIPAKAI"
    ws["A1"].font = JUDUL
    baris = [
        ("Naive Bayes (model utama)", "P(C|X) = P(C) x PROD_i P(x_i|C)  ;  kelas dipilih = argmax P(C|X)"),
        ("Kepadatan normal (Gaussian)", "P(x_i|C) = 1/(sqrt(2*pi*sigma_C^2)) x exp(-(x_i-mu_C)^2 / (2*sigma_C^2))"),
        ("Logistic Regression (pembanding)", "P(spam|X) = 1 / (1 + e^-(b0 + b1*x1 + ... + b5*x5))"),
        ("Decision Tree (Gini impurity)", "Gini = 1 - SUM_i p_i^2   ;   keputusan mengikuti cabang sesuai ambang fitur"),
        ("Accuracy", "(TP + TN) / (TP + TN + FP + FN)"),
        ("Precision", "TP / (TP + FP)   - seberapa tepat deteksi spam"),
        ("Recall", "TP / (TP + FN)   - seberapa banyak spam yang berhasil tertangkap"),
        ("F1-score", "2 x (Precision x Recall) / (Precision + Recall)"),
    ]
    for j, t in enumerate(["Model / Metrik", "Rumus"], 1):
        header(ws, 3, j, t)
    for i, (nama, rumus) in enumerate(baris, start=4):
        ws.cell(row=i, column=1, value=nama).font = TEBAL
        ws.cell(row=i, column=2, value=rumus)
        for j in (1, 2):
            ws.cell(row=i, column=j).border = BORDER
    ws["A13"] = ("Keterangan: C = kelas (0/1), X = seluruh fitur email, x_i = fitur ke-i, "
                 "mu_C/rata-rata & sigma_C/standar deviasi kelas C, b = bobot yang dipelajari logistic regression.")
    ws["A13"].font = Font(italic=True)
    lebar(ws, [34, 95])


def main():
    df = pd.read_csv(CSV)
    with open(JSON, encoding="utf-8") as f:
        payload = json.load(f)

    wb = Workbook()
    wb.remove(wb.active)
    sheet_variabel(wb, payload["dataset"])
    sheet_dataset(wb, df)
    sheet_studi_kasus(wb, payload["naive_bayes_stats"])
    sheet_evaluasi(wb, payload)
    sheet_formula(wb)
    wb.save(OUT)
    print(f"Excel dibuat: {OUT}")


if __name__ == "__main__":
    main()

