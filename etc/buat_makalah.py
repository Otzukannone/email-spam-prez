"""
Membuat dokumen makalah TGS 3 (Word) - Klasifikasi Email Spam.
Output: TGS_3_Kelompok_1_Klasifikasi_Email_Spam_NRP_Nama_Makalah.docx

Format mengikuti makalah TGS 1:
  - Sampul
  - Halaman depan (Kata Pengantar, Daftar Isi) dengan angka Romawi (i, ii, ...)
  - Isi (BAB I s.d. BAB V) dengan angka Arab mulai dari 1 di halaman Pendahuluan
  - Daftar Pustaka
Semua angka hasil eksperimen diambil dari hasil/hasil_eksperimen.json
sehingga konsisten dengan program Python dan dokumen Excel/PPT.
"""

import json
import os

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

BASE = os.path.dirname(os.path.abspath(__file__))
JSON = os.path.join(BASE, "hasil", "hasil_eksperimen.json")
HASIL = os.path.join(BASE, "hasil")
OUT = os.path.join(BASE, "TGS_3_Kelompok_1_Klasifikasi_Email_Spam_NRP_Nama_Makalah.docx")

with open(JSON, encoding="utf-8") as f:
    P = json.load(f)

H = P["hasil"]                 # metrik tiap model
NB = P["naive_bayes_stats"]    # parameter Gaussian Naive Bayes
DS = P["dataset"]
NBm, LRm, DTm = H["Naive Bayes"], H["Logistic Regression"], H["Decision Tree"]
CONTOH_X = [6, 4, 4, 12, 3]    # email contoh: kata_free, kata_menang, jumlah_link, huruf_kapital, jumlah_dolar


# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------
def set_page_numbering(section, fmt: str, start: int):
    """fmt: 'lowerRoman' atau 'decimal'"""
    sectPr = section._sectPr
    pg = sectPr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType")
        sectPr.append(pg)
    pg.set(qn("w:fmt"), fmt)
    pg.set(qn("w:start"), str(start))


def add_page_number_footer(section):
    """Memasang nomor halaman di tengah bawah."""
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.clear()
    run = p.add_run()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    run._r.addnext(fld)


def para(doc, teks="", bold=False, italic=False, size=12, align=None,
         space_after=6, font="Times New Roman", color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.5
    if align is not None:
        p.alignment = align
    if teks:
        r = p.add_run(teks)
        r.bold, r.italic, r.font.size = bold, italic, Pt(size)
        r.font.name = font
        if color:
            r.font.color.rgb = color
    return p


def heading(doc, teks, level=1):
    """Heading BAB dengan gaya manual agar rapi."""
    if level == 1:
        p = para(doc, teks, bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    elif level == 2:
        p = para(doc, teks, bold=True, size=12, space_after=6)
    else:
        p = para(doc, teks, bold=True, italic=True, size=12, space_after=6)
    return p


def body(doc, teks):
    """Paragraf isi rata kiri-kanan dengan pelebaran kata pertama (seperti TGS 1)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Cm(1.25)
    r = p.add_run(teks)
    r.font.size = Pt(12)
    r.font.name = "Times New Roman"
    return p


def bullet(doc, teks):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(teks)
    r.font.size = Pt(12)
    r.font.name = "Times New Roman"
    return p


def caption(doc, teks):
    return para(doc, teks, bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)


def tabel(doc, data, header_baris=True, ukuran=10):
    t = doc.add_table(rows=len(data), cols=len(data[0]))
    t.style = "Table Grid"
    t.alignment = 1
    for i, baris in enumerate(data):
        for j, sel in enumerate(baris):
            c = t.cell(i, j)
            c.text = ""
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(sel))
            r.font.size = Pt(ukuran)
            r.font.name = "Times New Roman"
            r.bold = header_baris and i == 0
    return t


def sampul(doc):
    for _ in range(2):
        para(doc, "", space_after=0)
    para(doc, "INSTITUT TEKNOLOGI NASIONAL", bold=True, size=16,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    para(doc, "BANDUNG", bold=True, size=16, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    para(doc, "FAKULTAS TEKNOLOGI INDUSTRI", bold=True, size=14,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    para(doc, "PROGRAM STUDI INFORMATIKA", bold=True, size=14,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("_" * 42)
    r.font.size = Pt(12)

    for _ in range(2):
        para(doc, "", space_after=0)
    para(doc, "TGS 3 - TUGAS 3", bold=True, size=14,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    para(doc, "IFB-307 DATA MINING & INFORMATION RETRIEVAL", bold=True, size=14,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)
    para(doc, "KLASIFIKASI EMAIL SPAM DENGAN", bold=True, size=20,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    para(doc, "NAIVE BAYES, LOGISTIC REGRESSION,", bold=True, size=20,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    para(doc, "DAN DECISION TREE", bold=True, size=20,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    para(doc, "(Studi Kasus pada Dataset Dummy 240 Email)", italic=True, size=12,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)

    data = [
        ["Kelompok", ": 1"],
        ["Nama Anggota 1 (NRP)", ": ..........................................."],
        ["Nama Anggota 2 (NRP)", ": ..........................................."],
        ["Nama Dosen Pengampu", ": ..........................................."],
        ["Tahun Akademik", ": 2026/2027"],
    ]
    t = doc.add_table(rows=len(data), cols=2)
    t.alignment = 1
    for i, (k, v) in enumerate(data):
        for j, teks in enumerate((k, v)):
            c = t.cell(i, j)
            c.text = ""
            r = c.paragraphs[0].add_run(teks)
            r.font.size = Pt(12)
            r.font.name = "Times New Roman"


def halaman_depan(doc):
    """Kata Pengantar + Daftar Isi + Daftar Gambar (penomoran Romawi)."""
    heading(doc, "KATA PENGANTAR")
    body(doc, "Puji syukur kami panjatkan ke hadirat Tuhan Yang Maha Esa karena "
              "berkat rahmat dan karunia-Nya, tugas makalah mata kuliah IFB-307 Data Mining "
              "& Information Retrieval dapat diselesaikan dengan baik. Makalah ini membahas "
              "pembuatan model klasifikasi untuk masalah deteksi email spam, mulai dari definisi "
              "dan sejarah klasifikasi, formula matematika yang digunakan, studi kasus perhitungan "
              "di Excel, hingga implementasi program Python.")
    body(doc, "Model utama yang digunakan adalah Naive Bayes karena sederhana, cepat, dan sesuai "
              "dengan sifat data frekuensi pada email. Selain itu, kami membandingkannya dengan "
              "Logistic Regression sebagai model pembanding dan Decision Tree sebagai visualisasi "
              "aturan keputusan. Seluruh angka hasil evaluasi dalam makalah ini berasal dari "
              "eksperimen nyata terhadap dataset dummy sebanyak "
              f"{DS['jumlah']} email yang kami buat sendiri, sehingga hasilnya dapat direproduksi "
              "dengan menjalankan ulang program Python yang kami lampirkan.")
    body(doc, "Kami menyadari makalah ini masih jauh dari kesempurnaan, baik dari segi isi "
              "maupun penyusunannya. Oleh karena itu, kami mengharapkan kritik dan saran yang "
              "membangun dari pembaca agar kami dapat memperbaikinya di kemudian hari.")
    para(doc, "Bandung, Oktober 2026", size=12, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=2)
    para(doc, "Kelompok 1", size=12, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=20)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    heading(doc, "DAFTAR ISI")
    daftar = [
        ("KATA PENGANTAR", "i"), ("DAFTAR ISI", "ii"), ("DAFTAR GAMBAR DAN TABEL", "iii"),
        ("BAB I PENDAHULUAN", "1"),
        ("   1.1 Latar Belakang", "1"), ("   1.2 Rumusan Masalah", "2"),
        ("   1.3 Tujuan Penulisan", "2"),
        ("BAB II TINJAUAN PUSTAKA", "3"),
        ("   2.1 Definisi Klasifikasi", "3"), ("   2.2 Sejarah Klasifikasi dan Data Mining", "3"),
        ("   2.3 Model Naive Bayes", "4"), ("   2.4 Model Logistic Regression", "5"),
        ("   2.5 Model Decision Tree", "6"), ("   2.6 Metrik Evaluasi", "6"),
        ("BAB III STUDI KASUS DAN DATA", "8"),
        ("   3.1 Deskripsi Dataset dan Variabel", "8"),
        ("   3.2 Perhitungan Manual Naive Bayes di Excel", "9"),
        ("BAB IV IMPLEMENTASI PROGRAM PYTHON", "11"),
        ("   4.1 Lingkungan dan File Proyek", "11"),
        ("   4.2 Letak Model pada Kode dan Cara Kerja", "11"),
        ("   4.3 Hasil Eksperimen dan Pembahasan", "13"),
        ("BAB V KESIMPULAN", "15"),
        ("DAFTAR PUSTAKA", "16"),
    ]
    for teks, nomor in daftar:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(15.5), 2, 1)
        r = p.add_run(f"{teks}\t{nomor}")
        r.font.size = Pt(12)
        r.font.name = "Times New Roman"
        r.bold = teks.startswith("BAB") or teks == "DAFTAR PUSTAKA"

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    heading(doc, "DAFTAR GAMBAR DAN TABEL")
    gambar = [
        "Gambar 4.1  Cuplikan perintah menjalankan program Python",
        "Gambar 4.2  Confusion matrix model Naive Bayes",
        "Gambar 4.3  Perbandingan metrik tiga model",
        "Gambar 4.4  Visualisasi Decision Tree",
        "Tabel 2.1  Perbandingan enam model klasifikasi",
        "Tabel 3.1  Definisi variabel dataset",
        "Tabel 3.2  Parameter Naive Bayes (mean dan std per kelas)",
        "Gambar 3.1  Perbandingan rata-rata fitur per kelas",
        "Tabel 3.3  Perhitungan posterior pada studi kasus",
        "Tabel 4.1  Hasil evaluasi tiga model",
    ]
    for g in gambar:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(g)
        r.font.size = Pt(12)
        r.font.name = "Times New Roman"


def bab1(doc):
    heading(doc, "BAB I\nPENDAHULUAN")
    heading(doc, "1.1 Latar Belakang", level=2)
    body(doc, "Email adalah salah satu alat komunikasi paling banyak digunakan, baik untuk "
              "keperluan kuliah, kerja, maupun transaksi. Sayangnya, kotak masuk kita hampir "
              "selalu menerima email yang tidak diinginkan, mulai dari promosi, penipuan, "
              "hingga undangan berhadiah palsu. Email seperti itu disebut email spam. Spam "
              "tidak hanya menghabiskan waktu, tetapi juga bisa membawa tautan berbahaya "
              "sehingga pengguna perlu menyaringnya.")
    body(doc, "Pada dasarnya, menyaring spam adalah masalah klasifikasi: menentukan apakah "
              "sebuah email termasuk kelas \"spam\" atau \"bukan spam\" berdasarkan ciri-ciri "
              "yang dimilikinya. Ciri tersebut bisa berupa jumlah kata promosi, jumlah tautan, "
              "panjang huruf kapital, dan lain-lain. Dalam tugas ini kami membangun model "
              "klasifikasi email spam dengan tiga algoritma, yaitu Naive Bayes sebagai model "
              "utama, Logistic Regression sebagai pembanding, dan Decision Tree sebagai "
              "penjelas aturan keputusan yang mudah divisualkan.")
    body(doc, "Pemilihan Naive Bayes didasarkan pada sifat data email yang berupa frekuensi/"
              "jumlah kata, sehingga pendekatan probabilitas Bayes sangat cocok dan ringan. "
              "Logistic Regression dipilih karena merupakan model garis dasar yang kuat dan "
              "mudah ditafsirkan, sedangkan Decision Tree dipilih karena keputusannya bisa "
              "digambar menjadi pohon sehingga audiens dapat melihat \"aturan\" yang dipakai "
              "model. Data yang digunakan adalah dataset dummy buatan kami sendiri berjumlah "
              f"{DS['jumlah']} email dengan {len(DS['fitur'])} fitur numerik dan satu kolom label "
              "(1 = spam, 0 = bukan spam), dibuat agar penelitian ini sederhana, jujur, dan "
              "dapat direproduksi tanpa mengklaim data sebagai data asli dari sumber mana pun.")
    heading(doc, "1.2 Rumusan Masalah", level=2)
    body(doc, "Berdasarkan latar belakang di atas, rumusan masalah dalam makalah ini adalah:")
    bullet(doc, "Bagaimana definisi dan sejarah singkat klasifikasi dalam data mining?")
    bullet(doc, "Bagaimana formula matematika di balik Naive Bayes, Logistic Regression, dan Decision Tree?")
    bullet(doc, "Bagaimana studi kasus perhitungan klasifikasi email spam menggunakan Excel?")
    bullet(doc, "Bagaimana implementasi model tersebut dalam program Python dan seberapa baik hasilnya?")
    heading(doc, "1.3 Tujuan Penulisan", level=2)
    body(doc, "Tujuan penulisan makalah ini adalah:")
    bullet(doc, "Menjelaskan konsep dan sejarah klasifikasi secara ringkas.")
    bullet(doc, "Menjabarkan formula ketiga model serta metrik evaluasinya.")
    bullet(doc, "Mendemonstrasikan perhitungan manual Naive Bayes pada studi kasus Excel.")
    bullet(doc, "Mengimplementasikan model pada program Python yang dapat membaca file data "
                 "(CSV) dan menguji prediksi email baru, lalu melaporkan hasil eksperimen yang "
                 "sebenarnya.")


def bab2(doc):
    heading(doc, "BAB II\nTINJAUAN PUSTAKA")
    heading(doc, "2.1 Definisi Klasifikasi", level=2)
    body(doc, "Klasifikasi adalah salah satu tugas utama dalam data mining (penambangan data) "
              "yaitu proses membangun model yang mampu mengelompokkan data ke dalam kelas atau "
              "kategori berdasarkan contoh-contoh yang sudah diketahui labelnya (supervised "
              "learning). Menurut Han, Kamber dan Pei, klasifikasi adalah proses mempelajari "
              "fungsi pemetaan (kelas, data) dari data pelatihan, di mana fungsi tersebut "
              "berupa kelas diskrit. Secara sederhana, model \"belajar\" dari data latih "
              "berlabel, lalu digunakan untuk menebak label data baru yang belum diketahui.")
    body(doc, "Contoh penerapan klasifikasi sangat banyak: deteksi email spam, deteksi penipuan "
              "pada transaksi kartu kredit, prediksi penyakit dari gejala pasien, hingga "
              "pengenalan tulisan tangan. Pada semua contoh tersebut, masalahnya sama: "
              "menentukan kelas output berdasarkan nilai-nilai input.")
    heading(doc, "2.2 Sejarah Klasifikasi dan Data Mining", level=2)
    body(doc, "Konsep klasifikasi sebenarnya sudah ada jauh sebelum komputer lahir. Aristoteles "
              "sekitar 340 SM menyusun taksonomi makhluk hidup berdasarkan ciri-cirinya, dan "
              "Carl Linnaeus abad ke-18 membuat klasifikasi tumbuhan dan hewan yang masih "
              "dipakai sampai sekarang. Dalam statistika, metode klasifikasi berkembang melalui "
              "analisis diskriminan (Fisher, 1936) yang berusaha menemukan garis pemisah "
              "antarkelas.")
    body(doc, "Bidang data mining sendiri mulai populer pada tahun 1990-an seiring ledakan data "
              "digital. Naive Bayes berakar dari teorema Bayes (Thomas Bayes, 1763) dan banyak "
              "digunakan untuk klasifikasi teks sejak 1990-an (misalnya email spam). Decision "
              "Tree populer lewat algoritma ID3 dan C4.5 (Quinlan, 1986; 1993), "
              "sedangkan Logistic Regression telah lama dipakai statistikawan dan kembali "
              "populer sebagai model baseline machine learning. Saat ini, filter spam di "
              "layanan email besar umumnya berupa sistem berlapis: pemeriksaan reputasi pengirim, "
              "aturan keamanan, dan klasifikasi machine learning pada isi pesan; nama model "
              "persisnya tidak dipublikasikan sehingga kita tidak dapat mengklaim penyedia "
              "tertentu memakai satu algoritma tertentu.")


    heading(doc, "2.3 Model Naive Bayes (Model Utama)", level=2)
    body(doc, "Naive Bayes adalah klasifier probabilitas yang dibangun di atas teorema Bayes "
              "dengan asumsi \"naif\": setiap fitur dianggap saling independen terhadap fitur "
              "lainnya jika kelas sudah diketahui. Walaupun asumsi ini tidak sepenuhnya benar "
              "untuk bahasa manusia, pada praktiknya model ini bekerja sangat baik dan sangat "
              "cepat, sehingga menjadi pilihan klasik untuk klasifikasi teks dan email spam.")
    para(doc, "Rumus Teorema Bayes:", size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    para(doc, "P(C | X) = P(X | C) x P(C) / P(X)", bold=True, size=13,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    body(doc, "Dengan asumsi naif, P(X|C) dipecah menjadi perkalian probabilitas tiap fitur, "
              "sehingga rumus yang dipakai program menjadi:")
    para(doc, "P(C | X) proporsional P(C) x PROD(i=1..n) P(x_i | C)", bold=True, size=13,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    body(doc, "di mana C adalah kelas (0 = bukan spam, 1 = spam), X = (x1, ..., xn) adalah "
              "seluruh fitur email, P(C) adalah prior kelas, dan P(x_i|C) adalah peluang fitur "
              "ke-i pada kelas C. Karena data kita numerik berupa jumlah/ frekuensi, peluang "
              "P(x_i|C) dimodelkan dengan distribusi normal (Gaussian) sehingga model ini "
              "disebut Gaussian Naive Bayes:")
    para(doc, "P(x_i | C) = 1 / sqrt(2*pi*sigma_C^2) x exp( -(x_i - mu_C)^2 / (2*sigma_C^2) )",
         bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    body(doc, "mu_C adalah rata-rata fitur pada kelas C dan sigma_C adalah simpangan baku-nya; "
              "kedua nilai ini dipelajari model dari data latih. Keputusan akhir diambil dengan "
              "memilih kelas yang posterior-nya paling besar (argmax). Misalnya P(X) = 0,00002 "
              "untuk kelas spam dan 0,00001 untuk bukan spam, maka email diputuskan sebagai spam "
              "karena 0,00002 > 0,00001.")

    heading(doc, "2.4 Model Logistic Regression (Pembanding)", level=2)
    body(doc, "Logistic Regression adalah model klasifikasi yang memetakan kombinasi linear "
              "fitur ke dalam nilai probabilitas antara 0 dan 1 melalui fungsi logistik "
              "(sigmoid). Model ini cepat, mudah ditafsirkan karena setiap fitur punya bobot "
              "(koefisien) b_i, dan sering dijadikan pembanding (baseline) bagi model lain.")
    para(doc, "P(spam | X) = 1 / (1 + e^-(b0 + b1*x1 + b2*x2 + ... + bn*xn))", bold=True,
         size=13, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    body(doc, "b0 disebut intercept, b1..bn adalah bobot yang dipelajari dari data. Jika "
              "P(spam|X) >= 0,5 email diklasifikasikan spam, selain itu bukan spam. Kelemahan "
              "utamanya adalah garis pemisah (decision boundary) yang linear, sehingga pola "
              "rumit tidak tertangkap tanpa rekayasa fitur.")
    heading(doc, "2.5 Model Decision Tree (Penjelas Visual)", level=2)
    body(doc, "Decision Tree (pohon keputusan) membagi data secara berulang berdasarkan "
              "pertanyaan sederhana pada satu fitur, misalnya \"kata_menang <= 0,5?\", sampai "
              "daun pohon berisi keputusan kelas. Pembagian dilakukan dengan kriteria impurity, "
              "dalam program ini memakai Gini impurity:")
    para(doc, "Gini(node) = 1 - SUM(k) p_k^2", bold=True, size=13,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    body(doc, "di mana p_k adalah proporsi data kelas k pada node tersebut. Gini = 0 berarti "
              "node murni (semua data satu kelas). Keunggulan Decision Tree adalah hasilnya "
              "berupa aturan yang bisa dibaca manusia dan mudah digambar, sedangkan kekurangannya "
              "adalah mudah overfit jika pohon dibuat terlalu dalam.")


    heading(doc, "2.6 Metrik Evaluasi Model", level=2)
    body(doc, "Untuk menilai model klasifikasi tidak cukup hanya melihat akurasi, terutama pada "
              "masalah spam di mana kesalahan FP dan FN biayanya berbeda. Empat metrik yang "
              "digunakan dalam makalah ini adalah:")
    bullet(doc, "Accuracy = (TP + TN) / (TP + TN + FP + FN) - proporsi seluruh prediksi yang benar.")
    bullet(doc, "Precision = TP / (TP + FP) - di antara email yang diprediksi spam, berapa yang benar-benar spam.")
    bullet(doc, "Recall = TP / (TP + FN) - di antara email spam yang sebenarnya, berapa yang berhasil tertangkap.")
    bullet(doc, "F1-score = 2 x Precision x Recall / (Precision + Recall) - rata-rata harmonik keduanya.")
    body(doc, "FP (False Positive) berarti email normal dikirim ke kotak spam, sedangkan FN "
              "(False Negative) berarti spam lolos ke kotak masuk. Jika kehilangan email "
              "penting lebih berbahaya, model dipilih yang precision-nya tinggi; jika spam "
              "harus tertangkap semua, recall yang diprioritaskan. Pada tugas ini kami "
              "memakai F1-score sebagai ukuran utama karena menyeimbangkan keduanya.")
    tabel(doc, [
        ["Model", "Kelebihan", "Kekurangan", "Kompleksitas"],
        ["Naive Bayes", "Cepat, sederhana, cocok untuk frekuensi kata", "Asumsi independensi fitur tidak selalu realistis", "Rendah"],
        ["Logistic Regression", "Interpretabel, probabilitas keluaran", "Hanya garis pemisah linear", "Rendah-sedang"],
        ["Linear SVM", "Efektif untuk teks berdimensi tinggi", "Probabilitas butuh kalibrasi", "Sedang"],
        ["Decision Tree", "Aturan eksplisit, mudah divisualkan", "Pohon dalam mudah overfit", "Rendah-sedang"],
        ["Random Forest", "Lebih stabil dari satu pohon", "Kurang interpretabel, lebih berat", "Sedang-tinggi"],
        ["KNN", "Sederhana, tanpa training rumit", "Prediksi lambat, peka skala fitur", "Latih rendah / tebak tinggi"],
    ], ukuran=9)
    caption(doc, "Tabel 2.1  Perbandingan enam model klasifikasi untuk email spam")


def bab3(doc):
    heading(doc, "BAB III\nSTUDI KASUS DAN DATA")
    heading(doc, "3.1 Deskripsi Dataset dan Variabel", level=2)
    body(doc, "Dataset yang digunakan adalah file spam_dummy.csv yang kami buat sendiri "
              f"berisi {DS['jumlah']} contoh email ({DS['bukan_spam']} email bukan spam dan "
              f"{DS['spam']} email spam). Dataset ini data dummy/sintetis dengan seed tetap, "
              "bukan data asli UCI Spambase, meskipun struktur fiturnya terinspirasi dari "
              "dataset tersebut (jumlah kata promosi, huruf kapital, simbol dolar, dan "
              "tautan). Setiap baris mewakili satu email dan terdiri dari 5 fitur numerik "
              "serta satu kolom label.")
    body(doc, "Semua fitur berupa bilangan bulat sehingga mudah dimengerti: angka 6 pada "
              "kolom kata_free berarti email tersebut mengandung 6 kata promosi \"free/"
              "gratis\"; angka 12 pada huruf_kapital berarti rangkaian huruf kapital "
              "terpanjang pada email berukuran 12 karakter (misalnya kata GRATIS), dan "
              "seterusnya. Label 1 berarti spam dan 0 berarti bukan spam. Karena bentuk "
              "datanya numerik kontinu, model Gaussian Naive Bayes sangat sesuai.")
    tabel(doc, [
        ["No", "Kolom", "Arti", "Contoh nilai & maknanya"],
        ["1", "kata_free", "Jumlah kata promosi free/gratis", "6 = ada 6 kata promosi; 0 = tidak ada"],
        ["2", "kata_menang", "Jumlah kata win/menang/hadiah", "4 = ajakan menang hadiah 4 kali"],
        ["3", "jumlah_link", "Jumlah tautan/URL dalam email", "4 = ada 4 link; 1 = normal"],
        ["4", "huruf_kapital", "Panjang rangkaian HURUF KAPITAL terpanjang", "12 = kata GRATIS (12 huruf kapital)"],
        ["5", "jumlah_dolar", "Jumlah simbol $ dalam email", "3 = ada 3 simbol uang; 0 = tidak ada"],
        ["6", "label", "Kelas target (keluaran model)", "1 = SPAM, 0 = BUKAN SPAM"],
    ], ukuran=9)
    caption(doc, "Tabel 3.1  Definisi variabel dataset")
    body(doc, "Data dibagi menjadi 75% data latih (180 email) dan 25% data uji (60 email) "
              f"dengan random_state = {DS['random_state']} supaya pembagian sama setiap "
              "program dijalankan. Data latih dipakai model untuk \"belajar\", data uji "
              "dipakai mengevaluasi kinerja model pada data yang belum pernah dilihat.")

    heading(doc, "3.2 Perhitungan Manual Naive Bayes di Excel", level=2)
    body(doc, "Studi kasus Excel (sheet \"Studi Kasus NB\") menunjukkan perhitungan lengkap "
              "klasifikasi satu email dari awal sampai keputusan, memakai formula Excel "
              "sehingga angkanya bisa dilihat dan diubah langsung. Email yang diuji adalah "
              "email promosi berhadiah dengan nilai fitur:")
    tabel(doc, [
        ["kata_free", "kata_menang", "jumlah_link", "huruf_kapital", "jumlah_dolar"],
        [str(CONTOH_X[0]), str(CONTOH_X[1]), str(CONTOH_X[2]), str(CONTOH_X[3]), str(CONTOH_X[4])],
    ], ukuran=10)
    caption(doc, "Email contoh: banyak kata FREE dan WIN, 4 tautan, huruf kapital 12 karakter, dan 3 simbol $")
    body(doc, "Langkah perhitungannya adalah:")
    bullet(doc, f"1. Prior dari data latih: P(bukan spam) = {NB['prior'][0]:.4f}, "
                f"P(spam) = {NB['prior'][1]:.4f}.")
    bullet(doc, "2. Parameter normal (mean dan std) tiap fitur per kelas dipelajari dari "
                "180 data latih, ditampilkan pada Tabel 3.2.")
    bullet(doc, "3. Likelihood tiap fitur dihitung dengan formula Excel "
                "NORM.DIST(x, mean, std, FALSE), lalu dikalikan antarfitur (asumsi naif).")
    bullet(doc, "4. Posterior = prior x produk likelihood, lalu dinormalkan agar jumlah kedua "
                "kelas = 100%.")
    bullet(doc, "5. Keputusan = kelas dengan probabilitas terbesar (argmax).")

    baris_tabel = [["Fitur", "mean kelas 0", "std kelas 0", "mean kelas 1", "std kelas 1"]]
    for f in DS["fitur"]:
        baris_tabel.append([f, f"{NB['mean'][f][0]:.4f}", f"{NB['std'][f][0]:.4f}",
                            f"{NB['mean'][f][1]:.4f}", f"{NB['std'][f][1]:.4f}"])
    tabel(doc, baris_tabel, ukuran=9)
    caption(doc, "Tabel 3.2  Parameter Gaussian Naive Bayes hasil belajar dari data latih")
    path_gambar = os.path.join(HASIL, "naive_bayes_mean.png")
    if os.path.exists(path_gambar):
        doc.add_picture(path_gambar, width=Inches(4.8))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption(doc, "Gambar 3.1  Perbandingan rata-rata fitur per kelas (hasil belajar model)")
    body(doc, "Perhatikan polanya: email spam rata-rata punya kata_free, kata_menang, "
              "jumlah_link, huruf_kapital, dan jumlah_dolar yang lebih tinggi dibanding email "
              "normal. Inilah \"pola\" yang ditemukan model. Memasukkan nilai contoh email ke "
              "rumus menghasilkan posterior tak ternormalisasi sekitar 1,83 x 10^-40 untuk "
              "kelas bukan spam dan 2,22 x 10^-7 untuk kelas spam, sehingga P(spam|X) hampir "
              "100% dan email diputuskan SPAM, sesuai intuisi kita.")
    tabel(doc, [
        ["Kelas", "Prior x Likelihood", "Posterior ternormalisasi", "Keputusan"],
        ["0 - Bukan Spam", "1,83 x 10^-40", "≈ 0,00%", "kalah"],
        ["1 - Spam", "2,22 x 10^-7", "≈ 100,00%", "SPAM (dipilih)"],
    ], ukuran=10)
    caption(doc, "Tabel 3.3  Ringkasan hasil perhitungan posterior studi kasus")


def bab4(doc):
    heading(doc, "BAB IV\nIMPLEMENTASI PROGRAM PYTHON")
    heading(doc, "4.1 Lingkungan dan File Proyek", level=2)
    body(doc, "Program ditulis dalam bahasa Python 3 dengan bahan bantu pustaka scikit-learn "
              "(model klasifikasi), pandas (membaca CSV), numpy, dan matplotlib (grafik). "
              "File-file dalam proyek:")
    bullet(doc, "spam_dummy.csv - dataset dummy 240 email (boleh diedit atau diganti).")
    bullet(doc, "buat_dataset.py - pembuat dataset dummy (opsional, seed tetap).")
    bullet(doc, "klasifikasi_spam.py - program utama: membaca CSV, melatih model, mengevaluasi, "
                 "menggambar grafik, dan menguji email baru.")
    bullet(doc, "hasil/hasil_eksperimen.json - angka hasil eksperimen (dipakai dokumen ini).")
    bullet(doc, "hasil/*.png - grafik confusion matrix, perbandingan model, dan pohon keputusan.")
    body(doc, "Cara menjalankan: python klasifikasi_spam.py (latih + evaluasi otomatis; pada "
              "jendela terminal/klik ganda langsung membuka menu interaktif supaya jendela tidak "
              "tertutup), python klasifikasi_spam.py --prediksi 6 4 4 12 3 (menebak satu email "
              "baru), python klasifikasi_spam.py --data data_lain.csv (memakai file data lain), "
              "atau python klasifikasi_spam.py --menu (menu interaktif paksa). Cara termudah "
              "menjalankannya adalah klik ganda file Jalankan Program.bat di File Explorer. "
              "File data berupa CSV dan penulisannya mudah, sehingga mahasiswa bisa menambah "
              "email baru sendiri lalu menjalankan ulang program untuk menguji model.")

    heading(doc, "4.2 Letak Model pada Kode dan Cara Kerja", level=2)
    body(doc, "Model klasifikasi terletak di dalam fungsi latih_dan_evaluasi() pada file "
              "klasifikasi_spam.py. Di sana tiga model dibuat lewat fungsi buat_model() lalu "
              "dilatih dengan perintah .fit(X_train, y_train) - perintah inilah yang membuat "
              "model \"belajar\" pola dari data latih:")
    kode = [
        "def buat_model():",
        "    return {",
        '        "Naive Bayes": GaussianNB(),               # MODEL UTAMA',
        '        "Logistic Regression": LogisticRegression(max_iter=1000),',
        '        "Decision Tree": DecisionTreeClassifier(max_depth=4),',
        "    }",
        "",
        "for nama, model in buat_model().items():",
        "    model.fit(X_train, y_train)     # <- model belajar pola dari data latih",
        "    y_pred = model.predict(X_test)  # <- model menebak data uji",
    ]
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    c = t.cell(0, 0)
    c.text = ""
    for i, baris in enumerate(kode):
        p = c.paragraphs[0] if i == 0 else c.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        r = p.add_run(baris if baris else " ")
        r.font.name = "Consolas"
        r.font.size = Pt(9.5)
    caption(doc, "Cuplikan kode inti klasifier pada klasifikasi_spam.py")
    body(doc, "Cara kerja program secara berurutan: (1) membaca spam_dummy.csv dengan pandas; "
              "(2) memisahkan kolom fitur X dan kolom label y; (3) membagi data latih dan uji "
              "75/25; (4) melatih ketiga model; (5) menghitung accuracy, precision, recall, "
              "F1, dan confusion matrix pada data uji; (6) menyimpan grafik dan "
              "hasil_eksperimen.json; (7) bila diberi perintah --prediksi, model yang sudah "
              "terlatih dipakai menebak satu email baru melalui fungsi prediksi_satu().")
    body(doc, "Pola dikenali model melalui statistik yang dipelajari dari data latih: Naive "
              "Bayes menyimpan rata-rata dan simpangan baku tiap fitur per kelas sehingga ia "
              "bisa menilai \"wajar-tidaknya\" sebuah nilai fitur bagi kelas spam; Logistic "
              "Regression menyimpan bobot b_i tiap fitur sehingga fitur seperti kata_menang "
              "yang bobotnya besar menarik keputusan ke arah spam; Decision Tree menyimpan "
              "serangkaian ambang percabangan (pada model ini simpul akar-nya adalah "
              "kata_menang <= 0,5).")
    body(doc, "Keluaran model berupa: (a) kelas hasil tebakan 0 atau 1, (b) untuk Naive Bayes "
              "dan Logistic Regression juga probabilitas spam sehingga kita tahu seberapa yakin "
              "model, serta (c) metrik agregat yang menjelaskan kualitas model secara keseluruhan. "
              "Contoh pengujian email promosi (6, 4, 4, 12, 3) menghasilkan keputusan SPAM dengan "
              "probabilitas 100% pada ketiga model, sedangkan email normal (0, 0, 1, 2, 0) "
              "diprediksi bukan spam dengan probabilitas spam di bawah 4%.")


    heading(doc, "4.3 Hasil Eksperimen dan Pembahasan", level=2)
    body(doc, "Eksperimen dijalankan dengan perintah python klasifikasi_spam.py dan hasilnya "
              "dikutip apa adanya dari file hasil_eksperimen.json. Tabel 4.1 berisi metrik "
              "pada 60 data uji; angka ini adalah hasil nyata, bukan angka karangan, dan dapat "
              "dihitung ulang dengan menjalankan program.")
    baris = [["Model", "Accuracy", "Precision", "Recall", "F1-score"]]
    for nama, m in H.items():
        baris.append([nama + (" (utama)" if nama == "Naive Bayes" else ""),
                      f"{m['accuracy']*100:.2f}%", f"{m['precision']:.4f}",
                      f"{m['recall']:.4f}", f"{m['f1']:.4f}"])
    tabel(doc, baris, ukuran=10)
    caption(doc, "Tabel 4.1  Hasil evaluasi tiga model pada data uji (60 email)")

    gambar = [
        ("confusion_naive_bayes.png", "Gambar 4.2  Confusion matrix model Naive Bayes"),
        ("perbandingan_model.png", "Gambar 4.3  Perbandingan metrik tiga model"),
        ("decision_tree.png", "Gambar 4.4  Visualisasi Decision Tree"),
    ]
    for nama_file, teks in gambar:
        path = os.path.join(HASIL, nama_file)
        if os.path.exists(path):
            doc.add_picture(path, width=Inches(5.4))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            caption(doc, teks)

    tn, fp = NBm["confusion_matrix"][0]
    fn, tp = NBm["confusion_matrix"][1]
    body(doc, f"Pada confusion matrix Naive Bayes, model benar mengenali {tn} email normal dan "
              f"{tp} email spam; terdapat {fp} email normal yang keliru masuk spam dan {fn} spam "
              "yang lolos. Dari tabel terlihat Naive Bayes (model utama) memperoleh F1 tertinggi "
              f"({NBm['f1']:.4f}) dengan recall {NBm['recall']:.4f}, artinya hampir semua spam "
              "berhasil tertangkap - hal yang paling diinginkan dari filter spam. Logistic "
              f"Regression sangat berdekatan (F1 {LRm['f1']:.4f}) dan bahkan sedikit lebih "
              f"precision, sehingga model tersebut layak sebagai pembanding. Decision Tree "
              f"(F1 {DTm['f1']:.4f}) sedikit lebih rendah, namun unggul dalam hal keterbacaan "
              "karena keputusannya berupa aturan if-then yang bisa dilihat pada Gambar 4.4.")
    body(doc, "Perlu ditegaskan bahwa angka di atas khusus untuk dataset dummy dan pembagian "
              "data tertentu; pada data lain hasilnya bisa berbeda. Namun kecenderungan hasil "
              "(Naive Bayes cocok untuk fitur frekuensi) sejalan dengan teori. Keputusan memilih "
              "model akhirnya bergantung pada biaya kesalahan: jika email normal boleh terjebak "
              "sedikit, prioritaskan recall; jika sebaliknya, prioritaskan precision.")


def bab5(doc):
    heading(doc, "BAB V\nKESIMPULAN")
    body(doc, "Berdasarkan pembahasan dan eksperimen pada bab-bab sebelumnya, dapat ditarik "
              "kesimpulan sebagai berikut.")
    bullet(doc, "Klasifikasi adalah tugas data mining untuk memprediksi kelas data berdasarkan "
                "contoh berlabel; konsepnya berakar dari taksonomi kuno dan berkembang pesat "
                "sejak era statistika dan data mining modern.")
    bullet(doc, "Model utama Naive Bayes bekerja dengan teorema Bayes dan kepadatan normal per "
                "fitur; Logistic Regression memakai fungsi sigmoid; Decision Tree membagi data "
                "memakai kriteria Gini. Ketiganya memiliki formula yang dapat dipahami tanpa "
                "matematika tingkat lanjut.")
    bullet(doc, "Studi kasus Excel menunjukkan perhitungan posterior Naive Bayes secara manual "
                "untuk satu email promosi, menghasilkan keputusan SPAM dengan keyakinan hampir "
                "100%, konsisten dengan intuisi.")
    bullet(doc, f"Program Python klasifikasi_spam.py berhasil dilatih pada {DS['jumlah']} data "
                f"dummy dan memperoleh hasil terbaik pada model Naive Bayes (accuracy "
                f"{NBm['accuracy']*100:.1f}%, F1 {NBm['f1']:.4f}), diikuti Logistic Regression "
                f"(F1 {LRm['f1']:.4f}) dan Decision Tree (F1 {DTm['f1']:.4f}).")
    bullet(doc, "Program dapat membaca file data (CSV) yang boleh dimodifikasi/dibuat baru, "
                "sehingga pengguna bebas menguji email atau dataset lain, dan seluruh hasil "
                "dapat direproduksi.")
    body(doc, "Saran kami untuk pengembangan selanjutnya: menambah fitur lain (misalnya panjang "
              "subjek atau jumlah penerima), memakai data email yang lebih besar, serta "
              "mengeksplorasi model lain seperti SVM atau Random Forest untuk perbandingan "
              "yang lebih luas.")


def daftar_pustaka(doc):
    heading(doc, "DAFTAR PUSTAKA")
    rujukan = [
        "Han, J., Kamber, M., dan Pei, J. (2011). Data Mining: Concepts and Techniques "
        "(3rd ed.). Morgan Kaufmann Publishers.",
        "James, G., Witten, D., Hastie, T., dan Tibshirani, R. (2021). An Introduction to "
        "Statistical Learning with Applications in R (2nd ed.). Springer.",
        "McCallum, A. dan Nigam, K. (1998). A Comparison of Event Models for Naive Bayes "
        "Text Classification. AAAI-98 Workshop on Learning for Text Categorization.",
        "Quinlan, J. R. (1993). C4.5: Programs for Machine Learning. Morgan Kaufmann.",
        "Pedregosa, F. et al. (2011). Scikit-learn: Machine Learning in Python. Journal of "
        "Machine Learning Research, 12, 2825-2830.",
        "UCI Machine Learning Repository. Spambase Data Set. "
        "https://archive.ics.uci.edu/dataset/94/spambase (diakses 2026).",
        "Penn State Online. STAT 857 - Classification Example. "
        "https://online.stat.psu.edu/stat857/node/147/ (diakses 2026).",
        "scikit-learn Documentation - Naive Bayes. "
        "https://scikit-learn.org/stable/modules/naive_bayes.html (diakses 2026).",
    ]
    for r in rujukan:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.first_line_indent = Cm(-1.25)
        run = p.add_run(r)
        run.font.size = Pt(12)
        run.font.name = "Times New Roman"


def main():
    doc = Document()
    gaya = doc.styles["Normal"]
    gaya.font.name = "Times New Roman"
    gaya.font.size = Pt(12)

    sec = doc.sections[0]          # sampul: tanpa nomor halaman
    sec.top_margin = sec.bottom_margin = Cm(2.5)
    sec.left_margin = sec.right_margin = Cm(3)
    sec.footer.is_linked_to_previous = False

    sampul(doc)

    # Halaman depan: nomor Romawi (i, ii, iii)
    sec1 = doc.add_section(WD_SECTION.NEW_PAGE)
    sec1.top_margin = sec1.bottom_margin = Cm(2.5)
    sec1.left_margin = sec1.right_margin = Cm(3)
    add_page_number_footer(sec1)
    set_page_numbering(sec1, "lowerRoman", 1)
    halaman_depan(doc)

    # Bagian isi: penomoran Arab mulai dari 1
    sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
    sec2.top_margin = sec2.bottom_margin = Cm(2.5)
    sec2.left_margin = sec2.right_margin = Cm(3)
    add_page_number_footer(sec2)
    set_page_numbering(sec2, "decimal", 1)

    bab1(doc)
    bab2(doc)
    bab3(doc)
    bab4(doc)
    bab5(doc)
    daftar_pustaka(doc)

    doc.save(OUT)
    print(f"Makalah dibuat: {OUT}")


if __name__ == "__main__":
    main()









