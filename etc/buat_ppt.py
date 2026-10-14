"""
Membuat dokumen presentasi PowerPoint TGS 3 - Klasifikasi Email Spam.
Output: TGS_3_Kelompok_1_Klasifikasi_Email_Spam_NRP_Nama_PPT.pptx

Isi slide disusun identik dengan makalah: definisi/sejarah, formula,
studi kasus Excel, implementasi Python, hasil eksperimen (angka asli dari
hasil/hasil_eksperimen.json), kesimpulan, dan daftar pustaka.
"""

import json
import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

BASE = os.path.dirname(os.path.abspath(__file__))
JSON = os.path.join(BASE, "hasil", "hasil_eksperimen.json")
HASIL = os.path.join(BASE, "hasil")
OUT = os.path.join(BASE, "TGS_3_Kelompok_1_Klasifikasi_Email_Spam_NRP_Nama_PPT.pptx")

with open(JSON, encoding="utf-8") as f:
    P = json.load(f)
H, NB, DS = P["hasil"], P["naive_bayes_stats"], P["dataset"]
NBm, LRm, DTm = H["Naive Bayes"], H["Logistic Regression"], H["Decision Tree"]

BIRU = RGBColor(0x1F, 0x4E, 0x79)
PUTIH = RGBColor(0xFF, 0xFF, 0xFF)
ABU = RGBColor(0x40, 0x40, 0x40)
ORANYE = RGBColor(0xC5, 0x5A, 0x11)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def kotak(slide, x, y, w, h, warna=RGBColor(0x1F, 0x4E, 0x79)):
    from pptx.enum.shapes import MSO_SHAPE
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = warna
    s.line.fill.background()
    return s


def teks(slide, x, y, w, h, isi, ukuran=18, bold=False, warna=ABU,
         align=PP_ALIGN.LEFT, spasi=1.0):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    baris = isi.split("\n") if isinstance(isi, str) else isi
    for i, b in enumerate(baris):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = b
        p.alignment = align
        p.space_after = Pt(6)
        p.line_spacing = spasi
        for r in p.runs:
            r.font.size = Pt(ukuran)
            r.font.bold = bold
            r.font.color.rgb = warna
            r.font.name = "Calibri"
    return tb


def slide_judul(judul, sub=None):
    s = prs.slides.add_slide(BLANK)
    kotak(s, 0, 0, prs.slide_width, Inches(0.15), BIRU)
    kotak(s, 0, Inches(1.05), Inches(0.28), Inches(0.7), ORANYE)
    teks(s, Inches(0.5), Inches(0.95), Inches(12.3), Inches(1), judul, ukuran=30,
         bold=True, warna=BIRU)
    if sub:
        teks(s, Inches(0.5), Inches(1.7), Inches(12.3), Inches(0.5), sub,
             ukuran=16, warna=ORANYE, bold=True)
    return s


def poin(slide, daftar, ukuran=17, y=Inches(2.0), h=Inches(5.0)):
    isi = "\n".join(("• " + b) if not b.startswith(" ") else b for b in daftar)
    return teks(slide, Inches(0.5), y, Inches(12.3), h, isi, ukuran=ukuran, spasi=1.15)


def slide_cover():
    s = prs.slides.add_slide(BLANK)
    kotak(s, 0, 0, prs.slide_width, prs.slide_height, BIRU)
    kotak(s, 0, Inches(3.05), prs.slide_width, Inches(0.06), ORANYE)
    teks(s, Inches(0.8), Inches(0.6), Inches(11.7), Inches(1),
         "INSTITUT TEKNOLOGI NASIONAL BANDUNG", ukuran=22, bold=True, warna=PUTIH,
         align=PP_ALIGN.CENTER)
    teks(s, Inches(0.8), Inches(1.15), Inches(11.7), Inches(0.6),
         "Fakultas Teknologi Industri - Program Studi Informatika", ukuran=16,
         warna=PUTIH, align=PP_ALIGN.CENTER)
    teks(s, Inches(0.8), Inches(1.75), Inches(11.7), Inches(0.6),
         "IFB-307 Data Mining & Information Retrieval - TGS 3", ukuran=16,
         warna=RGBColor(0xFF, 0xD9, 0x66), align=PP_ALIGN.CENTER, bold=True)
    teks(s, Inches(0.8), Inches(3.4), Inches(11.7), Inches(1.6),
         "KLASIFIKASI EMAIL SPAM DENGAN\nNAIVE BAYES, LOGISTIC REGRESSION,\nDAN DECISION TREE",
         ukuran=34, bold=True, warna=PUTIH, align=PP_ALIGN.CENTER, spasi=1.1)
    teks(s, Inches(0.8), Inches(5.6), Inches(11.7), Inches(1.2),
         "Kelompok 1 - Anggota: [Nama 1 (NRP)], [Nama 2 (NRP)]\n"
         f"Studi kasus pada dataset dummy {DS['jumlah']} email - Oktober 2026",
         ukuran=17, warna=PUTIH, align=PP_ALIGN.CENTER)


def slide_definisi():
    s = slide_judul("Definisi & Sejarah Klasifikasi", "Bab II - Tinjauan Pustaka")
    poin(s, [
        "Klasifikasi = tugas data mining untuk memprediksi kelas suatu data dari ciri-cirinya, "
        "berdasarkan contoh yang sudah berlabel (supervised learning).",
        "Tujuan: menentukan email termasuk SPAM (1) atau BUKAN SPAM (0).",
        "Sejarah: taksonomi Aristoteles (340 SM) & Linnaeus (abad ke-18) -> analisis diskriminan "
        "Fisher (1936) -> teorema Bayes (1763) -> algoritma ID3/C4.5 Quinlan (1986/1993).",
        "Era 1990-an: data mining populer; Naive Bayes menjadi klasik untuk teks & email spam.",
        "Di industri, filter spam berlapis: reputasi pengirim + aturan keamanan + klasifikasi "
        "machine learning (konfigurasi persis tiap penyedia tidak dipublikasikan).",
        "Contoh lain: deteksi penipuan kartu kredit, prediksi penyakit, pengenalan tulisan tangan.",
    ], ukuran=17)


def slide_naive_bayes():
    s = slide_judul("Model Utama: Naive Bayes", "Teorema Bayes + asumsi fitur independen")
    teks(s, Inches(0.5), Inches(1.9), Inches(12.3), Inches(0.5),
         "Rumus Teorema Bayes:", ukuran=17, bold=True, warna=ORANYE)
    teks(s, Inches(0.5), Inches(2.3), Inches(12.3), Inches(0.6),
         "P(C | X) = P(X | C) x P(C) / P(X)    ->    P(C | X) ∝ P(C) x PROD i P(x i | C)",
         ukuran=22, bold=True, warna=BIRU, align=PP_ALIGN.CENTER)
    poin(s, [
        "C = kelas (0/1), X = fitur email, P(C) = prior, P(x i|C) = peluang fitur ke-i.",
        "Data numerik -> Gunakan Gaussian Naive Bayes (kepadatan normal per fitur):",
    ], ukuran=16, y=Inches(3.15), h=Inches(1.4))
    teks(s, Inches(0.5), Inches(4.15), Inches(12.3), Inches(0.6),
         "P(x i | C) = 1 / sqrt(2 pi sigma²) x exp( -(x i - mu)² / (2 sigma²) )",
         ukuran=21, bold=True, warna=BIRU, align=PP_ALIGN.CENTER)
    poin(s, [
        "mu & sigma (rata-rata & simpangan baku) dipelajari dari data latih per kelas.",
        "Keputusan = kelas dengan posterior terbesar (argmax).",
        "Kelebihan: sangat cepat & cocok untuk fitur frekuensi kata. Kekurangan: asumsi "
        "independensi fitur tidak sepenuhnya realistis untuk bahasa.",
    ], ukuran=16, y=Inches(4.85), h=Inches(2.2))


def slide_logreg_tree():
    s = slide_judul("Model Pembanding & Penjelas Visual")
    kotak(s, Inches(0.5), Inches(1.95), Inches(6.0), Inches(4.9),
          RGBColor(0xEF, 0xF3, 0xF8))
    teks(s, Inches(0.7), Inches(2.1), Inches(5.6), Inches(0.5),
         "Logistic Regression (pembanding)", ukuran=18, bold=True, warna=BIRU)
    teks(s, Inches(0.7), Inches(2.6), Inches(5.6), Inches(0.9),
         "P(spam | X) = 1 / (1 + e^-(b0 + b1x1 + ... + bnxn))",
         ukuran=16, bold=True, warna=ORANYE, align=PP_ALIGN.CENTER)
    teks(s, Inches(0.7), Inches(3.5), Inches(5.6), Inches(3.2),
         "• b0 = intercept, b1..bn = bobot hasil belajar.\n"
         "• Sigmoid memetakan nilai linear ke probabilitas 0-1.\n"
         "• Ambang keputusan: P(spam|X) >= 0,5 -> spam.\n"
         "• Plus: cepat, bobot tiap fitur mudah ditafsirkan.\n"
         "• Minus: batas keputusan linear saja.",
         ukuran=15, spasi=1.15)
    kotak(s, Inches(6.8), Inches(1.95), Inches(6.0), Inches(4.9),
          RGBColor(0xFB, 0xF0, 0xE6))
    teks(s, Inches(7.0), Inches(2.1), Inches(5.6), Inches(0.5),
         "Decision Tree (penjelas visual)", ukuran=18, bold=True, warna=BIRU)
    teks(s, Inches(7.0), Inches(2.6), Inches(5.6), Inches(0.9),
         "Gini(node) = 1 - SUM k (p k)²",
         ukuran=16, bold=True, warna=ORANYE, align=PP_ALIGN.CENTER)
    teks(s, Inches(7.0), Inches(3.5), Inches(5.6), Inches(3.2),
         "• Data dibelah berulang: mis. \"kata_menang <= 0,5?\" (simpul akar model kita).\n"
         "• Gini = 0 berarti node murni (satu kelas).\n"
         "• Hasil berupa aturan if-then yang mudah dibaca.\n"
         "• Plus: mudah divisualkan & dipahami manusia.\n"
         "• Minus: pohon terlalu dalam mudah overfit.",
         ukuran=15, spasi=1.15)


def slide_metrik():
    s = slide_judul("Metrik Evaluasi", "Kenapa akurasi saja tidak cukup?")
    poin(s, [
        "Accuracy = (TP + TN) / semua data - proporsi prediksi benar.",
        "Precision = TP / (TP + FP) - ketepatan deteksi spam (email normal tidak salah tangkap).",
        "Recall = TP / (TP + FN) - spam yang berhasil tertangkap (spam tidak lolos).",
        "F1 = 2 x Precision x Recall / (Precision + Recall) - keseimbangan keduanya.",
        "FP: email normal masuk kotak spam (kehilangan email penting).",
        "FN: spam lolos ke kotak masuk (user menerima spam).",
        "Karena biaya FP & FN berbeda, model utama dipilih berdasarkan F1, bukan akurasi saja.",
    ], ukuran=17)


def slide_dataset():
    s = slide_judul("Dataset & Variabel", f"spam_dummy.csv - {DS['jumlah']} email (data dummy, "
                                          "bukan UCI Spambase asli)")
    baris = [["Fitur", "Arti", "Contoh nilai"],
             ["kata_free", "Jumlah kata promosi free/gratis", "6 = ada 6 kata promosi"],
             ["kata_menang", "Jumlah kata win/menang/hadiah", "4 = ajakan menang hadiah 4x"],
             ["jumlah_link", "Jumlah tautan/URL email", "4 = ada 4 link"],
             ["huruf_kapital", "Panjang rangkaian HURUF KAPITAL terpanjang", "12 = kata GRATIS (12 huruf)"],
             ["jumlah_dolar", "Jumlah simbol $", "3 = ada 3 simbol uang"],
             ["label (target)", "Kelas yang diprediksi", "1 = SPAM, 0 = BUKAN SPAM"]]
    t = s.shapes.add_table(len(baris), 3, Inches(0.5), Inches(2.1),
                           Inches(12.3), Inches(3.4)).table
    t.columns[0].width = Inches(2.6)
    t.columns[1].width = Inches(5.7)
    t.columns[2].width = Inches(4.0)
    for i, b in enumerate(baris):
        for j, sel in enumerate(b):
            c = t.cell(i, j)
            c.text = sel
            for p in c.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(14)
                    r.font.bold = i == 0
                    r.font.color.rgb = PUTIH if i == 0 else ABU
    teks(s, Inches(0.5), Inches(5.75), Inches(12.3), Inches(1.2),
         f"Komposisi: {DS['bukan_spam']} bukan spam / {DS['spam']} spam | "
         f"Pembagian 75% latih : 25% uji ({round(DS['jumlah']*DS['test_size'])} data uji) | "
         f"random_state = {DS['random_state']} (hasil selalu sama)",
         ukuran=16, bold=True, warna=ORANYE)


def slide_studi_kasus():
    s = slide_judul("Studi Kasus Excel", "Sheet \"Studi Kasus NB\" - perhitungan manual 1 email")
    teks(s, Inches(0.5), Inches(1.9), Inches(12.3), Inches(0.5),
         "Email diuji (promosi berhadiah): kata_free=6, kata_menang=4, jumlah_link=4, "
         "huruf_kapital=12, jumlah_dolar=3", ukuran=16, bold=True, warna=ORANYE)
    baris = [["Langkah", "Cara hitung di Excel", "Hasil"],
             ["1. Prior", "P(C) dari data latih",
              f"P(0)={NB['prior'][0]:.4f}; P(1)={NB['prior'][1]:.4f}"],
             ["2. Parameter", "mean & std tiap fitur per kelas (Tabel 3.2 makalah)",
              "mu/spam: kata_free=2,72; kata_menang=1,72; dst."],
             ["3. Likelihood", "NORM.DIST(x, mean, std, FALSE) per fitur",
              "dihitung per fitur, lalu dikalikan"],
             ["4. Posterior", "prior x produk likelihood, dinormalkan",
              "P(0)≈0,00%; P(1)≈100,00%"],
             ["5. Keputusan", "=IF(P1>=P0; \"SPAM\"; \"BUKAN SPAM\")", "SPAM (argmax)"]]
    t = s.shapes.add_table(len(baris), 3, Inches(0.5), Inches(2.5),
                           Inches(12.3), Inches(3.0)).table
    t.columns[0].width = Inches(2.4)
    t.columns[1].width = Inches(6.4)
    t.columns[2].width = Inches(3.5)
    for i, b in enumerate(baris):
        for j, sel in enumerate(b):
            c = t.cell(i, j)
            c.text = sel
            for p in c.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(13)
                    r.font.bold = i == 0
                    r.font.color.rgb = PUTIH if i == 0 else ABU
    teks(s, Inches(0.5), Inches(5.7), Inches(12.3), Inches(1.3),
         "Posterior tak ternormalisasi: bukan spam 1,83 x 10^-40 vs spam 2,22 x 10^-7\n"
         "=> P(spam|X) hampir 100% - konsisten dengan intuisi (email penuh kata FREE/WIN, "
         "link, huruf kapital, dan simbol $).", ukuran=16, warna=ABU)


def slide_implementasi():
    s = slide_judul("Implementasi Program Python", "File: klasifikasi_spam.py (scikit-learn)")
    kotak(s, Inches(0.5), Inches(1.95), Inches(7.4), Inches(4.6),
          RGBColor(0x26, 0x26, 0x26))
    teks(s, Inches(0.7), Inches(2.1), Inches(7.0), Inches(4.3),
         "def buat_model():\n"
         "    return {\n"
         '        "Naive Bayes": GaussianNB(),            # MODEL UTAMA\n'
         '        "Logistic Regression": LogisticRegression(max_iter=1000),\n'
         '        "Decision Tree": DecisionTreeClassifier(max_depth=4),\n'
         "    }\n\n"
         "for nama, model in buat_model().items():\n"
         "    model.fit(X_train, y_train)    # model BELAJAR pola\n"
         "    y_pred = model.predict(X_test) # model menebak data uji",
         ukuran=15, warna=RGBColor(0x6F, 0xD3, 0x6F), spasi=1.05)
    teks(s, Inches(8.2), Inches(2.0), Inches(4.7), Inches(4.8),
         "Cara menjalankan:\n"
         "1. Klik ganda \"Jalankan Program.bat\"\n"
         "   -> latih + evaluasi + menu (jendela tidak tertutup)\n"
         "2. python klasifikasi_spam.py --prediksi 6 4 4 12 3\n"
         "   -> tebak 1 email baru\n"
         "3. python klasifikasi_spam.py --data data_lain.csv\n"
         "   -> pakai file CSV lain (boleh dibuat sendiri)\n"
         "4. python klasifikasi_spam.py --menu\n"
         "   -> menu interaktif dari terminal\n\n"
         "Output: kelas 0/1 + probabilitas spam,\n"
         "metrik, grafik, hasil_eksperimen.json",
         ukuran=15, warna=ABU, spasi=1.1)


def slide_hasil():
    s = slide_judul("Hasil Eksperimen", "Angka asli dari hasil_eksperimen.json (data uji 60 email)")
    baris = [["Model", "Accuracy", "Precision", "Recall", "F1-score"]]
    for nama, m in H.items():
        baris.append([nama + (" (utama)" if nama == "Naive Bayes" else ""),
                      f"{m['accuracy']*100:.2f}%", f"{m['precision']:.4f}",
                      f"{m['recall']:.4f}", f"{m['f1']:.4f}"])
    t = s.shapes.add_table(len(baris), 5, Inches(0.5), Inches(2.0),
                           Inches(7.6), Inches(2.0)).table
    for i, b in enumerate(baris):
        for j, sel in enumerate(b):
            c = t.cell(i, j)
            c.text = sel
            for p in c.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for r in p.runs:
                    r.font.size = Pt(14)
                    r.font.bold = (i == 0) or (i == 1)
                    r.font.color.rgb = PUTIH if i == 0 else ABU
    gambar = os.path.join(HASIL, "perbandingan_model.png")
    if os.path.exists(gambar):
        s.shapes.add_picture(gambar, Inches(8.3), Inches(1.9), width=Inches(4.6))
    tn, fp = NBm["confusion_matrix"][0]
    fn, tp = NBm["confusion_matrix"][1]
    teks(s, Inches(0.5), Inches(4.3), Inches(7.6), Inches(2.6),
         f"Naive Bayes (model utama) unggul: F1 {NBm['f1']:.4f}, recall {NBm['recall']:.4f}.\n"
         f"Confusion matrix NB: TN={tn}, FP={fp}, FN={fn}, TP={tp}.\n"
         f"Logistic Regression dekat (F1 {LRm['f1']:.4f}) -> pembanding yang baik.\n"
         f"Decision Tree (F1 {DTm['f1']:.4f}) lebih rendah, tapi aturannya paling mudah dibaca.",
         ukuran=15, warna=ABU, spasi=1.15)
    teks(s, Inches(0.5), Inches(6.7), Inches(12.3), Inches(0.5),
         "Catatan: hasil khusus dataset dummy ini; pada data lain hasil bisa berbeda.",
         ukuran=13, warna=ORANYE)


def slide_visual():
    s = slide_judul("Visualisasi Model", "Confusion matrix & pohon keputusan")
    p1 = os.path.join(HASIL, "confusion_naive_bayes.png")
    p2 = os.path.join(HASIL, "decision_tree.png")
    if os.path.exists(p1):
        s.shapes.add_picture(p1, Inches(0.6), Inches(2.0), height=Inches(4.4))
    if os.path.exists(p2):
        s.shapes.add_picture(p2, Inches(5.4), Inches(2.0), width=Inches(7.4))
    teks(s, Inches(0.6), Inches(6.6), Inches(12.0), Inches(0.6),
         "Kiri: seberapa benar Naive Bayes per kelas. Kanan: aturan percabangan Decision Tree "
         "(tiap simpul = pertanyaan pada satu fitur).", ukuran=14, warna=ABU)


def slide_kesimpulan():
    s = slide_judul("Kesimpulan")
    poin(s, [
        "Klasifikasi = memprediksi kelas data dari ciri-cirinya; konsepnya tua (taksonomi) "
        "tapi berkembang pesat bersama data mining modern.",
        f"Naive Bayes (teorema Bayes + distribusi normal) terbukti paling sesuai untuk fitur "
        f"frekuensi email: F1 {NBm['f1']:.4f}, accuracy {NBm['accuracy']*100:.1f}%.",
        f"Logistic Regression (sigmoid) menjadi pembanding kuat (F1 {LRm['f1']:.4f}); "
        f"Decision Tree (Gini) memberi penjelasan visual aturan (F1 {DTm['f1']:.4f}).",
        "Studi kasus Excel membuktikan perhitungan posterior secara manual menghasilkan "
        "keputusan SPAM ≈ 100% untuk email promosi berhadiah.",
        "Program Python membaca file CSV (bisa diedit/dibuat baru), melatih model, mengevaluasi, "
        "dan menguji email baru - semua hasil dapat direproduksi.",
        "Saran: tambah fitur (panjang subjek, jumlah penerima), pakai data lebih besar, "
        "uji SVM/Random Forest.",
    ], ukuran=17)


def slide_pustaka():
    s = slide_judul("Daftar Pustaka")
    teks(s, Inches(0.5), Inches(1.9), Inches(12.3), Inches(5.0),
         "Han, J., Kamber, M., Pei, J. (2011). Data Mining: Concepts and Techniques (3rd ed.). "
         "Morgan Kaufmann.\n"
         "James, G. et al. (2021). An Introduction to Statistical Learning (2nd ed.). Springer.\n"
         "McCallum, A., Nigam, K. (1998). A Comparison of Event Models for Naive Bayes Text "
         "Classification. AAAI-98.\n"
         "Quinlan, J. R. (1993). C4.5: Programs for Machine Learning. Morgan Kaufmann.\n"
         "Pedregosa, F. et al. (2011). Scikit-learn: Machine Learning in Python. JMLR, 12.\n"
         "UCI Spambase Data Set - https://archive.ics.uci.edu/dataset/94/spambase\n"
         "Penn State STAT 857 - https://online.stat.psu.edu/stat857/node/147/\n"
         "Scikit-learn Naive Bayes - https://scikit-learn.org/stable/modules/naive_bayes.html",
         ukuran=15, spasi=1.25)


def main():
    slide_cover()
    slide_definisi()
    slide_naive_bayes()
    slide_logreg_tree()
    slide_metrik()
    slide_dataset()
    slide_studi_kasus()
    slide_implementasi()
    slide_hasil()
    slide_visual()
    slide_kesimpulan()
    slide_pustaka()
    prs.save(OUT)
    print(f"PPT dibuat: {OUT} ({len(prs.slides)} slide)")


if __name__ == "__main__":
    main()



