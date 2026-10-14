"""
TGS 3 - Data Mining & Information Retrieval (IFB-307)
Klasifikasi Email Spam dengan Naive Bayes (utama), Logistic Regression
(pembanding), dan Decision Tree (penjelas visual).

Cara pakai:
  python klasifikasi_spam.py
      -> melatih ketiga model dari spam_dummy.csv, mencetak metrik evaluasi,
         menyimpan grafik ke folder "hasil/" dan hasil_eksperimen.json,
         lalu membuka menu interaktif (jika dijalankan dari jendela terminal /
         klik ganda, supaya jendela tidak langsung tertutup)
  python klasifikasi_spam.py --data data_lain.csv
      -> melatih dari file CSV lain (boleh diubah/dibuat sendiri)
  python klasifikasi_spam.py --prediksi 4 0 2 7 1
      -> memprediksi satu email baru: kata_free kata_menang jumlah_link
         huruf_kapital jumlah_dolar
  python klasifikasi_spam.py --prediksi-teks "Congratulations, you WON a free prize!"
      -> memprediksi SATU pesan mentah (teks); fitur diekstrak otomatis
  python klasifikasi_spam.py --data sms_spam.csv
      -> data TEKS MENTAH asli (label,message) juga diterima; 5 fitur
         diekstrak otomatis lewat fungsi ekstrak_fitur()
  python klasifikasi_spam.py --menu
      -> menu interaktif di terminal

Letak model klasifikasi di kode ini:
  - latih_dan_evaluasi()           : membuat & melatih model, menghitung metrik
  - GaussianNB().fit(...)          : MODEL NAIVE BAYES (utama)
  - LogisticRegression().fit(...)  : MODEL LOGISTIC REGRESSION (pembanding)
  - DecisionTreeClassifier().fit() : MODEL DECISION TREE (penjelas visual)
  - prediksi_satu()                : memakai model terlatih untuk email baru
  - ekstrak_fitur()                : mengubah teks pesan mentah -> 5 fitur angka
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys

import matplotlib

matplotlib.use("Agg")  # backend non-interaktif supaya aman di terminal
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score,
    confusion_matrix,
    precision_recall_fscore_support,
)
from sklearn.model_selection import train_test_split
from sklearn.naive_bayes import GaussianNB
from sklearn.tree import DecisionTreeClassifier, plot_tree

# ---------------------------------------------------------------------------
# Konstanta
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DATA = os.path.join(BASE_DIR, "spam_dummy.csv")
HASIL_DIR = os.path.join(BASE_DIR, "hasil")
FITUR = ["kata_free", "kata_menang", "jumlah_link", "huruf_kapital", "jumlah_dolar"]
LABEL_NAMA = {0: "BUKAN SPAM (ham)", 1: "SPAM"}
RANDOM_STATE = 42

# Kata kunci kolom untuk format data TEKS MENTAH (huruf besar-kecil bebas)
TEKS_KOLOM = ("message", "text", "teks", "pesan", "body", "content", "msg", "sms", "v2")
LABEL_KOLOM = ("label", "category", "class", "kategori", "v1")
LABEL_KE_ANGKA = {
    "spam": 1, "spammail": 1, "1": 1, "yes": 1, "true": 1,
    "ham": 0, "0": 0, "no": 0, "false": 0, "normal": 0,
    "non-spam": 0, "not spam": 0, "bukan spam": 0,
}


# ---------------------------------------------------------------------------
# Ekstraksi fitur dari teks mentah (untuk data spam berupa teks asli)
# ---------------------------------------------------------------------------
def ekstrak_fitur(teks: str) -> list[int]:
    """Mengubah satu pesan/email TEKS MENTAH menjadi 5 fitur angka
    (formatnya persis sama dengan spam_dummy.csv):
      kata_free     : jumlah kata "free"/"gratis"      (case-insensitive)
      kata_menang   : jumlah kata win/winner/menang/hadiah
      jumlah_link   : jumlah URL (http/https/www/akhiran domain)
      huruf_kapital : panjang rangkaian HURUF KAPITAL terpanjang (mis. GRATIS)
      jumlah_dolar  : jumlah simbol '$'
    """
    t = str(teks)
    deret_besar = [len(m) for m in re.findall(r"\b[A-Z]{2,}\b", t)]
    jumlah_link = len(re.findall(
        r"(?:https?://\S+|www\.\S+|\b[\w-]+\.(?:com|net|org|id|co\.uk|xyz|info|biz|us|uk|me)\b)",
        t, flags=re.IGNORECASE))
    return [
        len(re.findall(r"\b(?:free|gratis)\b", t, flags=re.IGNORECASE)),
        len(re.findall(r"\b(?:win|winner|menang|hadiah)\b", t, flags=re.IGNORECASE)),
        jumlah_link,
        max(deret_besar, default=0),
        t.count("$"),
    ]


def _label_ke_angka(v) -> int | None:
    """Mengubah label teks ('spam'/'ham'/dst.) menjadi 1/0; None jika tak dikenal."""
    return LABEL_KE_ANGKA.get(str(v).strip().lower())


def _dari_teks(label_seri: pd.Series, teks_seri: pd.Series, sumber: str) -> pd.DataFrame:
    """Membuat tabel fitur dari pasangan label + teks pesan mentah."""
    df = pd.DataFrame([ekstrak_fitur(t) for t in teks_seri], columns=FITUR)
    df["label"] = [_label_ke_angka(v) for v in label_seri]
    dikenal = df["label"].notna()
    if not dikenal.all():
        contoh = sorted({str(v) for v in label_seri[~dikenal].unique()[:5]})
        print(f"PERINGATAN: {int((~dikenal).sum())} baris berlabel tak dikenal {contoh} - dilewati.")
        df = df[dikenal]
    df = df.astype({c: int for c in FITUR} | {"label": int}).reset_index(drop=True)
    print(f"Data TEKS MENTAH '{sumber}': {len(df)} pesan -> fitur diekstrak otomatis "
          f"dengan ekstrak_fitur() (spam={int((df['label'] == 1).sum())}, "
          f"ham={int((df['label'] == 0).sum())}).")
    return df


def muat_data(path: str) -> pd.DataFrame:
    """Membaca file data email. Menerima DUA format:
    (a) fitur numerik siap pakai (kolom kata_free ... label), ATAU
    (b) teks mentah + label (label/message) - fitur diekstrak otomatis."""
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"File '{path}' tidak ditemukan. Jalankan 'python buat_dataset.py' / "
            "'python unduh_sms_spam.py', atau sediakan CSV sendiri. Format yang diterima: "
            "(a) kolom " + ", ".join(FITUR + ["label"]) +
            " ATAU (b) kolom label + message (teks mentah)."
        )

    df = None
    try:
        df = pd.read_csv(path)
    except pd.errors.ParserError:
        # Kemungkinan besar bukan CSV koma-biasa (mis. TSV tanpa header) -> nanti di bawah
        pass

    # (a) Sudah berupa fitur numerik
    if df is not None and all(c in df.columns for c in FITUR + ["label"]):
        return df

    # (b) Format teks dengan header, mis. "label,message" atau "Category,Message"
    if df is not None:
        kolom_label = next((c for c in df.columns if str(c).strip().lower() in LABEL_KOLOM), None)
        kolom_teks = next((c for c in df.columns if str(c).strip().lower() in TEKS_KOLOM), None)
        if kolom_label is not None and kolom_teks is not None:
            return _dari_teks(df[kolom_label], df[kolom_teks], sumber=os.path.basename(path))

    # (b') Format teks TANPA header: 2 kolom "label<separator>teks" (mis. SMSSpamCollection)
    # Coba deteksi otomatis dulu, lalu eksplisit TAB, lalu koma (sniffer bisa salah pilih).
    for pemisah in ("auto", "\t", ","):
        try:
            if pemisah == "auto":
                df2 = pd.read_csv(path, header=None, sep=None, engine="python")
            else:
                df2 = pd.read_csv(path, header=None, sep=pemisah)
        except Exception:
            continue
        if df2.shape[1] == 2:
            return _dari_teks(df2.iloc[:, 0], df2.iloc[:, 1], sumber=os.path.basename(path))

    kolom_dikenal = list(df.columns) if df is not None else "(tidak terbaca sebagai tabel biasa)"
    raise ValueError(
        f"Format file '{path}' tidak dikenal. Kolom yang ada: {kolom_dikenal}.\n"
        "Dukungan: (a) " + ", ".join(FITUR + ["label"]) +
        " ATAU (b) label + teks pesan (kolom label/category + message/text)."
    )


def buat_model():
    """Membuat tiga model klasifikasi yang dibandingkan dalam tugas ini."""
    return {
        "Naive Bayes": GaussianNB(),  # MODEL UTAMA - teorema Bayes + distribusi normal
        "Logistic Regression": LogisticRegression(max_iter=1000, random_state=RANDOM_STATE),
        "Decision Tree": DecisionTreeClassifier(max_depth=4, random_state=RANDOM_STATE),
    }


def latih_dan_evaluasi(df: pd.DataFrame):
    """Melatih ketiga model dan menghitung accuracy, precision, recall, F1, confusion matrix."""
    X = df[FITUR].to_numpy()
    y = df["label"].to_numpy()
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.25, random_state=RANDOM_STATE, stratify=y
    )

    hasil, model_terlatih = {}, {}
    for nama, model in buat_model().items():
        model.fit(X_train, y_train)                      # <-- proses "belajar" pola dari data latih
        y_pred = model.predict(X_test)                   # <-- model menebak data uji
        p, r, f1, _ = precision_recall_fscore_support(
            y_test, y_pred, average="binary", zero_division=0
        )
        cm = confusion_matrix(y_test, y_pred, labels=[0, 1])
        hasil[nama] = {
            "accuracy": round(float(accuracy_score(y_test, y_pred)), 4),
            "precision": round(float(p), 4),
            "recall": round(float(r), 4),
            "f1": round(float(f1), 4),
            "confusion_matrix": cm.tolist(),
        }
        model_terlatih[nama] = model
        print(f"\n=== {nama} ===")
        print(f"  Accuracy : {hasil[nama]['accuracy']:.2%}")
        print(f"  Precision: {hasil[nama]['precision']:.4f}  (ketepatan deteksi spam)")
        print(f"  Recall   : {hasil[nama]['recall']:.4f}  (spam yang berhasil tertangkap)")
        print(f"  F1-score : {hasil[nama]['f1']:.4f}")
        print(f"  Confusion matrix [ [TN, FP], [FN, TP] ]: {cm.tolist()}")

    return hasil, model_terlatih, (X_train, X_test, y_train, y_test)


def statistik_naive_bayes(model: GaussianNB) -> dict:
    """Mengambil parameter Gauss (mean, std) dan prior tiap kelas dari Naive Bayes
    untuk ditampilkan pada studi kasus Excel/makalah."""
    return {
        "kelas": [int(c) for c in model.classes_],
        "prior": [float(v) for v in model.class_prior_],
        "mean": {
            f: [round(float(model.theta_[c][j]), 4) for c in range(len(model.classes_))]
            for j, f in enumerate(FITUR)
        },
        "std": {
            f: [round(float(np.sqrt(model.var_[c][j])), 4) for c in range(len(model.classes_))]
            for j, f in enumerate(FITUR)
        },
        "fitur": FITUR,
    }


def simpan_grafik(hasil, model_terlatih, data_split, nb_stats, prefix=""):
    """Menyimpan grafik: confusion matrix, perbandingan model, pohon keputusan, parameter NB.
    prefix: nama awal berkas sesuai dataset (kosong untuk spam_dummy / dataset resmi dokumen)."""
    os.makedirs(HASIL_DIR, exist_ok=True)

    # 1) Confusion matrix tiap model
    for nama, m in hasil.items():
        cm = np.array(m["confusion_matrix"])
        fig, ax = plt.subplots(figsize=(4.2, 3.6))
        ax.imshow(cm, cmap="Blues")
        for (i, j), v in np.ndenumerate(cm):
            ax.text(j, i, str(v), ha="center", va="center",
                    color="white" if v > cm.max() / 2 else "black", fontsize=15)
        ax.set_xlabel("Prediksi")
        ax.set_ylabel("Kenyataan")
        ax.set_xticks([0, 1], ["Bukan Spam", "Spam"])
        ax.set_yticks([0, 1], ["Bukan Spam", "Spam"])
        ax.set_title(f"Confusion Matrix - {nama}")
        fig.tight_layout()
        fig.savefig(os.path.join(HASIL_DIR, f"{prefix}confusion_{nama.lower().replace(' ', '_')}.png"), dpi=150)
        plt.close(fig)

    # 2) Bar chart perbandingan metrik
    metrik = ["accuracy", "precision", "recall", "f1"]
    fig, ax = plt.subplots(figsize=(7.5, 4.2))
    x = np.arange(len(metrik))
    width = 0.25
    for i, (nama, m) in enumerate(hasil.items()):
        ax.bar(x + (i - 1) * width, [m[k] for k in metrik], width, label=nama)
        for j, k in enumerate(metrik):
            ax.text(j + (i - 1) * width, m[k] + 0.015, f"{m[k]:.2f}", ha="center", fontsize=7)
    ax.set_xticks(x, ["Accuracy", "Precision", "Recall", "F1"])
    ax.set_ylim(0, 1.08)
    ax.set_title("Perbandingan Kinerja 3 Model Klasifikasi")
    ax.legend()
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(HASIL_DIR, f"{prefix}perbandingan_model.png"), dpi=150)
    plt.close(fig)

    # 3) Visualisasi Decision Tree
    fig, ax = plt.subplots(figsize=(12, 6))
    plot_tree(
        model_terlatih["Decision Tree"],
        feature_names=FITUR,
        class_names=["Bukan Spam", "Spam"],
        filled=True,
        rounded=True,
        fontsize=8,
        ax=ax,
    )
    ax.set_title("Decision Tree - aturan klasifikasi email spam")
    fig.tight_layout()
    fig.savefig(os.path.join(HASIL_DIR, f"{prefix}decision_tree.png"), dpi=150)
    plt.close(fig)

    # 4) Heatmap rata-rata fitur per kelas - bahan studi kasus Naive Bayes
    fig, ax = plt.subplots(figsize=(6.5, 3.4))
    means = np.array([nb_stats["mean"][f] for f in FITUR]).T  # 2 kelas x 5 fitur
    im = ax.imshow(means, cmap="Oranges")
    ax.set_xticks(range(len(FITUR)), FITUR, rotation=20, ha="right")
    ax.set_yticks([0, 1], ["Bukan Spam", "Spam"])
    for i in range(means.shape[0]):
        for j in range(means.shape[1]):
            ax.text(j, i, f"{means[i, j]:.2f}", ha="center", va="center", fontsize=9)
    ax.set_title("Rata-rata fitur per kelas (parameter Naive Bayes)")
    fig.colorbar(im, shrink=0.8)
    fig.tight_layout()
    fig.savefig(os.path.join(HASIL_DIR, f"{prefix}naive_bayes_mean.png"), dpi=150)
    plt.close(fig)

    print(f"\nGrafik disimpan di: {HASIL_DIR}")


def prediksi_satu(model_terlatih, nilai_fitur: list[float]) -> None:
    """Memprediksi satu email baru berdasarkan 5 nilai fitur yang dimasukkan pengguna."""
    baris = pd.DataFrame([nilai_fitur], columns=FITUR)
    print("\n--- Hasil prediksi email baru ---")
    print(f"Input: {dict(zip(FITUR, nilai_fitur))}")
    for nama, model in model_terlatih.items():
        kel = int(model.predict(baris.to_numpy())[0])
        teks = f"{kel} -> {LABEL_NAMA[kel]}"
        if hasattr(model, "predict_proba"):
            proba = model.predict_proba(baris.to_numpy())[0]
            teks += f" (probabilitas spam = {proba[1]:.2%})"
        print(f"  {nama:<22}: {teks}")


def prediksi_teks(model_terlatih, teks: str) -> None:
    """Memprediksi SATU pesan/email mentah (berupa teks bebas).
    Teks diubah dulu menjadi 5 fitur dengan ekstrak_fitur(), lalu dimasukkan
    ke model yang sama - jadi hasilnya sebanding dengan data numerik lain."""
    fitur = ekstrak_fitur(teks)
    print("\n--- Ekstraksi fitur dari teks mentah ---")
    for nama, nilai in zip(FITUR, fitur):
        print(f"  {nama:<14} = {nilai}")
    prediksi_satu(model_terlatih, fitur)


def menu_interaktif(hasil, model_terlatih):
    """Menu terminal sederhana (opsional)."""
    while True:
        print("\n========== MENU KLASIFIKASI EMAIL SPAM ==========")
        print("1. Bandingkan ketiga model")
        print("2. Uji email baru (isi nilai fitur angka)")
        print("3. Contoh prediksi cepat")
        print("4. Uji teks pesan/email langsung (tempel teks)")
        print("0. Keluar")
        p = input("Pilih [0-4]: ").strip()
        if p == "1":
            for nama, m in hasil.items():
                print(f"  {nama:<22}: akurasi={m['accuracy']:.4f}, "
                      f"precision={m['precision']:.4f}, recall={m['recall']:.4f}, F1={m['f1']:.4f}")
        elif p == "2":
            try:
                nilai = [float(input(f"  {f} = ").strip()) for f in FITUR]
                prediksi_satu(model_terlatih, nilai)
            except ValueError:
                print("  Input harus berupa angka.")
        elif p == "3":
            prediksi_satu(model_terlatih, [5, 3, 3, 9, 2])   # mirip email promosi
            prediksi_satu(model_terlatih, [0, 0, 1, 2, 0])   # mirip email normal
        elif p == "4":
            t = input("  Tempel teks pesan/email: ").strip()
            if t:
                prediksi_teks(model_terlatih, t)
            else:
                print("  Teks kosong.")
        elif p == "0":
            break


def stdin_adalah_konsol() -> bool:
    """True hanya jika stdin benar-benar konsol (klik ganda / jendela terminal),
    bukan pipe, file, atau perangkat NUL. Catatan: isatty() di Windows menilai
    perangkat NUL sebagai True, jadi perlu dicek dengan GetConsoleMode."""
    if not sys.stdin.isatty():
        return False
    if os.name != "nt":
        return True
    try:
        import ctypes
        import msvcrt

        h = msvcrt.get_osfhandle(0)
        mode = ctypes.c_uint()
        ok = ctypes.windll.kernel32.GetConsoleMode(h, ctypes.byref(mode))
        return bool(ok)
    except Exception:
        return True


def main() -> None:
    parser = argparse.ArgumentParser(description="Klasifikasi email spam (TGS 3)")
    parser.add_argument("--data", default=DEFAULT_DATA, help="path CSV dataset")
    parser.add_argument("--prediksi", nargs=5, type=float,
                        metavar=("KF", "KM", "JL", "HK", "JD"),
                        help="prediksi 1 email: kata_free kata_menang jumlah_link "
                             "huruf_kapital jumlah_dolar")
    parser.add_argument("--prediksi-teks", dest="prediksi_teks", default=None,
                        help="prediksi 1 pesan/email TEKS MENTAH, contoh: "
                             "--prediksi-teks \"You WON a free prize!\"")
    parser.add_argument("--menu", action="store_true", help="jalankan menu interaktif")
    args = parser.parse_args()

    df = muat_data(args.data)
    print(f"Dataset: {args.data}  ({len(df)} email, "
          f"spam={int((df['label'] == 1).sum())}, "
          f"bukan spam={int((df['label'] == 0).sum())})")

    hasil, model_terlatih, split = latih_dan_evaluasi(df)
    nb_stats = statistik_naive_bayes(model_terlatih["Naive Bayes"])

    # Nama berkas hasil disesuaikan dataset; spam_dummy (dataset resmi dokumen)
    # memakai nama lama supaya makalah/PPT/Excel tidak tercampur hasil dataset lain.
    stem = os.path.splitext(os.path.basename(args.data))[0]
    prefix = "" if stem == "spam_dummy" else f"{stem}_"
    simpan_grafik(hasil, model_terlatih, split, nb_stats, prefix=prefix)

    best = max(hasil, key=lambda k: hasil[k]["f1"])
    payload = {
        "dataset": {
            "nama": os.path.basename(args.data),
            "jumlah": int(len(df)),
            "spam": int((df["label"] == 1).sum()),
            "bukan_spam": int((df["label"] == 0).sum()),
            "fitur": FITUR,
            "test_size": 0.25,
            "random_state": RANDOM_STATE,
        },
        "hasil": hasil,
        "model_terbaik_f1": best,
        "naive_bayes_stats": nb_stats,
    }
    os.makedirs(HASIL_DIR, exist_ok=True)
    if stem == "spam_dummy":
        # Dataset resmi dokumen: tulis HANYA sumber angka dokumen (tunggal, tanpa duplikat)
        out_json = os.path.join(HASIL_DIR, "hasil_eksperimen.json")
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
        print(f"Hasil eksperimen disimpan: {out_json}  (sumber angka dokumen TGS)")
    else:
        # Dataset lain: hasil terpisah supaya tidak pernah menimpa angka dokumen
        out_json = os.path.join(HASIL_DIR, f"hasil_{stem}.json")
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
        print(f"Hasil eksperimen disimpan: {out_json}")
        print("Catatan: hasil_eksperimen.json (angka dokumen makalah/PPT/Excel) TIDAK "
              "diubah; jalankan tanpa --data untuk memakai dataset dummy resmi.")
    print(f"Model dengan F1 terbaik  : {best} (F1={hasil[best]['f1']:.4f})")

    if args.prediksi:
        prediksi_satu(model_terlatih, args.prediksi)
    if args.prediksi_teks:
        prediksi_teks(model_terlatih, args.prediksi_teks)
    if args.menu:
        menu_interaktif(hasil, model_terlatih)
    elif len(sys.argv) == 1 and stdin_adalah_konsol():
        # Dijalankan tanpa argumen dari jendela (klik ganda / terminal):
        # buka menu, dan tahan jendela sampai pengguna menutupnya.
        print("\nMode menu interaktif (pilih 0 untuk keluar).")
        try:
            menu_interaktif(hasil, model_terlatih)
            input("\nSelesai. Tekan ENTER untuk menutup jendela...")
        except EOFError:
            pass  # stdin tidak tersedia (jalankan otomatis) - keluar dengan rapi


if __name__ == "__main__":
    main()



