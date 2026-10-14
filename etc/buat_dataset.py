"""
TGS 3 - Data Mining & Information Retrieval (IFB-307)
Pembuat dataset DUMMY untuk klasifikasi email spam.

PENTING: Dataset ini adalah data buatan (sintetis) yang dibuat dengan seed
tetap agar bisa direproduksi. BUKAN dataset asli UCI Spambase.
Struktur fiturnya terinspirasi dari UCI Spambase (frekuensi kata, huruf
kapital, dll.) tetapi nilainya dibuat sederhana agar mudah dipahami.

Fitur (semua bilangan bulat):
  kata_free    : jumlah kata "free"/"gratis" dalam email
  kata_menang  : jumlah kata "win"/"menang"/"hadiah"
  jumlah_link  : jumlah tautan/URL dalam email
  huruf_kapital: panjang rangkaian huruf KAPITAL terpanjang (mis. "GRATIS")
  jumlah_dolar : jumlah simbol '$' dalam email
  label        : 1 = spam, 0 = bukan spam (ham)
"""

import os

import numpy as np
import pandas as pd

SEED = 42
N_EMAIL = 240          # jumlah email
P_SPAM = 0.45          # proporsi spam
OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "spam_dummy.csv")


def buat_email(spam: bool, rng: np.random.Generator) -> tuple[int, int, int, int, int]:
    """Menghasilkan satu baris fitur untuk satu email."""
    if spam:
        # Email spam cenderung banyak kata promosi, link, huruf kapital, dolar
        kata_free = rng.poisson(2.5)
        kata_menang = rng.poisson(1.8)
        jumlah_link = rng.poisson(2.0)
        huruf_kapital = rng.poisson(5) + 1
        jumlah_dolar = rng.poisson(1.3)
    else:
        # Email normal (ham) cenderung bersih dari unsur promosi
        kata_free = rng.poisson(0.6)
        kata_menang = rng.poisson(0.2)
        jumlah_link = rng.poisson(0.8)
        huruf_kapital = rng.poisson(3)
        jumlah_dolar = rng.poisson(0.2)

    # Tambahkan tumpang tindih agar data realistis (kelas tidak terpisah 100%)
    if not spam and rng.random() < 0.15:
        # email promosi resmi (newsletter/diskon toko) yang mirip spam
        kata_free += int(rng.integers(2, 5))
        jumlah_link += int(rng.integers(1, 3))
        huruf_kapital += int(rng.integers(2, 5))
    if spam and rng.random() < 0.12:
        # spam yang tampak "tenang" (sedikit teks, sedikit link)
        kata_free = max(0, kata_free - 3)
        kata_menang = max(0, kata_menang - 2)
        jumlah_link = max(0, jumlah_link - 2)

    # Sedikit "gangguan" acak pada tiap email agar tidak sempurna
    kata_free += int(rng.integers(0, 2))
    huruf_kapital += int(rng.integers(0, 2))

    return int(kata_free), int(kata_menang), int(jumlah_link), int(huruf_kapital), int(jumlah_dolar)


def main() -> None:
    rng = np.random.default_rng(SEED)
    data = []
    for _ in range(N_EMAIL):
        spam = bool(rng.random() < P_SPAM)
        fitur = buat_email(spam, rng)
        data.append((*fitur, int(spam)))

    df = pd.DataFrame(
        data,
        columns=["kata_free", "kata_menang", "jumlah_link", "huruf_kapital", "jumlah_dolar", "label"],
    )
    df.to_csv(OUT_PATH, index=False)
    print(f"Dataset dummy dibuat: {OUT_PATH}")
    print(f"Jumlah baris : {len(df)}")
    print(f"Spam (1)     : {(df['label'] == 1).sum()}")
    print(f"Bukan spam(0): {(df['label'] == 0).sum()}")
    print(df.head())


if __name__ == "__main__":
    main()
