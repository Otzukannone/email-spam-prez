"""
Mengunduh dataset SMS SPAM ASLI dari UCI Machine Learning Repository
(SMS Spam Collection, dataset 228) lalu menyimpannya sebagai sms_spam.csv.

Sumber : https://archive.ics.uci.edu/dataset/228/sms+spam+collection
Isi    : 5.574 pesan SMS nyata berlabel "spam" / "ham" (bukan data buatan).
Format : kolom "label" (spam/ham) dan "message" (teks pesan).

Catatan: file ini TEKS MENTAH (bukan fitur angka). Program
klasifikasi_spam.py akan mengekstrak 5 fitur (kata_free, kata_menang,
jumlah_link, huruf_kapital, jumlah_dolar) dari teksnya secara otomatis,
sehingga model dan studi kasus Excel tetap sama.
"""

import io
import os
import zipfile

import pandas as pd
import requests

URL = "https://archive.ics.uci.edu/static/public/228/sms+spam+collection.zip"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sms_spam.csv")


def main() -> None:
    print(f"Mengunduh {URL} ...")
    r = requests.get(URL, timeout=60)
    r.raise_for_status()
    z = zipfile.ZipFile(io.BytesIO(r.content))
    with z.open("SMSSpamCollection") as f:
        df = pd.read_csv(f, sep="\t", header=None, names=["label", "message"],
                         encoding="latin-1")
    df.to_csv(OUT, index=False)
    print(f"Tersimpan: {OUT}")
    print(f"Jumlah pesan : {len(df)}")
    print(f"spam         : {(df['label'] == 'spam').sum()}")
    print(f"ham          : {(df['label'] == 'ham').sum()}")
    print(df.head(3).to_string())


if __name__ == "__main__":
    main()
