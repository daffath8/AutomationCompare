import os
import json
import pandas as pd
import unicodedata

# Lokasi file
json_folder = "/Users/daffafatahillah/Documents/projectTes/Data/json_tadarus_new"
csv_path = "/Users/daffafatahillah/Documents/projectTes/Data/database.csv"

# Load CSV
df = pd.read_csv(csv_path)

results = []

def normalize(text):
    if not isinstance(text, str):
        return ""
    return unicodedata.normalize('NFKC', text.strip())

# Hitung jumlah baris CSV
total_csv_rows = len(df)

row_index = 0

# Loop semua file JSON
for i in range(1, 115):
    filename = f"{i}.json"
    json_path = os.path.join(json_folder, filename)

    if not os.path.exists(json_path):
        print(f"⚠️ File tidak ditemukan: {filename}")
        continue

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Ambil semua ayat di dalam file JSON
    ayat_list = data.get("ayat", [])

    for ayat in ayat_list:
        teks_arab = normalize(ayat.get("teksArab", ""))

        # Ambil teks dari CSV
        if row_index >= total_csv_rows:
            print(f"❗ Tidak cukup baris di CSV untuk ayat ke-{row_index+1}")
            break

        teks_usmani = normalize(df.iloc[row_index]["teks_msi_usmani"])
        status = "MATCH" if teks_arab == teks_usmani else "MISMATCH"

        # Tambahkan ke hasil
        results.append({
            "teksArab": teks_arab,
            "teks_msi_usmani": teks_usmani,
            "match": status
        })

        # Debug mismatch
        if status == "MISMATCH":
            print(f"\n❌ MISMATCH di file: {filename}, ayat ke-{ayat.get('nomorAyat')}")
            print(f"JSON : {teks_arab}")
            print(f"CSV  : {teks_usmani}")
            max_len = max(len(teks_arab), len(teks_usmani))
            for j in range(max_len):
                c_json = teks_arab[j] if j < len(teks_arab) else "[kosong]"
                c_csv = teks_usmani[j] if j < len(teks_usmani) else "[kosong]"
                if c_json != c_csv:
                    print(f"  Posisi {j+1}: JSON='{c_json}' | CSV='{c_csv}'")

        row_index += 1

# Simpan hasil ke Excel
hasil_df = pd.DataFrame(results)
hasil_df.to_excel("hasil_perbandingan_new.xlsx", index=False)
print("\n✅ Hasil disimpan ke 'hasil_perbandingan_new.xlsx'")
