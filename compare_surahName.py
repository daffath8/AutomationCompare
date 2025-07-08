import json
import pandas as pd
import os
import unicodedata

# Paths
json_path = "/Users/daffafatahillah/Documents/projectTes/Data/data-sura-updated-search.json"
csv_path = "//Users/daffafatahillah/Documents/projectTes/Data/surahName.csv"
output_dir = "/Users/daffafatahillah/Documents/projectTes/Result"
output_file = os.path.join(output_dir, "surahName_comparison.xlsx")

# Pastikan folder output ada
os.makedirs(output_dir, exist_ok=True)

# Normalisasi teks
def normalize(text):
    if not isinstance(text, str):
        return ""
    return unicodedata.normalize("NFKC", text.strip())

# Load data
with open(json_path, "r", encoding="utf-8") as f:
    json_data = json.load(f)

csv_data = pd.read_csv(csv_path)

# Siapkan hasil per baris
results = []

for i, json_row in enumerate(json_data):
    try:
        csv_row = csv_data.iloc[i]
    except IndexError:
        print(f"❗ Data CSV kurang baris untuk index {i}")
        break

    # Ambil dan normalisasi field
    j_namaLatin = normalize(json_row.get("namaLatin", ""))
    j_tempatTurun = normalize(json_row.get("tempatTurun", ""))
    j_arti = normalize(json_row.get("arti", ""))

    c_namaLatin = normalize(csv_row.get("nama_latin", ""))
    c_tempatTurun = normalize(csv_row.get("kategori", ""))
    c_arti = normalize(csv_row.get("terjemahan", ""))

    # Cek kesesuaian
    match = (
        j_namaLatin == c_namaLatin and
        j_tempatTurun == c_tempatTurun and
        j_arti == c_arti
    )

    status = "MATCH" if match else "NOT_MATCH"

    results.append({
        "namaLatin (JSON)": j_namaLatin,
        "nama_latin (CSV)": c_namaLatin,
        "tempatTurun (JSON)": j_tempatTurun,
        "kategori (CSV)": c_tempatTurun,
        "arti (JSON)": j_arti,
        "terjemahan (CSV)": c_arti,
        "validate": status
    })

# Simpan ke Excel
df_result = pd.DataFrame(results)
df_result.to_excel(output_file, index=False)

print(f"\n✅ Hasil perbandingan disimpan di:\n{output_file}")
