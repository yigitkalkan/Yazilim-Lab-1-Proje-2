import pandas as pd

# Excel dosyalarını okuma
degerlendirme = pd.read_excel('degerlendirme.xlsx')
notlar = pd.read_excel('notlar.xlsx')

degerlendirme.reset_index(drop=True, inplace=True)
notlar.reset_index(drop=True, inplace=True)

degerlendirme.columns = degerlendirme.columns.str.strip()
notlar.columns = notlar.columns.str.strip()

columns_to_process = ['Öd1', 'Öd2', 'Quiz', 'Vize', 'Fin']

# Ağırlıklı değerlendirme tablosunu oluştur
results = {col: [] for col in columns_to_process}
for col in columns_to_process:
    base_value = degerlendirme.loc[0, col]
    for i in range(1, len(degerlendirme)):
        result = (base_value * degerlendirme.loc[i, col]) / 100
        results[col].append(result)

data = {'Ders Çıktı': list(range(1, len(results[columns_to_process[0]]) + 1))}
data.update(results)

tablo3_agirlikli = pd.DataFrame(data)
tablo3_agirlikli['Toplam'] = tablo3_agirlikli[columns_to_process].sum(axis=1)
tablo3_agirlikli.to_excel('Tablo3_Ağırlıklı_Değerlendirme.xlsx', index=False)
print("Tablo3_Ağırlıklı_Değerlendirme.xlsx oluşturuldu.")

# Her öğrenci için ayrı bir tablo oluşturma
for index, row in notlar.iterrows():
    student_name = row['Öğrenci']
    results = []
    for i in range(len(tablo3_agirlikli)):
        ders_cikti = i + 1
        toplam = 0
        max_puan = 0

        row_data = {'Ders Çıktı': ders_cikti}
        for col in columns_to_process:
            ogrenci_notu = row[col]
            agirlikli_deger = tablo3_agirlikli.loc[i, col]
            hesaplanan_deger = ogrenci_notu * agirlikli_deger
            row_data[col] = hesaplanan_deger
            toplam += hesaplanan_deger
            max_puan += agirlikli_deger * 100

        row_data['TOPLAM'] = toplam
        row_data['MAX'] = max_puan
        row_data['% Başarı'] = round((toplam / max_puan) * 100, 1) if max_puan > 0 else 0
        results.append(row_data)

    student_df = pd.DataFrame(results)
    filename = f"{student_name}_Tablosu.xlsx"
    student_df.to_excel(filename, index=False)
    print(f"{student_name} için tablo oluşturuldu: {filename}")

weights = degerlendirme.iloc[0]
notlar['ORT'] = (
    notlar['Öd1'] * weights['Öd1'] +
    notlar['Öd2'] * weights['Öd2'] +
    notlar['Quiz'] * weights['Quiz'] +
    notlar['Vize'] * weights['Vize'] +
    notlar['Fin'] * weights['Fin']
) / 100
notlar.to_excel('notlar.xlsx', index=False)
print("Ağırlıklı ortalamalar hesaplandı ve notlar dosyasına yazıldı.")

# Prg Çıktı tablolarını oluşturma
derscikti = pd.read_excel('derscikti.xlsx')
tablo3_agirlikli = pd.read_excel('Tablo3_Ağırlıklı_Değerlendirme.xlsx')
notlar = pd.read_excel('notlar.xlsx')

derscikti.reset_index(drop=True, inplace=True)
tablo3_agirlikli.reset_index(drop=True, inplace=True)
notlar.reset_index(drop=True, inplace=True)

derscikti.columns = derscikti.columns.str.strip()
tablo3_agirlikli.columns = tablo3_agirlikli.columns.str.strip()
notlar.columns = notlar.columns.str.strip()

for index, row in notlar.iterrows():
    student_name = row['Öğrenci']
    ogrenci_tablosu = pd.read_excel(f"{student_name}_Tablosu.xlsx")
    basari_degerleri = ogrenci_tablosu['% Başarı']

    prg_cikti_results = []
    for prg_index in range(len(derscikti)):
        prg_cikti_row = {'Prg Çıktı': prg_index + 1}
        toplam_deger = 0
        iliski_toplami = 0

        for ders_cikti_index in range(len(basari_degerleri)):
            basari = basari_degerleri.iloc[ders_cikti_index]
            ders_cikti_iliski = derscikti.iloc[prg_index, ders_cikti_index + 1]

            hesaplanan_deger = basari * ders_cikti_iliski if not pd.isna(ders_cikti_iliski) else 0
            prg_cikti_row[f"Başarı {ders_cikti_index + 1}"] = round(hesaplanan_deger, 2)

            toplam_deger += hesaplanan_deger
            iliski_toplami += ders_cikti_iliski if not pd.isna(ders_cikti_iliski) else 0

        basari_orani = (toplam_deger / iliski_toplami) if iliski_toplami > 0 else 0
        prg_cikti_row['Başarı Oranı'] = round(basari_orani, 2)
        prg_cikti_results.append(prg_cikti_row)

    prg_cikti_df = pd.DataFrame(prg_cikti_results)

    expected_columns = ['Prg Çıktı'] + [f"Başarı {i + 1}" for i in range(len(basari_degerleri))] + ['Başarı Oranı']
    prg_cikti_df = prg_cikti_df.reindex(columns=expected_columns, fill_value=0)

    filename = f"{student_name}_Prg_Cikti_Tablosu.xlsx"
    prg_cikti_df.to_excel(filename, index=False)
    print(f"{student_name} için Prg Çıktı Tablosu oluşturuldu: {filename}")

print("Tüm öğrenciler için Prg Çıktı tabloları oluşturuldu.")
