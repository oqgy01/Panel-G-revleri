import requests
from bs4 import BeautifulSoup
url = "https://docs.google.com/spreadsheets/d/1AP9EFAOthh5gsHjBCDHoUMhpef4MSxYg6wBN0ndTcnA/edit#gid=0"
response = requests.get(url)
html_content = response.content
soup = BeautifulSoup(html_content, "html.parser")
first_cell = soup.find("td", {"class": "s2"}).text.strip()
if first_cell != "Aktif":
    exit()
first_cell = soup.find("td", {"class": "s1"}).text.strip()
print(first_cell)

import pandas as pd
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
from io import BytesIO
import os
import numpy as np
import shutil
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
from selenium.webdriver.chrome.service import Service
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm
from selenium.common.exceptions import TimeoutException, WebDriverException
import xml.etree.ElementTree as ET
import warnings
from colorama import init, Fore, Style
import openpyxl
from openpyxl import load_workbook
import threading
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import tkinter as tk
from tkinter import simpledialog
import chromedriver_autoinstaller
import gc
warnings.filterwarnings("ignore")
pd.options.mode.chained_assignment = None

print("Oturum Açma Başarılı Oldu")
print(" /﹋\ ")
print("(҂`_´)")
print(Fore.RED + "<,︻╦╤─ ҉ - -")
print("/﹋\\")
print("Mustafa ARI")
print(" ")
print(Fore.RED + "Zamanlı Panel Yönetimi")



driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
login_url = "https://task.haydigiy.com/kullanici-giris/?ReturnUrl=%2Fadmin"
driver.get(login_url)
email_input = driver.find_element("id", "EmailOrPhone")
email_input.send_keys("mustafa_kod@haydigiy.com")
password_input = driver.find_element("id", "Password")
password_input.send_keys("123456")
password_input.send_keys(Keys.RETURN)


#region DEDEMAX Tesettür Ürünlerinin Ana Kategorilerini Tesettür Yapma

# Sayfaya Git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

# Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

# Kategori ID'si (TESETTÜR)
category_select.select_by_value("502")

# Seçiniz Kısmının Tikini Kaldırma (Kategori Dahil)
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
second_remove_button = all_remove_buttons[1]
second_remove_button.click()

# Seçiniz Kısmının Tikini Kaldırma (Marka Dahil)
second_remove_button = all_remove_buttons[4]
second_remove_button.click()

# Marka Dahil Alan
category_select = Select(driver.find_element("id", "SearchInManufacturerIds"))

# Marka ID'si (ModelTesettür)
category_select.select_by_value("79")

# Ara
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (TESETTÜR)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("502")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma (Ana Kategori Yap)
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme (Ana Kategori Yap)
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("3")

# Kaydet
save_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary')))
driver.execute_script("arguments[0].click();", save_button)

print(Fore.GREEN + "DEDEMAX Tesettür Ürünlerinin Ana Kategorisi TESETTÜR Yapıldı")

#endregion

#region DEDEMAX Tesettür Ürünlerini Tesettür Kategorisine Ekleme

# Marka Dahil Alan
category_select = Select(driver.find_element("id", "SearchInManufacturerIds"))

# Marka ID'si (ModelTesettür)
category_select.select_by_value("79")

# Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
second_remove_button = all_remove_buttons[3]
second_remove_button.click()

# Ara
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (TESETTÜR)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("502")

# Kaydet
save_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary')))
driver.execute_script("arguments[0].click();", save_button)

print(Fore.GREEN + "DEDEMAX Tesettür Ürünleri TESETTÜR Kategorisine Alındı")

#endregion

#region İç Giyim Kategori Güncellemeleri (Fantezi İç Giyim Ürünlerini Fantezi İç Giyim Kategorisine Alma)

# Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

# Kategori ID'si (İç Giyim > Fantezi > Jartiyer)
category_select.select_by_value("244")

# Kategori ID'si (İç Giyim > Fantezi > Fantezi Gecelik)
category_select.select_by_value("249")

# Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
second_remove_button = all_remove_buttons[1]
second_remove_button.click()

# Ara
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İç Giyim > Fantezi > Fantezi İç Giyim)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("243")

# Kaydet
save_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary')))
driver.execute_script("arguments[0].click();", save_button)

# Kapat
driver.quit()

print(Fore.GREEN + "Fantezi İç Giyim Ürünleri Fantezi İç Giyim Kategorisine Alındı")

#endregion

#region Ürün Listesi İndirme

def download_and_merge_excel(urls, output_file='CalismaAlani.xlsx'):
    dataframes = []
    for i, url in enumerate(urls, start=1):
        response = requests.get(url)
        file_name = f'excel{i}.xlsx'
        with open(file_name, 'wb') as f:
            f.write(response.content)
        dataframes.append(pd.read_excel(file_name))
        os.remove(file_name)

    # Excel dosyalarını birleştir
    merged_df = pd.concat(dataframes, ignore_index=True)
    merged_df.to_excel(output_file, index=False)

if __name__ == "__main__":
    urls = [
        "https://task.haydigiy.com/FaprikaXls/TJ1TB6/1/",
        "https://task.haydigiy.com/FaprikaXls/TJ1TB6/2/",
        "https://task.haydigiy.com/FaprikaXls/TJ1TB6/3/"
    ]
    download_and_merge_excel(urls)

print(Fore.GREEN + "Ürün Listesi İndirilidi")

#endregion

#region Arama Terimlerindeki Tarihleri Tespit Edip Sütunu Güncelleme

# "CalismaAlani" Excel'ini Oku
df = pd.read_excel('CalismaAlani.xlsx')

# Tarihleri çıkarmak için regex deseni
date_pattern = r'(\d{1,2}\.\d{1,2}\.\d{4})'

# "AramaTerimleri" sütunundaki tarihleri temizle ve aynı zamanda kaç güne biter hesapla
def process_search_term(row):
    arama_terimi = str(row['AramaTerimleri'])
    
    # Eğer tarih varsa, temizle ve gün sayısını hesapla
    tarih_match = re.search(date_pattern, arama_terimi)
    if tarih_match:
        tarih_str = tarih_match.group(1)
        tarih = datetime.strptime(tarih_str, '%d.%m.%Y')
        bugun = datetime.today()
        uzaklik = (bugun - tarih).days
        return uzaklik
    else:
        return 0

# "AramaTerimleri" sütununu güncelleme ve hesaplama işlemi
df['AramaTerimleri'] = df.apply(process_search_term, axis=1)

# Güncellenmiş DataFrame'i aynı Excel dosyasının üzerine yaz
df.to_excel('CalismaAlani.xlsx', index=False)

print(Fore.GREEN + "Ürünlerin Kaç Gündür Aktif Oldukları Belirlendi")

#endregion    

#region Liste Fiyatlarını Belirleme

# Ekceli Okuma
df = pd.read_excel('CalismaAlani.xlsx')
def calculate_list_price(row):
    alis_fiyati = row['AlisFiyati']
    kategori = row['Kategori']

    if 0 <= alis_fiyati <= 24.99:
        result = alis_fiyati + 10
    elif 25 <= alis_fiyati <= 39.99:
        result = alis_fiyati + 13
    elif 40 <= alis_fiyati <= 59.99:
        result = alis_fiyati + 17
    elif 60 <= alis_fiyati <= 200.99:
        result = alis_fiyati * 1.30
    elif alis_fiyati >= 201:
        result = alis_fiyati * 1.25
    else:
        result = alis_fiyati

    # KDV
    if isinstance(kategori, str) and any(category in kategori for category in ["Parfüm", "Gözlük", "Saat", "Kolye", "Küpe", "Bileklik", "Bilezik"]):
        result *= 1.20
    else:
        result *= 1.10

    return result

# Liste Fiyatı 2 Sütunu Oluşturma ve Liste Fiyatlarını Yazma
df['ListeFiyati2'] = df.apply(calculate_list_price, axis=1)

# Exceli Kaydet
df.to_excel('CalismaAlani.xlsx', index=False)

print(Fore.GREEN + "Liste Fiyatları Belirlendi")

#endregion

#region Beden Durumu Hesaplama

# "CalismaAlani" Excel'i Oku'
df = pd.read_excel('CalismaAlani.xlsx')

# Aktif Beden Oranı Hesaplama
# "StokAdedi" Sütununda 0'dan Büyük Olan Değerlerin Adedi
df['StokAdedi_GT_0'] = df['StokAdedi'].apply(lambda x: 1 if x > 0 else 0)
stok_adedi_gt_0_adet = df.groupby('UrunAdi')['StokAdedi_GT_0'].sum().reset_index()

# "UrunAdi" Sütunundaki Toplam Yenilenme Adedi
toplam_yenilenme_adedi = df.groupby('UrunAdi').size().reset_index(name='ToplamYenilenmeAdedi')

# Gereksiz Sütunları Silme
df = df.drop(['StokAdedi_GT_0'], axis=1, errors='ignore')

# Oranı Hesapla ve Yeni Sütunu Ekle
df = pd.merge(df, stok_adedi_gt_0_adet, on='UrunAdi', how='left')
df = pd.merge(df, toplam_yenilenme_adedi, on='UrunAdi', how='left')
df['Beden Durumu'] = df['StokAdedi_GT_0'] / df['ToplamYenilenmeAdedi']

# Gereksiz Sütunları Silme
df = df.drop(['StokAdedi_GT_0', 'ToplamYenilenmeAdedi'], axis=1, errors='ignore')

# Oranı 100 ile Çarpma
df['Beden Durumu'] *= 100

# Excel'i kaydet
df.to_excel('CalismaAlani.xlsx', index=False)

print(Fore.GREEN + "Seri Sonu Kategorisi İçin Ürünlerin Aktif Beden Oranları Belirlendi")

#endregion

#region İndirim Yüzdesi Hesaplama

# "CalismaAlani" Excel'i Oku'
df = pd.read_excel('CalismaAlani.xlsx')

# "ListeFiyati2" sütunundaki verilerden "SatisFiyati" sütunundaki verileri çıkarma
df['İndirim Yüzdesi'] = df['ListeFiyati2'] - df['SatisFiyati']

# "ListeFiyati2" sütunundaki veriye bölme
df['İndirim Yüzdesi'] = df['İndirim Yüzdesi'] / df['ListeFiyati2']

# Sonucu 100 ile çarpma
df['İndirim Yüzdesi'] = df['İndirim Yüzdesi'] * 100

# Excel'i Kaydet
df.to_excel('CalismaAlani.xlsx', index=False)

print(Fore.GREEN + "İndirimli Ürünler Kategorisi İçin Ürünlerin İndirim Oranı Belirlendi")

#endregion

#region Büyük Beden Ürünleri Belileme

# "CalismaAlani" Excel'i Oku'
df = pd.read_excel('CalismaAlani.xlsx')

# Belirlenen değerlere eşitse "Kategori" sütununu "Kadın Büyük Beden" olarak belirleme
beden_degerleri = ['Beden:44', 'Beden:46', 'Beden:48', 'Beden:4XL', 'Beden:50', 'Beden:52', 'Beden:54', 'Beden:5XL', 'Beden:XXL', 'Beden:3XL', 'Beden:46-48']
df['Büyük Beden'] = df.apply(lambda row: 'Kadın Büyük Beden' if row['Varyasyon'] in beden_degerleri and row['StokAdedi'] > 0 else '', axis=1)

# "UrunAdi" sütununda yenilenen değerlere sahip ürünlerden en az birinin "Kategori" değeri "Kadın Büyük Beden" ise,
# diğer aynı değerlere de "Kategori" sütununda "Kadın Büyük Beden" yazma
yenilenen_urunler = df[df.duplicated(subset='UrunAdi', keep=False) & (df['Büyük Beden'] == 'Kadın Büyük Beden')]
for urun_adi in yenilenen_urunler['UrunAdi'].unique():
    df.loc[df['UrunAdi'] == urun_adi, 'Büyük Beden'] = 'Kadın Büyük Beden'

# Excel'i Kaydet
df.to_excel('CalismaAlani.xlsx', index=False)

print(Fore.GREEN + "Büyük Beden Ürünler Kategorisi İçin Ürünler Belirlendi")

#endregion

#region Beden Durumu %50'nin Altında Olan Ürünleri Seri Sonu Kategorisine Alma

# "CalismaAlani" Excel'i Oku'
df = pd.read_excel('CalismaAlani.xlsx')

# "Beden Durumu" sütunundaki 50 ve 50'den küçük olan değerleri "Seri Sonu" olarak değiştirme
df['Beden Durumu'] = pd.to_numeric(df['Beden Durumu'], errors='coerce')
df.loc[df['Beden Durumu'] <= 50, 'Beden Durumu'] = 'Seri Sonu'

# "Beden Durumu" sütunu "Seri Sonu" olmayanları silme
df.loc[df['Beden Durumu'] != 'Seri Sonu', 'Beden Durumu'] = ''

# Excel'i Kaydet
df.to_excel('CalismaAlani.xlsx', index=False)

print(Fore.GREEN + "Beden Durumu %50'nin Altında Olan Ürünlere Seri Sonu Kategorisi Verildi")

#endregion

#region İndirim Yüzdesi %5'in Üstünde ve Fark Değeri 5'in Üstünde Olan Ürünleri İndirimli Ürünler ve Fiyata Hamle'ye Alma

# "CalismaAlani" Excel dosyasını oku
df = pd.read_excel('CalismaAlani.xlsx')

# Sütunları sayısal değerlere dönüştür, hata varsa NaN olarak ayarla
df['İndirim Yüzdesi'] = pd.to_numeric(df['İndirim Yüzdesi'], errors='coerce', downcast='float')
df['ListeFiyati2'] = pd.to_numeric(df['ListeFiyati2'], errors='coerce')
df['SatisFiyati'] = pd.to_numeric(df['SatisFiyati'], errors='coerce')

# Koşullu işlem: "İndirim Yüzdesi" Fiyata Hamle 2 değilse ve diğer koşullar sağlanıyorsa işlem yap
df.loc[
    (~df['İndirim Yüzdesi'].astype(str).str.contains('Fiyata Hamle 2')) &  # Fiyata Hamle 2 olmayanlar
    (df['İndirim Yüzdesi'] > 5) &  # İndirim Yüzdesi > 5
    (abs(df['ListeFiyati2'] - df['SatisFiyati']) > 6),  # Liste Fiyatı 2 ile Satış Fiyatı arasındaki fark > 5
    'İndirim Yüzdesi'
] = 'İndirimli Ürünler;Fiyata Hamle'

# Güncellenmiş Excel dosyasını kaydet
df.to_excel('CalismaAlani.xlsx', index=False)

print(Fore.GREEN + "İndirimli Ürünler ve Fiyata Hamle Kategorisine Alınacak Ürünler Belirlendi")

#endregion

#region Fiyata Hamle 2 Kategorisi İçin Ürünleri Belirleme

# "CalismaAlani" Excel'i Oku'
df = pd.read_excel('CalismaAlani.xlsx')

# 'SatisFiyati' ve 'ListeFiyati2' sütunlarındaki verileri sayısal değere dönüştür
df['SatisFiyati'] = pd.to_numeric(df['SatisFiyati'], errors='coerce')
df['ListeFiyati2'] = pd.to_numeric(df['ListeFiyati2'], errors='coerce')

# 'SatisFiyati' ve 'ListeFiyati2' sütunlarından çıkarma işlemi yap
df['Fark'] = df['SatisFiyati'] - df['ListeFiyati2']

# 'Fark' sütunundaki değeri kontrol et ve koşula uygunsa 'İndirim Yüzdesi' sütunundaki veriyi "FiyataHamle2" olarak değiştir
df.loc[df['Fark'] >= 4, 'İndirim Yüzdesi'] = 'FiyataHamle2'

# 'Fark' sütununu ve artık gerekli olmayan 'SatisFiyati' ve 'ListeFiyati2' sütunlarını silelim
df = df.drop(['Fark'], axis=1)

# Excel'i Kaydet
df.to_excel('CalismaAlani.xlsx', index=False, engine='openpyxl')

print(Fore.GREEN + "Fiyata Hamle 2 Kategorisi İçin Bindirimli Ürünler Belirlendi")

#endregion

#region Ekstra Açılan Sütunlarda Gereksiz Olan Hücrelerin İçeriğini Temizleme

# "CalismaAlani" Excel'i Oku'
df = pd.read_excel('CalismaAlani.xlsx')

# 'İndirim Yüzdesi' sütununda "FiyataHamle2" veya "İndirimli Ürünler;Fiyata Hamle" ile eşit olmayan hücrelerin içeriğini temizle
df.loc[(df['İndirim Yüzdesi'] != 'FiyataHamle2') & (df['İndirim Yüzdesi'] != 'İndirimli Ürünler;Fiyata Hamle'), 'İndirim Yüzdesi'] = ''

# Excel'i Kaydet
df.to_excel('CalismaAlani.xlsx', index=False, engine='openpyxl')

print(Fore.GREEN + "Ekstra Açılan Sütunlarda Gereksiz İçerikler Temizlendi")

#endregion

#region Yeni Gelen Sütun Oluşturma ve Arama Terimlerinde 14 ve Daha Küçük Olan Değerleri Yeni Gelen Olarak Ayarlama

# "CalismaAlani" Excel'i Oku'
df = pd.read_excel('CalismaAlani.xlsx')

# 'AramaTerimleri' sütununda 14 ve 14'dan küçük olan değerler için "Yeni Gelen" sütununu oluştur
df['Yeni Gelen'] = df['AramaTerimleri'].apply(lambda x: 'Yeni Gelen' if 0 < x <= 14 else '')

# Excel'i Kaydet
df.to_excel('CalismaAlani.xlsx', index=False, engine='openpyxl')

print(Fore.GREEN + "Son 14 Günde Resmi Yüklenen Ürünler Yeni Gelenler Kategorisine Alındı")

#endregion

#region Yeni Sütun Adında Bir Sütun Oluşturma ve Tüm Kategori Kurgularını Birleştirme

# "CalismaAlani" Excel'i Oku'
df = pd.read_excel('CalismaAlani.xlsx')

# Belirtilen sütunlardaki verileri aralarında ";" olacak şekilde birleştir
df['YeniSutun'] = (
    df['Beden Durumu'].astype(str) + ';' +
    df['İndirim Yüzdesi'].astype(str) + ';' +
    df['Büyük Beden'].astype(str) + ';' +
    df['Yeni Gelen'].astype(str)
)

# Boş olan hücreleri es geç
df['YeniSutun'] = df['YeniSutun'].apply(lambda x: ';'.join(filter(lambda y: pd.notna(y) and y != 'nan', x.split(';'))) if pd.notna(x) else '')

# Excel'i Kaydet
df.to_excel('CalismaAlani.xlsx', index=False, engine='openpyxl')

print(Fore.GREEN + "Tüm Kurgular Tek Bir Sütunda Toplandı")

#endregion

#region Ürünlere Liste Fiyatı Verme

# Excel'i oku
df = pd.read_excel('CalismaAlani.xlsx')

# Kategori sütunundaki her bir hücre için kontrol yap
for index, row in df.iterrows():
    # "Fiyata Hamle" içeriyor mu kontrol et ve aynı zamanda "Kategori" sütunu "Kot Pantolon" içermiyor mu kontrol et
    if isinstance(row['YeniSutun'], str) and "Fiyata Hamle" in row['YeniSutun']:
        if "Kot Pantolon" in str(row['Kategori']):
            # Kategori "Kot Pantolon" içeriyorsa ListeFiyati 0 yap
            df.at[index, 'ListeFiyati'] = 0
        else:
            # ListeFiyati sütunundaki hücreyi güncelle ve 1.8 ile çarp, sonucu tam sayıya dönüştür
            df.at[index, 'ListeFiyati'] = int(row['AlisFiyati'] * 1.8)
    else:
        # Mevcut değeri koru
        df.at[index, 'ListeFiyati'] = row['ListeFiyati']

# Değişiklikleri kaydet
df.to_excel('CalismaAlani.xlsx', index=False)

print(Fore.GREEN + "Fiyata Hamle Kategorisindeki Ürünlere Liste Fiyatı Verildi")

#endregion

#region Arama Terimleri ve Kategori Sütunlarının İçeriğini Temizleme

# "CalismaAlani" Excel'i Oku'
df = pd.read_excel('CalismaAlani.xlsx')

# 'Kategori' ve 'AramaTerimleri' sütunlarındaki içeriği temizle
df['Kategori'] = ''
df['AramaTerimleri'] = ''

# Excel'i Kaydet
df.to_excel('CalismaAlani.xlsx', index=False, engine='openpyxl')

print(Fore.GREEN + "Arama Terimleri ve Kategori Sütunlarının İçeriği Temizlendi")

#endregion

#region Yeni Sütun Verilerini Kategori Sütununa Taşıma

# "CalismaAlani" Excel'i Oku'
df = pd.read_excel('CalismaAlani.xlsx')

# 'YeniSutun' sütunundaki verileri 'Kategori' sütununa taşı
df['Kategori'] = df['YeniSutun']

# 'YeniSutun' sütununu ve artık gerekli olmayan diğer sütunları sil
df = df.drop(['YeniSutun'], axis=1)

# Excel'i Kaydet
df.to_excel('CalismaAlani.xlsx', index=False, engine='openpyxl')

print(Fore.GREEN + "Yeni Sütun Sütunu Kategori Sütununa Taşındı")

#endregion

#region Ekstra Açılan Tüm Sütunları Silme

# "CalismaAlani" Excel'i Oku'
df = pd.read_excel('CalismaAlani.xlsx')

# Belirtilen sütunları sil
df = df.drop(['ListeFiyati2', 'Beden Durumu', 'İndirim Yüzdesi', 'Büyük Beden', 'Yeni Gelen'], axis=1)

# Excel'i Kaydet
df.to_excel('CalismaAlani.xlsx', index=False, engine='openpyxl')

print(Fore.GREEN + "Ekstra Açılan Tüm Sütunlar Silindi")

#endregion

#region Giriş Yapma ve Fiyata Hamle Kategorisini Boşaltma

# ChromeOptions oluştur
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

login_url = "https://task.haydigiy.com/kullanici-giris/?ReturnUrl=%2Fadmin"
driver.get(login_url)

email_input = driver.find_element("id", "EmailOrPhone")
email_input.send_keys("mustafa_kod@haydigiy.com")

password_input = driver.find_element("id", "Password")
password_input.send_keys("123456")
password_input.send_keys(Keys.RETURN)

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("187")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Fiyata Hamle)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("187")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme (Kategoriden Çıkar)
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Fiyata Hamle Kategorisi Boşaltıldı")
#endregion

#region İndirimli Ürünler Kategorisi Boşaltma ve İndirimli Ürünler Etiketini Kaldırma

# Giriş yaptıktan sonra belirtilen sayfaya git (İndirimli Ürünler Etiketini Kaldırma)
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Etiket Dahil Alan
category_select = Select(driver.find_element("id", "SearchInProductTagIds"))

#Etiket ID'si (İndirimli Ürünler)
category_select.select_by_value("144")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[7]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Etiket Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ProductTag_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Etiket Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "ProductTagId")

# Etiket Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("144")

# Etiket Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "ProductTagTransactionId")

# Etiket Güncelleme Alanında Yapılacak İşlemi Seçme (Etiketi Kaldırma)
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")





# Giriş yaptıktan sonra belirtilen sayfaya git (İndirimli Ürünler Kategorisi Boşaltma)
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (İndirimli Ürünler)
category_select.select_by_value("374")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("374")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "İndirimli Ürünler Kategorisi Boşaltıldı ve Etiketleri Kaldırıldı")
#endregion

#region Seri Sonu Kategorisi Boşaltma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Seri Sonu)
category_select.select_by_value("74")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("74")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Seri Sonu Kategorisi Boşaltıldı")
#endregion

#region Yeni Gelenler Kategorisi Boşaltma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Yeni Gelenler)
category_select.select_by_value("347")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Yeni Gelenler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("347")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Yeni Gelenler Kategorisi Boşaltıldı")
#endregion

#region Excelle Ürün Yükleme Alanı
desired_url = "https://task.haydigiy.com/admin/importproductxls/edit/24"
driver.get(desired_url)

# Yükle Butonunu Bul
file_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="qqfile"]')))

# CalismaAlani Excel dosyasının mevcut çalışma dizininde olduğunu varsay
file_path = os.path.join(os.getcwd(), "CalismaAlani.xlsx")

# Dosyayı seç
file_input.send_keys(file_path)

# "İşlemler" düğmesine tıkla
operations_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'btn-success')))
operations_button.click()

# Dosya yükleme işlemi bittikten sonra çalıştır butonuna tıkla
execute_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'import-product-xls-execute-confirm')))
execute_button.click()

# 10 saniye bekle
time.sleep(10)

def wait_for_element_and_click(driver, by, value, timeout=10):
    try:
        element = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
        element.click()
        return True
    except (TimeoutException, WebDriverException) as e:
        print(f"Hata: {e}")
        return False

def wait_for_page_load(driver):
    while True:
        if driver.title:  # Tarayıcı başlığı varsa, sayfa yüklenmiş demektir
            break
        time.sleep(2)

# "Evet" butonunu tıkla
if wait_for_element_and_click(driver, By.ID, 'import-product-xls-execute'):
    # Yüklenmeyi Bekle
    wait_for_page_load(driver)

# Dinamik Bekleme İşlevi
def wait_for_success_or_timeout(driver, timeout=360):
    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            # Eğer <span data-notify="message">faprika.import.execute.success</span> bulunursa devam et
            if driver.find_elements(By.XPATH, '//span[@data-notify="message" and text()="faprika.import.execute.success"]'):
                pass
                return
        except Exception as e:
            pass
        time.sleep(1)  # 1 saniyelik aralıklarla kontrol et
    print(Fore.YELLOW + "Başarı mesajı bulunamadı, bekleme süresi doldu.")

# Dinamik bekleme çağrısı
wait_for_success_or_timeout(driver)

print(Fore.GREEN + "Excelle Ürün Yükleme Yapıldı")

#endregion

#region Çocuk Ürünlerini Yeni Gelenlerden Çıkarma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Çocuk & Bebek)
category_select.select_by_value("440")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Yeni Gelenler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("347")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Çocuk Ürünleri Yeni Gelenler Kategorisinden Çıkarıldı")
#endregion

#region Çocuk Ürünlerini İndirimli Ürünlerden Çıkarma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Çocuk & Bebek)
category_select.select_by_value("440")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("374")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Çocuk Ürünleri İndirimli Ürünler Kategorisinden Çıkarıldı")
#endregion

#region Çocuk Ürünlerini Seri Sonundan Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Çocuk & Bebek)
category_select.select_by_value("440")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("74")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Çocuk Ürünleri Seri Sonu Kategorisinden Çıkarıldı")
#endregion

#region Çocuk Ürünlerini Tekrar Stoktadan Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Çocuk & Bebek)
category_select.select_by_value("440")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("297")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Çocuk Ürünleri Tekrar Stokta Kategorisinden Çıkarıldı")
#endregion

#region Çocuk Ürünlerini Büyük Bedenden Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Çocuk & Bebek)
category_select.select_by_value("440")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("128")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Çocuk Ürünleri Büyük Beden Kategorisinden Çıkarıldı")
#endregion

#region DEDEMAX Tesettür Ürünlerini Sadece Tesettür Kategorisinde Tutma (ALT GİYİM)

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("502")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass


#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInManufacturerIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("39")

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("79")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[3]
    second_remove_button.click()
else:
    pass


#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Fiyata Hamle)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("13")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme (Kategoriden Çıkar)
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

#endregion

#region DEDEMAX Tesettür Ürünlerini Sadece Tesettür Kategorisinde Tutma (ÜST GİYİM)

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("502")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass


#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInManufacturerIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("39")

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("79")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[3]
    second_remove_button.click()
else:
    pass


#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Fiyata Hamle)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("79")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme (Kategoriden Çıkar)
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

#endregion

#region DEDEMAX Tesettür Ürünlerini Sadece Tesettür Kategorisinde Tutma (EŞOFMAN - PİJAMA)

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("502")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass


#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInManufacturerIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("39")

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("79")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[3]
    second_remove_button.click()
else:
    pass


#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Fiyata Hamle)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("34")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme (Kategoriden Çıkar)
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

#endregion

#region DEDEMAX Tesettür Ürünlerini Sadece Tesettür Kategorisinde Tutma (ELBİSE - TULUM)

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("502")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass


#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInManufacturerIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("39")

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("79")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[3]
    second_remove_button.click()
else:
    pass


#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Fiyata Hamle)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("27")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme (Kategoriden Çıkar)
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

#endregion

#region DEDEMAX Tesettür Ürünlerini Yeni Gelenlerden Çıkarma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("502")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass


#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInManufacturerIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("39")

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("79")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[3]
    second_remove_button.click()
else:
    pass


#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Fiyata Hamle)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("347")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme (Kategoriden Çıkar)
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

#endregion

#region ÜST GİYİM Kategorisindeki Ürünleri ÜST GİYİM Kategorisine Alma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("79")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("79")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("0")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "ÜST GİYİM Kategorisindeki Ürünler ÜST GİYİM kategorisine alındı")

#endregion

#region DIŞ GİYİM Kategorisindeki Ürünleri DIŞ GİYİM Kategorisine Alma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("532")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("532")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("0")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "DIŞ GİYİM Kategorisindeki Ürünler DIŞ GİYİM kategorisine alındı")

#endregion

#region EŞOFMAN - PİJAMA Kategorisindeki Ürünleri EŞOFMAN - PİJAMA Kategorisine Alma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("34")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("34")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("0")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "EŞOFMAN PİJAMA Kategorisindeki Ürünler EŞOFMAN PİJAMA kategorisine alındı")

#endregion

#region ALT GİYİM Kategorisindeki Ürünleri ALT GİYİM Kategorisine Alma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("13")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("13")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("0")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "ALT GİYİM Kategorisindeki Ürünler ALT GİYİM kategorisine alındı")

#endregion

#region ELBİSE - TULUM Kategorisindeki Ürünleri ELBİSE - TULUM Kategorisine Alma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("27")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("27")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("0")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "ELBİSE - TULUM Kategorisindeki Ürünler ELBİSE - TULUM kategorisine alındı")

#endregion

#region AYAKKABI Kategorisindeki Ürünleri AYAKKABI Kategorisine Alma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("18")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("18")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("0")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "AYAKKABI Kategorisindeki Ürünler AYAKKABI kategorisine alındı")

#endregion

#region AKSESUAR Kategorisindeki Ürünleri AKSESUAR Kategorisine Alma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("4")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("4")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("0")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "AKSESUAR Kategorisindeki Ürünler AKSESUAR kategorisine alındı")

#endregion

#region ÇOCUK - BEBEK Kategorisindeki Ürünleri ERKEK Kategorisine Alma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("440")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("440")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("0")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "ÇOCUK - BEBEK Kategorisindeki Ürünler ÇOCUK - BEBEK kategorisine alındı")

#endregion

#region ERKEK Kategorisindeki Ürünleri ERKEK Kategorisine Alma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("469")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("469")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("0")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "ERKEK Kategorisindeki Ürünler ERKEK kategorisine alındı")

#endregion

#region DIŞ GİYİM Kategorisindeki Ürünleri DIŞ GİYİM Kategorisine Alma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("541")
category_select.select_by_value("165")
category_select.select_by_value("85")


#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

time.sleep(3)

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("532")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("0")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "DIŞ GİYİM Kategorisindeki Ürünler DIŞ GİYİM kategorisine alındı")

#endregion

#region Aksesuarları Seri Sonundan Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("4")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Seri Sonu)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("74")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Aksesuarlar Seri Sonu Kategorisinden Çıkarıldı")
#endregion

#region Aksesuarları Yeni Gelenlerden Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("4")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Yeni Gelenler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("347")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Aksesuarlar Yeni Gelenler Kategorisinden Çıkarıldı")
#endregion

#region Aksesuarları İndirimli Ürünlerden Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("4")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("374")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Aksesuarlar İndirimli Ürünler Kategorisinden Çıkarıldı")
#endregion

#region Aksesuarları Tekrar Stoktadan Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("4")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("297")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Aksesuarlar Tekrar Stokta Kategorisinden Çıkarıldı")
#endregion

#region Aksesuarları Büyük Bedenden Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("4")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("128")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Aksesuarlar Büyük Beden Kategorisinden Çıkarıldı")
#endregion

#region Erkek Ürünlerini İndirimli Ürünlerden Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("469")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("374")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Erkek Ürünleri İndirimli Ürünler Kategorisinden Çıkarıldı")

#endregion

#region Erkek Ürünlerini Yeni Gelenlerden Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("469")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("347")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Erkek Ürünleri Yeni Gelenler Kategorisinden Çıkarıldı")
#endregion

#region Erkek Ürünlerini Tekrar Stoktadan Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("469")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("297")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Erkek Ürünleri Tekrar Stokta Kategorisinden Çıkarıldı")
#endregion

#region Erkek Ürünlerini Seri Sonundan Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("469")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("74")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Erkek Ürünleri Seri Sonu Kategorisinden Çıkarıldı")
#endregion

#region Erkek Ürünlerini Büyük Bedenden Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("469")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("128")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Erkek Ürünleri Büyük Beden Kategorisinden Çıkarıldı")
#endregion

#region İç Giyim Ürünlerini İndirimli Ürünlerden Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("172")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("374")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Erkek Ürünleri İndirimli Ürünler Kategorisinden Çıkarıldı")
#endregion

#region İç Giyim Ürünlerini Yeni Gelenlerden Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("172")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("347")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Erkek Ürünleri Yeni Gelenler Kategorisinden Çıkarıldı")
#endregion

#region İç Giyim Ürünlerini Tekrar Stoktadan Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("172")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("297")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Erkek Ürünleri Tekrar Stokta Kategorisinden Çıkarıldı")
#endregion

#region İç Giyim Ürünlerini Seri Sonundan Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("172")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("74")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Erkek Ürünleri Seri Sonu Kategorisinden Çıkarıldı")
#endregion

#region İç Giyim Ürünlerini Büyük Bedenden Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("172")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("128")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Erkek Ürünleri Büyük Beden Kategorisinden Çıkarıldı")
#endregion

#region Penti Ürünlerini Sütyen Takımından Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("483")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("255")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Erkek Ürünleri Bunları da Görmek İsteyebilirsiniz Kategorisinden Çıkarıldı")
#endregion   

#region Fiyata Hamle Ürünlere Etiket Verme

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

# Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

# Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("187")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

# Ara
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Etiket Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ProductTag_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Etiket Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "ProductTagId")

# Etiket Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürün)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("144")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Fiyata Hamle Kategorisinde Olan Ürünlere İndirimli Ürünler Etiketi Verildi")
#endregion

#region Yeni Gelenler Markasını Boşaltma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)



#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInManufacturerIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("144")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[3]
    second_remove_button.click()
else:
    pass


#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Manufacturer_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "ManufacturerId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Fiyata Hamle)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("144")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "ManufacturerTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme (Kategoriden Çıkar)
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

#endregion

#region Yeni Gelenler Kategorisindeki Ürünlere Yeni Gelenler Markası Verme

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)



#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Fiyata Hamle)
category_select.select_by_value("347")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass


#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Manufacturer_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "ManufacturerId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (Fiyata Hamle)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("144")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "ManufacturerTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme (Kategoriden Çıkar)
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("0")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

#endregion

#region Yeni Gelenler Kategorisinden Fiyata Hamleyi Hariç Tutma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("187")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("347")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Fiyata Hamle Ürünler Yeni Gelenlerden Çıkarıldı")
#endregion

#region Büyük Beden Etiketi Olan Ürünlerin Etiketini Kaldırma

# Giriş yaptıktan sonra belirtilen sayfaya git (İndirimli Ürünler Etiketini Kaldırma)
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Etiket Dahil Alan
category_select = Select(driver.find_element("id", "SearchInProductTagIds"))

#Etiket ID'si (İndirimli Ürünler)
category_select.select_by_value("113")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[7]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Etiket Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ProductTag_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Etiket Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "ProductTagId")

# Etiket Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("113")

# Etiket Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "ProductTagTransactionId")

# Etiket Güncelleme Alanında Yapılacak İşlemi Seçme (Etiketi Kaldırma)
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("1")

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Büyük Beden Etiketi Olan Ürünlerin Etiketleri Kaldırıldı")

#endregion

#region Büyük Beden Ürünlere Etiket Verme

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Büyük Beden)
category_select.select_by_value("128")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Etiket Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ProductTag_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Etiket Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "ProductTagId")

# Etiket Güncelleme Alanında Kategori ID'si Seçme (Büyük Beden)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("113")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Büyük Beden Kategorisinde Olan Ürünlere Büyük Beden Etiketi Verildi")
#endregion

#region Kot Pantolon Kategorisindeki Ürünlere Etiket Verme

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Büyük Beden)
category_select.select_by_value("11")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Etiket Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ProductTag_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Etiket Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "ProductTagId")

# Etiket Güncelleme Alanında Kategori ID'si Seçme (Büyük Beden)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("225")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Kot Pantolon Kategorisinde Olan Ürünlere Kot Pantolon Etiketi Verildi")
#endregion

#region Seri Sonundaki Ürünleri Olduğu Gibi Dev İndirimler Kategorisine Alma

# Giriş yaptıktan sonra belirtilen sayfaya git
desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
driver.get(desired_page_url)

#Kategori Dahil Alan
category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

#Kategori ID'si (Aksesuar)
category_select.select_by_value("74")

#Seçiniz Kısmının Tikini Kaldırma
all_remove_buttons = driver.find_elements(By.XPATH, "//span[@class='select2-selection__choice__remove']")
if len(all_remove_buttons) > 1:
    second_remove_button = all_remove_buttons[1]
    second_remove_button.click()
else:
    pass

#Ara Butonuna Tıklama
search_button = driver.find_element(By.ID, "search-products")
search_button.click()

# Sayfanın en sonuna git
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# Kategori Güncelleme Tikine Tıklama
checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "Category_Update")))
driver.execute_script("arguments[0].click();", checkbox)

# Kategori Güncelleme Alanını Bulma
category_id_select = driver.find_element(By.ID, "CategoryId")

# Kategori Güncelleme Alanında Kategori ID'si Seçme (İndirimli Ürünler)
category_id_select = Select(category_id_select)
category_id_select.select_by_value("374")

# Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
category_transaction_select = driver.find_element(By.ID, "CategoryTransactionId")

# Kategori Güncelleme Alanında Yapılacak İşlemi Seçme
category_transaction_select = Select(category_transaction_select)
category_transaction_select.select_by_value("0")

time.sleep(3)

try:
    #Kaydet Butonunu Bulma
    save_button = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, 'btn-primary'))
    )

    #Kaydet Butonununa Tıklama
    driver.execute_script("arguments[0].click();", save_button)
except Exception as e:
    print(f"Hata: {e}")

print(Fore.GREEN + "Seri Sonu Kategorisindeki Ürünler Dev İndirimler Kategorisine Alındı")

#endregion

#region Faturasız Siparişleri Faturaya Gönderme

# Sipariş Aktarım Sayfasına Gitme ve Tarih Aralığını 1 Haftaya Göre Ayarlama
desired_page_url = "https://task.haydigiy.com/admin/exportorder/edit/129/"
driver.get(desired_page_url)

# 7 gün önceki tarihi al tarihini al
yesterday = datetime.now() - timedelta(days=7)
formatted_date = yesterday.strftime("%d.%m.%Y")

# Input alanını bulma ve tarih değerini giriş yapma
end_date_input = driver.find_element("id", "StartDate")
end_date_input.clear()  # Eğer mevcut bir değer varsa temizleyin
end_date_input.send_keys(formatted_date)


# Dünün tarihini al
yesterday = datetime.now() - timedelta(days=1)
formatted_date = yesterday.strftime("%d.%m.%Y")

# Input alanını bulma ve tarih değerini giriş yapma
end_date_input = driver.find_element("id", "EndDate")
end_date_input.clear()  # Eğer mevcut bir değer varsa temizleyin
end_date_input.send_keys(formatted_date)

# Buttonu bulma ve tıklama
save_button = driver.find_element("css selector", 'button.btn.btn-primary[name="save"]')
save_button.click()

driver.quit()




#Seleniumsuz Giriş Yapma

# Kullanıcı adı ve şifre
username = "mustafa_kod@haydigiy.com"
password = "123456"

# Oturum Aç
login_url = "https://task.haydigiy.com/kullanici-giris/?ReturnUrl=%2Fadmin"

# İstek başlıkları
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
    "Referer": "https://task.haydigiy.com/",
}

# Oturum açma sayfasına GET isteği gönderme
session = requests.Session()
response = session.get(login_url, headers=headers)
soup = BeautifulSoup(response.text, "html.parser")

# __RequestVerificationToken değerini alma
token = soup.find("input", {"name": "__RequestVerificationToken"}).get("value")

# POST isteği için istek verilerini ayarlama
login_data = {
    "EmailOrPhone": username,
    "Password": password,
    "__RequestVerificationToken": token,
}

# Oturum açma isteği gönderme
response = session.post(login_url, data=login_data, headers=headers)

# Her Bir ID'nin Başına Link Eklyerek Sırayla İstek Gönderme
def get_order_ids_from_xml(xml_url):
    try:
        response = requests.get(xml_url)
        if response.status_code != 200:
            print(f"GET request failed with status code: {response.status_code}")
            return

        root = ET.fromstring(response.content)
        order_ids = [order.get('Id') for order in root.findall('Order')]
        return order_ids

    except requests.exceptions.RequestException as e:
        print(f"Error occurred during the request: {e}")
        return

def send_order_ids_to_shipment_integration(order_ids):
    base_url = "https://task.haydigiy.com//admin/order/sendinvoicetointegration/?orderId="
    for order_id in order_ids:
        url = base_url + order_id
        response = session.get(url, headers=headers)
        if response.status_code == 200:
            pass
        else:
            print(f"Hata: {order_id} with status code: {response.status_code}")

if __name__ == "__main__":
    xml_url = "https://task.haydigiy.com/FaprikaOrderXml/VWONC5/1/"
    order_ids = get_order_ids_from_xml(xml_url)
    if order_ids:
        send_order_ids_to_shipment_integration(order_ids)

print(Fore.GREEN + "Faturasız Siparişler Entegrasyona Yeniden Gönderildi")

#endregion

gc.collect()
os.remove('CalismaAlani.xlsx')


