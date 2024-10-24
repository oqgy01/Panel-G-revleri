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
print(Fore.RED + "Panel Yönetimi")
print(Fore.RED + "İndirim Raporu")

indirim_raporu = input("İndirim Raporu (E/H): ").strip().upper()
panel_gorevleri = input("Panel Görevleri (E/H): ").strip().upper()

if indirim_raporu == "E":

    # Tecrübeleri Yarıştırmak Neye Yarar Aklını Karıştırmaktan Başka
    # İnsanları Savaştırmak Kolay Peki Sonra Barıştırmak Çok Saçma
    # Kapalı Kapılar Penecereler Hep Yine Düştüm Bak Neredeyim Acep
    # Bir Bulsam Kendimi Eskisi Gibi Olamam Asla
    # Biliyorum Fazla Bu İstediğim

    #region Ürün Listesi İndirme

    def download_and_merge_excel(url1, url2, url3):
        response1 = requests.get(url1)
        with open('excel1.xlsx', 'wb') as f1:
            f1.write(response1.content)
        response2 = requests.get(url2)
        with open('excel2.xlsx', 'wb') as f2:
            f2.write(response2.content)
        response3 = requests.get(url3)
        with open('excel3.xlsx', 'wb') as f3:
            f3.write(response3.content)
        
        df1 = pd.read_excel('excel1.xlsx')
        df2 = pd.read_excel('excel2.xlsx')
        df3 = pd.read_excel('excel3.xlsx')
        
        merged_df = pd.concat([df1, df2, df3], ignore_index=True)
        merged_df.to_excel('UrunListesi.xlsx', index=False)
        
        os.remove('excel1.xlsx')
        os.remove('excel2.xlsx')
        os.remove('excel3.xlsx')

    if __name__ == "__main__":
        url1 = "https://task.haydigiy.com/FaprikaXls/D7TU34/1/"
        url2 = "https://task.haydigiy.com/FaprikaXls/D7TU34/2/"
        url3 = "https://task.haydigiy.com/FaprikaXls/D7TU34/3/"
        download_and_merge_excel(url1, url2, url3)

    print(Fore.YELLOW + "Ürün Listesi İndirildi")

    #endregion

    #region Belli Sütunlar Hariç Diğerlerini Silme

    # Exceli Okuma
    df_merged = pd.read_excel('UrunListesi.xlsx')

    # Sütunları Belirle
    columns_to_keep = ["StokKodu", "UrunAdi", "StokAdedi", "AlisFiyati", "SatisFiyati", "AramaTerimleri", "MorhipoKodu", "VaryasyonMorhipoKodu", "HepsiBuradaKodu", "VaryasyonHepsiBuradaKodu", "VaryasyonN11Kodu", "Kategori"]

    # Sil
    df_merged = df_merged[columns_to_keep]

    # Exceli Kaydet
    df_merged.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "Calışma Alanı Oluşturuldu ve Gereksiz Sütunlar Silindi")

    #endregion

    #region Arama Terimlerindeki Tarihleri Tespit Edip Çıkarma

    # Exceli Oku
    df_calisma_alani = pd.read_excel('CalismaAlani.xlsx')

    # Tarihleri çıkar
    date_pattern = r'(\d{1,2}\.\d{1,2}\.\d{4})'

    # "AramaTerimleri" sütunundaki tarihleri temizle
    df_calisma_alani['AramaTerimleri'] = df_calisma_alani['AramaTerimleri'].apply(lambda x: re.search(date_pattern, str(x)).group(1) if re.search(date_pattern, str(x)) else None)

    # Exceli Kaydet
    with pd.ExcelWriter('CalismaAlani.xlsx', engine='xlsxwriter') as writer:
        df_calisma_alani.to_excel(writer, index=False, sheet_name='Sheet1')

    print(Fore.YELLOW + "Ürünlerin Resim Yüklenme Tarihleri Ayrıştırıldı")
    #endregion

    #region XML İndirme ve Ürün Listesine Kategorileri Çektirme

    # XML'den Ürün Bilgilerini Çekme ve Temizleme
    xml_url = "https://task.haydigiy.com/FaprikaXml/R28I7Z/1/"
    response = requests.get(xml_url)
    xml_data = response.text
    soup = BeautifulSoup(xml_data, 'xml')

    product_data = []
    for item in soup.find_all('item'):
        title = item.find('title').text
        # ' - H' ile başlayan tüm kısımları kaldırmak için düzenli ifade kullanıyoruz
        title_cleaned = re.sub(r' - H.*', '', title)
        
        product_id = item.find('g:id').text if item.find('g:id') else None
        product_data.append({'UrunAdi': title_cleaned, 'ID': product_id})

    df_xml = pd.DataFrame(product_data)

    # Excel ile Birleştirme
    df_calisma_alani = pd.read_excel('CalismaAlani.xlsx')
    df_merged = pd.merge(df_calisma_alani, df_xml, how='left', left_on='UrunAdi', right_on='UrunAdi')

    # Exceli Kaydet
    df_merged.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "XML İndirildi ve Ürünlerin Kategorileri Tespit Edildi")
    #endregion

    #region Liste Fiyatlarını Hesaplama

    # Ekceli Okuma
    df = pd.read_excel('CalismaAlani.xlsx')

    # Alış Fiyatına Göre İşlemler ve Kategori Kontrolü
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
        if isinstance(kategori, str) and any(category in kategori for category in ["Parfüm", "Gözlük", "Saat"]):
            result *= 1.20
        else:
            result *= 1.10

        return result

    # Yeni Sütun Oluşturma
    df['ListeFiyati'] = df.apply(calculate_list_price, axis=1)

    # Exceli Kaydet
    df.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "Liste Fiyatları Hesaplandı")

    #endregion

    #region Arama Terimlerindeki Tarihleri Güne Çevirme

    # Ekceli Okuma
    df = pd.read_excel('CalismaAlani.xlsx')

    # AramaTerimleri Sütununda Tarih Olanları İşleme Alma
    def calculate_days_to_today(row):
        arama_terimi = row['AramaTerimleri']

        # Eğer hücre boşsa veya tarih içermiyorsa 0 döndür
        if pd.isna(arama_terimi) or not any(char.isdigit() for char in str(arama_terimi)):
            return 0

        # Tarihi çıkartma
        tarih = datetime.strptime(arama_terimi.split(';')[0], '%d.%m.%Y')
        
        # Bugünkü tarihten uzaklık hesapla
        bugun = datetime.today()
        uzaklik = (bugun - tarih).days

        return uzaklik

    # "AramaTerimleri" Sütununu Güncelleme
    df['AramaTerimleri'] = df.apply(calculate_days_to_today, axis=1)

    # Exceli Kaydet
    df.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "Ürünlerin Resim Yüklenme Tarihleri Gün Olarak Çevilirdi")

    #endregion

    #region Aktif Beden Oranını Tespit Etme

    # Ekceli Okuma
    df = pd.read_excel('CalismaAlani.xlsx')

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

    # Exceli Kaydet
    df.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "Ürünlerin Aktif Beden Oranları Hesaplandı")

    #endregion

    #region Stok Adedine Göre ETOPLA Yapma

    # Ekceli Okuma
    df = pd.read_excel('CalismaAlani.xlsx')

    # "UrunAdi" Sütunundaki "StokAdedi" Değerlerinin Toplamını Hesapla
    df['StokAdedi2'] = df.groupby('UrunAdi')['StokAdedi'].transform('sum')

    # "StokAdedi" Sütununu Sil
    df = df.drop(['StokAdedi'], axis=1, errors='ignore')

    # Yenilenen değerleri teke düşür
    df = df.drop_duplicates()

    # Exceli Kaydet
    df.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "Stok Adetleri Beden Bazlıdan Renk Bazlıya Göre Hesaplandı")

    #endregion

    #region GMT ve SİTA'dan Veri Çekip Stok Adedine İşleme (Yeni)

    google_sheet_url = "https://docs.google.com/spreadsheets/d/1aA5LhkQYgtwHLcKRV1mKl9Lb6VeOgUNIC9zy2kRagrs/gviz/tq?tqx=out:csv"

    try:
        google_df = pd.read_csv(google_sheet_url)
        google_excel_file = "GMT ve SİTA.xlsx"
        google_df.to_excel(google_excel_file, index=False)
    except requests.exceptions.RequestException as e:
        pass





    # Excel dosyalarını okuma
    calisma_alani_df = pd.read_excel('CalismaAlani.xlsx')
    gmt_sita_df = pd.read_excel('GMT ve SİTA.xlsx')

    # GMT sütununu oluşturma ve doldurma
    def get_gmt_etopla(urun_adi):
        matched_row = gmt_sita_df[gmt_sita_df['GMT Ürün Adı'] == urun_adi]
        if not matched_row.empty:
            return matched_row.iloc[0]['GMT Etopla']
        else:
            return 0

    calisma_alani_df['GMT'] = calisma_alani_df['UrunAdi'].apply(get_gmt_etopla)

    # Sonucu yeni bir Excel dosyasına kaydetme
    calisma_alani_df.to_excel('CalismaAlani.xlsx', index=False)







    # Excel dosyalarını okuma
    calisma_alani_df = pd.read_excel('CalismaAlani.xlsx')
    gmt_sita_df = pd.read_excel('GMT ve SİTA.xlsx')

    # GMT sütununu oluşturma ve doldurma
    def get_gmt_etopla(urun_adi):
        matched_row = gmt_sita_df[gmt_sita_df['SİTA Ürün Adı'] == urun_adi]
        if not matched_row.empty:
            return matched_row.iloc[0]['SİTA Etopla']
        else:
            return 0

    calisma_alani_df['SİTA'] = calisma_alani_df['UrunAdi'].apply(get_gmt_etopla)

    # StokAdedi2 sütununu güncelleme
    calisma_alani_df['StokAdedi2'] = calisma_alani_df['StokAdedi2'] + calisma_alani_df['GMT'] + calisma_alani_df['SİTA']

    # GMT ve SİTA sütunlarını silme
    calisma_alani_df.drop(columns=['GMT', 'SİTA'], inplace=True)

    # Sonucu yeni bir Excel dosyasına kaydetme
    calisma_alani_df.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "GMT ve SİTA'dan Ürün Adetleri Tespit Edilip Stok Adetlerine Dahil Edildi")

    #endregion

    #region Diğer Depo Adetleriyle Sitedeki Adetleri Toplama ve Diğer Depo Adetleri Sütununu Silme

    # Ekceli Okuma
    df = pd.read_excel('CalismaAlani.xlsx')

    # "TrendyolKodu" Sütunundaki Boş Değerleri Sıfır ile Doldur
    df.fillna({'VaryasyonMorhipoKodu': 0}, inplace=True)

    # "StokAdedi2" Sütunundaki Hücrelere "TrendyolKodu" Sütunundaki Değerleri Toplayarak Güncelle
    df['StokAdedi2'] = df['StokAdedi2'] + df['VaryasyonMorhipoKodu']

    # "TrendyolKodu" Sütununu Sil
    df.drop(['VaryasyonMorhipoKodu'], axis=1, inplace=True, errors='ignore')

    # Exceli Kaydet
    df.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "Sitedeki Stok Adetleri İle Diğer Depodaki Stok Adetleri Toplandı ve Sütunlar Silindi")

    #endregion

    #region Sütunları Yeniden Adlandırma

    # Ekceli Okuma
    df = pd.read_excel('CalismaAlani.xlsx')

    # Sütun adlarını yeniden adlandırma
    df.rename(columns={'MorhipoKodu': 'Günlük Satış Adedi', 'HepsiBuradaKodu': 'Görüntülenme Adedi', 'VaryasyonHepsiBuradaKodu': 'Raf Aralığı', 'AramaTerimleri': 'Kaç Gündür Aktif', 'VaryasyonN11Kodu': 'Son 1 Haftada Kaç Gündür Aktif Satışta'}, inplace=True)

    # Exceli Kaydet
    df.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "Sütun İsimleri Yeniden Ayarlandı")

    #endregion

    #region Raf Aralığı Olmayan Ürünleri Temizleme

    # Ekceli Okuma
    df = pd.read_excel('CalismaAlani.xlsx')

    # 'Raf Aralığı' sütununda boş olan hücrelerin satırlarını sil
    df.dropna(subset=['Raf Aralığı'], inplace=True)

    # Kaydet
    df.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "Raf Aralığı Girilmemiş Ürünler Rapordan Hariç Tutuldu")

    #endregion

    #region Raf Aralığı Sütunundaki Tarihleri Başlangıç ve Bitiş Olarak İki Sütuna Ayırma

    # Ekceli Okuma
    df = pd.read_excel('CalismaAlani.xlsx')

    # "Satışa Başlayacak Tarih" sütununu oluşturma ve tarih formatına dönüştürme
    df['Satışa Başlayacak Tarih'] = pd.to_datetime(df['Raf Aralığı'].str.split('-').str[0], format='%d.%m.%Y')

    # "Raf Ömrü" sütununu oluşturma ve tarih formatına dönüştürme
    df['Raf Ömrü'] = pd.to_datetime(df['Raf Aralığı'].str.split('-').str[1], format='%d.%m.%Y')

    # "TrendyolKodu" Sütununu Sil
    df.drop(['Raf Aralığı'], axis=1, inplace=True, errors='ignore')

    # Exceli Kaydet
    df.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "Raf Aralıkları Başlangıç ve Bitiş Olarak İki Sütuna Ayrıldı")

    #endregion

    #region Satışa Başlama Tarihi Bugünden Sonra Olan Ürünleri Listeden Temizleme

    # Ekceli Okuma
    df = pd.read_excel('CalismaAlani.xlsx')

    # Bugünün tarihini al
    bugun = datetime.today()

    # Tarih sütununu datetime formatına dönüştür
    df['Satışa Başlayacak Tarih'] = pd.to_datetime(df['Satışa Başlayacak Tarih'])

    # Bugünden büyük olan satırları filtrele ve sil
    df = df[df['Satışa Başlayacak Tarih'] <= bugun]

    # "TrendyolKodu" Sütununu Sil
    df.drop(['Satışa Başlayacak Tarih'], axis=1, inplace=True, errors='ignore')

    # Exceli Kaydet
    df.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "Satışa Başlaması Gereken Tarihi Henüz Gelmemiş Ürünler Listeden Temizlendi")

    #endregion

    #region Günlük Satış Adedi Sütunu

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # Yeni bir sütun oluşturarak tüm hücrelere 1 değeri atama
    veri['Yeni Günlük Satış Adedi'] = 1

    # "Günlük Satış Adedi" sütunundan sadece 0'dan büyük olan değerleri yeni sütuna aktarma
    veri.loc[veri['Günlük Satış Adedi'] > 0, 'Yeni Günlük Satış Adedi'] = veri['Günlük Satış Adedi']

    # "Günlük Satış Adedi" sütununu silme ve yeni sütunun adını "Günlük Satış Adedi" olarak değiştirme
    veri.drop(columns=['Günlük Satış Adedi'], inplace=True)
    veri.rename(columns={'Yeni Günlük Satış Adedi': 'Günlük Satış Adedi'}, inplace=True)

    # Güncellenmiş veriyi kaydet
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Ürünlerin Günlük Ortalama Satış Adetleri Hesaplandı")

    #endregion

    #region Son 1 Haftada Kaç Gündür Aktif Satıştasına Göre Ürünleri Temizleme

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # "Son 1 Haftada Kaç Gündür Aktif Satışta" sütununda 2'den küçük olan satırları filtrele ve sil
    veri = veri[veri['Son 1 Haftada Kaç Gündür Aktif Satışta'] >= 2]

    # "TrendyolKodu" Sütununu Sil
    veri.drop(['Son 1 Haftada Kaç Gündür Aktif Satışta'], axis=1, inplace=True, errors='ignore')

    # Güncellenmiş veriyi kaydet
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Ürünler Son 1 Haftada Kaç Gündür Aktif Satışta Tespit Edildi ve 2 Gün Altı Temizlendi")

    #endregion

    #region Görüntülenme Adedi Belirli Bir Rakamın Altında Olan Ürünleri Listeden Temizleme

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # "Görüntülenme Adedi" sütununda 40'dan küçük olan satırları filtrele ve sil
    veri = veri[veri['Görüntülenme Adedi'] >= 40]

    # Güncellenmiş veriyi kaydet
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Görüntülenme Adedi Belli Bir Rakam Altında Olan Ürünler Temzilendi")

    #endregion

    #region Kaç Güne Biter Hesaplama

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # "Kaç Güne Biter" adında yeni bir sütun oluştur
    veri['Kaç Güne Biter'] = (veri['StokAdedi2'] / veri['Günlük Satış Adedi']).round()

    # Sonucu en yakın tam sayıya yuvarla
    veri['Kaç Güne Biter'] = veri['Kaç Güne Biter'].astype(int)

    # Güncellenmiş veriyi kaydet
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Ürünlerin Kaç Güne Biteceği Hesaplandı")

    #endregion

    #region Kaç Gündür Aktif Sütununda 7'den Küçük Olan Ürünleri Listeden Temizleme

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # "Görüntülenme Adedi" sütununda 40'dan küçük olan satırları filtrele ve sil
    veri = veri[veri['Kaç Gündür Aktif'] >= 7]

    # Güncellenmiş veriyi kaydet
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Resim Yüklenme Tarihi 7 ve 7 Günün Altında Olan Ürünler Temizlendi")

    #endregion

    #region Görüntülenmenin Satışa Dönüş Oranını Hesaplama

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # "Görüntülenmenin Satış Dönüş Oranı" adında yeni bir sütun oluştur
    veri['Görüntülenmenin Satış Dönüş Oranı'] = veri['Günlük Satış Adedi'] / veri['Görüntülenme Adedi'] * 100


    # Güncellenmiş veriyi kaydet
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Ürünlerin Görüntülenmesinin Satışa Dönüş Oranları Tespit Edildi")

    #endregion

    #region Ürünün Raf Ömrüne Olan Uzaklığını Hesaplama ve Önce Bitiyorsa 0 Olarak Baz Alma

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # Bugünün tarihini al
    bugun = datetime.today()

    # "Raf Ömrü" sütunundaki tarih verilerini datetime nesnelerine dönüştür
    veri['Raf Ömrü'] = pd.to_datetime(veri['Raf Ömrü'])

    # "Raf Ömrüne Olan Uzaklık" adında yeni bir sütun oluştur
    veri['Raf Ömrüne Olan Uzaklık'] = (veri['Raf Ömrü'] - bugun).dt.days - veri['Kaç Güne Biter']

    # "Raf Ömrüne Olan Uzaklık" sütunundaki 0'dan büyük olan değerleri 0 olarak değiştirme
    veri.loc[veri['Raf Ömrüne Olan Uzaklık'] > 0, 'Raf Ömrüne Olan Uzaklık'] = 0

    # Güncellenmiş veriyi kaydet
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Ürünlerin Raf Ömrüne Olan Uzaklığı Hesaplandı Eğer Önce Bitiyorsa 0 Olarak Baz Alındı")

    #endregion

    #region Kategorisi Olmayan Ürünleri Listeden Temizleme

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # Kategori sütununda boş olan satırları filtrele ve sil
    veri = veri.dropna(subset=['Kategori'])

    # Güncellenmiş veriyi kaydet
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Kategorisi Olmayan Ürünler Listeden Temizlendi")

    #endregion

    #region Ürünün Aktif Beden Oranına Göre Başarısını Hesaplama

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # Beden Oranı Başarısı Kategori Toplamı adında yeni bir sütun oluşturma
    veri['Beden Oranı Başarısı Kategori Toplamı'] = pd.cut(veri['Beden Durumu'], bins=[0, 25, 50, 75, 100], labels=[9.39, 12.52, 18.78, 37.56], include_lowest=True)

    # Güncellenmiş veriyi kaydet
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Aktif Beden Oranına Göre Ürüne Başarı Oranı Verildi")

    #endregion

    #region Ürünün Kategori Ortalamasını Bulup Raf Ömrüne Olan Uzaklığına Göre Başarısını Hesaplama

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # "Raf Ömrüne Olan Uzaklık" sütunundaki tüm değerleri sıfır ile doldurma
    veri['Raf Ömrü Başarısı Kategori Toplamı'] = 0

    # "Beden Oranı Başarısı Kategori Toplamı" adında yeni bir sütun oluştur
    veri['Raf Ömrü Başarısı Kategori Toplamı'] = veri.groupby('Kategori')['Raf Ömrüne Olan Uzaklık'].transform('sum')

    # Eğer Raf Ömrüne Olan Uzaklık 0 ise, "Raf Ömrü Başarısı Kategori Toplamı" sütununu sıfır yap
    veri.loc[veri['Raf Ömrüne Olan Uzaklık'] == 0, 'Raf Ömrü Başarısı Kategori Toplamı'] = 0

    # Eğer Raf Ömrüne Olan Uzaklık 0 değilse, verilen kodu uygula
    veri.loc[veri['Raf Ömrüne Olan Uzaklık'] != 0, 'Raf Ömrü Başarısı Kategori Toplamı'] = (100 * veri['Raf Ömrüne Olan Uzaklık']) / veri['Raf Ömrü Başarısı Kategori Toplamı'] * 100 / 2 * -1

    # Güncellenmiş veriyi kaydet
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Raf Ömrüne Olan Uzaklık Başarısı Hesaplandı")

    #endregion

    #region Ürünün Kategori Ortalamasını Bulup Görüntülenmenin Satışa Dönüş Oranına Göre Başarısını Hesaplama 

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # "Beden Oranı Başarısı Kategori Toplamı" adında yeni bir sütun oluştur
    veri['Görüntülenmenin Satışa Dönüş Oranı Başarısı Kategori Toplamı'] = veri.groupby('Kategori')['Görüntülenmenin Satış Dönüş Oranı'].transform('sum')

    # "Raf Ömrü Başarısı Kategori Toplamı" sütununu güncelleme
    veri['Görüntülenmenin Satışa Dönüş Oranı Başarısı Kategori Toplamı'] = (100 * veri['Görüntülenmenin Satış Dönüş Oranı']) / veri['Görüntülenmenin Satışa Dönüş Oranı Başarısı Kategori Toplamı'] * 100

    # Güncellenmiş veriyi kaydet
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Görüntülenmenin Satışa Dönüş Oranı Başarısı Hesaplandı")

    #endregion
    
    #region Tüm Başarı Ortalamalarına Göre Belli Bir Önemde Yüzde Çıkarma

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # "Başarı Oranı" adında yeni bir sütun oluşturma ve "Kategori" sütunundaki her kategorinin tekrar sayısını yazma
    veri['Ürün Başarı Oranı'] = veri.groupby('Kategori')['Kategori'].transform('count')

    # "Ürün Başarı Oranı" sütununu güncelleme
    veri['Ürün Başarı Oranı'] = ((veri['Beden Oranı Başarısı Kategori Toplamı'] * 0.10) + (veri['Raf Ömrü Başarısı Kategori Toplamı'] * 0.40) + (veri['Görüntülenmenin Satışa Dönüş Oranı Başarısı Kategori Toplamı'] * 0.50)) * veri['Ürün Başarı Oranı'] / 100

    # Ürün Başarı Oranı sütunundaki 0'dan küçük olan değerleri 0 olarak değiştirme
    veri.loc[veri['Ürün Başarı Oranı'] < 0, 'Ürün Başarı Oranı'] = 0

    # Güncellenmiş veriyi kaydetme
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Tüm Başarı Alanları Önem Sırasına Göre Oranlandı ve Ürün Başarı Ortalaması Hesaplandı")

    #endregion

    #region Çıkan Sonuç Yüzdeye Göre Kategorinin Başarı Ortalamasını Bulma

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # "Kategori Başarı Ortalaması" adında yeni bir sütun oluşturma ve "Kategori" sütunundaki aynı olan değerlerin "Ürün Başarı Oranı" sütunundaki verilerini toplayıp bölme
    veri['Kategori Başarı Ortalaması'] = veri.groupby('Kategori')['Ürün Başarı Oranı'].transform('sum') / veri.groupby('Kategori')['Kategori'].transform('count')

    # Güncellenmiş veriyi kaydetme
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Ürünlerin Başarı Ortalamasına Göre Kategorinin Ortalama Başarısı Hesaplandı")

    #endregion

    #region Ürünün Başarısının Kategori Ortalamasına Olan Uzaklığını Bulma

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # "Yapılan İndirim Yüzdesi" adında yeni bir sütun oluşturma ve "Kategori Başarı Ortalaması" sütunundaki değerler ile "Ürün Başarı Oranı" sütunundaki değerlerin farkını yazma
    veri['Yapılan İndirim Yüzdesi'] = veri['Kategori Başarı Ortalaması'] - veri['Ürün Başarı Oranı']

    # Güncellenmiş veriyi kaydetme
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Kategori Başarısı ve Ürün Başarısı Arasındaki Fark Hesaplandı")

    #endregion

    #region Çıkan Başarı Farkına Göre Ürünün Karına İndirim Yapma

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # "Yapılan İndirim Tutarı" adında yeni bir sütun oluşturma ve "ListeFiyati" sütunundaki değerler ile "AlisFiyati" sütunundaki değerlerin farkını yazma
    veri['Yapılan İndirim Tutarı'] = veri['ListeFiyati'] - veri['AlisFiyati']

    # "Yapılan İndirim Tutarı" sütunundaki veriyi "Yapılan İndirim Yüzdesi" sütunundaki yüzde kadar azaltma
    veri['Yapılan İndirim Tutarı'] *= veri['Yapılan İndirim Yüzdesi'] / 100

    # Çıkan sonucu en yakın tam sayıya yuvarlama (Yapılan İndirimlerin Sadece Yarısı Baz Alındı)
    veri['Yapılan İndirim Tutarı'] = veri['Yapılan İndirim Tutarı'].round() * 2

    # Güncellenmiş veriyi kaydetme
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Çıkan Farka Göre Ürünlere İndirim Uygulandı")

    #endregion

    #region İndirimli Ürün Kategori Ortalama Başarısının Üstüne Çıktığında Çıktığı Kadar Artırım Yapma

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # Yeni sütunun koşullarına göre değer atama, sonucu -1 ile çarpma ve en yakın tam sayıya yuvarlama
    veri['Yapılan Artırım'] = np.where((veri['SatisFiyati'] < veri['ListeFiyati']) & (veri['Yapılan İndirim Yüzdesi'] < 0),
                                    np.round((veri['ListeFiyati'] - veri['AlisFiyati']) * (veri['Yapılan İndirim Yüzdesi'] / 100) * -1),
                                    0)

    # Güncellenmiş veriyi kaydetme
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Kategori Başarısının Üstüne Geçen İndirimli Ürünlere Bindirim Yapıldı ")

    #endregion

    #region Yeni Satış Fiyatını Hesaplama ve Koşullar

    # Excel dosyasını oku
    veri = pd.read_excel('CalismaAlani.xlsx')

    # Yeni Satış Fiyatı sütununu oluşturma
    veri['Yeni Satış Fiyatı'] = np.where((veri['Yapılan İndirim Tutarı'] > 0) & (veri['SatisFiyati'] > veri['AlisFiyati']),
                                        np.where((veri['SatisFiyati'] - veri['Yapılan İndirim Tutarı']) < veri['AlisFiyati'],
                                                veri['AlisFiyati'],
                                                veri['SatisFiyati'] - veri['Yapılan İndirim Tutarı']),
                                        np.where((veri['SatisFiyati'] + veri['Yapılan Artırım']) > veri['ListeFiyati'],
                                                veri['ListeFiyati'],
                                                veri['SatisFiyati'] + veri['Yapılan Artırım']))


    # Güncellenmiş veriyi kaydetme
    veri.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Yapılan İndirim ve Bindirim Alış Fiyatını Altına ya da Liste Fiyatının Üstüne Çıkma Engeli Koyuldu")

    #endregion

    #region Ürünlere Bindirim Yapma (Yeni) (Pasif)
    '''
    # Excel dosyasını okuma
    file_path = 'CalismaAlani.xlsx'
    df = pd.read_excel(file_path)

    # Yeni sütun oluşturma ve şartlara göre değer atama
    def yeni_satis_fiyati_hesapla(row):
        if row['Yapılan İndirim Yüzdesi'] < 0 and row['SatisFiyati'] >= row['ListeFiyati']:
            if row['Günlük Satış Adedi'] in [2, 3]:
                return row['ListeFiyati'] * 1.02
            elif row['Günlük Satış Adedi'] in [4, 5]:
                return row['ListeFiyati'] * 1.04
            elif row['Günlük Satış Adedi'] in [6, 7]:
                return row['ListeFiyati'] * 1.05
            elif row['Günlük Satış Adedi'] in [8, 9]:
                return row['ListeFiyati'] * 1.06
            elif row['Günlük Satış Adedi'] in [10, 11]:
                return row['ListeFiyati'] * 1.07
            elif row['Günlük Satış Adedi'] in [12, 13]:
                return row['ListeFiyati'] * 1.08
            elif row['Günlük Satış Adedi'] in [14, 15]:
                return row['ListeFiyati'] * 1.09
            elif row['Günlük Satış Adedi'] in [16, 17]:
                return row['ListeFiyati'] * 1.10
            elif row['Günlük Satış Adedi'] >= 18:
                return row['ListeFiyati'] * 1.11
        return row['Yeni Satış Fiyatı']

    df['Yeni Satış Fiyatı'] = df.apply(yeni_satis_fiyati_hesapla, axis=1)

    # Sonuçları aynı Excel dosyasına yazma
    output_file_path = 'CalismaAlani.xlsx'
    df.to_excel(output_file_path, index=False)

    print(Fore.YELLOW + "Ürünlere Bindirim Yapıldı")
    '''
    #endregion

    #region İndirimli ve Bindirimli Ürünlerin Yuvarlamalarını Ayarlama (Pasif)
    '''
    # Excel dosyasını oku
    df = pd.read_excel("CalismaAlani.xlsx")

    # "Yeni Satış Fiyatı" ve "ListeFiyati" sütunlarını tam sayıya çevir
    df["Yeni Satış Fiyatı"] = df["Yeni Satış Fiyatı"].astype(int)
    df["ListeFiyati"] = df["ListeFiyati"].astype(int)

    # Verilen koşullara göre veriyi güncelle
    for i in range(len(df)):
        yeni_satis_fiyati = df.at[i, "Yeni Satış Fiyatı"]

        # Yeni Satış Fiyatı belirli aralıklarda ise fiyatı sabitle
        if 100 <= yeni_satis_fiyati <= 105:
            df.at[i, "Yeni Satış Fiyatı"] = 99
        elif 200 <= yeni_satis_fiyati <= 207:
            df.at[i, "Yeni Satış Fiyatı"] = 199
        elif 300 <= yeni_satis_fiyati <= 309:
            df.at[i, "Yeni Satış Fiyatı"] = 299
        elif 400 <= yeni_satis_fiyati <= 412:
            df.at[i, "Yeni Satış Fiyatı"] = 399
        elif 500 <= yeni_satis_fiyati <= 520:
            df.at[i, "Yeni Satış Fiyatı"] = 499
        else:
            # Diğer koşullar için mevcut güncellemeleri uygula
            liste_fiyati = df.at[i, "ListeFiyati"]

            if -1 <= (liste_fiyati - yeni_satis_fiyati) <= 1:
                continue
            
            if yeni_satis_fiyati < 9999:
                last_digit = yeni_satis_fiyati % 10
                if last_digit == 0:
                    df.at[i, "Yeni Satış Fiyatı"] -= 1
                elif last_digit == 1:
                    df.at[i, "Yeni Satış Fiyatı"] -= 2
                elif last_digit == 2:
                    df.at[i, "Yeni Satış Fiyatı"] -= 3
                elif last_digit == 3:
                    df.at[i, "Yeni Satış Fiyatı"] += 1            
                elif last_digit == 5:
                    df.at[i, "Yeni Satış Fiyatı"] -= 1                    
                elif last_digit == 6:
                    df.at[i, "Yeni Satış Fiyatı"] -= 2
                elif last_digit == 7:
                    df.at[i, "Yeni Satış Fiyatı"] += 2
                elif last_digit == 8:
                    df.at[i, "Yeni Satış Fiyatı"] += 1

        # Çıkan sonucu 0,99 ile topla
        df["Yeni Satış Fiyatı"] += 0.99

    # Güncellenmiş DataFrame'i aynı Excel dosyasına kaydet
    df.to_excel("CalismaAlani.xlsx", index=False)

    print(Fore.YELLOW + "Sadece İndirimli ya da Bindirimli Ürünler İçin Fiyat Yuvarlama Yapıldı")
    '''
    #endregion

    #region Yapılan İndirim Tutarı Sütununu Güncelleme

    # Excel dosyasını oku
    calisma_alani = pd.read_excel('CalismaAlani.xlsx')

    # 'Yapılan İndirim Tutarı' sütununu güncelle
    calisma_alani['Yapılan İndirim Tutarı'] = calisma_alani['Yeni Satış Fiyatı'] - calisma_alani['SatisFiyati']

    # Güncellenmiş veriyi Excel dosyasına yaz
    calisma_alani.to_excel('CalismaAlani.xlsx', index=False)

    print(Fore.YELLOW + "Yapılan İndirim Tutarı Sütunu Güncellendi")

    #endregion

    #region Yeni Satış Fiyatlarını Ürün Listesine Çektirme ve Çıkmayanları Silme

    # Excel dosyalarını oku
    df_urun_listesi = pd.read_excel('UrunListesi.xlsx')
    df_calisma_alani = pd.read_excel('CalismaAlani.xlsx')

    # UrunListesi dosyasındaki her bir Stok Kodu için döngü yap
    for index, row in df_urun_listesi.iterrows():
        stok_kodu = row['StokKodu']
        
        # Stok Kodunu CalismaAlani dosyasında ara
        calisma_alani_row = df_calisma_alani[df_calisma_alani['StokKodu'] == stok_kodu]
        
        if not calisma_alani_row.empty:
            # Eşleşme bulunduysa Yeni Satış Fiyatı değerini al
            yeni_satis_fiyati = calisma_alani_row.iloc[0]['Yeni Satış Fiyatı']
            
            # UrunListesi dosyasındaki ilgili satıra Yeni Satış Fiyatı değerini yaz
            df_urun_listesi.at[index, 'SatisFiyati'] = yeni_satis_fiyati
        else:
            # Eşleşme bulunamadıysa "Yok" yaz
            df_urun_listesi.at[index, 'SatisFiyati'] = 'Yok'

    # "Yok" değeri olan satırları sil
    df_urun_listesi = df_urun_listesi[df_urun_listesi['SatisFiyati'] != 'Yok']

    # Sonucu başka bir Excel dosyasına yaz
    df_urun_listesi.to_excel('UrunListesi.xlsx', index=False)

    print(Fore.YELLOW + "Yeni Satış Fiyatıları Çalışma Alanından Çıkarılıp Ürün Listesine İşlendi")

    #endregion

    #region Excelle Ürün Yükleme Alanı

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    login_url = "https://task.haydigiy.com/kullanici-giris/?ReturnUrl=%2Fadmin"
    driver.get(login_url)

    email_input = driver.find_element("id", "EmailOrPhone")
    email_input.send_keys("mustafa_kod@haydigiy.com")

    password_input = driver.find_element("id", "Password")
    password_input.send_keys("123456")
    password_input.send_keys(Keys.RETURN)

    desired_url = "https://task.haydigiy.com/admin/importproductxls/edit/24"
    driver.get(desired_url)

    # Yükle Butonunu Bul
    file_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="qqfile"]')))

    # CalismaAlani Excelini Bul
    file_path = "C:/Users/Public/Panel Görevleri/UrunListesi.xlsx"

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

    time.sleep(360)

    # Tarayıcıyı kapat
    driver.quit()

    print(Fore.GREEN + "İndirimler ve Bindirimler Siteye İşlendi")
    
    #endregion

    #region Özet Dosyaları Oluşturma ve Saklama

    # "CalismaAlani.xlsx" Excel dosyasını oku
    df_calisma_alani = pd.read_excel('CalismaAlani.xlsx')

    # Excel dosyasını güncelleyerek genişlikleri ve ortalamayı ayarla
    with pd.ExcelWriter('CalismaAlani.xlsx', engine='xlsxwriter') as writer:
        df_calisma_alani.to_excel(writer, index=False, sheet_name='Sheet1')

        # ExcelWriter objesinden workbook ve worksheet'e eriş
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        # DataFrame sütun genişliklerini al
        column_widths = [max(df_calisma_alani[col].astype(str).apply(len).max(), len(col)) + 2 for col in df_calisma_alani.columns]

        # Sütun genişliklerini Excel worksheet'e ayarla
        for i, width in enumerate(column_widths):
            worksheet.set_column(i, i, width)

        # Tabloyu ortala
        center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        for i, col in enumerate(df_calisma_alani.columns):
            worksheet.write(0, i, col, center_format)
            
        # Sütun başlıklarının rengini gri yap
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'align': 'center', 'valign': 'vcenter'})
        for col_num, value in enumerate(df_calisma_alani.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Verileri tabloya yazarken ortala
        for i, col in enumerate(df_calisma_alani.columns):
            for j, value in enumerate(df_calisma_alani[col]):
                worksheet.write(j + 1, i, value, center_format)


        # Filtreyi ekle
        worksheet.autofilter(0, 0, df_calisma_alani.shape[0], df_calisma_alani.shape[1] - 1)

    # Excel dosyasını yükle
    wb = load_workbook('CalismaAlani.xlsx')
    ws = wb.active

    # Gizlenecek sütunların harflerini bul
    columns_to_hide = ["Beden Oranı Başarısı Kategori Toplamı", "StokKodu", "Beden Oranı Başarısı Kategori Toplamı", "Raf Ömrü Başarısı Kategori Toplamı", "Görüntülenmenin Satışa Dönüş Oranı Başarısı Kategori Toplamı", "Yapılan İndirim Yüzdesi"]
    column_letters = []

    for column in columns_to_hide:
        for cell in ws[1]:
            if cell.value == column:
                column_letters.append(cell.column_letter)

    # Sütunları gizle
    for col in column_letters:
        ws.column_dimensions[col].hidden = True

    # Değişiklikleri kaydet
    wb.save('CalismaAlani.xlsx')






    # Excel dosyasını yükle
    workbook = openpyxl.load_workbook('CalismaAlani.xlsx')

    # "Sheet1" sayfasını seç
    sheet1 = workbook['Sheet1']

    # "Sheet1" sayfasının kopyasını oluştur ve adını "Alış Fiyatında Olan Ürünler" yap
    workbook.copy_worksheet(sheet1).title = 'Alış Fiyatında Olan Ürünler'

    # "Sheet1" sayfasının adını "İndirim Raporu Özet" olarak değiştir
    sheet1.title = 'İndirim Raporu Özet'

    # Değişiklikleri kaydet
    workbook.save('CalismaAlani.xlsx')






    # Excel dosyasını yükle
    workbook = openpyxl.load_workbook('CalismaAlani.xlsx')

    # "Alış Fiyatında Olan Ürünler" sayfasını seç
    sheet = workbook['Alış Fiyatında Olan Ürünler']

    # Sütun başlıklarını bulmak için ilk satırı oku
    header = {cell.value: idx for idx, cell in enumerate(sheet[1], start=1)}

    # "SatisFiyati" ve "AlisFiyati" sütunlarının indekslerini al
    satis_fiyati_col = header['Yeni Satış Fiyatı']
    alis_fiyati_col = header['AlisFiyati']

    # Silinecek satırları toplamak için liste
    rows_to_delete = []

    # İkinci satırdan başlayarak satırları kontrol et
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        satis_fiyati = row[satis_fiyati_col - 1].value
        alis_fiyati = row[alis_fiyati_col - 1].value

        # Şartları karşılamayan satırları belirle
        if satis_fiyati is not None and alis_fiyati is not None:
            if not (alis_fiyati <= satis_fiyati <= alis_fiyati + 2):
                rows_to_delete.append(row[0].row)

    # Silme işlemini tersten yaparak gerçekleştirelim
    for row_idx in sorted(rows_to_delete, reverse=True):
        sheet.delete_rows(row_idx, 1)

    # Değişiklikleri kaydet
    workbook.save('CalismaAlani.xlsx')











    # Excel dosyasını yükle
    workbook = openpyxl.load_workbook('CalismaAlani.xlsx')

    # "Alış Fiyatında Olan Ürünler" sayfasını seç
    sheet = workbook['Alış Fiyatında Olan Ürünler']

    # Sütun başlıklarını bulmak için ilk satırı oku
    header = {cell.value: idx for idx, cell in enumerate(sheet[1], start=1)}

    # Gerekli başlıkların mevcut olup olmadığını kontrol et
    required_columns = ['Yeni Satış Fiyatı', 'AlisFiyati', 'Yapılan İndirim Tutarı']
    if not all(column in header for column in required_columns):
        raise ValueError(f"Gerekli sütun başlıkları eksik: {', '.join(required_columns)}")

    # "Yapılan İndirim Tutarı" sütununun indeksini al
    indirim_tutari_col = header['Yapılan İndirim Tutarı']

    # Yeni bir sütun ekle ve başlığı "Şimdi mi Alış Fiyatına Çekildi" olarak ayarla
    new_col_idx = sheet.max_column + 1
    sheet.cell(row=1, column=new_col_idx).value = "Şimdi mi Alış Fiyatına Çekildi"

    # İkinci satırdan başlayarak satırları kontrol et ve yeni sütunu doldur
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        indirim_tutari = row[indirim_tutari_col - 1].value

        # "Yapılan İndirim Tutarı" verisini kontrol et ve yeni sütuna değer yaz
        if indirim_tutari is not None:
            if indirim_tutari < 0:
                sheet.cell(row=row[0].row, column=new_col_idx).value = "Evet"
            else:
                sheet.cell(row=row[0].row, column=new_col_idx).value = "Hayır"

    # Değişiklikleri kaydet
    workbook.save('CalismaAlani.xlsx')









    # Yeni dosya adları için tarih bilgisini al
    bugunun_tarihi = datetime.today().strftime('%Y-%m-%d')

    # Dosya adlarını güncelle
    os.rename('CalismaAlani.xlsx', f'{bugunun_tarihi} İndirim Raporu Özet.xlsx')
    os.rename('UrunListesi.xlsx', f'{bugunun_tarihi} İndirim Yükleme.xlsx')


    # Yeni klasör adı için tarih bilgisini al
    bugunun_tarihi = datetime.today().strftime('%Y-%m-%d')

    # Klasörü oluştur
    klasor_yolu = os.path.join(os.getcwd(), bugunun_tarihi)
    os.makedirs(klasor_yolu, exist_ok=True)

    # Dosyaları yeni klasöre taşı
    dosya_adlari = [f'{bugunun_tarihi} İndirim Raporu Özet.xlsx', f'{bugunun_tarihi} İndirim Yükleme.xlsx']
    for dosya_adı in dosya_adlari:
        eski_yol = os.path.join(os.getcwd(), dosya_adı)
        yeni_yol = os.path.join(klasor_yolu, dosya_adı)
        os.rename(eski_yol, yeni_yol)

    #endregion

    os.remove('GMT ve SİTA.xlsx')

else:
    pass




if panel_gorevleri == "E":
    
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
        (abs(df['ListeFiyati2'] - df['SatisFiyati']) > 5),  # Liste Fiyatı 2 ile Satış Fiyatı arasındaki fark > 5
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

    time.sleep(360)
    
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

    #region Unisex Kategorisindeki Ürünleri Unisex Markasına Ekleme

    # Giriş yaptıktan sonra belirtilen sayfaya git
    desired_page_url = "https://task.haydigiy.com/admin/product/bulkedit/"
    driver.get(desired_page_url)



    #Kategori Dahil Alan
    category_select = Select(driver.find_element("id", "SearchInCategoryIds"))

    #Kategori ID'si (Fiyata Hamle)
    category_select.select_by_value("384")

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
    category_id_select.select_by_value("77")

    # Kategori Güncelleme Alanında Yapılacak İşlem Alanını Bulma
    category_transaction_select = driver.find_element(By.ID, "ManufacturerTransactionId")

    # Kategori Güncelleme Alanında Yapılacak İşlemi Seçme (Kategoriden Çıkar)
    category_transaction_select = Select(category_transaction_select)
    category_transaction_select.select_by_value("0")

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

else:
    pass
