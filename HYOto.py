import requests
from bs4 import BeautifulSoup
from io import StringIO
import time
from datetime import datetime
import pandas as pd
from lazyme.string import color_print
from colorama import Fore
import os
import tkinter as tk
from tkinter import messagebox
import sys



url="https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/sirket-karti.aspx?hisse=ACSEL"
r=requests.get(url)
s=BeautifulSoup(r.text,"html.parser")
s1=s.find("select",{"id":"ddlAddCompare"})
s2=s1.find("optgroup").find_all("option")

hisseler=[]

for i in s2:
    hisseler.append(i.text)


def site():
    link=f"https://analizim.halkyatirim.com.tr/Financial/ScoreCardDetail?hisseKod={hisse}"
    r=requests.get(link,headers={'User-Agent': 'XYZ/3.0'})
    soup=BeautifulSoup(r.content,"html.parser")
    return soup
    
def bir(print_table=False):
    tablo=site().find("div",{"id":"pazar-endeskleri"})
    tablo=pd.read_html(StringIO(str(tablo)),flavor="bs4")[0]
    tablo.columns=["Özellikler","Bilgiler"]
    if print_table:
        print(tablo)
    return tablo

def iki(print_table=False):
    tablo=site().find("div",{"id":"fiyat-performansi"})
    tablo=pd.read_html(StringIO(str(tablo)),flavor="bs4")[0]
    tablo.columns.values[0]=""
    if print_table:
        print(tablo)
    return tablo

def uc(print_table=False):
    tablo=site().find("div",{"id":"piyasa-degeri"})
    tablo=pd.read_html(StringIO(str(tablo)),flavor="bs4")[0]
    tablo.columns=["",""]
    if print_table:
        print(tablo)
    return tablo

def dort(print_table=False):
    tablo=site().find("div",{"id":"teknik-veriler"})
    tablo=pd.read_html(StringIO(str(tablo)),flavor="bs4")[0]
    tablo.columns=["Teknik","Değer","Sonuç"]
    if print_table:
        print(tablo)
    return tablo

def bes(print_table=False):
    tablo=site().find("div",{"id":"temel-veri-analizleri"})
    tablo=pd.read_html(StringIO(str(tablo)),flavor="bs4")[0]
    tablo.columns=["Kalemler","Değerler"]
    if print_table:
        print(tablo)
    return tablo

def altı(print_table=False):
    tablo=site().find("div",{"id":"fiyat-ozeti"})
    tablo=pd.read_html(StringIO(str(tablo)),flavor="bs4")[0]
    tablo.columns=["Kalemler","Değerler"]
    if print_table:
        print(tablo)
    return tablo

def yedi(print_table=False):
    tablo=site().find("div",{"id":"finanslar"})
    tablo=pd.read_html(StringIO(str(tablo)),flavor="bs4")[0]
    if print_table:
        print(tablo)
    return tablo

def sekiz(print_table=False):
    tablo=site().find("div",{"id":"karlilik"})
    tablo=pd.read_html(StringIO(str(tablo)),flavor="bs4")[0]
    if print_table:
        print(tablo)
    return tablo

def dokuz(print_table=False):
    tablo=site().find("div",{"id":"carpanlar"})
    tablo=pd.read_html(StringIO(str(tablo)),flavor="bs4")[0]
    if print_table:
        print(tablo)
    return tablo


def on():
    masaustu=os.path.join(os.path.expanduser('~'),"Desktop")
    excel=masaustu + "/{}.xlsx".format(hisse)
    yaz=pd.ExcelWriter(excel,engine="openpyxl")
    
    bir().to_excel(yaz,sheet_name="Pazar Endeskleri",index=False)
    iki().to_excel(yaz,sheet_name="Fiyat Performansı",index=False)
    uc().to_excel(yaz,sheet_name="Piyasa Değeri",index=False)
    dort().to_excel(yaz,sheet_name="Teknik Veriler",index=False)
    bes().to_excel(yaz,sheet_name="Temel Veri Analizleri",index=False)
    altı().to_excel(yaz,sheet_name="Fiyat Özeti",index=False)
    yedi().to_excel(yaz,sheet_name="Finanslar",index=False)
    sekiz().to_excel(yaz,sheet_name="Karlılık",index=False)
    dokuz().to_excel(excel,sheet_name="Çarpanlar",index=False)

    yaz._save()
    
    def uyarı(mesaj):
        root=tk.Tk()
        root.withdraw()
        root.attributes("-topmost",True)
        son=messagebox.showinfo("Uyarı",mesaj)
        root.destroy()
        if son == "ok":
            sys.exit()
   
    mesaj=f"Dosya '{hisse}.xlsx' olarak masaüstüne kaydedildi."
    uyarı(mesaj)

 
color_print(f"******Halk Yatırım Skor Kart******   Hoşgeldiniz @AnlasEkonomi {datetime.now().year} \n\n",color="blue",underline=True)

while True: 
    hisse=input(f"{Fore.RED}Lütfen Hisse Kodu Giriniz: {Fore.RESET}")
    hisse=hisse.upper()

    if hisse in hisseler:
        color_print("Devam Ediliyor...",color="green")
        time.sleep(2)

        s=["Pazar ve Endeksleri","Fiyat Performansı","Piyasa Değeri",
        "Teknik Veriler","Temel Analiz Verileri","Fiyat Özeti","Finansallar",
        "Karlılık","Çarpanlar","Toplu Excel"]

        print("1-{} \n2-{} \n3-{} \n4-{} \n5-{} \n6-{} \n7-{} \n8-{} \n9-{} \n10-{}"
        .format(s[0],s[1],s[2],s[3],s[4],s[5],s[6],s[7],s[8],s[9]))

        while True:
            giris=input("Lütfen İstediğiniz Tablo Kodunu Giriniz...")

            if giris in ["1","2","3","4","5","6","7","8","9","10"]:
                color_print("Devam Ediliyor...",color="green")
                time.sleep(2)
                print("\033c")
                
                if giris=="1":
                    bir(print_table=True)
                    break
                elif giris=="2":
                    iki(print_table=True)
                    break
                elif giris=="3":
                    uc(print_table=True)
                    break 
                elif giris=="4":
                    dort(print_table=True)
                    break
                elif giris=="5":
                    bes(print_table=True)
                    break
                elif giris=="6":
                    altı(print_table=True)
                    break
                elif giris=="7":
                    yedi(print_table=True)
                    break
                elif giris=="8":
                    sekiz(print_table=True)
                    break
                elif giris=="9":
                    dokuz(print_table=True)
                    break
                elif giris=="10":
                    on()
                    break
            else:
                print("Lütfen Geçerli Bir Tablo Kodu Giriniz!!!\n")
                time.sleep(2)
        break
    else:
        print("Lütfen Geçerli Bir Hisse Kodu Giriniz...\n")
        time.sleep(2)