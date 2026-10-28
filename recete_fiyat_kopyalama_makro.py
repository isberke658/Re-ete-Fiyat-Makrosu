# -*- coding: utf-8 -*-
"""
================================================================================
REÇETE MALİYETİ (TİGER / NETUPP) BİRİM FİYAT KOPYALAMA MAKROSU (v2.0)
Hücre 38 Standartlarında - Bağımsız Standalone & Sunucu Uyumlu Sürüm
================================================================================
Özellikler:
1. Gömülü 182 Mamül Manifesti: Sunucuda 'reçete sıra.xlsx' olmasa bile EXE
   aynı dizinde 182 mamüllük Excel dosyasını otomatik oluşturur ve bağlar!
2. Çift Kademeli Kopyalama: Pano temizleme + Donanım Ctrl+C / Ctrl+A + Çift kontrol.
3. Canlı Pano & Ham Veri Takibi: Panodan okunan her metin adli olarak loglanır.
4. Kayıtlı Koordinatlar Hazır: Fiyat (1630, 351), Mamul Seçici (1194, 295).
5. Excel Açık Koruması: Dosya Excel'de kilitliyse bile '_GUNCEL.xlsx' olarak yazar.
================================================================================
"""

import sys
import re
import os
import time
import json
import threading
import datetime
import ctypes
import ctypes.wintypes as wt
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# Windows API
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# 64-bit Güvenli API Tanımlamaları
kernel32.GlobalAlloc.restype = ctypes.c_void_p
kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
user32.GetClipboardData.restype = ctypes.c_void_p
user32.GetClipboardData.argtypes = [ctypes.c_uint]
user32.SetClipboardData.restype = ctypes.c_void_p
user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]

try:
    user32.SetProcessDPIAware()
except Exception:
    pass

# Sabitler
CF_UNICODETEXT = 13
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP   = 0x0004
KEYEVENTF_KEYUP      = 0x0002

VK_RETURN  = 0x0D
VK_DOWN    = 0x28
VK_CONTROL = 0x11
VK_A       = 0x41
VK_C       = 0x43

# Hotkeys
VK_F6      = 0x75
VK_F7      = 0x76
VK_F8      = 0x77
VK_F9      = 0x78
VK_F10     = 0x79
VK_ESCAPE  = 0x1B

if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(APP_DIR, "recete_fiyat_kopyalama_ayarlar.json")
LOG_FILE    = os.path.join(APP_DIR, "recete_fiyat_kopyalama_log.txt")

# Gömülü 182 Mamül Verisi (Offline & Sunucu Yedek Koruması)
EMBEDDED_MAMUL_182 = [
  {
    "sira": 1,
    "mk": "109000000018",
    "uk": "DKR-MS-SS-10",
    "model": "ARTEMİS",
    "ad": "ORTA SEHPA",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 2,
    "mk": "109000000025",
    "uk": "DKR-MS-TT-10",
    "model": "ARTEMİS",
    "ad": "ORTA SEHPA",
    "renk": "Traverten-Traverten"
  },
  {
    "sira": 3,
    "mk": "109000000032",
    "uk": "DKR-MX-AA-14",
    "model": "ARYA",
    "ad": "ZİGON SEHPA",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 4,
    "mk": "109000000049",
    "uk": "DKR-MX-BY-14",
    "model": "ARYA",
    "ad": "ZİGON SEHPA",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 136,
    "mk": "109000000056",
    "uk": "DKR-MX-SA-14",
    "model": "ARYA",
    "ad": "ZİGON SEHPA",
    "renk": "Sonomo-Antrasit"
  },
  {
    "sira": 135,
    "mk": "109000000063",
    "uk": "DKR-MX-SB-14",
    "model": "ARYA",
    "ad": "ZİGON SEHPA",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 8,
    "mk": "109000000070",
    "uk": "DKR-MS-SS-01",
    "model": "BAFA",
    "ad": "SEHPA",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 9,
    "mk": "109000000087",
    "uk": "KTP-AA-15.5",
    "model": "BİLİ",
    "ad": "KİTAPLIK 3 RAFLI 2 KAPAKLI",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 10,
    "mk": "109000000094",
    "uk": "KTP-BY-15.5",
    "model": "BİLİ",
    "ad": "KİTAPLIK 3 RAFLI 2 KAPAKLI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 11,
    "mk": "109000000100",
    "uk": "KTP-SS-15.5",
    "model": "BİLİ",
    "ad": "KİTAPLIK 3 RAFLI 2 KAPAKLI",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 12,
    "mk": "109000000117",
    "uk": "KTP-AA-15.4",
    "model": "BİLİ",
    "ad": "KİTAPLIK 3 RAFLI DAR",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 13,
    "mk": "109000000124",
    "uk": "KTP-BY-15.4",
    "model": "BİLİ",
    "ad": "KİTAPLIK 3 RAFLI DAR",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 14,
    "mk": "109000000131",
    "uk": "KTP-SS-15.4",
    "model": "BİLİ",
    "ad": "KİTAPLIK 3 RAFLI DAR",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 15,
    "mk": "109000000148",
    "uk": "KTP-AA-15.2",
    "model": "BİLİ",
    "ad": "KİTAPLIK 3 RAFLI GENİŞ",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 16,
    "mk": "109000000155",
    "uk": "KTP-BY-15.2",
    "model": "BİLİ",
    "ad": "KİTAPLIK 3 RAFLI GENİŞ",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 17,
    "mk": "109000000162",
    "uk": "KTP-SS-15.2",
    "model": "BİLİ",
    "ad": "KİTAPLIK 3 RAFLI GENİŞ",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 18,
    "mk": "109000000179",
    "uk": "KTP-AA-15.3",
    "model": "BİLİ",
    "ad": "KİTAPLIK 5 RAFLI 2 KAPAKLI",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 19,
    "mk": "109000000186",
    "uk": "KTP-BY-15.3",
    "model": "BİLİ",
    "ad": "KİTAPLIK 5 RAFLI 2 KAPAKLI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 20,
    "mk": "109000000193",
    "uk": "KTP-SS-15.3",
    "model": "BİLİ",
    "ad": "KİTAPLIK 5 RAFLI 2 KAPAKLI",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 21,
    "mk": "109000000209",
    "uk": "KTP-AA-15.1",
    "model": "BİLİ",
    "ad": "KİTAPLIK 5 RAFLI DAR",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 22,
    "mk": "109000000216",
    "uk": "KTP-BY-15.1",
    "model": "BİLİ",
    "ad": "KİTAPLIK 5 RAFLI DAR",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 23,
    "mk": "109000000223",
    "uk": "KTP-SS-15.1",
    "model": "BİLİ",
    "ad": "KİTAPLIK 5 RAFLI DAR",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 24,
    "mk": "109000000230",
    "uk": "KTP-AA-15",
    "model": "BİLİ",
    "ad": "KİTAPLIK 5 RAFLI GENİŞ",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 25,
    "mk": "109000000247",
    "uk": "KTP-BY-15",
    "model": "BİLİ",
    "ad": "KİTAPLIK 5 RAFLI GENİŞ",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 26,
    "mk": "109000000254",
    "uk": "KTP-SS-15",
    "model": "BİLİ",
    "ad": "KİTAPLIK 5 RAFLI GENİŞ",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 27,
    "mk": "109000000261",
    "uk": "DKR-AA-07",
    "model": "BUCAK",
    "ad": "ÇOK AMAÇLI KİTAPLIK",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 28,
    "mk": "109000000278",
    "uk": "DKR-BY-07",
    "model": "BUCAK",
    "ad": "ÇOK AMAÇLI KİTAPLIK",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 29,
    "mk": "109000000285",
    "uk": "DKR-SS-07",
    "model": "BUCAK",
    "ad": "ÇOK AMAÇLI KİTAPLIK",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 30,
    "mk": "109000000292",
    "uk": "DKR-TT-07",
    "model": "BUCAK",
    "ad": "ÇOK AMAÇLI KİTAPLIK",
    "renk": "Traverten-Traverten"
  },
  {
    "sira": 137,
    "mk": "109000000308",
    "uk": "KTP-MX-SB-10.01",
    "model": "DERGA",
    "ad": "2 RAFLI KİTAPLIK",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 138,
    "mk": "109000000315",
    "uk": "KTP-MX-SB-10",
    "model": "DERGA",
    "ad": "3 RAFLI KİTAPLIK",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 35,
    "mk": "109000000322",
    "uk": "SFY-AA-02.02",
    "model": "DERİNKUYU",
    "ad": "2 ÇEKMECELİKOMODİN",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 36,
    "mk": "109000000339",
    "uk": "SFY-BY-02.02",
    "model": "DERİNKUYU",
    "ad": "2 ÇEKMECELİKOMODİN",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 37,
    "mk": "109000000346",
    "uk": "SFY-SS-02.02",
    "model": "DERİNKUYU",
    "ad": "2 ÇEKMECELİKOMODİN",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 38,
    "mk": "109000000353",
    "uk": "SFY-AA-02.01",
    "model": "DERİNKUYU",
    "ad": "3 ÇEKMECELİ ŞİFONYER",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 39,
    "mk": "109000000360",
    "uk": "SFY-BY-02.01",
    "model": "DERİNKUYU",
    "ad": "3 ÇEKMECELİ ŞİFONYER",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 40,
    "mk": "109000000377",
    "uk": "SFY-SS-02.01",
    "model": "DERİNKUYU",
    "ad": "3 ÇEKMECELİ ŞİFONYER",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 181,
    "mk": "109000000384",
    "uk": "SFY-AM-02",
    "model": "DERİNKUYU",
    "ad": "4 ÇEKMECELİ ŞİFONYER",
    "renk": "Amalfi-Amalfi"
  },
  {
    "sira": 41,
    "mk": "109000000391",
    "uk": "SFY-AA-02",
    "model": "DERİNKUYU",
    "ad": "4 ÇEKMECELİ ŞİFONYER",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 42,
    "mk": "109000000407",
    "uk": "SFY-BY-02",
    "model": "DERİNKUYU",
    "ad": "4 ÇEKMECELİ ŞİFONYER",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 43,
    "mk": "109000000414",
    "uk": "SFY-SS-02",
    "model": "DERİNKUYU",
    "ad": "4 ÇEKMECELİ ŞİFONYER",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 44,
    "mk": "109000000421",
    "uk": "SFY-AM-03-02",
    "model": "DERİNKUYU",
    "ad": "5 ÇEKMECELİ ŞİFONYER",
    "renk": "Amalfi-Amalfi"
  },
  {
    "sira": 45,
    "mk": "109000000438",
    "uk": "SFY-AA-03-02",
    "model": "DERİNKUYU",
    "ad": "5 ÇEKMECELİ ŞİFONYER",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 46,
    "mk": "109000000445",
    "uk": "SFY-BY-03-02",
    "model": "DERİNKUYU",
    "ad": "5 ÇEKMECELİ ŞİFONYER",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 47,
    "mk": "109000000452",
    "uk": "SFY-SS-03-02",
    "model": "DERİNKUYU",
    "ad": "5 ÇEKMECELİ ŞİFONYER",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 48,
    "mk": "109000000469",
    "uk": "MTF-MX-BY-01",
    "model": "DORTA",
    "ad": "MUTFAK DOLABI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 139,
    "mk": "109000000476",
    "uk": "MTF-MX-SB-01",
    "model": "DORTA",
    "ad": "MUTFAK DOLABI",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 54,
    "mk": "109000000483",
    "uk": "DLP-BY-3",
    "model": "ELMAR",
    "ad": "ÇOK AMAÇLI DOLAP",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 141,
    "mk": "109000000490",
    "uk": "DLP-SA-3",
    "model": "ELMAR",
    "ad": "ÇOK AMAÇLI DOLAP",
    "renk": "Sonomo-Antrasit"
  },
  {
    "sira": 140,
    "mk": "109000000506",
    "uk": "DLP-SB-3",
    "model": "ELMAR",
    "ad": "ÇOK AMAÇLI DOLAP",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 55,
    "mk": "109000000513",
    "uk": "MTF-MX-BY-03",
    "model": "EMSAR",
    "ad": "MUTFAK DOLABI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 143,
    "mk": "109000000520",
    "uk": "MTF-MX-SA-03",
    "model": "EMSAR",
    "ad": "MUTFAK DOLABI",
    "renk": "Sonomo-Antrasit"
  },
  {
    "sira": 157,
    "mk": "109000000537",
    "uk": "MTF-MX-SA-08",
    "model": "PIDAS",
    "ad": "MUTFAK DOLABI",
    "renk": "Sonomo-Antrasit"
  },
  {
    "sira": 142,
    "mk": "109000000544",
    "uk": "MTF-MX-SB-03",
    "model": "EMSAR",
    "ad": "MUTFAK DOLABI",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 144,
    "mk": "109000000551",
    "uk": "MTF-MX-SY-03",
    "model": "EMSAR",
    "ad": "MUTFAK DOLABI",
    "renk": "Sonomo-Yeşil"
  },
  {
    "sira": 57,
    "mk": "109000000568",
    "uk": "AYK-MX-AA-14",
    "model": "HARRAN",
    "ad": "AYAKKABILIK",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 58,
    "mk": "109000000575",
    "uk": "AYK-MX-BY-14",
    "model": "HARRAN",
    "ad": "AYAKKABILIK",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 59,
    "mk": "109000000582",
    "uk": "AYK-MX-SS-14",
    "model": "HARRAN",
    "ad": "AYAKKABILIK",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 60,
    "mk": "109000000599",
    "uk": "TVU-BY-05",
    "model": "HARTA",
    "ad": "TV ÜNİTESİ",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 145,
    "mk": "109000000605",
    "uk": "TVU-SB-05",
    "model": "HARTA",
    "ad": "TV ÜNİTESİ",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 61,
    "mk": "109000000612",
    "uk": "PRT-MS-BY-02",
    "model": "IVRIZ",
    "ad": "VESTİYER",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 62,
    "mk": "109000000629",
    "uk": "PRT-MS-SS-02",
    "model": "IVRIZ",
    "ad": "VESTİYER",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 63,
    "mk": "109000000636",
    "uk": "AYK-MX-BY-07",
    "model": "JERMA",
    "ad": "AYAKKABILIK",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 147,
    "mk": "109000000643",
    "uk": "AYK-MX-SA-07",
    "model": "JERMA",
    "ad": "AYAKKABILIK",
    "renk": "Sonomo-Antrasit"
  },
  {
    "sira": 146,
    "mk": "109000000650",
    "uk": "AYK-MX-SB-07",
    "model": "JERMA",
    "ad": "AYAKKABILIK",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 76,
    "mk": "109000000667",
    "uk": "YK-DLP-BY-06",
    "model": "LENTO",
    "ad": "1K GARDROP",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 148,
    "mk": "109000000674",
    "uk": "YK-DLP-SB-06",
    "model": "LENTO",
    "ad": "1K GARDROP",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 77,
    "mk": "109000000681",
    "uk": "YK-DLP-BY-05",
    "model": "LENTO",
    "ad": "2K GARDROP",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 149,
    "mk": "109000000698",
    "uk": "YK-DLP-SB-05",
    "model": "LENTO",
    "ad": "2K GARDROP",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 78,
    "mk": "109000000704",
    "uk": "YK-KMD-BY-03",
    "model": "LENTO",
    "ad": "KOMODİN",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 150,
    "mk": "109000000711",
    "uk": "YK-KMD-SB-03",
    "model": "LENTO",
    "ad": "KOMODİN",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 79,
    "mk": "109000000728",
    "uk": "YK-KMD-BY-04",
    "model": "LENTO",
    "ad": "ŞİFONYER",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 151,
    "mk": "109000000735",
    "uk": "YK-KMD-SB-04",
    "model": "LENTO",
    "ad": "ŞİFONYER",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 80,
    "mk": "109000000742",
    "uk": "DKR-MX-BY-03",
    "model": "LİDYA",
    "ad": "ÇOK AMAÇLI RAF",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 152,
    "mk": "109000000759",
    "uk": "DKR-MX-SB-03",
    "model": "LİDYA",
    "ad": "ÇOK AMAÇLI RAF",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 153,
    "mk": "109000000766",
    "uk": "KTP-MS-SB-03",
    "model": "LUMA",
    "ad": "KİTAPLIK",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 81,
    "mk": "109000000773",
    "uk": "CMS-BY-01",
    "model": "MAZI",
    "ad": "ÇALIŞMA MASASI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 82,
    "mk": "109000000780",
    "uk": "AYK-MX-BY-11",
    "model": "ORFA",
    "ad": "AYAKKABILIK",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 154,
    "mk": "109000000797",
    "uk": "AYK-MX-SB-11",
    "model": "ORFA",
    "ad": "AYAKKABILIK",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 83,
    "mk": "109000000803",
    "uk": "AYK-MX-BY-12",
    "model": "PALAZ",
    "ad": "AYAKKABILIK",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 155,
    "mk": "109000000810",
    "uk": "AYK-MX-SB-12",
    "model": "PALAZ",
    "ad": "AYAKKABILIK",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 84,
    "mk": "109000000827",
    "uk": "CM-AN-AA-03",
    "model": "PERA",
    "ad": "ÇALIŞMA MASASI",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 85,
    "mk": "109000000834",
    "uk": "CM-AN-BY-03",
    "model": "PERA",
    "ad": "ÇALIŞMA MASASI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 158,
    "mk": "109000000841",
    "uk": "CM-AN-SB-03",
    "model": "PERA",
    "ad": "ÇALIŞMA MASASI",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 86,
    "mk": "109000000858",
    "uk": "AYK-MS-BY-13",
    "model": "PERGE",
    "ad": "AYAKKABILIK",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 160,
    "mk": "109000000865",
    "uk": "AYK-MS-SA-13",
    "model": "PERGE",
    "ad": "AYAKKABILIK",
    "renk": "Sonomo-Antrasit"
  },
  {
    "sira": 159,
    "mk": "109000000872",
    "uk": "AYK-MS-SB-13",
    "model": "PERGE",
    "ad": "AYAKKABILIK",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 87,
    "mk": "109000000889",
    "uk": "AYK-MS-SS-13",
    "model": "PERGE",
    "ad": "AYAKKABILIK",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 88,
    "mk": "109000000896",
    "uk": "MTF-MX-BY-08",
    "model": "PIDAS",
    "ad": "MUTFAK DOLABI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 182,
    "mk": "109000000902",
    "uk": "MTF-MX-SA-08-2",
    "model": "PIDAS",
    "ad": "MUTFAK DOLABI",
    "renk": "Sonomo-Antrasit"
  },
  {
    "sira": 156,
    "mk": "109000000919",
    "uk": "MTF-MX-SB-08",
    "model": "PIDAS",
    "ad": "MUTFAK DOLABI",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 92,
    "mk": "109000000926",
    "uk": "KMD-MX-BY-06.1",
    "model": "RİVA",
    "ad": "1 ÇEKMECELİ KOMODİN",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 93,
    "mk": "109000000933",
    "uk": "KMD-MX-BY-06",
    "model": "RİVA",
    "ad": "2 ÇEKMECELİ KOMODİN",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 94,
    "mk": "109000000940",
    "uk": "DLP-MX-BY-04",
    "model": "RİVA",
    "ad": "2 KAPILI GARDROP",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 95,
    "mk": "109000000957",
    "uk": "DLP-MX-BY-05",
    "model": "RİVA",
    "ad": "3 KAPILI GARDROP",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 96,
    "mk": "109000000964",
    "uk": "SFY-MX-BY-08",
    "model": "RİVA",
    "ad": "4 ÇEKMECELİ ŞİFONYER",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 97,
    "mk": "109000000971",
    "uk": "SFY-MS-AA-06",
    "model": "RİVA",
    "ad": "5 ÇEKMECELİ ŞİFONYER",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 98,
    "mk": "109000000988",
    "uk": "SFY-MS-BY-06",
    "model": "RİVA",
    "ad": "5 ÇEKMECELİ ŞİFONYER",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 99,
    "mk": "109000000995",
    "uk": "SFY-MS-SS-06",
    "model": "RİVA",
    "ad": "5 ÇEKMECELİ ŞİFONYER",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 100,
    "mk": "109000001008",
    "uk": "TVU-MX-BY-07",
    "model": "SELYA",
    "ad": "TV ÜNİTESİ",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 161,
    "mk": "109000001015",
    "uk": "TVU-MX-SB-07",
    "model": "SELYA",
    "ad": "TV ÜNİTESİ",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 162,
    "mk": "109000001022",
    "uk": "DKR-MX-SB-08",
    "model": "SİMAV",
    "ad": "YAN SEHPA",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 104,
    "mk": "109000001039",
    "uk": "DKR-MS-AA-06",
    "model": "SÜMELA",
    "ad": "RAFLI SEHPA",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 105,
    "mk": "109000001046",
    "uk": "DKR-MS-SS-06",
    "model": "SÜMELA",
    "ad": "RAFLI SEHPA",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 106,
    "mk": "109000001053",
    "uk": "DKR-MS-TT-06",
    "model": "SÜMELA",
    "ad": "RAFLI SEHPA",
    "renk": "Traverten-Traverten"
  },
  {
    "sira": 107,
    "mk": "109000001060",
    "uk": "DKR-MX-BY-13",
    "model": "TARSA",
    "ad": "ORTA SEHPA",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 108,
    "mk": "109000001077",
    "uk": "DKR-MX-SS-13",
    "model": "TARSA",
    "ad": "ORTA SEHPA",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 109,
    "mk": "109000001084",
    "uk": "DKR-MS-AA-11",
    "model": "TİLDA",
    "ad": "KOLTUK ARKASI SEHPA",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 110,
    "mk": "109000001091",
    "uk": "DKR-MS-BY-11",
    "model": "TİLDA",
    "ad": "KOLTUK ARKASI SEHPA",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 111,
    "mk": "109000001107",
    "uk": "DKR-MS-SS-11",
    "model": "TİLDA",
    "ad": "KOLTUK ARKASI SEHPA",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 112,
    "mk": "109000001114",
    "uk": "DKR-MS-TT-11",
    "model": "TİLDA",
    "ad": "KOLTUK ARKASI SEHPA",
    "renk": "Traverten-Traverten"
  },
  {
    "sira": 113,
    "mk": "109000001121",
    "uk": "BNY-BY-03",
    "model": "TİRZA",
    "ad": "ÇAMAŞIR MAK. DOLABI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 163,
    "mk": "109000001138",
    "uk": "BNY-SB-03",
    "model": "TİRZA",
    "ad": "ÇAMAŞIR MAK. DOLABI",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 120,
    "mk": "109000001145",
    "uk": "DKR-MX-AA-04",
    "model": "TRUVA",
    "ad": "YAN SEHPA",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 121,
    "mk": "109000001152",
    "uk": "DKR-MX-BY-04",
    "model": "TRUVA",
    "ad": "YAN SEHPA",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 164,
    "mk": "109000001169",
    "uk": "DKR-MX-SB-04",
    "model": "TRUVA",
    "ad": "YAN SEHPA",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 123,
    "mk": "109000001176",
    "uk": "KMD-AA-04",
    "model": "URARTU",
    "ad": "KOMODİN",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 124,
    "mk": "109000001183",
    "uk": "KMD-BY-04",
    "model": "URARTU",
    "ad": "KOMODİN",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 125,
    "mk": "109000001190",
    "uk": "KMD-SS-04",
    "model": "URARTU",
    "ad": "KOMODİN",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 165,
    "mk": "109000001206",
    "uk": "AYK-MX-SB-13",
    "model": "WERNA",
    "ad": "AYAKKABILIK",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 133,
    "mk": "109000001213",
    "uk": "CMS-AA-04",
    "model": "ZEYNA",
    "ad": "ÇALIŞMA MASASI",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 134,
    "mk": "109000001220",
    "uk": "CMS-BY-04",
    "model": "ZEYNA",
    "ad": "ÇALIŞMA MASASI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 167,
    "mk": "109000001237",
    "uk": "CMS-SA-04",
    "model": "ZEYNA",
    "ad": "ÇALIŞMA MASASI",
    "renk": "Sonomo-Antrasit"
  },
  {
    "sira": 166,
    "mk": "109000001244",
    "uk": "CMS-SB-04",
    "model": "ZEYNA",
    "ad": "ÇALIŞMA MASASI",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 49,
    "mk": "109000001251",
    "uk": "DKR-MS-AM-09",
    "model": "EFES",
    "ad": "YAN SEHPA",
    "renk": "Amalfi-Amalfi"
  },
  {
    "sira": 50,
    "mk": "109000001268",
    "uk": "DKR-MS-AA-09",
    "model": "EFES",
    "ad": "YAN SEHPA",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 52,
    "mk": "109000001275",
    "uk": "DKR-MS-SS-09",
    "model": "EFES",
    "ad": "YAN SEHPA",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 64,
    "mk": "109000001282",
    "uk": "KNZ-KMD-MX-BY-02",
    "model": "KİNZEY",
    "ad": "1 ÇEKMECELİ KOMODİN",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 65,
    "mk": "109000001299",
    "uk": "KNZ-KMD-MX-VV-02",
    "model": "KİNZEY",
    "ad": "1 ÇEKMECELİ KOMODİN",
    "renk": "Vizon- Vizon"
  },
  {
    "sira": 66,
    "mk": "109000001305",
    "uk": "KNZ-KMD-MX-BY-01",
    "model": "KİNZEY",
    "ad": "1 KAPAKLI KOMODİN",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 67,
    "mk": "109000001312",
    "uk": "KNZ-KMD-MX-VV-01",
    "model": "KİNZEY",
    "ad": "1 KAPAKLI KOMODİN",
    "renk": "Vizon- Vizon"
  },
  {
    "sira": 68,
    "mk": "109000001329",
    "uk": "KNZ-DLP-MX-BY-01",
    "model": "KİNZEY",
    "ad": "2K GARDROP",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 69,
    "mk": "109000001336",
    "uk": "KNZ-DLP-MX-VV-01",
    "model": "KİNZEY",
    "ad": "2K GARDROP",
    "renk": "Vizon- Vizon"
  },
  {
    "sira": 70,
    "mk": "109000001343",
    "uk": "KNZ-DLP-MX-BY-02",
    "model": "KİNZEY",
    "ad": "3K GARDROP",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 71,
    "mk": "109000001350",
    "uk": "KNZ-DLP-MX-VV-02",
    "model": "KİNZEY",
    "ad": "3K GARDROP",
    "renk": "Vizon- Vizon"
  },
  {
    "sira": 72,
    "mk": "109000001367",
    "uk": "KNZ-SFY-MX-BY-01",
    "model": "KİNZEY",
    "ad": "7 ÇEKMECELİ ŞİFONYER",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 73,
    "mk": "109000001374",
    "uk": "KNZ-SFY-MX-VV-01",
    "model": "KİNZEY",
    "ad": "7 ÇEKMECELİ ŞİFONYER",
    "renk": "Vizon- Vizon"
  },
  {
    "sira": 74,
    "mk": "109000001381",
    "uk": "KNZ-KRY-MX-BY-01",
    "model": "KİNZEY",
    "ad": "ÇOCUK KARYOLASI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 75,
    "mk": "109000001398",
    "uk": "KNZ-KRY-MX-VV-01",
    "model": "KİNZEY",
    "ad": "ÇOCUK KARYOLASI",
    "renk": "Vizon- Vizon"
  },
  {
    "sira": 5,
    "mk": "109000001404",
    "uk": "CMS-AA-06",
    "model": "ATELY",
    "ad": "CALISMA MASASI",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 6,
    "mk": "109000001411",
    "uk": "CMS-BY-06",
    "model": "ATELY",
    "ad": "CALISMA MASASI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 7,
    "mk": "109000001428",
    "uk": "CMS-SS-06",
    "model": "ATELY",
    "ad": "CALISMA MASASI",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 32,
    "mk": "109000001435",
    "uk": "PRT-MS-AA-09",
    "model": "ÇINAR",
    "ad": "Portmanto",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 33,
    "mk": "109000001442",
    "uk": "PRT-MS-BY-09",
    "model": "ÇINAR",
    "ad": "Portmanto",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 34,
    "mk": "109000001459",
    "uk": "PRT-MS-SS-09",
    "model": "ÇINAR",
    "ad": "Portmanto",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 89,
    "mk": "109000001466",
    "uk": "CMS-AA-05",
    "model": "POYRAZ",
    "ad": "CALISMA MASASI",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 90,
    "mk": "109000001473",
    "uk": "CMS-BY-05",
    "model": "POYRAZ",
    "ad": "CALISMA MASASI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 91,
    "mk": "109000001480",
    "uk": "CMS-SS-05",
    "model": "POYRAZ",
    "ad": "CALISMA MASASI",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 114,
    "mk": "109000001497",
    "uk": "KMD-MX-AA-09.1",
    "model": "TOPRAK",
    "ad": "1 ÇEKMECELİ KOMODİN",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 115,
    "mk": "109000001503",
    "uk": "KMD-MX-BY-09.1",
    "model": "TOPRAK",
    "ad": "1 ÇEKMECELİ KOMODİN",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 116,
    "mk": "109000001510",
    "uk": "KMD-MX-SS-09.1",
    "model": "TOPRAK",
    "ad": "1 ÇEKMECELİ KOMODİN",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 117,
    "mk": "109000001527",
    "uk": "KMD-MX-AA-09",
    "model": "TOPRAK",
    "ad": "KOMODIN",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 118,
    "mk": "109000001534",
    "uk": "KMD-MX-BY-09",
    "model": "TOPRAK",
    "ad": "KOMODIN",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 119,
    "mk": "109000001541",
    "uk": "KMD-MX-SS-09",
    "model": "TOPRAK",
    "ad": "KOMODIN",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 126,
    "mk": "109000001558",
    "uk": "OYD-AA-01",
    "model": "VEGA",
    "ad": "DIKEY OYUNCAK DOLABI",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 127,
    "mk": "109000001565",
    "uk": "OYD-BY-01",
    "model": "VEGA",
    "ad": "DIKEY OYUNCAK DOLABI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 128,
    "mk": "109000001572",
    "uk": "OYD-SS-01",
    "model": "VEGA",
    "ad": "DIKEY OYUNCAK DOLABI",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 129,
    "mk": "109000001589",
    "uk": "OYD-AA-02",
    "model": "VEGA",
    "ad": "YATAY OYUNCAK DOLABI",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 130,
    "mk": "109000001596",
    "uk": "OYD-BY-02",
    "model": "VEGA",
    "ad": "YATAY OYUNCAK DOLABI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 131,
    "mk": "109000001602",
    "uk": "OYD-SS-02",
    "model": "VEGA",
    "ad": "YATAY OYUNCAK DOLABI",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 168,
    "mk": "109000001619",
    "uk": "CMS-AS-04",
    "model": "ZEYNA",
    "ad": "CALISMA MASASI",
    "renk": "Antrasit-Sonomo"
  },
  {
    "sira": 51,
    "mk": "109000001626",
    "uk": "DKR-MS-BY-09",
    "model": "EFES",
    "ad": "YAN SEHPA",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 53,
    "mk": "109000001633",
    "uk": "DKR-MS-TT-09",
    "model": "EFES",
    "ad": "YAN SEHPA",
    "renk": "Traverten-Traverten"
  },
  {
    "sira": 122,
    "mk": "109000001640",
    "uk": "DKR-MX-TT-04",
    "model": "TRUVA",
    "ad": "YAN SEHPA",
    "renk": "Traverten-Traverten"
  },
  {
    "sira": 56,
    "mk": "109000001657",
    "uk": "KMD-BY-02",
    "model": "FIRTINA",
    "ad": "KOMODIN",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 31,
    "mk": "109000001664",
    "uk": "PRT-MS-AM-09",
    "model": "ÇINAR",
    "ad": "Portmanto",
    "renk": "Amalfi-Amalfi"
  },
  {
    "sira": 178,
    "mk": "109000001671",
    "uk": "STB25B05",
    "model": "EVIDEA",
    "ad": "KD BED SIDE TABLE",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 101,
    "mk": "109000001688",
    "uk": "SVN-AA-01",
    "model": "SEVINSE",
    "ad": "CALISMA MASASI",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 102,
    "mk": "109000001695",
    "uk": "SVN-BY-01",
    "model": "SEVINSE",
    "ad": "CALISMA MASASI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 103,
    "mk": "109000001701",
    "uk": "SVN-SS-01",
    "model": "SEVINSE",
    "ad": "CALISMA MASASI",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 179,
    "mk": "109000001718",
    "uk": "WG25B69",
    "model": "EVIDEA",
    "ad": "KD WALL SHELF-2",
    "renk": "Sonomo-Beyaz"
  },
  {
    "sira": 180,
    "mk": "109000001725",
    "uk": "WG26B37",
    "model": "EVIDEA",
    "ad": "KD WALL SHELF-1",
    "renk": "Amalfi-Amalfi"
  },
  {
    "sira": 132,
    "mk": "109000001732",
    "uk": "YK-KRY-BY-02",
    "model": "YASAM",
    "ad": "KARYOLA ALTI CEKMECE",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 169,
    "mk": "109000001749",
    "uk": "CMS-MX-AA-07",
    "model": "KEYIF",
    "ad": "OYUN MASASI",
    "renk": "Antrasit-Antrasit"
  },
  {
    "sira": 170,
    "mk": "109000001756",
    "uk": "CMS-MX-BY-07",
    "model": "KEYIF",
    "ad": "OYUN MASASI",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 171,
    "mk": "109000001763",
    "uk": "CMS-MX-SS-07",
    "model": "KEYIF",
    "ad": "OYUN MASASI",
    "renk": "Sonomo-Sonomo"
  },
  {
    "sira": 174,
    "mk": "109000001770",
    "uk": "NS-KNS-BY-01",
    "model": "NOVA",
    "ad": "KONSOL",
    "renk": "Beyaz-Beyaz"
  },
  {
    "sira": 173,
    "mk": "109000001787",
    "uk": "TVU-MS-SA-12",
    "model": "MELRA",
    "ad": "TV UNITESI",
    "renk": "Sonomo-Antrasit"
  },
  {
    "sira": 175,
    "mk": "109000001787",
    "uk": "NS-KNS-MM-01",
    "model": "NOVA",
    "ad": "KONSOL",
    "renk": "Milano Ceviz-Milano Ceviz"
  },
  {
    "sira": 172,
    "mk": "109000001794",
    "uk": "TVU-MS-MC-12",
    "model": "MELRA",
    "ad": "TV UNITESI",
    "renk": "Mılano Ceviz-Mılano Ceviz"
  },
  {
    "sira": 176,
    "mk": "109000001800",
    "uk": "TVU-MS-MT-13",
    "model": "ULMAR",
    "ad": "TV UNITESI",
    "renk": "Mılano Ceviz-Traverten"
  },
  {
    "sira": 177,
    "mk": "109000001824",
    "uk": "TVU-MS-SA-13",
    "model": "ULMAR",
    "ad": "TV UNITESI",
    "renk": "Sonomo-Antrasit"
  }
]


def create_empty_live_excel_file(file_path: str) -> bool:
    """
    Sıfırdan tertemiz boş bir Excel dosyası oluşturur:
    Sıra | Mamül Kodu | Mamül Adı | Birim Fiyat (TL) | Kayıt Zamanı
    """
    if not OPENPYXL_AVAILABLE:
        return False
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Reçete_Fiyatları"

        headers = ["Sıra", "Mamül Kodu", "Mamül Adı", "Birim Fiyat (TL)", "Kayıt Zamanı"]
        ws.append(headers)

        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        align_center = Alignment(horizontal="center", vertical="center")

        for col_idx in range(1, 6):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center

        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 48
        ws.column_dimensions["D"].width = 20
        ws.column_dimensions["E"].width = 16

        wb.save(file_path)
        return True
    except Exception as e:
        print(f"Boş Excel oluşturma hatası: {e}")
        return False


def find_default_live_excel_path() -> str:
    """Masaüstünde 'Reçete_Maliyeti_Canlı_Liste.xlsx' dosya yolunu belirler."""
    candidates = [
        r"Reçete_Maliyeti_Canlı_Liste.xlsx",
        os.path.join(os.path.expanduser("~"), "OneDrive", "Masaüstü", "Reçete_Maliyeti_Canlı_Liste.xlsx"),
        os.path.join(os.path.expanduser("~"), "Desktop", "Reçete_Maliyeti_Canlı_Liste.xlsx"),
        os.path.join(APP_DIR, "Reçete_Maliyeti_Canlı_Liste.xlsx")
    ]
    for c in candidates:
        if os.path.exists(os.path.dirname(c)):
            return c
    return os.path.join(APP_DIR, "Reçete_Maliyeti_Canlı_Liste.xlsx")


def find_default_excel_path() -> str:
    """Masaüstünde temiz canlı liste Excel dosyasını (Reçete_Maliyeti_Canlı_Liste.xlsx) döndürür."""
    return find_default_live_excel_path()


def create_default_excel_file(file_path: str):
    """Eğer 'reçete sıra.xlsx' bulunamazsa gömülü 182 mamülle sıfırdan oluşturur."""
    if not OPENPYXL_AVAILABLE:
        return False
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sayfa1"

        headers = ["Sütun1", "Mamül Kod", "Ürün Kodu", "Model", "Modül Adı", "Renk", "Birim Fiyat (TL)", "Reçete Durumu", "Kayıt Zamanı"]
        ws.append(headers)
        ws.row_dimensions[1].font = Font(name="Segoe UI", size=11, bold=True)

        for item in EMBEDDED_MAMUL_182:
            ws.append([
                item["sira"],
                item["mk"],
                item["uk"],
                item["model"],
                item["ad"],
                item["renk"],
                "",
                "",
                ""
            ])

        wb.save(file_path)
        return True
    except Exception as e:
        print(f"Varsayılan Excel üretme hatası: {e}")
        return False


def log_to_file(msg: str):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}\n"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


def get_clipboard_text() -> str:
    if not user32.OpenClipboard(0):
        return ""
    try:
        h = user32.GetClipboardData(CF_UNICODETEXT)
        if not h:
            return ""
        p = kernel32.GlobalLock(h)
        if not p:
            return ""
        text = ctypes.c_wchar_p(p).value
        kernel32.GlobalUnlock(h)
        return text or ""
    finally:
        user32.CloseClipboard()


def set_clipboard_text(text: str) -> bool:
    if not user32.OpenClipboard(0):
        return False
    try:
        user32.EmptyClipboard()
        b = (text + '\0').encode('utf-16le')
        h = kernel32.GlobalAlloc(0x0042, len(b))
        if not h:
            return False
        p = kernel32.GlobalLock(h)
        if not p:
            return False
        ctypes.memmove(p, b, len(b))
        kernel32.GlobalUnlock(h)
        user32.SetClipboardData(CF_UNICODETEXT, h)
        return True
    finally:
        user32.CloseClipboard()


def parse_tr_price(val) -> float | None:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace('₺', '').replace('TL', '').replace('tl', '').replace('$', '').replace('€', '').strip()
    if not s or s.startswith('#'):
        return None
    if s.count(',') > 1 and '.' not in s:
        parts = s.split(',')
        try:
            return float(''.join(parts[:-1]) + '.' + parts[-1])
        except Exception:
            return None
    if '.' in s and ',' in s:
        if s.find(',') < s.find('.'):
            clean = s.replace(',', '')
        else:
            clean = s.replace('.', '').replace(',', '.')
        try:
            return float(clean)
        except Exception:
            return None
    if ',' in s:
        clean = s.replace(',', '.')
        try:
            return float(clean)
        except Exception:
            return None
    if '.' in s:
        try:
            return float(s)
        except Exception:
            return None
    try:
        return float(s)
    except Exception:
        return None


def parse_grid_row_clip(text: str) -> dict | None:
    """
    ERP tablosundan (Kodun altı satırından) kopyalanan satırı ayrıştırır.
    """
    if not text:
        return None

    lines = [l.strip() for l in text.strip().split('\n') if l.strip()]
    for line in lines:
        parts = [p.strip() for p in line.split('\t')]
        if len(parts) < 2:
            parts = [p.strip() for p in re.split(r'\s{2,}', line) if p.strip()]

        kod_idx = -1
        for i, part in enumerate(parts):
            clean_p = part.replace('.', '').replace(' ', '')
            if clean_p.startswith('109') and len(clean_p) >= 12:
                kod_idx = i
                break

        if kod_idx != -1:
            raw_kod = parts[kod_idx]
            clean_kod = raw_kod.replace('.', '').replace(' ', '')
            ad = parts[kod_idx + 1] if len(parts) > kod_idx + 1 else ""

            fiyat = None
            fiyat_str = ""

            if len(parts) >= 7:
                cand = parts[6].replace('₺', '').strip()
                try:
                    fiyat = float(cand.replace('.', '').replace(',', '.'))
                    fiyat_str = cand
                except ValueError:
                    pass

            if fiyat is None:
                for p in parts[2:]:
                    cand = p.replace('₺', '').replace('TL', '').strip()
                    if re.match(r'^\d{1,3}(\.\d{3})*,\d{2}$', cand) or re.match(r'^\d+,\d{2}$', cand):
                        try:
                            fiyat = float(cand.replace('.', '').replace(',', '.'))
                            fiyat_str = cand
                            break
                        except ValueError:
                            pass

            if fiyat is None and len(parts) >= 2:
                for cand in reversed(parts):
                    c_clean = cand.replace('₺', '').replace('TL', '').strip()
                    num = parse_tr_price(c_clean)
                    if num is not None and num > 0:
                        fiyat = num
                        fiyat_str = c_clean
                        break

            if fiyat is not None:
                return {
                    'raw_kod': raw_kod,
                    'clean_kod': clean_kod,
                    'ad': ad,
                    'fiyat': fiyat,
                    'fiyat_str': fiyat_str
                }
    return None


VK_F4 = 0x73


class ReceteFiyatEngine:
    def __init__(self, ui_callback=None):
        self.ui_callback = ui_callback
        self.running = False
        self.paused = False
        self.thread = None

        # Varsayılan Fabrika Koordinatları
        self.coord_mamul_box  = (1194, 295)    # Üstteki 'Mamul: [ 109... v ]' kutusu (F9)
        self.coord_price_cell = (1556, 382)    # Tablodaki 1. Satır / Mamül Satırı ("Kodun Altı") (F8)

        # Excel Parametreleri
        self.live_excel_path = find_default_live_excel_path()
        self.excel_path = self.live_excel_path
        self.sheet_name = "Sayfa1"
        self.start_row = 2
        self.price_col_letter = "G"
        self.current_idx = 0
        self.total_rows = 182
        self.last_copied_code = ""
        self.captured_count = 0

        # Zamanlama & Seçenekler
        self.wait_after_down_sec = 1.0         # Mamül değiştikten sonra hesaplama beklemesi
        self.copy_delay_sec = 0.35             # Kopyalama beklemesi
        self.skip_on_fail = True               # Fiyat okunamadığında sonraki mamüle devam et
        self.press_enter_after_down = True     # Aşağı ok sonrası DevExpress Enter onayı
        self.target_repeat_count = 182         # Kullanıcının belirlediği tekrar / işlem sayısı
        self.consecutive_same_count = 0

        self.mode = "auto"
        self.mamul_list = []

        self.load_excel_manifest()
        self.load_config()

    def log(self, text: str):
        log_to_file(text)
        if self.ui_callback:
            self.ui_callback("log", text)

    def update_status(self, text: str):
        if self.ui_callback:
            self.ui_callback("status", text)

    def update_progress(self, current: int, total: int):
        if self.ui_callback:
            self.ui_callback("progress", (current, total))

    def load_excel_manifest(self, custom_path: str = None) -> bool:
        if custom_path:
            self.excel_path = custom_path

        is_live = ("canlı" in self.excel_path.lower() or 
                   "canli" in self.excel_path.lower() or 
                   self.excel_path.endswith("Reçete_Maliyeti_Canlı_Liste.xlsx"))

        if is_live:
            if not os.path.exists(self.excel_path):
                create_empty_live_excel_file(self.excel_path)
            try:
                wb = openpyxl.load_workbook(self.excel_path, data_only=True)
                ws = wb.active
                self.captured_count = max(0, ws.max_row - 1)
                self.current_idx = self.captured_count
                self.total_rows = getattr(self, "target_repeat_count", 182) or 182
                self.mamul_list = []
                self.log(f"📁 Temiz Canlı Liste Excel bağlandı: {os.path.basename(self.excel_path)} ({self.captured_count} kayıtlı mamül)")
                return True
            except Exception as e:
                self.log(f"Canlı liste açma uyarısı: {e}")
                create_empty_live_excel_file(self.excel_path)
                return True

        if not os.path.exists(self.excel_path):
            create_default_excel_file(self.excel_path)

        if not os.path.exists(self.excel_path) or not OPENPYXL_AVAILABLE:
            self.mamul_list = []
            for i, it in enumerate(EMBEDDED_MAMUL_182, 1):
                self.mamul_list.append({
                    "row": i + 1,
                    "mk": it["mk"],
                    "uk": it["uk"],
                    "model": it["model"],
                    "ad": it["ad"],
                    "renk": it["renk"],
                    "fiyat": None
                })
            self.total_rows = len(self.mamul_list)
            self.log(f"Gömülü yedek listeden {self.total_rows} mamül hafızaya alındı.")
            return True

        try:
            wb = openpyxl.load_workbook(self.excel_path, data_only=True)
            ws = wb[self.sheet_name] if self.sheet_name in wb.sheetnames else wb.active
            self.mamul_list = []
            for r in range(self.start_row, ws.max_row + 1):
                mk = ws.cell(r, 2).value
                uk = ws.cell(r, 3).value
                model = ws.cell(r, 4).value
                ad = ws.cell(r, 5).value
                renk = ws.cell(r, 6).value
                mevcut_fiyat = ws.cell(r, 7).value
                if mk:
                    self.mamul_list.append({
                        "row": r,
                        "mk": str(mk).replace('.', '').strip(),
                        "uk": str(uk or "").strip(),
                        "model": str(model or "").strip(),
                        "ad": str(ad or "").strip(),
                        "renk": str(renk or "").strip(),
                        "fiyat": mevcut_fiyat
                    })

            self.total_rows = len(self.mamul_list) if self.mamul_list else 182
            self.log(f"Excel dosyasından {self.total_rows} mamül yüklendi: {os.path.basename(self.excel_path)}")
            return True
        except Exception as e:
            self.log(f"Excel okuma hatası: {e}.")
            self.mamul_list = []
            self.total_rows = 182
            return True

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    self.coord_price_cell = tuple(cfg.get("coord_price_cell", self.coord_price_cell))
                    self.coord_mamul_box = tuple(cfg.get("coord_mamul_box", self.coord_mamul_box))
                    self.start_row = int(cfg.get("start_row", self.start_row))
                    self.current_idx = int(cfg.get("current_idx", 0))
                    saved_path = cfg.get("excel_path", "")
                    if saved_path and os.path.exists(saved_path) and "reçete sıra" not in saved_path.lower():
                        self.excel_path = saved_path
                    else:
                        self.excel_path = self.live_excel_path
                    self.price_col_letter = cfg.get("price_col_letter", self.price_col_letter)
                    self.wait_after_down_sec = float(cfg.get("wait_after_down_sec", self.wait_after_down_sec))
                    self.press_enter_after_down = bool(cfg.get("press_enter_after_down", self.press_enter_after_down))
                    self.target_repeat_count = int(cfg.get("target_repeat_count", self.target_repeat_count))
            except Exception:
                pass

    def save_config(self):
        cfg = {
            "coord_price_cell": list(self.coord_price_cell),
            "coord_mamul_box": list(self.coord_mamul_box),
            "start_row": self.start_row,
            "current_idx": self.current_idx,
            "excel_path": self.excel_path,
            "price_col_letter": self.price_col_letter,
            "wait_after_down_sec": self.wait_after_down_sec,
            "press_enter_after_down": self.press_enter_after_down,
            "target_repeat_count": self.target_repeat_count
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

    def mouse_click(self, x: int, y: int, double: bool = False):
        user32.SetCursorPos(int(x), int(y))
        time.sleep(0.06)
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.04)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        if double:
            time.sleep(0.08)
            user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            time.sleep(0.04)
            user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.08)

    def send_key(self, vk_code: int):
        user32.keybd_event(vk_code, 0, 0, 0)
        time.sleep(0.04)
        user32.keybd_event(vk_code, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.06)

    def send_ctrl_c(self):
        user32.keybd_event(VK_CONTROL, 0, 0, 0)
        time.sleep(0.04)
        user32.keybd_event(VK_C, 0, 0, 0)
        time.sleep(0.05)
        user32.keybd_event(VK_C, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.04)
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.1)

    def send_ctrl_a_c(self):
        user32.keybd_event(VK_CONTROL, 0, 0, 0)
        time.sleep(0.04)
        user32.keybd_event(VK_A, 0, 0, 0)
        time.sleep(0.04)
        user32.keybd_event(VK_A, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.04)
        user32.keybd_event(VK_C, 0, 0, 0)
        time.sleep(0.05)
        user32.keybd_event(VK_C, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.04)
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.1)

    ERP_RECETESI_OLMAYAN_MAMULLER = {
        "109000000384",  # DERİNKUYU 4 ÇEKMECELİ ŞİFONYER (Amalfi-Amalfi) - Satır 39
        "109000000902",  # PIDAS MUTFAK DOLABI (Sonomo-Antrasit) - Satır 91
        "109000001671",  # EVIDEA KD BED SIDE TABLE (Sonomo-Sonomo) - Satır 168
        "109000001718",  # EVIDEA KD WALL SHELF-2 (Sonomo-Beyaz) - Satır 172
        "109000001725",  # EVIDEA KD WALL SHELF-1 (Amalfi-Amalfi) - Satır 173
    }

    def save_price_to_excel(self, row_idx: int, price_val: float | None, status: str = "") -> bool:
        if not os.path.exists(self.excel_path):
            create_default_excel_file(self.excel_path)
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb[self.sheet_name] if self.sheet_name in wb.sheetnames else wb.active
            cell_ref = f"{self.price_col_letter.upper()}{row_idx}"
            ws[cell_ref] = price_val
            if status:
                ws[f"H{row_idx}"] = status
            ws[f"I{row_idx}"] = datetime.datetime.now().strftime("%H:%M:%S")
            wb.save(self.excel_path)
            return True
        except PermissionError:
            alt_path = self.excel_path.replace(".xlsx", "_GUNCEL.xlsx")
            self.log(f"UYARI: Excel dosyası açık! Alternatif dosyaya yazılıyor: {os.path.basename(alt_path)}")
            try:
                wb = openpyxl.load_workbook(self.excel_path)
                ws = wb[self.sheet_name] if self.sheet_name in wb.sheetnames else wb.active
                ws[f"{self.price_col_letter.upper()}{row_idx}"] = price_val
                if status:
                    ws[f"H{row_idx}"] = status
                ws[f"I{row_idx}"] = datetime.datetime.now().strftime("%H:%M:%S")
                wb.save(alt_path)
                return True
            except Exception:
                return False
        except Exception as e:
            self.log(f"Excel kayıt hatası: {e}")
            return False

    def reset_all_and_clear_excel(self) -> bool:
        """
        Kullanıcı 'Başa Sar' dediğinde:
        1. Sayaçları sıfırlar (current_idx = 0, captured_count = 0, last_copied_code = '', consecutive_same_count = 0).
        2. Canlı Excel dosyasını (Reçete_Maliyeti_Canlı_Liste.xlsx) TAMAMEN SIFIRLAR ve boş şablon yazar!
        3. Eğer 'reçete sıra.xlsx' hedef dosya ise veya mevcutsa, onun da G (Fiyat), H (Durum), I (Saat) sütunlarını sıfırlar.
        """
        self.current_idx = 0
        self.captured_count = 0
        self.last_copied_code = ""
        self.consecutive_same_count = 0
        self.save_config()

        # 1. Canlı Excel Dosyasını (Reçete_Maliyeti_Canlı_Liste.xlsx) sıfırla:
        create_empty_live_excel_file(self.live_excel_path)

        # 2. Eğer hedef excel yolu (self.excel_path) farklı bir dosya ise:
        if self.excel_path and os.path.exists(self.excel_path) and os.path.abspath(self.excel_path) != os.path.abspath(self.live_excel_path):
            try:
                wb = openpyxl.load_workbook(self.excel_path)
                ws = wb[self.sheet_name] if self.sheet_name in wb.sheetnames else wb.active
                if ws.cell(1, 1).value == "Sıra" and ws.cell(1, 2).value == "Mamül Kodu":
                    create_empty_live_excel_file(self.excel_path)
                else:
                    for r in range(self.start_row, ws.max_row + 1):
                        ws.cell(r, 7, None)
                        ws.cell(r, 8, None)
                        ws.cell(r, 9, None)
                    wb.save(self.excel_path)
            except Exception as e:
                self.log(f"Hedef Excel sıfırlama uyarısı: {e}")

        self.log("🔄 [SIFIRLAMA] Excel içeriği tamamen silindi ve sıfırlandı. 1. sıradan başlamaya hazır!")
        return True

    def save_grid_row_to_excel(self, data: dict) -> tuple[int, bool]:
        """
        Kopyalanan satırdan çekilen Mamül Kodu, Adı ve Fiyatını doğrudan
        Reçete_Maliyeti_Canlı_Liste.xlsx dosyasına ve varsa hedef Excel'e yazar.
        """
        live_path = self.live_excel_path
        target_updated = False
        sira = 1

        # 1. Temiz Canlı Liste Dosyası (Reçete_Maliyeti_Canlı_Liste.xlsx)
        try:
            if not os.path.exists(live_path):
                create_empty_live_excel_file(live_path)

            wb_live = openpyxl.load_workbook(live_path)
            ws_live = wb_live.active

            # Başlık 1. satırda olduğu için sıra = ws_live.max_row
            sira = ws_live.max_row
            ws_live.append([
                sira,
                data["raw_kod"],
                data["ad"],
                data["fiyat"],
                datetime.datetime.now().strftime("%H:%M:%S")
            ])
            # Sayı formatı (#,##0.00)
            cell_f = ws_live.cell(row=ws_live.max_row, column=4)
            cell_f.number_format = '#,##0.00'
            cell_f.alignment = Alignment(horizontal="right")

            wb_live.save(live_path)
        except PermissionError:
            alt_live = live_path.replace(".xlsx", "_GUNCEL.xlsx")
            self.log(f"UYARI: Canlı Excel açık! Alternatif dosyaya yazılıyor: {os.path.basename(alt_live)}")
            try:
                if not os.path.exists(alt_live):
                    create_empty_live_excel_file(alt_live)
                wb_alt = openpyxl.load_workbook(alt_live)
                ws_alt = wb_alt.active
                sira = ws_alt.max_row
                ws_alt.append([
                    sira,
                    data["raw_kod"],
                    data["ad"],
                    data["fiyat"],
                    datetime.datetime.now().strftime("%H:%M:%S")
                ])
                wb_alt.save(alt_live)
            except Exception as e2:
                self.log(f"Alternatif Canlı Excel kayıt hatası: {e2}")
        except Exception as e:
            self.log(f"Canlı Excel kayıt uyarısı: {e}")

        # 2. Varsa Hedef Excel (reçete sıra.xlsx)
        if self.excel_path and os.path.exists(self.excel_path) and os.path.abspath(self.excel_path) != os.path.abspath(live_path):
            try:
                wb_t = openpyxl.load_workbook(self.excel_path)
                ws_t = wb_t[self.sheet_name] if self.sheet_name in wb_t.sheetnames else wb_t.active
                clean_target = data["clean_kod"]

                for r in range(2, ws_t.max_row + 1):
                    cell_mk = str(ws_t.cell(r, 2).value or "").replace(".", "").strip()
                    if cell_mk == clean_target:
                        ws_t.cell(r, 7, data["fiyat"])
                        ws_t.cell(r, 8, "Aktarıldı (OK)")
                        ws_t.cell(r, 9, datetime.datetime.now().strftime("%H:%M:%S"))
                        target_updated = True
                        break

                if not target_updated:
                    new_r = ws_t.max_row + 1
                    ws_t.append([
                        new_r - 1,
                        data["raw_kod"],
                        "",
                        "",
                        data["ad"],
                        "",
                        data["fiyat"],
                        "Aktarıldı (Yeni)",
                        datetime.datetime.now().strftime("%H:%M:%S")
                    ])
                    target_updated = True

                wb_t.save(self.excel_path)
            except Exception as e:
                self.log(f"Hedef Excel ({os.path.basename(self.excel_path)}) kayıt uyarısı: {e}")

        return sira, target_updated

    def copy_and_record_current(self):
        """
        Tablodaki 1. Satıra ('Kodun Altı') tıklar, kopyalar ve ANINDA Excel'e kaydeder.
        """
        set_clipboard_text("")
        time.sleep(0.05)

        px, py = self.coord_price_cell
        self.mouse_click(px, py)
        time.sleep(0.15)
        self.send_ctrl_c()
        time.sleep(self.copy_delay_sec)

        raw_text = get_clipboard_text().strip()

        if not raw_text:
            time.sleep(0.10)
            self.mouse_click(px, py, double=True)
            time.sleep(0.12)
            self.send_ctrl_c()
            time.sleep(self.copy_delay_sec)
            raw_text = get_clipboard_text().strip()

        if not raw_text:
            time.sleep(0.08)
            self.send_ctrl_a_c()
            time.sleep(self.copy_delay_sec)
            raw_text = get_clipboard_text().strip()

        # 1. Akıllı Grid/Satır Kopyalamayı (Kodun Altı) dene:
        grid_data = parse_grid_row_clip(raw_text)
        if grid_data is not None:
            # KRİTİK: Eğer aynı mamül kodu geldiyse hemen durma!
            if self.last_copied_code and grid_data["clean_kod"] == self.last_copied_code:
                self.consecutive_same_count += 1
                if self.consecutive_same_count >= 5:
                    self.log(f"🎉 ERP LİSTESİ SONUNA ULAŞILDI (Son Mamül: {grid_data['raw_kod']} - {grid_data['ad']})")
                    self.stop()
                    return "FINISHED"
                else:
                    self.log(f"⏳ Henüz sıradaki reçeteye geçilmedi (Hala: {grid_data['raw_kod']}). F4 ile liste zorla açılıp tekrar Aşağı Ok deneniyor...")
                    return "SAME_ITEM_RETRY"

            self.consecutive_same_count = 0
            self.last_copied_code = grid_data["clean_kod"]
            self.captured_count += 1

            sira, target_updated = self.save_grid_row_to_excel(grid_data)
            target_info = " + reçete sıra.xlsx Güncellendi" if target_updated else ""
            self.log(f"✅ [Canlı #{sira}] {grid_data['raw_kod']} | {grid_data['ad']} ➔ {grid_data['fiyat_str']} ₺ (Excel'e Eklendi{target_info})")

            self.current_idx += 1
            self.save_config()
            self.update_progress(self.current_idx, self.total_rows)
            return grid_data["fiyat"]

        # 2. Klasik Yöntem Yedek:
        num_val = parse_tr_price(raw_text)
        if num_val is not None:
            if self.current_idx < len(self.mamul_list):
                mamul = self.mamul_list[self.current_idx]
                target_row = mamul["row"]
                mk_code = mamul["mk"]
                model_name = f"{mamul['model']} {mamul['ad']} ({mamul['renk']})"
                success = self.save_price_to_excel(target_row, num_val, status="Aktarıldı (OK)")
                formatted_str = f"{num_val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') + " ₺"
                if success:
                    self.log(f"✅ [{self.current_idx + 1}/{self.total_rows}] {mk_code} | {model_name} ➔ {formatted_str} (Hücre: G{target_row})")
                else:
                    self.log(f"❌ [{self.current_idx + 1}/{self.total_rows}] {mk_code} | Excel kayıt hatası!")

            self.current_idx += 1
            self.save_config()
            self.update_progress(self.current_idx, self.total_rows)
            return num_val

        p_preview = raw_text[:40] if raw_text else "BOŞ (Kopyalanamadı)"
        self.log(f"⚠️ [{self.current_idx + 1}/{self.total_rows}] Fiyat okunamadı! (Pano: '{p_preview}' | Boyut: {len(raw_text)})")
        return None

    def advance_to_next_mamul(self):
        """
        Programdaki Mamül seçici kutusuna tıklar, F4 ile açılır listeyi anında açar,
        Aşağı Ok (VK_DOWN) ile sıradaki mamüle geçer ve Enter ile seçimi onaylar.
        """
        mx, my = self.coord_mamul_box
        # 1. Mamul kutusunu kesin olarak odakla
        self.mouse_click(mx, my)
        time.sleep(0.12)

        # 2. LookUpEdit listesini tek seferde açmak için F4 gönder
        self.send_key(VK_F4)
        time.sleep(0.10)

        # 3. Aşağı Ok (VK_DOWN) basarak sıradaki mamüle geç
        self.send_key(VK_DOWN)
        time.sleep(0.10)

        # 4. Enter (VK_RETURN) basarak seçimi onayla ve listeyi kapat
        if getattr(self, "press_enter_after_down", True):
            self.send_key(VK_RETURN)

        # 5. Reçetenin ekrana yüklenmesini ve hesaplanmasını bekle
        time.sleep(self.wait_after_down_sec)

    def start_auto_loop(self):
        if self.running:
            return
        if not self.mamul_list:
            self.load_excel_manifest()

        self.running = True
        self.paused = False
        self.consecutive_same_count = 0
        self.mode = "auto"
        self.thread = threading.Thread(target=self._auto_loop_worker, daemon=True)
        self.thread.start()

    def start_watch_mode(self):
        if self.running:
            return
        if not self.mamul_list:
            self.load_excel_manifest()

        self.running = True
        self.paused = False
        self.mode = "watch"
        self.thread = threading.Thread(target=self._watch_loop_worker, daemon=True)
        self.thread.start()

    def pause(self):
        self.paused = True
        self.update_status("DURAKLATILDI (F6 ile devam et)")
        self.log("Döngü duraklatıldı.")

    def resume(self):
        self.paused = False
        self.update_status("ÇALIŞIYOR...")
        self.log("Döngü devam ettiriliyor.")

    def find_first_empty_excel_row(self) -> int:
        """
        Canlı Excel dosyasındaki kayıt sayısını okur ve kalınan sırayı (indeksi) döndürür.
        Örn: 8 kayıt varsa slot = 8 (yani 9. mamülden devam eder). Boşsa slot = 0 (1. mamül).
        """
        target_file = self.live_excel_path if os.path.exists(self.live_excel_path) else self.excel_path
        if not os.path.exists(target_file):
            return 0
        try:
            wb = openpyxl.load_workbook(target_file, data_only=True)
            ws = wb.active
            if ws.cell(1, 1).value == "Sıra" and ws.cell(1, 2).value == "Mamül Kodu":
                return max(0, ws.max_row - 1)
            else:
                for idx, mamul in enumerate(self.mamul_list):
                    row_num = mamul["row"]
                    val = ws.cell(row=row_num, column=7).value
                    if val is None or str(val).strip() == "":
                        return idx
                return len(self.mamul_list)
        except Exception as e:
            self.log(f"Excel boş satır tarama uyarısı: {e}")
            return self.current_idx

    def stop(self):
        self.running = False
        self.paused = False
        self.save_config()
        sira_no = max(1, self.captured_count + 1)
        son_kod = self.last_copied_code or "Yok"
        self.update_status(f"DURDURULDU · Kaldığı Sıra: {sira_no} · F6 ile Devam Et")
        self.log(f"Makro durduruldu. Kaldığı sıra: {sira_no} (Son Mamül: {son_kod}). Çekilen veriler Excel'e işlendi. Tekrar F6 basarak devam edebilirsiniz.")

    def _auto_loop_worker(self):
        target = getattr(self, "target_repeat_count", 182)
        target_desc = f"{target} Adet Tekrar" if target > 0 else "Sınırsız (0)"
        self.log(f"=== TAM OTOMATİK AKTARIM BAŞLATILDI [Başlangıç: {self.current_idx + 1}. Mamül | Hedef: {target_desc}] ===")
        self.update_status(f"ÇALIŞIYOR (Hedef: {target_desc})...")

        completed_in_this_run = 0

        while self.running:
            while self.paused and self.running:
                time.sleep(0.2)

            if not self.running:
                break

            # 1. Kopyala ve Excel'e Anında Yaz
            val = self.copy_and_record_current()

            if val == "FINISHED":
                break

            if not self.running:
                break

            # Eğer aynı mamül geldiyse sıradakine geçmek için tekrar aşağı in
            if val == "SAME_ITEM_RETRY":
                time.sleep(0.2)
                self.advance_to_next_mamul()
                continue

            # Başarılı bir mamül çekildiyse sayacı artır
            if val is not None:
                completed_in_this_run += 1

            # Kullanıcının belirlediği işlem sayısına ulaşıldıysa dur!
            if target > 0 and completed_in_this_run >= target:
                self.log(f"🎯 HEDEFE ULAŞILDI: Belirlenen {target} adet mamül başarıyla aktarıldı ve Excel'e kaydedildi!")
                self.stop()
                break

            # ERP'de reçetesi olmayan mamül ise ERP ekranında Aşağı Ok basılmaz
            if val == "SKIPPED_NO_RECETE":
                time.sleep(0.1)
                continue

            # Eğer fiyat okunamadıysa ve skip_on_fail aktifse sonraki mamüle devam et
            if val is None and self.skip_on_fail:
                self.current_idx += 1
                self.update_progress(self.current_idx, max(self.current_idx, self.total_rows))

            # 2. Sıradaki Mamüle İlerle (F4 + Aşağı Ok + Enter)
            self.advance_to_next_mamul()

        self.running = False
        self.update_status("TAMAMLANDI")
        self.log(f"🎉 İşlem Bitti. Bu turda {completed_in_this_run} mamül aktarıldı. Toplam {self.captured_count or self.current_idx} mamül Excel'de!")

    def _watch_loop_worker(self):
        self.log("=== PANO İZLEYİCİ BAŞLATILDI (AKILLI SATIR DESTEKLİ) ===")
        self.log("Siz programda gezerken tablodaki satıra (veya fiyata) Ctrl+C yaptıkça otomatik yakalanıp Excel'e yazılacaktır.")
        self.update_status("PANO DİNLENİYOR...")
        last_clip = get_clipboard_text().strip()

        while self.running:
            clip = get_clipboard_text().strip()
            if clip and clip != last_clip:
                last_clip = clip
                grid_data = parse_grid_row_clip(clip)
                if grid_data is not None:
                    self.captured_count += 1
                    sira, target_updated = self.save_grid_row_to_excel(grid_data)
                    target_info = " + Excel Güncellendi" if target_updated else ""
                    self.log(f"📋 Pano Algılandı [#{self.captured_count}]: {grid_data['raw_kod']} | {grid_data['ad']} ➔ {grid_data['fiyat_str']} ₺ (reçete sıra.xlsx{target_info})")
                    self.current_idx += 1
                    self.update_progress(self.current_idx, self.total_rows)
                else:
                    num = parse_tr_price(clip)
                    if num is not None and self.current_idx < len(self.mamul_list):
                        mamul = self.mamul_list[self.current_idx]
                        target_row = mamul["row"]
                        mk_code = mamul["mk"]
                        model_name = f"{mamul['model']} {mamul['ad']} ({mamul['renk']})"

                        success = self.save_price_to_excel(target_row, num, status="Aktarıldı (Pano)")
                        formatted_str = f"{num:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') + " ₺"
                        if success:
                            self.log(f"📋 Pano Algılandı [{self.current_idx + 1}/{self.total_rows}]: {mk_code} ➔ {formatted_str} (G{target_row})")
                        self.current_idx += 1
                        self.save_config()
                        self.update_progress(self.current_idx, self.total_rows)

            time.sleep(0.2)

        self.running = False
        self.update_status("TAMAMLANDI")


class ReceteFiyatApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Reçete Maliyeti Birim Fiyat Excel Aktarım Makrosu (v2.8 Hücre 38)")
        self.geometry("690x850")
        self.minsize(640, 780)
        self.attributes("-topmost", True)

        self.capturing_coord = None
        self.engine = ReceteFiyatEngine(ui_callback=None)

        self._build_ui()
        self.engine.ui_callback = self.on_engine_event

        self._start_hotkey_listener()
        self.engine.log(f"Program hazır. {len(self.engine.mamul_list)} mamül yüklendi: {os.path.basename(self.engine.excel_path)}")
        self.engine.update_progress(self.engine.current_idx, self.engine.total_rows)

    def _build_ui(self):
        bg_dark = "#18181b"
        fg_white = "#ffffff"
        card_bg = "#27272a"
        accent_blue = "#2563eb"
        accent_green = "#16a34a"
        accent_red = "#dc2626"
        accent_purple = "#9333ea"

        self.configure(bg=bg_dark)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", background=bg_dark, foreground=fg_white, font=("Segoe UI", 9))
        style.configure("Card.TFrame", background=card_bg, relief="flat")
        style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"), foreground="#38bdf8", background=card_bg)
        style.configure("Sub.TLabel", font=("Segoe UI", 8), foreground="#9ca3af", background=card_bg)

        # Başlık Kartı
        header_frame = ttk.Frame(self, style="Card.TFrame", padding=12)
        header_frame.pack(fill="x", padx=12, pady=(10, 4))

        title_lbl = ttk.Label(header_frame, text="⚡ Reçete Maliyeti ➔ Excel Akıllı Aktarım Makrosu (v3.0)", style="Header.TLabel")
        title_lbl.pack(anchor="w")
        sub_lbl = ttk.Label(header_frame, text="Kodun Altındaki Satırı Kopyalar ➔ Kod + İsim + Fiyatı Anında 'reçete sıra.xlsx'e Kaydeder (Sıfır Kayma)", style="Sub.TLabel")
        sub_lbl.pack(anchor="w", pady=(2, 0))

        # 0. Excel Dosya Seçim Kartı
        file_card = ttk.Frame(self, style="Card.TFrame", padding=10)
        file_card.pack(fill="x", padx=12, pady=4)

        f_title = ttk.Label(file_card, text="📁 Hedef Excel Dosyası (Otomatik Kayıt)", font=("Segoe UI", 10, "bold"), foreground="#38bdf8", background=card_bg)
        f_title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))

        self.var_excel_path = tk.StringVar(value=self.engine.excel_path)
        ent_file = tk.Entry(file_card, textvariable=self.var_excel_path, bg="#18181b", fg=fg_white, insertbackground=fg_white, width=42)
        ent_file.grid(row=1, column=0, sticky="ew", padx=(0, 6))

        btn_browse = tk.Button(file_card, text="📂 Gözat...", bg="#3f3f46", fg=fg_white, font=("Segoe UI", 8, "bold"), relief="flat", command=self.browse_excel)
        btn_browse.grid(row=1, column=1, sticky="e", padx=(0, 4))

        btn_open = tk.Button(file_card, text="📊 Excel'i Aç", bg="#16a34a", fg=fg_white, font=("Segoe UI", 8, "bold"), relief="flat", padx=6, command=self.open_excel_file)
        btn_open.grid(row=1, column=2, sticky="e")

        self.lbl_file_status = ttk.Label(file_card, text=f"✅ Durum: {self.engine.captured_count} Mamül Kayıtlı ({os.path.basename(self.engine.excel_path)})", font=("Segoe UI", 8, "bold"), foreground="#4ade80", background=card_bg)
        self.lbl_file_status.grid(row=2, column=0, columnspan=3, sticky="w", pady=(3, 0))

        # 1. Koordinat Kartı
        coord_card = ttk.Frame(self, style="Card.TFrame", padding=10)
        coord_card.pack(fill="x", padx=12, pady=4)

        c_header = ttk.Frame(coord_card, style="Card.TFrame")
        c_header.pack(fill="x", pady=(0, 6))
        c_title = ttk.Label(c_header, text="📍 1. Program Ekran Koordinatları (2 Nokta)", font=("Segoe UI", 10, "bold"), foreground="#38bdf8", background=card_bg)
        c_title.pack(side="left")

        btn_reset_coords = tk.Button(c_header, text="🔄 Fabrika Koordinatları (1194, 295 / 1556, 382)", bg="#3f3f46", fg="#93c5fd", font=("Segoe UI", 8), relief="flat", padx=6, command=self.reset_factory_coords)
        btn_reset_coords.pack(side="right")

        coord_grid = ttk.Frame(coord_card, style="Card.TFrame")
        coord_grid.pack(fill="x")

        # 1. Nokta: Mamul Seçici Dropdown Kutusu (F9)
        ttk.Label(coord_grid, text="1. Mamul Kutusu (Üstteki Açılır Liste):", background=card_bg).grid(row=0, column=0, sticky="w", pady=2)
        self.lbl_mamul_coord = ttk.Label(coord_grid, text=f"X: {self.engine.coord_mamul_box[0]}, Y: {self.engine.coord_mamul_box[1]}", font=("Segoe UI", 9, "bold"), foreground="#a78bfa", background=card_bg)
        self.lbl_mamul_coord.grid(row=0, column=1, padx=8, sticky="w", pady=2)
        btn_pick_mamul = tk.Button(coord_grid, text="📍 Mamul Seç (F9)", bg="#6366f1", fg=fg_white, font=("Segoe UI", 8, "bold"), relief="flat", padx=8, command=lambda: self.start_coord_picker("mamul"))
        btn_pick_mamul.grid(row=0, column=2, sticky="e", pady=2)

        # 2. Nokta: 1. Satır / Mamül Satırı ("Kodun Altı") (F8)
        ttk.Label(coord_grid, text="2. Tablo 1. Satırı ('Kodun Altı'):", background=card_bg).grid(row=1, column=0, sticky="w", pady=2)
        self.lbl_price_coord = ttk.Label(coord_grid, text=f"X: {self.engine.coord_price_cell[0]}, Y: {self.engine.coord_price_cell[1]}", font=("Segoe UI", 9, "bold"), foreground="#4ade80", background=card_bg)
        self.lbl_price_coord.grid(row=1, column=1, padx=8, sticky="w", pady=2)
        btn_pick_price = tk.Button(coord_grid, text="📍 Satır Seç (F8)", bg=accent_blue, fg=fg_white, font=("Segoe UI", 8, "bold"), relief="flat", padx=8, command=lambda: self.start_coord_picker("price"))
        btn_pick_price.grid(row=1, column=2, sticky="e", pady=2)

        # 2. İlerleme ve Parametre Kartı
        param_card = ttk.Frame(self, style="Card.TFrame", padding=10)
        param_card.pack(fill="x", padx=12, pady=4)

        p_title = ttk.Label(param_card, text="📊 2. İlerleme, İşlem Sayısı ve Seçenekler", font=("Segoe UI", 10, "bold"), foreground="#38bdf8", background=card_bg)
        p_title.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 4))

        # Progress Bar
        self.progress_bar = ttk.Progressbar(param_card, orient="horizontal", mode="determinate", length=400)
        self.progress_bar.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(0, 4))

        self.lbl_progress_text = ttk.Label(param_card, text=f"İlerleme: {self.engine.current_idx} / {self.engine.total_rows} Mamül", font=("Segoe UI", 9, "bold"), foreground="#fbbf24", background=card_bg)
        self.lbl_progress_text.grid(row=2, column=0, columnspan=2, sticky="w")

        btn_box = ttk.Frame(param_card, style="Card.TFrame")
        btn_box.grid(row=2, column=2, columnspan=2, sticky="e")

        btn_find_empty = tk.Button(btn_box, text="🎯 Excel'de Kaldığı Yeri Bul", bg="#0284c7", fg=fg_white, font=("Segoe UI", 8, "bold"), relief="flat", padx=4, command=self.find_and_set_empty_slot)
        btn_find_empty.pack(side="left", padx=2)

        btn_reset = tk.Button(btn_box, text="🔄 Başa Sar (Excel'i Sıfırla)", bg="#dc2626", fg=fg_white, font=("Segoe UI", 8, "bold"), relief="flat", padx=6, command=self.reset_to_start)
        btn_reset.pack(side="left", padx=2)

        # Satır 3: Bekleme ve Sıra No
        ttk.Label(param_card, text="Hesaplama Beklemesi (sn):", background=card_bg).grid(row=3, column=0, sticky="w", pady=(6, 0))
        self.var_wait_sec = tk.StringVar(value=str(self.engine.wait_after_down_sec))
        tk.Entry(param_card, textvariable=self.var_wait_sec, bg="#18181b", fg=fg_white, width=6, justify="center").grid(row=3, column=1, sticky="w", pady=(6, 0), padx=6)

        ttk.Label(param_card, text="Kaldığı Sıra No:", background=card_bg).grid(row=3, column=2, sticky="w", pady=(6, 0))
        self.var_current_idx = tk.StringVar(value=str(self.engine.current_idx + 1))
        ent_idx = tk.Entry(param_card, textvariable=self.var_current_idx, bg="#18181b", fg=fg_white, width=6, justify="center")
        ent_idx.grid(row=3, column=3, sticky="w", pady=(6, 0))
        ent_idx.bind("<Return>", lambda e: self.set_custom_idx())

        # Satır 4: İŞLEM / TEKRAR SAYISI
        ttk.Label(param_card, text="İşlem / Tekrar Sayısı:", background=card_bg, font=("Segoe UI", 9, "bold")).grid(row=4, column=0, sticky="w", pady=(8, 0))
        repeat_box = ttk.Frame(param_card, style="Card.TFrame")
        repeat_box.grid(row=4, column=1, columnspan=3, sticky="w", pady=(8, 0))

        self.var_target_count = tk.StringVar(value=str(getattr(self.engine, "target_repeat_count", 182)))
        ent_target = tk.Entry(repeat_box, textvariable=self.var_target_count, bg="#18181b", fg="#38bdf8", font=("Segoe UI", 9, "bold"), width=6, justify="center")
        ent_target.pack(side="left", padx=(0, 6))

        btn_20 = tk.Button(repeat_box, text="20", bg="#3f3f46", fg=fg_white, font=("Segoe UI", 8), relief="flat", padx=4, command=lambda: self.set_repeat_count(20))
        btn_20.pack(side="left", padx=2)

        btn_50 = tk.Button(repeat_box, text="50", bg="#3f3f46", fg=fg_white, font=("Segoe UI", 8), relief="flat", padx=4, command=lambda: self.set_repeat_count(50))
        btn_50.pack(side="left", padx=2)

        btn_182 = tk.Button(repeat_box, text="182 (Tam)", bg="#2563eb", fg=fg_white, font=("Segoe UI", 8, "bold"), relief="flat", padx=4, command=lambda: self.set_repeat_count(182))
        btn_182.pack(side="left", padx=2)

        btn_unlimited = tk.Button(repeat_box, text="∞ Sınırsız (0)", bg="#4b5563", fg="#a78bfa", font=("Segoe UI", 8), relief="flat", padx=4, command=lambda: self.set_repeat_count(0))
        btn_unlimited.pack(side="left", padx=2)

        # Satır 5: Enter Onayı Checkbox
        self.var_press_enter = tk.BooleanVar(value=self.engine.press_enter_after_down)
        chk_enter = tk.Checkbutton(param_card, text="Geçişte Enter Tuşuna Bas (Önerilen: Açık)", variable=self.var_press_enter, bg=card_bg, fg="#e2e8f0", selectcolor="#18181b", activebackground=card_bg, activeforeground=fg_white, font=("Segoe UI", 8), command=self.toggle_enter_setting)
        chk_enter.grid(row=5, column=0, columnspan=4, sticky="w", pady=(6, 0))

        # 3. Kontrol Butonları Kartı
        btn_card = ttk.Frame(self, style="Card.TFrame", padding=10)
        btn_card.pack(fill="x", padx=12, pady=4)

        self.btn_start_auto = tk.Button(btn_card, text="🚀 TAM OTOMATİK BAŞLAT (F6)", bg=accent_green, fg=fg_white, font=("Segoe UI", 10, "bold"), relief="flat", padx=10, pady=6, command=self.run_auto)
        self.btn_start_auto.pack(side="left", expand=True, fill="x", padx=3)

        self.btn_single_step = tk.Button(btn_card, text="⚡ TEK KAYIT AL (F10)", bg="#0284c7", fg=fg_white, font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=6, command=self.run_single_step)
        self.btn_single_step.pack(side="left", padx=3)

        self.btn_start_watch = tk.Button(btn_card, text="📋 Pano İzle", bg=accent_purple, fg=fg_white, font=("Segoe UI", 9, "bold"), relief="flat", padx=8, pady=6, command=self.run_watch)
        self.btn_start_watch.pack(side="left", padx=3)

        self.btn_pause = tk.Button(btn_card, text="⏸️ Duraklat (F7)", bg="#f59e0b", fg=fg_white, font=("Segoe UI", 9, "bold"), relief="flat", padx=8, pady=6, command=self.pause_action)
        self.btn_pause.pack(side="left", padx=3)

        self.btn_stop = tk.Button(btn_card, text="🛑 Durdur (ESC)", bg=accent_red, fg=fg_white, font=("Segoe UI", 9, "bold"), relief="flat", padx=8, pady=6, command=self.stop_action)
        self.btn_stop.pack(side="left", padx=3)

        # Durum Çubuğu
        self.lbl_status = tk.Label(self, text="HAZIR · İşlem sayısını belirleyip F6'ya basın (Çekilen her mamül anında Excel'e kaydedilir)", bg="#09090b", fg="#38bdf8", font=("Segoe UI", 9, "bold"), pady=6)
        self.lbl_status.pack(fill="x", padx=12, pady=(2, 2))

        # Canlı Log Kutusu
        log_card = ttk.Frame(self, style="Card.TFrame", padding=8)
        log_card.pack(fill="both", expand=True, padx=12, pady=(2, 10))

        l_title = ttk.Label(log_card, text="📜 Canlı Aktarım Günlüğü (Hücre 38 Adli Takip)", font=("Segoe UI", 9, "bold"), foreground="#9ca3af", background=card_bg)
        l_title.pack(anchor="w", pady=(0, 4))

        self.txt_log = tk.Text(log_card, bg="#09090b", fg="#e2e8f0", insertbackground=fg_white, font=("Consolas", 8), relief="flat")
        self.txt_log.pack(fill="both", expand=True)

    def set_repeat_count(self, count: int):
        self.var_target_count.set(str(count))
        self.engine.target_repeat_count = count
        self.engine.save_config()
        desc = f"{count} adet işlem" if count > 0 else "Sınırsız (0)"
        self.engine.log(f"İşlem / Tekrar Sayısı ayarlandı: {desc}")
        self.lbl_status.config(text=f"🎯 İşlem Sayısı: {desc} (F6 ile Başlatın)", fg="#38bdf8")

    def toggle_enter_setting(self):
        self.engine.press_enter_after_down = self.var_press_enter.get()
        self.engine.save_config()

    def reset_factory_coords(self):
        self.engine.coord_mamul_box = (1194, 295)
        self.engine.coord_price_cell = (1556, 382)
        self.lbl_mamul_coord.config(text=f"X: 1194, Y: 295")
        self.lbl_price_coord.config(text=f"X: 1556, Y: 382")
        self.engine.save_config()
        self.lbl_status.config(text="✅ Fabrika Koordinatları Yüklendi: Mamul=(1194, 295), Satır=(1556, 382)", fg="#4ade80")
        self.engine.log("Fabrika Koordinatları Yüklendi: Mamul=(1194, 295), Satır=(1556, 382)")

    def open_excel_file(self):
        target = self.engine.excel_path
        if not os.path.exists(target):
            create_empty_live_excel_file(target)
        try:
            os.startfile(target)
            self.engine.log(f"Excel dosyası açıldı: {os.path.basename(target)}")
        except Exception as e:
            self.engine.log(f"Excel açma hatası: {e}")

    def browse_excel(self):
        f = filedialog.askopenfilename(title="Hedef Excel Dosyasını Seç", filetypes=[("Excel Dosyaları", "*.xlsx *.xls")])
        if f:
            self.var_excel_path.set(f)
            self.engine.load_excel_manifest(f)
            self.lbl_file_status.config(text=f"✅ Durum: {len(self.engine.mamul_list)} Mamül Hazır ({os.path.basename(f)})", foreground="#4ade80")
            self.engine.update_progress(self.engine.current_idx, self.engine.total_rows)
            self.engine.save_config()

    def reset_to_start(self):
        self.engine.reset_all_and_clear_excel()
        self.var_current_idx.set("1")
        self.progress_bar["value"] = 0
        self.lbl_progress_text.config(text="İlerleme: 0 Mamül (Excel Sıfırlandı)")
        self.lbl_status.config(text="🔄 EXCEL SIFIRLANDI · Sıfırdan 1. Mamülle Başlamaya Hazır (F6)", fg="#4ade80")
        if hasattr(self, 'btn_start_auto'):
            self.btn_start_auto.config(text="🚀 TAM OTOMATİK BAŞLAT (F6)")
        self.lbl_file_status.config(text=f"✅ Durum: Excel Sıfırlandı (0 Kayıt) · {os.path.basename(self.engine.excel_path)}", foreground="#4ade80")

    def set_custom_idx(self):
        try:
            val = int(self.var_current_idx.get()) - 1
            if val >= 0:
                self.engine.current_idx = val
                self.engine.save_config()
                self.engine.update_progress(val, max(val + 1, self.engine.total_rows))
                self.engine.log(f"İlerleme {val + 1}. mamüle ayarlandı.")
        except Exception:
            pass

    def start_coord_picker(self, target: str):
        self.capturing_coord = target
        prompt_txt = "MAMÜL DROPDOWN KUTUSUNA" if target == "mamul" else "KODUN ALTINDAKİ 1. SATIRA"
        self.lbl_status.config(text=f"📍 FAREYİ {prompt_txt} GÖTÜRÜP 2 SANİYE BEKLEYİN...", fg="#facc15")
        threading.Thread(target=self._coord_picker_worker, args=(target,), daemon=True).start()

    def _coord_picker_worker(self, target: str):
        time.sleep(2.0)
        pt = wt.POINT()
        user32.GetCursorPos(ctypes.byref(pt))
        x, y = pt.x, pt.y

        if target == "price":
            self.engine.coord_price_cell = (x, y)
            self.lbl_price_coord.config(text=f"X: {x}, Y: {y}")
            self.engine.log(f"Mamül Satırı Koordinatı ('Kodun Altı') Kaydedildi: X={x}, Y={y}")
        elif target == "mamul":
            self.engine.coord_mamul_box = (x, y)
            self.lbl_mamul_coord.config(text=f"X: {x}, Y: {y}")
            self.engine.log(f"Mamul Seçici Koordinatı Kaydedildi: X={x}, Y={y}")

        self.engine.save_config()
        t_name = "MAMÜL KUTUSU (Dropdown)" if target == "mamul" else "MAMÜL SATIRI (Kodun Altı)"
        self.lbl_status.config(text=f"✅ {t_name} Koordinatı Alındı: ({x}, {y})", fg="#4ade80")

        # Mantık Denetimi (Ters Seçim Koruması)
        mx, my = self.engine.coord_mamul_box
        px, py = self.engine.coord_price_cell
        if my >= py:
            warn_msg = f"⚠️ DİKKAT: Mamul Kutusu (Y={my}), Tablo Satırından (Y={py}) daha aşağıda! F9 ve F8'i ters seçmiş olabilirsiniz."
            self.engine.log(warn_msg)
            self.lbl_status.config(text=warn_msg, fg="#f87171")

        self.capturing_coord = None

    def find_and_set_empty_slot(self):
        slot = self.engine.find_first_empty_excel_row()
        self.engine.current_idx = slot
        self.engine.captured_count = slot
        self.engine.save_config()
        self.var_current_idx.set(str(slot + 1))
        self.engine.update_progress(slot, max(slot + 1, self.engine.total_rows))
        if slot > 0:
            self.engine.log(f"🎯 Excel tarandı: Excel'de kayıtlı {slot} mamül var. Kaldığı sıra: {slot + 1}.")
            self.lbl_status.config(text=f"🎯 Excel'de {slot} mamül kayıtlı · Kaldığı yer: {slot + 1}. Mamül (F6 ile devam edin)", fg="#38bdf8")
            if hasattr(self, 'btn_start_auto'):
                self.btn_start_auto.config(text=f"🚀 KALDIĞI YERDEN DEVAM ET (F6) [{slot + 1}. Mamül]")
        else:
            self.engine.log("🎯 Excel tarandı: Excel dosyası tamamen boş (0 mamül). 1. sıradan başlanacak.")
            self.lbl_status.config(text="🎯 Excel boş · 1. Mamülden başlayacak (F6)", fg="#4ade80")
            if hasattr(self, 'btn_start_auto'):
                self.btn_start_auto.config(text="🚀 TAM OTOMATİK BAŞLAT (F6)")

    def run_auto(self):
        try:
            self.engine.wait_after_down_sec = float(self.var_wait_sec.get())
        except Exception:
            pass
        try:
            custom_idx = int(self.var_current_idx.get()) - 1
            if custom_idx >= 0:
                self.engine.current_idx = custom_idx
        except Exception:
            pass
        try:
            target_cnt = int(self.var_target_count.get())
            self.engine.target_repeat_count = max(0, target_cnt)
        except Exception:
            self.engine.target_repeat_count = 182

        self.engine.press_enter_after_down = self.var_press_enter.get()
        if hasattr(self, 'btn_start_auto'):
            target_lbl = f"{self.engine.target_repeat_count} Adet" if self.engine.target_repeat_count > 0 else "Sınırsız"
            self.btn_start_auto.config(text=f"🚀 ÇALIŞIYOR (Hedef: {target_lbl})...")
        self.engine.start_auto_loop()

    def run_single_step(self):
        threading.Thread(target=self.engine.copy_and_record_current, daemon=True).start()

    def run_watch(self):
        self.engine.start_watch_mode()

    def pause_action(self):
        if self.engine.running:
            if not self.engine.paused:
                self.engine.pause()
            else:
                self.engine.resume()

    def stop_action(self):
        self.engine.stop()
        if hasattr(self, 'btn_start_auto'):
            next_num = self.engine.current_idx + 1
            self.btn_start_auto.config(text=f"🚀 KALDIĞI YERDEN DEVAM ET (F6) [{next_num}. Mamül]")

    def on_engine_event(self, ev_type: str, data):
        if ev_type == "log":
            if hasattr(self, "txt_log") and self.txt_log:
                self.txt_log.insert("end", f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {data}\n")
                self.txt_log.see("end")
        elif ev_type == "status":
            if hasattr(self, "lbl_status") and self.lbl_status:
                self.lbl_status.config(text=data)
        elif ev_type == "progress":
            if hasattr(self, "progress_bar") and self.progress_bar:
                curr, tot = data
                self.progress_bar["maximum"] = max(tot, curr, 1)
                self.progress_bar["value"] = curr
            if hasattr(self, "lbl_progress_text") and self.lbl_progress_text:
                if curr <= tot:
                    self.lbl_progress_text.config(text=f"İlerleme: {curr} / {tot} Mamül")
                else:
                    self.lbl_progress_text.config(text=f"İlerleme: {curr}. Mamül (Sınır Yok · Canlı Akış)")
            if hasattr(self, "var_current_idx") and self.var_current_idx:
                self.var_current_idx.set(str(curr + 1))

    def _start_hotkey_listener(self):
        def listener():
            while True:
                # F6: Başlat / Devam Et
                if user32.GetAsyncKeyState(VK_F6) & 0x8000:
                    if not self.engine.running:
                        self.after(10, self.run_auto)
                    elif self.engine.paused:
                        self.after(10, self.engine.resume)
                    time.sleep(0.3)

                # F7: Duraklat
                if user32.GetAsyncKeyState(VK_F7) & 0x8000:
                    self.after(10, self.pause_action)
                    time.sleep(0.3)

                # F8: Satır Koordinatı Al
                if user32.GetAsyncKeyState(VK_F8) & 0x8000:
                    self.after(10, lambda: self.start_coord_picker("price"))
                    time.sleep(0.3)

                # F9: Mamul Kutusu Koordinatı Al
                if user32.GetAsyncKeyState(VK_F9) & 0x8000:
                    self.after(10, lambda: self.start_coord_picker("mamul"))
                    time.sleep(0.3)

                # F10: Tek Kayıt Al (Manuel Geçişte Tek Tuş)
                if user32.GetAsyncKeyState(VK_F10) & 0x8000:
                    self.after(10, self.run_single_step)
                    time.sleep(0.3)

                # ESC: Acil Durdurma
                if user32.GetAsyncKeyState(VK_ESCAPE) & 0x8000:
                    if self.engine.running:
                        self.after(10, self.stop_action)
                    time.sleep(0.3)

                time.sleep(0.05)

        t = threading.Thread(target=listener, daemon=True)
        t.start()


if __name__ == "__main__":
    app = ReceteFiyatApp()
    app.mainloop()
