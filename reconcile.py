"""
reconcile.py
Mesin rekonsiliasi antar rekening untuk Stoa Space.

Input : workbook xlsx berisi 1 sheet per rekening (BRI, BCA, Jago, Kasir, dst)
        dengan kolom baku: Tanggal | Keterangan Transaksi | Kategori Transaksi |
        Debit | Kredit | Saldo Kumulatif | Subjek Transaksi | Objek Transaksi |
        Keterangan Tambahan

Output: workbook baru dengan:
        - Setiap sheet rekening asli disalin apa adanya, plus kolom bantu
          "Nominal Bersih" (formula, bukan angka mati)
        - Sheet "Rekonsiliasi" berisi hasil pencocokan transfer antar rekening
          dan daftar minus/selisih yang perlu verifikasi
        - Sheet "Laporan Laba Rugi" (Income Statement)
        - Sheet "Neraca" (Balance Sheet)
        - Sheet "Laporan Arus Kas" (Cash Flow Statement)

Semua angka di sheet laporan dibuat dengan rumus Excel beralamat absolut
($Kolom$Baris), merujuk langsung ke sheet rekening. Tidak ada angka hasil
kalkulasi Python yang ditulis sebagai nilai mati kecuali memang tidak
mungkin direpresentasikan sebagai rumus (contoh: catatan naratif audit).
"""

import re
import copy
from collections import Counter
import shared_rules
import datetime
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from dataclasses import dataclass, field

# ---------------------------------------------------------------------------
# Konfigurasi
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SECTION_FILL = PatternFill("solid", fgColor="E5E7EB")
SECTION_FONT = Font(bold=True)
HIGH_FILL = PatternFill("solid", fgColor="C6EFCE")
MED_FILL = PatternFill("solid", fgColor="FFEB9C")
LOW_FILL = PatternFill("solid", fgColor="FFC7CE")
TRANSFER_MATCH_FILL = PatternFill("solid", fgColor="BDD7EE")
DATE_FORMAT = "d-mmm-yy"
NUMBER_FORMAT = "#,##0"
THIN = Side(style="thin", color="D1D5DB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Kategori yang menandakan perpindahan uang ANTAR rekening sendiri
# (bukan pendapatan/beban riil) -> tidak masuk Laba Rugi, harus saling
# menutup nol di Rekonsiliasi.
TRANSFER_KEYWORDS = shared_rules.get("transfer_keywords", [
    "pindah rekening internal",
    "pindang rekening internal",
    "transfer internal",
    "transfer lainnya",
    "transaksi internal",
    "pemindahbukuan",
])

# Kata kunci di KETERANGAN (bukan kategori) yang juga menandakan transfer
# antar rekening sendiri, buat jaga-jaga kalau kategorinya salah/tidak
# konsisten diisi di sumber data (mis. "Setoran Tunai" kadang tidak diberi
# kategori "Pindah/Transfer Internal" padahal itu kasir->bank sendiri).
# Tidak dipakai kalau kategorinya sudah eksplisit "Modal & Setoran Pemilik"
# (setoran modal dari luar, bukan pindah antar rekening sendiri).
DESC_TRANSFER_KEYWORDS = shared_rules.get("desc_transfer_keywords", [
    "setoran tunai",
    "setor tunai",
    "setoran via cdm",
    "pindah rekening",
    "transfer internal",
])

# Kategori setoran/penarikan modal pemilik -> masuk Neraca (ekuitas),
# bukan Laba Rugi.
CAPITAL_KEYWORDS = shared_rules.get("capital_keywords", [
    "modal & setoran pemilik",
    "modal dan setoran pemilik",
    "laba ditahan bulanan",
])

# Pengeluaran pribadi owner - DIKELUARKAN dari Laba Rugi (bukan beban
# bisnis) di SEMUA level (bulanan, kuartal, tahunan), diperlakukan
# seperti prive/penarikan modal (mengurangi Ekuitas, BUKAN Beban) supaya
# Neraca tetap balance tanpa menganggu Laba Rugi bisnis. Setiap transaksi
# kategori ini akan diverifikasi manual (lihat sheet Rekonsiliasi bagian
# 5) - user menegaskan ini perlu diaudit satu-satu, bukan otomatis
# dipercaya begitu saja.
PERSONAL_EXPENSE_KEYWORDS = shared_rules.get("personal_expense_keywords", [
    "pengeluaran pribadi", "keperluan pribadi", "kepentingan pribadi", "milik pribadi",
])

# Nama depan pegawai yang DIKETAHUI dipakai lebih dari satu orang (mis.
# "Baiq" - gelar/awalan umum, dipakai oleh "Baiq Sabrina Ameli" DAN
# "Baiq Widiani Rintis Sari" di EMPLOYEE_ALIASES quarterly.py) - kalau
# nama depan hasil ekstraksi dari Objek/Keterangan transaksi Gaji
# persis salah satu dari daftar ini, TIDAK BOLEH dianggap identitas
# pasti (lihat _gaji_rekon_lokal_info) - datanya sendiri (Objek cuma
# "Baiq" tanpa nama belakang) genuinely tidak cukup untuk membedakan,
# menebak dari nominal beresiko salah dan terlihat pasti padahal cuma
# tebakan. Ditandai eksplisit "(?)" untuk verifikasi manual, BUKAN
# dipaksakan ke salah satu nama lengkap.
AMBIGUOUS_FIRST_NAMES = shared_rules.get("ambiguous_first_names", ["baiq"])

# Alias pemilik (owner) - kalau Subjek DAN Objek transaksi SAMA-SAMA
# salah satu dari daftar ini, transaksi ini kemungkinan besar transfer
# antar rekening pribadi/bisnis owner sendiri (mis. owner transfer dari
# rekening pribadinya ke rekening bisnis) - diperlakukan sebagai
# kandidat Transaksi Internal (lihat Txn.is_transfer), lalu tetap lewat
# pencocokan nominal+tanggal normal (find_matches) seperti transfer
# lain - bukan otomatis "pasti benar" tanpa verifikasi.
OWNER_ALIASES = shared_rules.get("owner_aliases", [
    "owner", "ojan", "kak ojan", "ozan", "pakdok", "roziyan", "ahmad roziyan hidayat",
    "ahmad roziyan", "ahmad roziya",
])

# Alias pegawai+owner -> nama lengkap (sumber SAMA dengan quarterly.py,
# lewat shared_rules.json - dict {alias huruf kecil: "Nama Lengkap"}).
# Dipakai untuk melengkapi Objek jadi nama lengkap Title Case kalau
# dikenali (lihat pass "lengkapi nama Objek" di run_rekon_lokal).
# Default lengkap disertakan (BUKAN dict kosong) supaya tetap berfungsi
# kalau shared_rules.json kebetulan belum ada/kosong isinya.
_DEFAULT_EMPLOYEE_ALIASES = {
    "ahmad roziyan hidayat": "Ahmad Roziyan Hidayat", "ahmad roziyan h.": "Ahmad Roziyan Hidayat",
    "ahmad roziyan": "Ahmad Roziyan Hidayat", "roziyan hidayat": "Ahmad Roziyan Hidayat",
    # "ahmad roziya" (tanpa "n") - alias/typo owner yang sering muncul di
    # data bank mentah (mis. Subjek "AHMAD ROZIYA"/"Ahmad Roziya"). BEDA
    # dengan "ahmad rizan" (lihat CATEGORY_OVERRIDE_RULES) - nama itu
    # sengaja TIDAK dianggap owner, dipastikan Modal Masuk murni.
    "ahmad roziya": "Ahmad Roziyan Hidayat",
    "roziyan": "Ahmad Roziyan Hidayat", "ojan": "Ahmad Roziyan Hidayat",
    "kak ojan": "Ahmad Roziyan Hidayat", "ozan": "Ahmad Roziyan Hidayat",
    "pakdok": "Ahmad Roziyan Hidayat", "owner": "Ahmad Roziyan Hidayat",
    "lusiana valufi n": "Lusiana Valufi N", "lusiana": "Lusiana Valufi N",
    "upi": "Lusiana Valufi N", "upiw": "Lusiana Valufi N",
    "alun ayumi fiqha p": "Alun Ayumi Fiqha P", "alun": "Alun Ayumi Fiqha P", "luna": "Alun Ayumi Fiqha P",
    "baiq sabrina ameli": "Baiq Sabrina Ameli", "sabrina": "Baiq Sabrina Ameli",
    "amel": "Baiq Sabrina Ameli", "amelia": "Baiq Sabrina Ameli",
    "viona winda octavia": "Viona Winda Octavia", "viona": "Viona Winda Octavia",
    "vio": "Viona Winda Octavia", "vivi": "Viona Winda Octavia", "ivi": "Viona Winda Octavia",
    "ismayanti": "Ismayanti", "maya": "Ismayanti",
    "adinda nurusshafwa": "Adinda Nurusshafwa", "dinda": "Adinda Nurusshafwa",
    "baiq widiani rintis sari": "Baiq Widiani Rintis Sari", "baiq widiani rinti": "Baiq Widiani Rintis Sari",
    "widia": "Baiq Widiani Rintis Sari", "sari": "Baiq Widiani Rintis Sari",
    "kurnia utami nur": "Kurnia Utami Nur", "kurnia": "Kurnia Utami Nur",
    "kur": "Kurnia Utami Nur", "puput": "Kurnia Utami Nur",
    "oriegia shativa mulia": "Oriegia Shativa Mulyadi", "oriegia shativa mulyadi": "Oriegia Shativa Mulyadi",
    "oriegia shativa mu": "Oriegia Shativa Mulyadi", "gia": "Oriegia Shativa Mulyadi", "origia": "Oriegia Shativa Mulyadi",
    "naura lutfia": "Naura Lutfia", "naura": "Naura Lutfia", "ola": "Naura Lutfia",
    "mila septiana": "Mila Septiana", "mila": "Mila Septiana",
    "ridiaton rizki": "Ridiaton Rizki", "kik": "Ridiaton Rizki",
    "royyan paice sangwenang": "Royyan Paice Sangwenang", "royyan": "Royyan Paice Sangwenang",
    "irma aprianti": "Irma Aprianti", "irma": "Irma Aprianti",
    "latifatul husna": "Latifatul Husna", "eva": "Latifatul Husna",
    "panji anjanis pran": "Panji Anjanis Pran", "panji": "Panji Anjanis Pran", "pandji": "Panji Anjanis Pran",
}
EMPLOYEE_ALIASES = shared_rules.get("employee_aliases", _DEFAULT_EMPLOYEE_ALIASES) or _DEFAULT_EMPLOYEE_ALIASES

# Deklarasi HUTANG BARU yang masuk lewat transfer bank - kata kunci ini
# TIDAK mengubah kategori (transaksi "hutang"/"pinjaman" + arah masuk
# SUDAH otomatis jadi Modal & Setoran Pemilik lewat aturan yang ada di
# CATEGORY_OVERRIDE_RULES), tujuannya CUMA menandai transaksi ini untuk
# daftar audit terpisah (Rekonsiliasi bagian 6) - supaya user diingatkan
# menambahkan baris baru di Buku Hutang (laporan kuartal/tahunan) tanpa
# harus scroll manual cari sendiri transaksi mana yang perlu didaftarkan.
NEW_DEBT_KEYWORDS = shared_rules.get("new_debt_keywords", [
    "hutang baru", "pinjaman baru", "terima pinjaman", "terima hutang",
    "pinjaman cair", "hutang cair", "pencairan pinjaman", "pencairan hutang",
])

# "Transfer Masuk" sendirian terlalu umum untuk langsung dianggap Modal
# (transfer masuk dari pelanggan/pihak luar seharusnya Penjualan, bukan
# Modal) - jadi HANYA dianggap setara Modal & Setoran Pemilik kalau
# keterangannya eksplisit bilang "dari rekening sendiri" (uang milik
# owner sendiri yang dipindah antar rekening, mis. pencairan investasi
# pribadi yang disetor ke rekening bisnis), sesuai kasus nyata yang
# ditemukan: "BI-Fast Transfer Masuk ... (dari rekening sendiri)".
CAPITAL_SELF_TRANSFER_KEYWORDS = shared_rules.get("capital_self_transfer_keywords", ["dari rekening sendiri", "rekening sendiri"])

# Aturan override kategori berdasarkan kata kunci di Keterangan/Keterangan
# Tambahan/Objek, ditegaskan langsung oleh user berdasarkan pengalaman
# nyata bisnisnya - dicek SEBELUM aturan-aturan lain (is_transfer, dst),
# dan MENGGANTIKAN Kategori asli untuk semua perhitungan (lihat
# Txn.effective_kategori). Urutan penting: yang pertama cocok menang.
# "any": salah satu kata kunci cukup. "all": semua kata kunci harus ada
# (tidak harus berdekatan). "sheet_contains": opsional, cuma berlaku
# kalau nama sheet mengandung teks ini (case-insensitive).
_DEFAULT_CATEGORY_OVERRIDE_RULES = [
    {"any": ["pengeluaran pribadi", "keperluan pribadi", "kepentingan pribadi", "milik pribadi"],
     "category": "Pengeluaran Pribadi", "sheet_contains": None},
    # "Ahmad Rizan" SENGAJA dibedakan dari owner ("Ahmad Roziyan
    # Hidayat"/alias "Ahmad Roziya") - nama beda orang, dipastikan Modal
    # Masuk murni (modal/investasi masuk), BUKAN transfer internal owner.
    {"any": ["ahmad rizan"], "category": "Modal & Setoran Pemilik", "sheet_contains": None},
    {"all": ["briva", "tokopedia"], "amount_min": 900000, "amount_max": 1100000,
     "category": "Overhead", "sheet_contains": None},
    {"any": ["tokopedia"], "category": "Belanja Bahan", "sheet_contains": None},
    {"all": ["cashback", "qris"], "category": "Biaya Admin Bank", "sheet_contains": None},
    {"any": ["cashback mdr"], "category": "Biaya Admin Bank", "sheet_contains": None},
    {"any": ["cashback jago"], "category": "Biaya Admin Bank", "sheet_contains": "jago"},
    {"any": ["dr koreksi bunga", "cr koreksi bunga", "koreksi bunga"], "category": "Biaya Admin Bank", "sheet_contains": None},
    {"any": ["interest on account"], "category": "Biaya Admin Bank", "sheet_contains": None},
    {"any": ["layanan"], "category": "Overhead", "sheet_contains": "jago"},
    {"any": ["fb", "facebook", "meta ads"], "category": "Marketing", "sheet_contains": None},
    {"any": ["sponsorship", "charity", "donasi"], "category": "Marketing", "sheet_contains": None},
    {"any": ["masuya graha trikencana"], "category": "Belanja Bahan", "sheet_contains": None},
    {"any": ["masuya graha trike"], "category": "Belanja Bahan", "sheet_contains": None},
    {"any": ["sukanda", "dineta"], "category": "Belanja Bahan", "sheet_contains": None},
    {"any": ["muh yani sh", "muh. yani sh", "muhammad yani sh"], "category": "Pembayaran Hutang", "sheet_contains": None},
    {"any": ["sahabudin"], "category": "Overhead", "sheet_contains": None},
    {"any": ["modal & setoran pemilik", "modal dan setoran pemilik"], "category": "Modal & Setoran Pemilik", "sheet_contains": None},
    {"any": ["hutang", "pinjaman"], "none_of": ["bayar hutang", "bayar pinjaman", "cicilan hutang", "cicilan pinjaman"],
     "direction": "masuk", "category": "Hutang Masuk", "sheet_contains": None},
    {"any": ["hutang", "pinjaman"], "direction": "keluar", "category": "Pembayaran Hutang", "sheet_contains": None},
    {"any": ["setoran via cdm"], "category": "Transaksi Internal", "sheet_contains": None},
    {"any": ["pemindahbukuan", "transfer internal"], "category": "Transaksi Internal", "sheet_contains": None},
    {"any": ["visionet"], "category": "Penjualan Grabfood", "sheet_contains": None},
    {"any": ["tarik tunai qris"], "category": "Penjualan", "sheet_contains": None},
    {"any": ["tarik tunai"], "category": "Penjualan", "sheet_contains": "kas"},
    {"any": ["visionet"], "category": "Penjualan", "sheet_contains": None},
    # Sewa dan Mantenantce Bangunan (nama kategori SENGAJA ejaan ini,
    # sesuai kontrak kategori v3 dari bot konversi) - dicek SEBELUM aturan
    # Belanja Utilitas/Overhead yang lebih generik, karena kata kunci di
    # sini (Kabel/Listrik, dst) sering tumpang tindih dengan utilitas
    # umum - urutan menang duluan penting.
    {"any": ["sewa bangunan", "renovasi bangunan", "biaya tukang", "ongkos tukang", "bayar tukang",
             "bahan bangunan",
             "renovasi kabel", "kabel", "lampu", "toren", "besi", "keramik", "pipa", "westafel",
             "wc", "keran", "depo bangunan", "mitra 10", "toko bangunan"],
     "category": "Sewa dan Maintenance Bangunan", "sheet_contains": None},
    {"any": ["sewa", "utilitas", "web", "spotify"],
     "category": "Overhead", "sheet_contains": None},
    {"any": ["parkir", "penyetakan", "stiker", "sticker", "print", "cetak", "sablon"],
     "category": "Overhead", "sheet_contains": None},
    {"any": ["pulsa", "my telkomsel", "pulsa simpati", "telkomsel", "telkom", "air pdam", "pdam",
             "listrik", "pln"],
     "category": "Belanja Utilitas", "sheet_contains": None},
    {"any": ["konsumsi"], "category": "Konsumsi dan Liburan", "sheet_contains": None},
    {"any": ["belanja tools", "tools", "cutleries", "mr diy"], "category": "Tools dan Equipments", "sheet_contains": None},
    {"any": ["seakun.id", "apple", "adobe"], "category": "Subscription", "sheet_contains": None},
    {"any": ["riset", "pelatihan", "training"], "category": "Riset dan Development", "sheet_contains": None},
    {"any": ["plastik"], "category": "Kemasan", "sheet_contains": None},
    {"any": ["nanda audia agustin", "nanda audia agusti"], "category": "Kemasan", "sheet_contains": None},
    {"any": ["madam baha", "madam bahan kue", "toko madam"], "category": "Belanja Bahan", "sheet_contains": None},
    {"any": ["yulia indah pratiwi", "yulia indah pratiw", "anugerah plastik"], "category": "Kemasan", "sheet_contains": None},
    {"any": ["beli masker", "shopee", "ovo", "gopay", "dana", "top up", "isi saldo", "tarikan atm",
             "ganti uang belanja", "es batu"],
     "category": "Overhead", "sheet_contains": None},
    {"any": ["sisa belanja", "sisa set"], "category": "Overhead", "sheet_contains": None},
    {"any": ["tukang", "reparasi", "service ac", "service mesin", "perbaikan ac", "perbaikan mesin",
             "perbaikan bangunan", "perbaiki ac", "perbaiki mesin", "uang ac", "maintenance"],
     "category": "Reparasi dan Maintenance Tools dan Mesin", "sheet_contains": None},
    {"any": ["minus", "lebih", "cust", "tip", "tips"], "category": "Tip/Minus/Lebih", "sheet_contains": None},
    # Jaring pengaman terakhir: SEMUA transaksi yang Kategori ASLI-nya
    # (bukan Keterangan) masih literal kategori LAMA (sebelum di-rename/
    # digabung sesuai kontrak kategori v3) dan tidak ketangkap aturan
    # spesifik manapun di atas - default ke kategori BARU sesuai
    # penegasan user, bukan dibiarkan jatuh ke Kategori Baru cuma
    # karena tidak ada kata kunci lain yang cocok. Ditaruh PALING
    # TERAKHIR (kalah prioritas dari SEMUA aturan kata kunci spesifik
    # di atas).
    {"kategori_asli": "belanja operasional", "category": "Overhead", "sheet_contains": None},
    {"kategori_asli": "belanja konsumsi", "category": "Konsumsi dan Liburan", "sheet_contains": None},
    {"kategori_asli": "reparasi dan maintenance", "category": "Reparasi dan Maintenance Tools dan Mesin", "sheet_contains": None},
    {"kategori_asli": "reparasi", "category": "Reparasi dan Maintenance Tools dan Mesin", "sheet_contains": None},
    {"kategori_asli": "pajak daerah", "category": "Pajak dan Administrasi", "sheet_contains": None},
    {"kategori_asli": "biaya administrasi", "category": "Pajak dan Administrasi", "sheet_contains": None},
    {"kategori_asli": "administrasi", "category": "Pajak dan Administrasi", "sheet_contains": None},
    {"kategori_asli": "biaya renovasi atap", "category": "Sewa dan Maintenance Bangunan", "sheet_contains": None},
    {"kategori_asli": "renovasi bangunan", "category": "Sewa dan Maintenance Bangunan", "sheet_contains": None},
    {"kategori_asli": "renovasi", "category": "Sewa dan Maintenance Bangunan", "sheet_contains": None},
]
# Dimuat dari shared_rules.json (dipakai bersama reconbot & bank-statement-bot)
# kalau ada; kalau file/kunci tidak ada, pakai daftar default di atas.


def _build_transfer_masuk_rules():
    """Bangun 3 aturan 'Transfer Masuk' yang ambigu (Kategori/Keterangan
    menyebut 'transfer masuk' tapi tidak jelas dari siapa), berdasarkan
    daftar alias pegawai/owner yang sedang aktif (shared_rules ->
    employee_aliases). User menegaskan urutan logika: (1) dari owner ->
    Modal & Setoran Pemilik, (2) dari pegawai (bukan owner) -> Belanja
    Operasional dengan nilai positif (sisa belanja/reimbursement), (3)
    bukan keduanya DAN nominal di bawah Rp300.000 -> Penjualan (kemungkinan
    besar pembayaran pelanggan kecil via QRIS/transfer). Nominal besar
    yang bukan owner/pegawai TIDAK di-assign otomatis - tetap jadi
    'Kategori Baru' untuk diaudit manual, sesuai penegasan user."""
    aliases = shared_rules.get("employee_aliases", {})
    owner_keywords = sorted({k for k, v in aliases.items() if v == "Ahmad Roziyan Hidayat"})
    employee_keywords = sorted({k for k, v in aliases.items() if v != "Ahmad Roziyan Hidayat"})
    rules = []
    if owner_keywords:
        rules.append({"all": ["transfer masuk"], "any": owner_keywords,
                       "category": "Modal & Setoran Pemilik", "sheet_contains": None})
    if employee_keywords:
        rules.append({"all": ["transfer masuk"], "any": employee_keywords,
                       "category": "Overhead", "sheet_contains": None})
    rules.append({"all": ["transfer masuk"], "none_of": owner_keywords + employee_keywords,
                  "amount_max": 300000, "category": "Penjualan", "sheet_contains": None})
    # Fallback KHUSUS untuk "transfer masuk" yang TIDAK match salah satu
    # dari 3 aturan di atas (bukan owner, bukan pegawai, nominal >=
    # Rp300rb) - paksa jadi 'Kategori Baru' supaya benar-benar diaudit.
    # Tanpa ini, kategori ASLI (mis. 'Transfer Lainnya'/'Transaksi
    # Internal') akan dianggap "sudah dikenal" oleh _is_recognized_category
    # (karena memang salah satu TRANSFER_KEYWORDS) dan lolos begitu saja
    # tanpa audit, padahal justru inilah yang paling perlu diaudit.
    rules.append({"all": ["transfer masuk"], "category": "Kategori Baru", "sheet_contains": None})
    return rules


_STATIC_CATEGORY_OVERRIDE_RULES = shared_rules.get("category_override_rules", _DEFAULT_CATEGORY_OVERRIDE_RULES)
CATEGORY_OVERRIDE_RULES = _STATIC_CATEGORY_OVERRIDE_RULES + _build_transfer_masuk_rules()

_PROTECTED_FROM_CATEGORY_OVERRIDE = set(shared_rules.get("protected_from_category_override", [
    "modal & setoran pemilik", "modal dan setoran pemilik", "laba ditahan bulanan",
    "saldo awal", "saldo awal bulan", "modal", "pengeluaran pribadi",
    "hutang masuk", "pembayaran hutang",
    # Kategori Layer 1 SPESIFIK dari bot konversi (kontrak v3) - kalau
    # Kategori ASLI SUDAH salah satu dari ini, PERCAYAI apa adanya,
    # JANGAN dicoba ditimpa aturan kata kunci generik lain. Tanpa ini,
    # kategori seperti "Sewa dan Maintenance Bangunan" bisa keliru
    # ketimpa aturan "sewa"->Overhead yang lebih umum, cuma karena kata
    # "sewa" ikut muncul di teks gabungan (termasuk Kategori aslinya
    # sendiri) tanpa ada kata kunci bangunan spesifik lain di Keterangan.
    "belanja bahan", "overhead", "konsumsi dan liburan", "belanja utilitas",
    "tools dan equipments", "kemasan", "subscription", "sewa dan maintenance bangunan",
    "reparasi dan maintenance tools dan mesin", "pajak dan administrasi", "belanja assets",
    "penjualan grabfood", "penjualan shopeefood",
    # Kategori biaya jasa perbankan - transaksi ini SENDIRI adalah biaya
    # admin/bunga bank, terlepas dari APA yang disebut di catatan
    # referensinya (mis. "Biaya terkait transaksi: Top Up Gopay ..." -
    # itu keterangan TENTANG transaksi apa yang memicu biaya ini, BUKAN
    # berarti biayanya sendiri harus dikategorikan sebagai belanja/top up
    # itu). Tanpa proteksi ini, kata kunci seperti "gopay"/"top up" di
    # catatan referensi bisa keliru menimpa kategori biaya bank yang
    # sudah benar jadi kategori lain (mis. Overhead).
    "biaya admin bank", "biaya admin & pajak bank", "biaya admin dan bunga bank",
    "bunga dan admin bank",
]))


def _override_keyword_found(pattern, text):
    return re.search(r"\b" + re.escape(pattern) + r"\b", text) is not None


# Sebagian bank menulis Keterangan Transaksi sebagai KODE ANGKA PANJANG
# (mis. nomor rekening pengirim diulang) alih-alih teks deskriptif -
# kadang malah rusak jadi notasi ilmiah kalau kolomnya kebetulan
# terbaca sebagai angka oleh Excel/software lain (mis. "1.57e+33").
# User menegaskan pola ini = transfer internal MASUK (kredit).
_LONG_NUMERIC_KETERANGAN_RE = re.compile(r"^\d{10,}$")
_SCI_NOTATION_KETERANGAN_RE = re.compile(r"^\d(\.\d+)?e\+?\d+$", re.IGNORECASE)
# Kode referensi Fliptech di BRI kadang berupa huruf+angka (mis.
# "FLP686872907"), bukan angka murni - user menegaskan pola ini BIASANYA
# transfer masuk, tapi tetap harus divalidasi manual (bukan match otomatis
# dengan percaya diri) - jadi cukup dikategorikan Transaksi Internal
# supaya masuk proses pencocokan normal, bukan diserahkan ke aturan lain.
_FLP_CODE_RE = re.compile(r"^flp\d{6,}$", re.IGNORECASE)


def _looks_like_long_numeric_code(desc):
    text = str(desc if desc is not None else "").strip()
    if not text:
        return False
    return (bool(_LONG_NUMERIC_KETERANGAN_RE.match(text)) or bool(_SCI_NOTATION_KETERANGAN_RE.match(text))
            or bool(_FLP_CODE_RE.match(text)))


# Kategori saldo awal -> dipakai untuk saldo awal Neraca, dilewati saat
# menjumlah transaksi berjalan.
OPENING_KEYWORDS = ["saldo awal"]

# Alias yang dipakai di kolom Subjek/Objek untuk merujuk rekening lain.
# Kunci = potongan teks yang mungkin muncul (huruf kecil), nilai = None
# (akan dicocokkan dengan resolve_account_alias terhadap nama sheet asli).
ACCOUNT_HINTS = ["bri-507", "bri-567", "bca-887", "bca-", "jago", "kasir"]

TOLERANCI_HARI = 30  # jendela pencarian pasangan transfer, sesuai prinsip audit
TOLERANSI_NOMINAL_PERSEN = 0.02  # 2% -> untuk toleransi biaya admin/pembulatan
TOLERANSI_NOMINAL_ABS = 5000  # atau selisih absolut di bawah ini dianggap wajar


@dataclass
class Txn:
    sheet: str
    row: int  # nomor baris di sheet asal (1-based, termasuk header)
    date: object
    desc: str
    kategori: str
    debit: float
    kredit: float
    saldo: float
    subjek: str
    objek: str
    ket: str

    @property
    def nominal(self):
        """Nilai transaksi bertanda: negatif jika debit (uang keluar),
        positif jika kredit (uang masuk). Kalau SATU baris punya Debit
        DAN Kredit sekaligus (mis. biaya admin dipotong langsung dari
        transaksi masuk, digabung satu baris) - jumlahkan keduanya untuk
        efek bersihnya, JANGAN cuma ambil salah satu (debit sudah negatif
        di data sumber, jadi penjumlahan otomatis menghasilkan net yang
        benar)."""
        return (self.debit or 0) + (self.kredit or 0)

    @property
    def is_transfer(self):
        k = (self.effective_kategori or "").lower()
        if any(kw in k for kw in TRANSFER_KEYWORDS):
            return True
        subjek_k = (self.subjek or "").strip().lower()
        objek_k = (self.objek or "").strip().lower()
        # Subjek DAN Objek SAMA-SAMA alias owner - kemungkinan besar
        # transfer antar rekening pribadi/bisnis owner sendiri, jadikan
        # kandidat pencocokan transfer terlepas dari Kategori aslinya
        # (lihat OWNER_ALIASES) - masih lewat pencocokan nominal+tanggal
        # normal, bukan otomatis dianggap benar tanpa verifikasi.
        if subjek_k in OWNER_ALIASES and objek_k in OWNER_ALIASES:
            return True
        melibatkan_owner = subjek_k in OWNER_ALIASES or objek_k in OWNER_ALIASES
        # Modal & Setoran Pemilik yang melibatkan owner (salah satu dari
        # Subjek/Objek) - JADIKAN kandidat pencocokan transfer juga
        # (bukan langsung dianggap False seperti Modal biasa dari luar).
        # User menegaskan: setelah verifikasi manual, banyak transaksi
        # yang tadinya dikira "Setoran Pemilik" (modal baru masuk dari
        # luar) ternyata SEBENARNYA transfer internal (owner pindahkan
        # uang antar rekening miliknya sendiri) - kalau memang ketemu
        # pasangan valid (nominal+tanggal cocok) di rekening lain, itu
        # LEBIH DIPERCAYA daripada anggapan awal "Setoran Pemilik".
        # Kalau TIDAK ketemu pasangan, tetap jatuh ke Modal seperti biasa
        # (penanganan itu terjadi di proses koreksi /rekonlokal, bukan
        # di sini - di sini cuma menentukan APAKAH masuk kandidat dulu).
        if any(kw in k for kw in CAPITAL_KEYWORDS) and melibatkan_owner:
            return True
        # fallback ke keterangan kalau kategori tidak/salah diisi, kecuali
        # sudah eksplisit dikategorikan sebagai modal (setoran dari luar,
        # bukan pindah antar rekening sendiri), pengeluaran pribadi, atau
        # hutang masuk/pembayaran hutang (uang dari/ke PIHAK LUAR, bukan
        # pindah antar rekening sendiri - jangan dicoba dicocokkan sebagai
        # transfer internal)
        if (any(kw in k for kw in CAPITAL_KEYWORDS) or any(kw in k for kw in PERSONAL_EXPENSE_KEYWORDS)
                or k == "hutang masuk" or k == "pembayaran hutang"):
            return False
        d = (self.desc or "").lower()
        return any(kw in d for kw in DESC_TRANSFER_KEYWORDS)

    @property
    def is_personal_expense(self):
        k = (self.effective_kategori or "").lower()
        return any(kw in k for kw in PERSONAL_EXPENSE_KEYWORDS)

    @property
    def is_new_debt_declaration(self):
        """True kalau transaksi ini kemungkinan besar PENCAIRAN HUTANG
        BARU (bukan cicilan/pembayaran hutang yang sudah ada) - dicek
        dari effective_kategori == 'Hutang Masuk' (konvensi bot konversi
        terbaru: kategori ini SELALU ditulis eksplisit untuk pencairan
        hutang, bukan cuma diselipkan sebagai kata kunci di Keterangan)
        ATAU kata kunci eksplisit (jaring pengaman untuk data lama/
        sumber lain yang belum ikut konvensi ini), DAN uangnya masuk
        (kredit). Tidak mengubah kategori efektif - cuma dipakai untuk
        daftar audit terpisah (Rekonsiliasi bagian 6) yang mengingatkan
        user menambahkan baris baru di Buku Hutang (laporan kuartal/
        tahunan)."""
        if self.nominal <= 0:
            return False
        if (self.effective_kategori or "").strip().lower() == "hutang masuk":
            return True
        text = f"{self.desc or ''} {self.ket or ''} {self.objek or ''} {self.subjek or ''}".lower()
        return any(_override_keyword_found(kw, text) for kw in NEW_DEBT_KEYWORDS)

    @property
    def is_tip_minus_variant(self):
        return (self.effective_kategori or "").strip().lower() == "tip/minus/lebih"

    @property
    def category_override(self):
        """Kategori pengganti berdasarkan CATEGORY_OVERRIDE_RULES (kata
        kunci di Keterangan/Keterangan Tambahan/Objek/Subjek, ditegaskan
        user berdasarkan pengalaman nyata bisnisnya) - None kalau tidak
        ada aturan yang cocok atau kategori aslinya sudah deliberate/tegas
        (modal, laba ditahan, saldo awal, gaji - lihat
        _PROTECTED_FROM_CATEGORY_OVERRIDE) dan tidak boleh ditimpa."""
        k = (self.kategori or "").strip().lower()
        if k in _PROTECTED_FROM_CATEGORY_OVERRIDE or k.startswith("gaji") or self.is_opening:
            return None
        if self.nominal > 0 and _looks_like_long_numeric_code(self.desc):
            return "Transaksi Internal"
        text = f"{self.kategori or ''} {self.desc or ''} {self.ket or ''} {self.objek or ''} {self.subjek or ''}".lower()
        for rule in CATEGORY_OVERRIDE_RULES:
            sheet_filter = rule.get("sheet_contains")
            if sheet_filter and sheet_filter not in (self.sheet or "").lower():
                continue
            amount_min = rule.get("amount_min")
            amount_max = rule.get("amount_max")
            if amount_min is not None and abs(self.nominal) < amount_min:
                continue
            if amount_max is not None and abs(self.nominal) > amount_max:
                continue
            direction = rule.get("direction")  # "masuk" (kredit) / "keluar" (debit) / None (keduanya)
            if direction == "masuk" and self.nominal <= 0:
                continue
            if direction == "keluar" and self.nominal >= 0:
                continue
            kategori_asli = rule.get("kategori_asli")
            if kategori_asli is not None and k != kategori_asli:
                continue
            if "any" in rule and not any(_override_keyword_found(kw, text) for kw in rule["any"]):
                continue
            if "all" in rule and not all(_override_keyword_found(kw, text) for kw in rule["all"]):
                continue
            if "none_of" in rule and any(_override_keyword_found(kw, text) for kw in rule["none_of"]):
                continue
            return rule["category"]
        return None

    @property
    def effective_kategori(self):
        """Kategori yang SEBENARNYA dipakai untuk semua perhitungan (Laba
        Rugi, Neraca, pencocokan transfer, dst) - urutan: (1)
        category_override kalau ada aturan yang cocok, (2) Kategori asli
        apa adanya kalau itu SUDAH salah satu kategori yang dikenal
        sistem (lihat _is_recognized_category), (3) 'Kategori Baru' kalau
        Kategori aslinya genuinely tidak dikenal sama sekali - supaya
        transaksi yang benar-benar belum dikenal TERLIHAT JELAS untuk
        diaudit, bukan diam-diam hilang dari Laba Rugi seperti kasus
        'SETORAN VIA CDM'/'Modal' pendek yang pernah kejadian. SEMUA
        rumus SUMIF/SUMIFS kategori di laporan merujuk ke kolom bantu M
        (ditulis dari nilai ini), bukan langsung ke kolom C."""
        override = self.category_override
        if override:
            return override
        if self.is_opening or _is_recognized_category(self.kategori):
            return self.kategori
        return "Kategori Baru"

    @property
    def is_capital(self):
        k = (self.effective_kategori or "").lower()
        if any(kw in k for kw in CAPITAL_KEYWORDS):
            return True
        # kategori pendek "Modal" (tanpa "& Setoran Pemilik") tetap
        # dianggap modal - user kadang menyingkat, mis. Keterangan "Modal
        # Masuk" dengan Kategori cuma "Modal"
        if k.strip() == "modal" or k.strip().startswith("modal "):
            return True
        if "transfer masuk" in k:
            combined = f"{self.desc or ''} {self.ket or ''}".lower()
            if any(kw in combined for kw in CAPITAL_SELF_TRANSFER_KEYWORDS):
                return True
        return False

    @property
    def is_opening(self):
        k = (self.kategori or "").lower()
        return any(kw in k for kw in OPENING_KEYWORDS)

    def cell_ref(self, col):
        return f"'{self.sheet}'!${col}${self.row}"


# ---------------------------------------------------------------------------
# Membaca sheet rekening
# ---------------------------------------------------------------------------

def is_closing_summary_row(tanggal, kategori, keterangan, is_first_row=False):
    """Deteksi baris rekap penutup (Saldo Awal/Saldo Akhir/Total Debit/Total
    Kredit) yang kadang ada di baris-baris akhir sheet rekening sebagai
    ringkasan, BUKAN transaksi. Kalau ikut dimasukkan ke rekonstruksi saldo
    berjalan (kolom J/K/L), nilainya (yang merupakan TOTAL/ringkasan, bukan
    nominal transaksi tunggal) akan merusak saldo kumulatif dan jadi sumber
    selisih di Neraca.

    Syarat UTAMA: baris TIDAK punya tanggal (dan bukan baris data pertama -
    baris Saldo Awal/Saldo Awal Bulan yang legitimate di baris pertama
    kadang memang tidak diisi tanggal, itu bukan blok penutup). Transaksi
    asli SELALU bertanggal, termasuk checkpoint tengah bulan seperti 'Saldo
    akhir sesi' (checkpoint akhir shift kasir, muncul berkali-kali per
    bulan, tapi tetap punya tanggal asli) - kalau syarat tanpa-tanggal ini
    tidak dijadikan gerbang wajib untuk kata kunci juga, baris seperti itu
    akan salah dikira blok penutup HANYA karena mengandung teks 'saldo
    akhir', dan memotong rekonstruksi saldo jauh sebelum akhir bulan
    sesungguhnya."""
    if is_first_row or tanggal is not None:
        return False
    text = f"{kategori or ''} {keterangan or ''}".strip().lower()
    return bool(text)


def read_account_sheet(ws):
    """Baca satu sheet rekening jadi list[Txn], berhenti di baris kosong
    pertama setelah header (baris trailing kosong diabaikan), ATAU begitu
    ketemu baris rekap penutup (Saldo Akhir/Total Debit/Total Kredit) -
    baris itu dan seterusnya tidak dianggap transaksi, tapi nilainya
    ditangkap terpisah sebagai acuan cross-check (lihat closing_info)."""
    txns = []
    closing_info = {}
    seen_first_row = False
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        tanggal, ket, kategori, debit, kredit, saldo, subjek, objek, ket_tambahan = (
            (c.value for c in row[:9])
        )
        if tanggal is None and ket is None and debit is None and kredit is None:
            continue
        is_first = not seen_first_row
        seen_first_row = True
        if is_closing_summary_row(tanggal, kategori, ket, is_first_row=is_first):
            label = f"{kategori or ''} {ket or ''}".strip().lower()
            value = None
            for cand in (debit, kredit, saldo):
                if isinstance(cand, (int, float)):
                    value = cand
                    break
            if "total debit" in label:
                closing_info["total_debit"] = value
            elif "total kredit" in label:
                closing_info["total_kredit"] = value
            elif "saldo akhir" in label:
                closing_info["saldo_akhir"] = value
            continue  # bukan transaksi, jangan dimasukkan ke txns
        txns.append(
            Txn(
                sheet=ws.title,
                row=row[0].row,
                date=tanggal,
                # Semua kolom teks di-str()-kan eksplisit - beberapa file
                # sumber (mis. Numbers/Excel yang salah menebak tipe sel)
                # kadang menyimpan kolom teks (Keterangan dst) sebagai
                # ANGKA MENTAH (bahkan notasi ilmiah kalau angkanya sangat
                # panjang, mis. nomor rekening yang diulang di Keterangan)
                # - tanpa str() eksplisit, downstream code yang
                # mengasumsikan string (mis. .lower()) akan crash.
                desc=str(ket) if ket is not None else "",
                kategori=str(kategori) if kategori is not None else "",
                debit=float(debit) if isinstance(debit, (int, float)) else 0,
                kredit=float(kredit) if isinstance(kredit, (int, float)) else 0,
                saldo=float(saldo) if isinstance(saldo, (int, float)) else None,
                subjek=str(subjek) if subjek is not None else "",
                objek=str(objek) if objek is not None else "",
                ket=str(ket_tambahan) if ket_tambahan is not None else "",
            )
        )
    return txns, closing_info


FLIPTECH_SPLIT_REMAINDER_MAX = 1000  # sisa di bawah ini dianggap biaya
                                       # admin/bunga gabungan, bukan nominal genuine


def split_fliptech_combined_rows(ws):
    """Sebagian transaksi Fliptech tercatat sebagai SATU baris dengan
    nominal gabungan (mis. -131103 = -131000 transfer + -103 biaya admin,
    -67105 = -67000 + -105) - bukan dua baris terpisah seperti konvensi
    normal ("Bagian dari transaksi Fliptech: Biaya Admin" di baris
    nol-nominal terpisah). Kalau dibiarkan satu baris, transfer TIDAK
    AKAN PERNAH cocok dengan pasangannya di rekening lain (yang biasanya
    nominal genap/dibulatkan), karena selisih kecil (103/105/dst) itu di
    luar toleransi pencocokan normal.

    Deteksi pola ini (kategori transfer-like, Objek/Keterangan/Keterangan
    Tambahan menyebut 'fliptech', sisa nominal di bawah Rp1.000 dan bukan
    0) dan PECAH jadi 2 baris: baris asli jadi nominal genap (dibulatkan
    ke kelipatan 1000 terdekat ke arah nol), baris baru (disisipkan tepat
    sesudahnya) untuk sisanya sebagai Biaya Admin Bank (kalau debit) atau
    Bunga Bank (kalau kredit) - konsisten dengan konvensi Fliptech yang
    sudah ada di file lain.

    HARUS dipanggil PALING AWAL, sebelum last_data_row/add_helper_column/
    read_account_sheet - supaya penyisipan baris tidak merusak rumus/
    kolom bantu yang sudah ditulis di baris-baris sesudahnya."""
    row = 2
    while True:
        tanggal = ws.cell(row=row, column=1).value
        keterangan = ws.cell(row=row, column=2).value
        kategori = ws.cell(row=row, column=3).value
        if tanggal is None and keterangan is None and kategori is None:
            break  # sudah lewat baris data terakhir
        objek = ws.cell(row=row, column=8).value
        ket_tambahan = ws.cell(row=row, column=9).value
        text = f"{keterangan or ''} {objek or ''} {ket_tambahan or ''}".lower()
        is_transfer_like = isinstance(kategori, str) and any(kw in kategori.lower() for kw in TRANSFER_KEYWORDS)
        if is_transfer_like and "fliptech" in text:
            debit = ws.cell(row=row, column=4).value
            kredit = ws.cell(row=row, column=5).value
            is_debit = isinstance(debit, (int, float)) and debit
            nominal = debit if is_debit else (kredit if isinstance(kredit, (int, float)) else None)
            if isinstance(nominal, (int, float)) and nominal != 0:
                remainder = round(abs(nominal) % 1000, 2)
                if 0 < remainder < FLIPTECH_SPLIT_REMAINDER_MAX:
                    sign = 1 if nominal > 0 else -1
                    main_amount = sign * round(abs(nominal) - remainder, 2)
                    fee_amount = sign * remainder
                    subjek = ws.cell(row=row, column=7).value
                    saldo_tercatat = ws.cell(row=row, column=6).value
                    # baris asli jadi nominal genap; Saldo Kumulatif
                    # dikosongkan (bukan snapshot resmi bank lagi, cuma
                    # posisi ANTARA - baris baru di bawah yang bawa
                    # Saldo Kumulatif resmi hasil rekaman bank)
                    if is_debit:
                        ws.cell(row=row, column=4, value=main_amount)
                    else:
                        ws.cell(row=row, column=5, value=main_amount)
                    ws.cell(row=row, column=6, value=None)

                    ws.insert_rows(row + 1)
                    # ws.insert_rows() TIDAK mewarisi format (font/
                    # number_format) dari baris sekitarnya - openpyxl
                    # kasih default kosong, beda dari gaya asli file
                    # (mis. Arial 9 + format tanggal DD/MM/YYYY jadi
                    # Calibri 11 + yyyy-mm-dd h:mm:ss). Salin dari baris
                    # asal (row) supaya baris baru konsisten visual
                    # dengan baris lain, bukan menonjol keliru.
                    for col in range(1, 10):
                        src_cell = ws.cell(row=row, column=col)
                        new_cell = ws.cell(row=row + 1, column=col)
                        new_cell.font = copy.copy(src_cell.font)
                        new_cell.number_format = src_cell.number_format
                        new_cell.alignment = copy.copy(src_cell.alignment)
                        new_cell.border = copy.copy(src_cell.border)
                    fee_label = "Biaya Admin Bank" if sign < 0 else "Bunga Bank"
                    note = (f"Bagian dari transaksi Fliptech: {fee_label} "
                            "(dipisah otomatis dari nominal gabungan oleh reconcile.py)")
                    ws.cell(row=row + 1, column=1, value=tanggal)
                    ws.cell(row=row + 1, column=2, value=note)
                    ws.cell(row=row + 1, column=3, value="Biaya Admin Bank")
                    if sign < 0:
                        ws.cell(row=row + 1, column=4, value=fee_amount)
                    else:
                        ws.cell(row=row + 1, column=5, value=fee_amount)
                    ws.cell(row=row + 1, column=6, value=saldo_tercatat)
                    ws.cell(row=row + 1, column=7, value="-")
                    ws.cell(row=row + 1, column=8, value=subjek)
                    ws.cell(row=row + 1, column=9, value=note)
                    row += 1  # lewati baris baru yang baru disisipkan
        row += 1


def last_data_row(ws):
    last = 1
    seen_first_row = False
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        tanggal, ket, kategori, debit, kredit = (c.value for c in row[:5])
        if not any(v is not None for v in (tanggal, ket, kategori, debit, kredit)):
            continue  # baris kosong, jangan dianggap baris data pertama
        is_first = not seen_first_row
        seen_first_row = True
        if is_closing_summary_row(tanggal, kategori, ket, is_first_row=is_first):
            break  # blok rekap penutup (Saldo Akhir/Total Debit/Kredit) -
            # berhenti di sini, jangan ikut dihitung sebagai baris transaksi
        last = row[0].row
    return last


_STANDARD_HEADER = ["Tanggal", "Keterangan Transaksi", "Kategori Transaksi", "Debit", "Kredit",
                     "Saldo Kumulatif", "Subjek Transaksi", "Objek Transaksi"]


def _looks_like_account_sheet(ws):
    """True kalau sheet ini berformat standar 9-kolom (Tanggal/
    Keterangan/Kategori/Debit/Kredit/Saldo/Subjek/Objek/Ket Tambahan) -
    dicek dari 8 kolom header pertama. Dipakai untuk fitur yang menerima
    file rekening BEBAS (bukan file Rekonsiliasi multi-sheet biasa),
    supaya sheet non-rekening (kalau ada) tidak ikut diproses."""
    header = [ws.cell(row=1, column=c).value for c in range(1, len(_STANDARD_HEADER) + 1)]
    return header == _STANDARD_HEADER


def resolve_account_sheet(hint, sheet_names):
    """Cocokkan teks bebas di kolom Subjek/Objek (mis. 'BRI-567(Biz)',
    'Rekening Jago + Admin Rp200') ke nama sheet rekening sebenarnya.

    Pakai token PALING SPESIFIK (terpanjang) yang cocok, BUKAN sheet
    pertama yang kebetulan punya token apapun yang cocok - penting kalau
    ada beberapa rekening berbagi token generik yang sama (mis. dua
    rekening sama-sama berakhiran '(Biz)': 'BRI-567(Biz)' dan
    'BCA-292(Biz)' - token 'Biz' pendek dan generik, cocok ke KEDUANYA;
    tanpa preferensi token terpanjang, sheet yang kebetulan lebih dulu
    di urutan wb.sheetnames yang menang, walau salah)."""
    if not hint:
        return None
    h = hint.lower()
    best = None
    best_len = 0
    for name in sheet_names:
        n = name.lower()
        # ambil token pembeda dari nama sheet, contoh nomor rekening / "jago" / "kasir"
        tokens = [t for t in n.replace("(", " ").replace(")", " ").split() if len(t) >= 3]
        for tok in tokens:
            if tok in h and len(tok) > best_len:
                best = name
                best_len = len(tok)
    return best


# ---------------------------------------------------------------------------
# Pencocokan transfer antar rekening ("Rekonsiliasi" inti)
# ---------------------------------------------------------------------------

@dataclass
class Match:
    src: Txn
    dst: Txn = None
    confidence: str = "Needs manual verification"
    reasoning: str = ""
    date_diff: int = None
    nominal_diff: float = None
    is_fliptech: bool = False


def days_between(a, b):
    """Selisih hari antara dua tanggal, menerima datetime.datetime,
    datetime.date, ATAU teks tanggal format umum (lihat coerce_date) - kalau
    cuma cek isinstance(datetime.date), file dengan kolom tanggal bertipe
    teks (mis. hasil parser 'preformatted' tertentu) akan selalu dianggap
    selisih 9999 hari (di luar toleransi), bikin SEMUA transfer di file itu
    gagal matched walau tanggal & nominalnya persis sama."""
    a = coerce_date(a)
    b = coerce_date(b)
    if a is None or b is None:
        return 9999
    return abs((a - b).days)


def _find_fliptech_loan_companion(src, all_txns):
    """Cari baris pendamping nol-nominal 'Bagian dari transaksi Fliptech: ...'
    di sheet & tanggal yang sama dengan src, yang keterangannya menyebut
    kata kunci pinjaman/cicilan/utang - sinyal bahwa transaksi src ini
    sebenarnya CICILAN PINJAMAN ke pihak luar (Fliptech Lentera Inspirasi
    sebagai penyedia pembiayaan), bukan transfer antar rekening sendiri,
    sehingga TIDAK AKAN PERNAH ketemu pasangannya di rekening manapun -
    beda dengan transfer internal biasa yang genuinely belum ketemu.
    Return teks catatan itu kalau ketemu, None kalau tidak."""
    loan_keywords = ["cicilan pinjaman", "cicilan", "pinjaman", "pembayaran utang", "angsuran"]
    src_date = coerce_date(src.date)
    for t in all_txns:
        if t is src or t.sheet != src.sheet or t.nominal != 0:
            continue
        if coerce_date(t.date) != src_date:
            continue
        text = (t.ket or "").lower()
        if not text.startswith("bagian dari transaksi fliptech"):
            continue
        for kw in loan_keywords:
            if kw in text:
                return t.ket
    return None


def _near_miss_candidates(src, all_txns, consumed_ids, max_hasil=3):
    """Cari transaksi lain (di rekening MANAPUN, tanggal & nominal
    mendekati src) yang KEMUNGKINAN sebenarnya pasangan src, tapi tidak
    terpakai find_matches karena satu dan lain hal - dipakai untuk
    memperkaya alasan 'Needs manual verification' supaya user bisa audit
    KENAPA bot menolaknya (kategori tidak sesuai, sudah kepakai transaksi
    lain, dst), bukan cuma 'tidak ditemukan' tanpa penjelasan."""
    hasil = []
    for t in all_txns:
        if t is src or t.sheet == src.sheet:
            continue
        if t.nominal == 0:
            continue
        date_diff = days_between(src.date, t.date)
        if date_diff is None or date_diff > TOLERANCI_HARI:
            continue
        toleransi = max(TOLERANSI_NOMINAL_ABS, abs(src.nominal) * TOLERANSI_NOMINAL_PERSEN)
        nominal_diff = abs(abs(src.nominal) - abs(t.nominal))
        if nominal_diff > toleransi:
            continue
        # sejauh ini nominal & tanggal cocok - kenapa tidak terpakai?
        if not t.is_transfer:
            sebab = (
                f"Kategori efektifnya '{t.effective_kategori}' (kategori asli: '{t.kategori}') - "
                "dianggap BUKAN transfer internal, jadi tidak dipertimbangkan sebagai pasangan."
            )
        elif id(t) in consumed_ids:
            sebab = "Sudah terpakai sebagai pasangan transaksi transfer lain - satu transaksi tidak bisa jadi pasangan dua transfer sekaligus."
        else:
            sebab = "Tidak jelas kenapa tidak terpilih - kemungkinan ada kandidat lain yang skornya lebih baik."
        hasil.append({
            "sheet": t.sheet, "row": t.row, "desc": t.desc, "nominal": abs(t.nominal),
            "date_diff": date_diff, "nominal_diff": nominal_diff, "sebab": sebab,
        })
        if len(hasil) >= max_hasil:
            break
    return hasil


def _format_near_miss_note(near_miss_list):
    """Format daftar kandidat dekat jadi teks BER-BARIS (bukan satu
    paragraf padat) - tiap kandidat: judul (rekening) lalu poin-poin
    detail, dipisah baris kosong antar kandidat supaya gampang dipindai
    mata saat audit."""
    if not near_miss_list:
        return ""
    rp = lambda n: f"Rp{n:,.0f}".replace(",", ".")
    blok = [f"{len(near_miss_list)} KANDIDAT DEKAT ditemukan:"]
    for i, c in enumerate(near_miss_list, start=1):
        blok.append(
            f"\n({i}) {c['sheet']}\n"
            f"- Baris: {c['row']}\n"
            f"- Keterangan: {c['desc']}\n"
            f"- Nominal: {rp(c['nominal'])}\n"
            f"- Selisih tanggal: {c['date_diff']} hari\n"
            f"- Selisih nominal: {rp(c['nominal_diff'])}\n"
            f"- Sebab: {c['sebab']}"
        )
    return "\n".join(blok)


def find_matches(all_txns, sheet_names):
    """Untuk setiap transaksi bertanda transfer internal, cari pasangan di
    rekening tujuan (berdasarkan Subjek/Objek) dalam jendela +-30 hari,
    dengan toleransi selisih nominal untuk biaya admin/pembulatan.
    Tidak pernah menyimpulkan 'tidak ditemukan' tanpa mencoba seluruh
    kandidat di rekening tujuan terlebih dahulu."""
    # nominal 0 dikecualikan dari pencocokan: baris seperti ini biasanya
    # penanda/referensi dari parser sumber (mis. "Bagian dari transaksi
    # Fliptech: ...") bukan perpindahan uang sungguhan - tidak ada nominal
    # untuk dicocokkan, jadi kalau ikut diproses selalu nyangkut sebagai
    # "Needs manual verification" tanpa nilai informasi apapun
    transfers = [t for t in all_txns if t.is_transfer and t.nominal != 0]
    # Urutkan transfers supaya src yang PUNYA kandidat match SEMPURNA
    # (nominal persis sama + tanggal persis sama) diproses LEBIH DULU -
    # tanpa ini, src yang diproses lebih awal (urutan baris di file, BUKAN
    # relevansi) bisa "mencuri" kandidat terbaik dari src LAIN yang
    # SEBENARNYA lebih cocok dengan kandidat itu (mis. src A tanggal 13
    # diproses duluan, ambil kandidat tanggal 18 karena itu yang terdekat
    # SAAT ITU - padahal src B tanggal 18 yang genuinely match SEMPURNA
    # dengan kandidat tanggal 18 itu, jadi kebagian sisa yang lebih jauh).
    # Src TANPA kandidat sempurna tetap diproses (urutan asli dipertahankan,
    # sort Python stabil), cuma belakangan - supaya tidak menghalangi src
    # LAIN yang punya kandidat sempurna mendapatkan haknya lebih dulu.
    def _has_exact_candidate(src):
        for t in transfers:
            if t is src or t.sheet == src.sheet or t.nominal == 0:
                continue
            if (t.nominal > 0) == (src.nominal > 0):
                continue
            if abs(abs(t.nominal) - abs(src.nominal)) == 0 and days_between(src.date, t.date) == 0:
                return True
        return False

    transfers = sorted(transfers, key=lambda s: not _has_exact_candidate(s))
    matched_dst_ids = set()
    consumed_ids = set()  # baik src maupun dst yang sudah punya pasangan
    results = []

    for src in transfers:
        if id(src) in consumed_ids:
            # sudah tercatat sebagai pasangan (dst) dari transaksi lain,
            # tidak perlu dilaporkan dua kali dari sisi yang berlawanan
            continue
        # tentukan rekening lawan dari Subjek/Objek (siapa pun yang BUKAN
        # rekening sheet sumber itu sendiri)
        counterpart_hint = None
        for hint in (src.objek, src.subjek):
            resolved = resolve_account_sheet(hint, sheet_names)
            if resolved and resolved != src.sheet:
                counterpart_hint = resolved
                break

        candidates = [
            t
            for t in all_txns
            if t is not src
            and t.is_transfer
            and t.nominal != 0
            and id(t) not in consumed_ids
            and t.sheet != src.sheet  # TIDAK PERNAH sheet yang sama - lihat catatan bug di bawah
            and (counterpart_hint is None or t.sheet == counterpart_hint)
            and (
                (t.nominal > 0) != (src.nominal > 0)  # tanda berlawanan (normal)
                or counterpart_hint is not None  # atau rekening tujuan sudah
                # pasti dari Subjek/Objek -> longgarkan syarat tanda, karena
                # sebagian pencatatan "Pindah Rekening Internal" tidak
                # konsisten memakai tanda negatif untuk uang keluar
            )
        ]
        # BUG SERIUS yang diperbaiki: sebelum ada baris "t.sheet !=
        # src.sheet" di atas, filter "(counterpart_hint is None or
        # t.sheet == counterpart_hint)" jadi SELALU True (vacuous)
        # begitu counterpart_hint None - artinya TIDAK ADA filter
        # rekening SAMA SEKALI, termasuk membolehkan transaksi dari
        # SHEET YANG SAMA (rekening sendiri) match dengan dirinya
        # sendiri kalau kebetulan tanggal+nominal berlawanan cocok.
        # Ditemukan dari laporan user: 2 baris di SATU rekening yang
        # sama (kredit +1.200.000 & debit -1.200.000 tanggal sama)
        # saling matched, padahal jelas harus rekening BERBEDA.
        if not candidates and counterpart_hint is None:
            # tidak ada petunjuk rekening tujuan -> perluas ke semua sheet lain
            candidates = [
                t
                for t in all_txns
                if t is not src
                and t.is_transfer
                and t.nominal != 0
                and id(t) not in consumed_ids
                and t.sheet != src.sheet
                and (t.nominal > 0) != (src.nominal > 0)
            ]

        # skor tiap kandidat: utamakan yang tandanya berlawanan (pencatatan
        # normal), baru selisih tanggal, baru selisih nominal
        scored = []
        for c in candidates:
            nominal_diff = abs(abs(src.nominal) - abs(c.nominal))
            date_diff = days_between(src.date, c.date)
            if date_diff > TOLERANCI_HARI:
                continue
            toleransi = max(TOLERANSI_NOMINAL_ABS, abs(src.nominal) * TOLERANSI_NOMINAL_PERSEN)
            if nominal_diff > toleransi:
                continue
            same_sign = (c.nominal > 0) == (src.nominal > 0)
            # Kas-Buku/Kas Kasir sengaja jadi prioritas PALING TERAKHIR
            # dipilih - setoran tunai/tarik tunai sering settle bertahap/
            # tercampur (lihat catatan di banyak kasus rekonsiliasi
            # sebelumnya), jadi kalau ada kandidat NON-Kas-Buku dengan
            # tanda yang sama cocoknya, itu didahulukan meski selisih
            # tanggal/nominalnya sedikit lebih besar dari kandidat Kas-Buku.
            is_kas_buku = 1 if c.sheet.strip().lower().startswith("kas") else 0
            scored.append((1 if same_sign else 0, is_kas_buku, date_diff, nominal_diff, c, same_sign))

        scored.sort(key=lambda x: (x[0], x[1], x[2], x[3]))

        if scored:
            _, _, date_diff, nominal_diff, dst, same_sign = scored[0]
            matched_dst_ids.add(id(dst))
            consumed_ids.add(id(src))
            consumed_ids.add(id(dst))
            is_fliptech = (
                "fliptech" in (src.desc or "").lower()
                or "fliptech" in (dst.desc or "").lower()
                or "fliptech" in (src.ket or "").lower()
                or "fliptech" in (dst.ket or "").lower()
            )
            if nominal_diff == 0 and date_diff == 0:
                conf = "High"
                reason = "Nominal dan tanggal sama persis di kedua rekening."
            elif nominal_diff == 0 and date_diff <= 3:
                conf = "High"
                reason = f"Nominal sama persis, selisih tanggal {date_diff} hari (wajar untuk settlement bank)."
            elif is_fliptech and nominal_diff <= FLIPTECH_FEE_THRESHOLD:
                conf = "High"
                reason = (
                    f"Transaksi via Fliptech, selisih nominal Rp{nominal_diff:,.0f} "
                    f"dipastikan biaya admin transfer (bukan keraguan), selisih tanggal "
                    f"{date_diff} hari."
                ).replace(",", ".")
            elif nominal_diff > 0 and nominal_diff <= TOLERANSI_NOMINAL_ABS:
                conf = "Medium"
                reason = (
                    f"Selisih nominal Rp{nominal_diff:,.0f} kemungkinan biaya admin/transfer, "
                    f"selisih tanggal {date_diff} hari."
                ).replace(",", ".")
            elif date_diff > 3:
                conf = "Medium"
                reason = f"Nominal cocok (toleransi Rp{nominal_diff:,.0f}) tapi tanggal berbeda {date_diff} hari, kemungkinan delayed posting/settlement.".replace(",", ".")
            else:
                conf = "Low"
                reason = f"Kecocokan hanya berdasarkan toleransi umum, perlu verifikasi manual (selisih Rp{nominal_diff:,.0f}, {date_diff} hari).".replace(",", ".")
            if same_sign:
                # tanda sama-sama positif/negatif di kedua rekening -
                # dicocokkan lewat Subjek/Objek + nominal, bukan lewat tanda,
                # karena penulisan tanda di salah satu sisi kemungkinan keliru
                reason += (
                    " Catatan: kedua sisi tercatat dengan tanda yang sama "
                    "(bukan berlawanan) - kemungkinan salah input tanda "
                    "debit/kredit di salah satu rekening, cek manual."
                )
                if conf == "High":
                    conf = "Medium"
            results.append(
                Match(src=src, dst=dst, confidence=conf, reasoning=reason,
                      date_diff=date_diff, nominal_diff=nominal_diff, is_fliptech=is_fliptech)
            )
        else:
            loan_note = _find_fliptech_loan_companion(src, all_txns)
            if loan_note is not None:
                results.append(
                    Match(
                        src=src,
                        dst=None,
                        confidence="Not applicable (bukan transfer internal)",
                        reasoning=(
                            "Baris pendamping Fliptech nol-nominal di baris yang sama "
                            f"eksplisit menyebut \"{loan_note}\" - ini kemungkinan besar "
                            "CICILAN PINJAMAN/PEMBIAYAAN ke pihak luar (Fliptech Lentera "
                            "Inspirasi bertindak sebagai penyedia pembiayaan, bukan sesama "
                            "rekening internal yang kita lacak), BUKAN transfer antar "
                            "rekening sendiri. Transaksi seperti ini TIDAK AKAN PERNAH "
                            "punya pasangan di rekening manapun - pertimbangkan minta "
                            "bank-statement-bot mengkategorikan ulang jadi kategori "
                            "cicilan/utang tersendiri (bukan Transaksi Internal) supaya "
                            "tidak terus muncul di sini."
                        ),
                    )
                )
                continue
            results.append(
                Match(
                    src=src,
                    dst=None,
                    confidence="Needs manual verification",
                    reasoning=(
                        "Tidak ada kandidat dengan nominal berlawanan dalam jendela "
                        f"±{TOLERANCI_HARI} hari di rekening tujuan yang terindikasi "
                        f"({counterpart_hint or 'tidak teridentifikasi dari Subjek/Objek'}).\n"
                        "Kemungkinan: dana masih dalam perjalanan (in-transit), tercatat di "
                        "bulan berikutnya, atau salah kategori."
                        + (("\n\n" + note) if (note := _format_near_miss_note(_near_miss_candidates(src, all_txns, consumed_ids))) else "")
                    ),
                )
            )
    combo_results = find_split_merge_matches(results)
    return results, combo_results


def find_split_merge_matches(results):
    """Prinsip audit #4/#6: transaksi bisa terpecah (1 keluar -> 2 masuk)
    atau tergabung (2 keluar -> 1 masuk). Cari di antara sisa transfer yang
    belum ketemu pasangannya (Needs manual verification), apakah kombinasi
    2 transaksi lain menjumlah ke nominal yang cocok, dalam jendela waktu
    yang wajar."""
    import itertools

    unmatched = [m for m in results if m.dst is None]
    pool = [m.src for m in unmatched]
    used = set()
    combos = []

    # urutkan berdasarkan nominal terbesar dulu supaya transaksi induk
    # (yang paling mungkin "dipecah") diproses lebih dulu
    unmatched_sorted = sorted(unmatched, key=lambda m: -abs(m.src.nominal))

    for m in unmatched_sorted:
        src = m.src
        if id(src) in used:
            continue
        candidates = [
            t for t in pool
            if id(t) not in used
            and t is not src
            and (t.nominal > 0) != (src.nominal > 0)
            and days_between(src.date, t.date) <= TOLERANCI_HARI
        ]
        best = None
        for a, b in itertools.combinations(candidates, 2):
            total = abs(a.nominal) + abs(b.nominal)
            diff = abs(total - abs(src.nominal))
            toleransi = max(TOLERANSI_NOMINAL_ABS, abs(src.nominal) * TOLERANSI_NOMINAL_PERSEN)
            if diff <= toleransi:
                if best is None or diff < best[0]:
                    best = (diff, a, b)
        if best:
            diff, a, b = best
            used.add(id(src))
            used.add(id(a))
            used.add(id(b))
            conf = "High" if diff == 0 else "Medium"
            reason = (
                f"Kemungkinan transaksi terpecah/tergabung: {src.sheet} Rp{abs(src.nominal):,.0f} "
                f"~= {a.sheet} Rp{abs(a.nominal):,.0f} + {b.sheet} Rp{abs(b.nominal):,.0f} "
                f"(selisih Rp{diff:,.0f})."
            ).replace(",", ".")
            if isinstance(src.ket, str) and src.ket.strip():
                reason += f" Catatan asal: \"{src.ket}\""
            combos.append({"src": src, "parts": [a, b], "diff": diff, "confidence": conf, "reasoning": reason})

    return combos


FLIPTECH_FEE_THRESHOLD = shared_rules.get("fliptech_fee_threshold", 2000)
TIP_MINUS_THRESHOLD = shared_rules.get("tip_minus_threshold", 100000)


def compute_balance_status(all_txns_by_sheet):
    """Hitung ulang di Python (bukan tunggu Excel/LibreOffice recalculate
    rumus Neraca) apakah tiap rekening bakal balanced atau masih ada
    selisih - dipakai untuk menulis status ini LANGSUNG ke sheet
    Rekonsiliasi supaya user tidak perlu buka sheet Neraca terpisah untuk
    tahu ada masalah atau tidak; begitu laporan digenerate, statusnya
    sudah kelihatan di satu tempat. Mirror persis logika Excel formula di
    write_income_statement/write_balance_sheet (basis kas penuh, TIDAK
    ada penyusutan Aset Tetap - beda dengan quarterly.py/annual.py yang
    mengkapitalisasi 'Belanja Assets')."""
    def sum_exact(txns, category):
        cat = category.strip().lower()
        return sum(t.nominal for t in txns if (t.effective_kategori or "").strip().lower() == cat)

    def sum_multi(txns, categories):
        cats = {c.strip().lower() for c in categories}
        return sum(t.nominal for t in txns if (t.effective_kategori or "").strip().lower() in cats)

    def sum_gaji(txns):
        return sum(t.nominal for t in txns if (t.effective_kategori or "").strip().lower().startswith("gaji"))

    def sum_tip_minus(txns):
        return sum(t.nominal for t in txns if t.is_tip_minus_variant)

    def sum_modal(txns):
        return sum(t.nominal for t in txns if t.is_capital)

    def sum_personal_expense(txns):
        return sum(t.nominal for t in txns if t.is_personal_expense)

    results = {}
    for sheet, txns in all_txns_by_sheet.items():
        if not txns:
            continue
        # rekonstruksi K persis seperti rumus Excel (K2=F2, K(n)=K(n-1)+J(n))
        opening_f = txns[0].saldo if isinstance(txns[0].saldo, (int, float)) else (txns[0].nominal or 0)
        k_values = [opening_f]
        for t in txns[1:]:
            k_values.append(k_values[-1] + t.nominal)
        total_aset = k_values[-1]

        # "Saldo Awal Bulan" di Neraca MERUJUK ke K pada baris yang eksplisit
        # bertanda is_opening (bisa jadi BUKAN baris pertama - lihat kasus
        # nyata: sheet dengan transaksi biasa tercatat sebelum baris "Saldo
        # Awal"-nya sendiri) - kalau langsung pakai txns[0].saldo mentah
        # tanpa rekonstruksi, hasilnya bisa beda dari yang sungguhan dipakai
        # rumus Excel, dan prediksi Selisih ini jadi tidak akurat.
        opening_idx = next((i for i, t in enumerate(txns) if t.is_opening), 0)
        saldo_awal = k_values[opening_idx]

        revenue = sum_multi(txns, INCOME_CATEGORIES_REVENUE)
        expense = sum(sum_exact(txns, cat) for cat in INCOME_CATEGORIES_EXPENSE)
        expense += sum_multi(txns, MARKETING_RND_CATEGORY_TEXTS)
        expense += sum_gaji(txns)
        expense += sum_multi(txns, BANK_FEE_CATEGORY_TEXTS)
        other = sum(
            sum_tip_minus(txns) if cat == "Tip/Minus/Lebih" else sum_exact(txns, cat)
            for cat in OTHER_CATEGORIES
        )
        laba_bersih = revenue + expense + other
        modal = sum_modal(txns)
        pengeluaran_pribadi = sum_personal_expense(txns)
        ekuitas = saldo_awal + modal + pengeluaran_pribadi + laba_bersih
        # Liabilitas (Hutang) - TERPISAH dari Ekuitas (hutang ke pihak
        # luar, bukan modal pemilik). Pembayaran Hutang (pelunasan pokok)
        # SENGAJA dikeluarkan dari OTHER_CATEGORIES/Laba Rugi supaya tidak
        # dobel hitung di sini - itu pengurang liabilitas, bukan beban.
        hutang_masuk = sum_exact(txns, "Hutang Masuk")
        pembayaran_hutang = sum_exact(txns, "Pembayaran Hutang")
        liabilitas = hutang_masuk + pembayaran_hutang
        transfer_bersih = sum_multi(txns, TRANSFER_CATEGORY_TEXTS)
        selisih = round((total_aset - ekuitas - liabilitas) - transfer_bersih, 2)
        kategori_baru_txns = [t for t in txns if t.effective_kategori == "Kategori Baru"]
        kategori_baru_total = round(sum(t.nominal for t in kategori_baru_txns), 2)
        results[sheet] = {
            "total_aset": round(total_aset, 2),
            "ekuitas": round(ekuitas, 2),
            "liabilitas": round(liabilitas, 2),
            "transfer_bersih": round(transfer_bersih, 2),
            "selisih": selisih,
            "n_kategori_baru": len(kategori_baru_txns),
            "kategori_baru_total": kategori_baru_total,
        }
    return results


def find_minus_flags(all_txns_by_sheet):
    """Kumpulkan indikasi 'minus' yang perlu verifikasi manual: kategori
    Tip/Minus/Lebih dengan nominal > Rp100.000, atau baris berpenanda flag
    (⚑) di keterangan.

    Catatan: Tip/Minus/Lebih dengan nominal <= Rp100.000 dianggap valid
    (wajar terjadi dari pembulatan/kembalian kasir sehari-hari), tidak perlu
    ditandai untuk verifikasi manual.

    Saldo kumulatif negatif di tengah data juga TIDAK dianggap indikasi
    minus - itu cuma efek sementara dari urutan penulisan transaksi (baris
    tertulis belum tentu urut kronologis sempurna), bukan minus riil.
    Kesehatan saldo yang sebenarnya dicek di level akhir bulan lewat kolom
    K (Saldo Kumulatif Rekonstruksi) tiap sheet dan sheet Neraca.

    Cross-check baris-per-baris terhadap saldo hasil rekonstruksi (opening +
    akumulasi nominal) juga sengaja TIDAK dipakai di sini karena pada
    sebagian sheet sumber (mis. BCA) kolom Saldo Kumulatif tidak selalu
    diisi berurutan per baris (beberapa transaksi bertanggal sama dikelompokkan
    dulu), sehingga cross-check per baris menghasilkan banyak false positive."""
    flags = []
    for sheet, txns in all_txns_by_sheet.items():
        for t in txns:
            reasons = []
            if t.is_tip_minus_variant and abs(t.nominal) > TIP_MINUS_THRESHOLD:
                reasons.append(
                    f"Kategori/keterangan menyebut varian Tip/Minus/Lebih (mis. minus/lebih/tip/"
                    f"uang cust) dengan nominal Rp{abs(t.nominal):,.0f} "
                    f"(di atas ambang wajar Rp{TIP_MINUS_THRESHOLD:,.0f})."
                    .replace(",", ".")
                )
            if isinstance(t.ket, str) and "⚑" in t.ket:
                reasons.append(f"Ditandai perlu verifikasi manual: {t.ket}")
            if reasons:
                flags.append((t, reasons))
    return flags


def find_new_category_flags(all_txns_by_sheet):
    """Kumpulkan semua transaksi yang effective_kategori-nya jadi
    'Kategori Baru' - genuinely tidak dikenal sistem sama sekali (bukan
    override yang gagal, bukan yang sudah dikenal seperti Tip/Minus).
    Dipakai untuk daftar audit di sheet Rekonsiliasi bagian 4, supaya
    transaksi yang belum dikenal terlihat jelas dan bisa ditelusuri,
    bukan diam-diam hilang dari Laba Rugi."""
    flags = []
    for sheet, txns in all_txns_by_sheet.items():
        for t in txns:
            if t.effective_kategori == "Kategori Baru":
                flags.append(t)
    return flags


def find_personal_expense_flags(all_txns_by_sheet):
    """Kumpulkan semua transaksi berkategori 'Pengeluaran Pribadi' -
    dikeluarkan dari Laba Rugi (diperlakukan seperti prive/penarikan
    modal, mengurangi Ekuitas), TAPI user menegaskan tiap transaksi ini
    perlu diverifikasi manual satu-satu. Dipakai untuk daftar audit di
    sheet Rekonsiliasi bagian 5."""
    flags = []
    for sheet, txns in all_txns_by_sheet.items():
        for t in txns:
            if t.is_personal_expense:
                flags.append(t)
    return flags


def find_new_debt_flags(all_txns_by_sheet):
    """Kumpulkan transaksi yang kemungkinan besar pencairan HUTANG BARU
    (lihat NEW_DEBT_KEYWORDS) - dipakai untuk daftar audit di sheet
    Rekonsiliasi bagian 6, mengingatkan user menambahkan baris baru di
    Buku Hutang (laporan kuartal/tahunan) - reconcile.py sendiri TIDAK
    punya Buku Hutang (itu cuma ada di quarterly.py/annual.py)."""
    flags = []
    for sheet, txns in all_txns_by_sheet.items():
        for t in txns:
            if t.is_new_debt_declaration:
                flags.append(t)
    return flags


# ---------------------------------------------------------------------------
# Penulisan sheet Rekonsiliasi
# ---------------------------------------------------------------------------

def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
        cell.alignment = Alignment(vertical="center", wrap_text=True)


def conf_fill(conf):
    return {"High": HIGH_FILL, "Medium": MED_FILL, "Low": LOW_FILL}.get(conf, LOW_FILL)


def write_rekonsiliasi_sheet(wb, matches, combo_matches, minus_flags, balance_status=None, new_category_flags=None, personal_expense_flags=None, new_debt_flags=None):
    if "Rekonsiliasi" in wb.sheetnames:
        del wb["Rekonsiliasi"]
    ws = wb.create_sheet("Rekonsiliasi")

    # Urutkan tiap bagian berdasarkan tanggal transaksi (bukan urutan sheet
    # rekening lalu baris) - memudahkan audit karena bisa ditelusuri
    # kronologis lintas semua rekening sekaligus, bukan per-rekening dulu
    # baru pindah ke rekening berikutnya. coerce_date dipakai karena
    # sebagian file punya kolom tanggal bertipe teks; None/tak terbaca
    # ditaruh paling akhir (bukan dianggap "paling awal") biar tidak
    # menyembunyikan baris bermasalah di atas.
    def _sort_key(d):
        parsed = coerce_date(d)
        return (parsed is None, parsed or datetime.date.max)

    matches = sorted(matches, key=lambda m: _sort_key(m.src.date))
    combo_matches = sorted(combo_matches, key=lambda cm: _sort_key(cm["src"].date))
    minus_flags = sorted(minus_flags, key=lambda tf: _sort_key(tf[0].date))

    ws["A1"] = "REKONSILIASI ANTAR REKENING"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Dibuat otomatis. Setiap baris merujuk langsung (link rumus) ke sel asal di sheet "
                "rekening. Tiap bagian diurutkan berdasarkan tanggal (lintas semua rekening) "
                "supaya gampang ditelusuri kronologis saat audit.")
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    # --- Bagian 1: pencocokan transfer antar rekening ---
    r = 4
    ws.cell(row=r, column=1, value="1. PENCOCOKAN TRANSFER ANTAR REKENING")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1

    headers = [
        "Rekening Asal", "Tanggal Asal", "Keterangan Asal", "Nominal Asal",
        "Rekening Tujuan", "Tanggal Tujuan", "Keterangan Tujuan", "Nominal Tujuan",
        "Selisih Tanggal (hari)", "Selisih Nominal (Rp)", "Confidence", "Alasan Audit",
        "Via Fliptech?", "Rekening Penanggung Biaya", "Biaya Admin Teridentifikasi (Rp)",
    ]
    hdr_row = r
    for i, h in enumerate(headers, start=1):
        ws.cell(row=hdr_row, column=i, value=h)
    style_header(ws, hdr_row, len(headers))
    r += 1

    for m in matches:
        ws.cell(row=r, column=1, value=m.src.sheet)
        ws.cell(row=r, column=2, value=f"='{m.src.sheet}'!$A${m.src.row}")
        ws.cell(row=r, column=2).number_format = DATE_FORMAT
        ws.cell(row=r, column=3, value=f"='{m.src.sheet}'!$B${m.src.row}")
        ws.cell(row=r, column=4, value=f"='{m.src.sheet}'!${'D' if m.src.debit else 'E'}${m.src.row}")
        ws.cell(row=r, column=4).number_format = NUMBER_FORMAT
        if m.dst:
            ws.cell(row=r, column=5, value=m.dst.sheet)
            ws.cell(row=r, column=6, value=f"='{m.dst.sheet}'!$A${m.dst.row}")
            ws.cell(row=r, column=6).number_format = DATE_FORMAT
            ws.cell(row=r, column=7, value=f"='{m.dst.sheet}'!$B${m.dst.row}")
            ws.cell(row=r, column=8, value=f"='{m.dst.sheet}'!${'D' if m.dst.debit else 'E'}${m.dst.row}")
            ws.cell(row=r, column=8).number_format = NUMBER_FORMAT
            ws.cell(row=r, column=9, value=m.date_diff)
            # Selisih Nominal dibuat rumus (bukan angka mati) supaya tetap
            # akurat kalau nominal di sheet rekening diedit ulang
            ws.cell(row=r, column=10, value=f"=ABS(ABS($D{r})-ABS($H{r}))")
            ws.cell(row=r, column=10).number_format = NUMBER_FORMAT
        else:
            ws.cell(row=r, column=5, value="(belum ditemukan)")
        ws.cell(row=r, column=11, value=m.confidence)
        ws.cell(row=r, column=12, value=m.reasoning)
        # kolom biaya admin (auto): dipakai Laporan Laba Rugi untuk membukukan
        # selisih transfer via Fliptech sebagai beban riil, bukan dibiarkan
        # menghilang sebagai selisih Neraca yang perlu intervensi manual
        ws.cell(row=r, column=13, value=m.is_fliptech if m.dst else False)
        if m.dst:
            fee_payer = m.src.sheet if m.src.debit else m.dst.sheet
            ws.cell(row=r, column=14, value=fee_payer)
            ws.cell(row=r, column=15,
                    value=f'=IF(AND($M{r}=TRUE,$H{r}<>""),ABS(ABS($D{r})-ABS($H{r})),0)')
        else:
            ws.cell(row=r, column=14, value="-")
            ws.cell(row=r, column=15, value=0)
        ws.cell(row=r, column=15).number_format = NUMBER_FORMAT
        fill = conf_fill(m.confidence)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(c == 12))
        ws.cell(row=r, column=11).fill = fill
        # tinggi baris eksplisit mengikuti jumlah baris teks Alasan Audit
        # (kolom L bisa multi-baris kalau ada kandidat dekat) - supaya
        # langsung kelihatan penuh saat dibuka, tidak perlu resize manual
        n_baris_teks = (m.reasoning or "").count("\n") + 1
        if n_baris_teks > 1:
            ws.row_dimensions[r].height = min(15 * n_baris_teks, 400)
        r += 1

    section1_data_start = hdr_row + 1
    section1_last_row = max(r - 1, section1_data_start)  # dipakai Laporan Laba Rugi (SUMIFS biaya admin)
    r += 1
    # --- Bagian 1b: transfer terpecah / tergabung (split & merge) ---
    ws.cell(row=r, column=1, value="1b. TRANSFER TERPECAH / TERGABUNG (SPLIT & MERGE)")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    headers1b = [
        "Rekening Asal", "Tanggal Asal", "Nominal Asal (Total)",
        "Pasangan 1", "Nominal 1", "Pasangan 2", "Nominal 2",
        "Selisih (Rp)", "Confidence", "Alasan Audit",
    ]
    hdr_row1b = r
    for i, h in enumerate(headers1b, start=1):
        ws.cell(row=hdr_row1b, column=i, value=h)
    style_header(ws, hdr_row1b, len(headers1b))
    r += 1
    for cm in combo_matches:
        src, a, b = cm["src"], cm["parts"][0], cm["parts"][1]
        ws.cell(row=r, column=1, value=src.sheet)
        ws.cell(row=r, column=2, value=f"='{src.sheet}'!$A${src.row}")
        ws.cell(row=r, column=2).number_format = DATE_FORMAT
        ws.cell(row=r, column=3, value=f"='{src.sheet}'!${'D' if src.debit else 'E'}${src.row}")
        ws.cell(row=r, column=3).number_format = NUMBER_FORMAT
        ws.cell(row=r, column=4, value=f"{a.sheet} (brs {a.row})")
        ws.cell(row=r, column=5, value=f"='{a.sheet}'!${'D' if a.debit else 'E'}${a.row}")
        ws.cell(row=r, column=5).number_format = NUMBER_FORMAT
        ws.cell(row=r, column=6, value=f"{b.sheet} (brs {b.row})")
        ws.cell(row=r, column=7, value=f"='{b.sheet}'!${'D' if b.debit else 'E'}${b.row}")
        ws.cell(row=r, column=7).number_format = NUMBER_FORMAT
        ws.cell(row=r, column=8, value=round(cm["diff"], 2))
        ws.cell(row=r, column=8).number_format = NUMBER_FORMAT
        ws.cell(row=r, column=9, value=cm["confidence"])
        ws.cell(row=r, column=10, value=cm["reasoning"])
        for c in range(1, len(headers1b) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(c == 10))
        ws.cell(row=r, column=9).fill = conf_fill(cm["confidence"])
        r += 1
    if not combo_matches:
        ws.cell(row=r, column=1, value="(tidak ada indikasi transfer terpecah/tergabung)")
        ws.cell(row=r, column=1).font = Font(italic=True, color="6B7280")
        r += 1

    r += 1
    # --- Bagian 2: minus / selisih kas perlu verifikasi ---
    ws.cell(row=r, column=1, value="2. INDIKASI MINUS / SELISIH KAS")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    headers2 = ["Rekening", "Tanggal", "Keterangan", "Nominal", "Saldo Kumulatif", "Alasan Perlu Verifikasi"]
    hdr_row2 = r
    for i, h in enumerate(headers2, start=1):
        ws.cell(row=hdr_row2, column=i, value=h)
    style_header(ws, hdr_row2, len(headers2))
    r += 1
    for t, reasons in minus_flags:
        ws.cell(row=r, column=1, value=t.sheet)
        ws.cell(row=r, column=2, value=f"='{t.sheet}'!$A${t.row}")
        ws.cell(row=r, column=2).number_format = DATE_FORMAT
        ws.cell(row=r, column=3, value=f"='{t.sheet}'!$B${t.row}")
        ws.cell(row=r, column=4, value=f"='{t.sheet}'!${'D' if t.debit else 'E'}${t.row}")
        ws.cell(row=r, column=4).number_format = NUMBER_FORMAT
        ws.cell(row=r, column=5, value=f"='{t.sheet}'!$F${t.row}")
        ws.cell(row=r, column=5).number_format = NUMBER_FORMAT
        ws.cell(row=r, column=6, value=" | ".join(reasons))
        for c in range(1, len(headers2) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(c == 6))
        ws.cell(row=r, column=6).fill = LOW_FILL
        r += 1

    r += 1
    # --- Bagian 3: status keseimbangan Neraca (dihitung Python, prediksi
    # sebelum file dibuka/di-recalculate Excel) - supaya begitu laporan
    # digenerate, status "sudah selesai/masih ada selisih" langsung
    # kelihatan di sini tanpa perlu buka sheet Neraca terpisah ---
    ws.cell(row=r, column=1, value="3. STATUS KESEIMBANGAN NERACA")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    if balance_status:
        headers3 = ["Rekening", "Total Aset", "Total Ekuitas", "Transfer Bersih", "Selisih", "Dari Kategori Baru", "Status"]
        hdr_row3 = r
        for i, h in enumerate(headers3, start=1):
            ws.cell(row=hdr_row3, column=i, value=h)
        style_header(ws, hdr_row3, len(headers3))
        r += 1
        any_selisih = False
        total_selisih = 0.0
        total_kategori_baru = 0.0
        total_n_kategori_baru = 0
        for sheet, s in balance_status.items():
            balanced = abs(s["selisih"]) < 1  # toleransi Rp1 (noise pembulatan)
            if not balanced:
                any_selisih = True
            total_selisih += s["selisih"]
            total_kategori_baru += s.get("kategori_baru_total", 0)
            total_n_kategori_baru += s.get("n_kategori_baru", 0)
            ws.cell(row=r, column=1, value=sheet)
            ws.cell(row=r, column=2, value=s["total_aset"])
            ws.cell(row=r, column=3, value=s["ekuitas"])
            ws.cell(row=r, column=4, value=s["transfer_bersih"])
            ws.cell(row=r, column=5, value=s["selisih"])
            n_kb = s.get("n_kategori_baru", 0)
            ws.cell(row=r, column=6, value=f"Rp{s.get('kategori_baru_total', 0):,.0f} ({n_kb} transaksi)".replace(",", ".") if n_kb else "-")
            ws.cell(row=r, column=7, value="Balanced" if balanced else f"ADA SELISIH Rp{abs(s['selisih']):,.0f}".replace(",", "."))
            for c in (2, 3, 4, 5):
                ws.cell(row=r, column=c).number_format = NUMBER_FORMAT
            for c in range(1, len(headers3) + 1):
                ws.cell(row=r, column=c).border = BORDER
            ws.cell(row=r, column=7).fill = HIGH_FILL if balanced else LOW_FILL
            ws.cell(row=r, column=7).font = Font(bold=True)
            r += 1
        r += 1
        if any_selisih:
            kb_note = ""
            if total_n_kategori_baru:
                kb_amount_fmt = f"{abs(total_kategori_baru):,.0f}".replace(",", ".")
                kb_note = (
                    f" Dari selisih ini, Rp{kb_amount_fmt} berasal dari {total_n_kategori_baru} "
                    "transaksi 'Kategori Baru' yang SENGAJA dikeluarkan dari perhitungan Laba Rugi/Neraca "
                    "(belum jelas kategorinya, harus diaudit dulu - lihat bagian 4 di bawah) - bukan bug, ini "
                    "dibiarkan tidak balance supaya selalu terlihat masih ada yang perlu diaudit."
                )
            ws.cell(row=r, column=1,
                    value=(f"BELUM SELESAI: masih ada selisih total Rp{abs(total_selisih):,.0f} yang belum "
                           "terjelaskan (lihat kolom Selisih per rekening di atas)."
                           .replace(",", ".") + kb_note))
            ws.cell(row=r, column=1).font = Font(bold=True, color="B91C1C")
        else:
            ws.cell(row=r, column=1, value="SELESAI: semua rekening balanced, tidak ada selisih yang perlu ditelusuri lebih lanjut.")
            ws.cell(row=r, column=1).font = Font(bold=True, color="15803D")
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(headers3))
        ws.row_dimensions[r].height = 40
        r += 1
    else:
        ws.cell(row=r, column=1, value="(status keseimbangan tidak dihitung)")
        ws.cell(row=r, column=1).font = Font(italic=True, color="6B7280")
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="4. TRANSAKSI KATEGORI BARU (perlu diaudit)")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    if new_category_flags:
        ws.cell(row=r, column=1, value=(
            "Kategori aslinya tidak dikenal sistem sama sekali (bukan kesalahan pencocokan, bukan yang "
            "sudah dikenal seperti Tip/Minus) - ditandai 'Kategori Baru' di kolom M sheet rekening masing-"
            "masing supaya kelihatan jelas dan tetap terhitung di Laba Rugi (bukan diam-diam hilang). "
            "Audit tiap baris: kategorikan manual di sumber data, atau minta tambahkan aturan baru."
        ))
        ws.cell(row=r, column=1).font = Font(italic=True, size=9, color="6B7280")
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 40
        r += 1
        headers4 = ["Rekening", "Tanggal", "Keterangan", "Kategori Asli", "Objek", "Nominal"]
        hdr_row4 = r
        for i, h in enumerate(headers4, start=1):
            ws.cell(row=hdr_row4, column=i, value=h)
        style_header(ws, hdr_row4, len(headers4))
        r += 1
        for t in sorted(new_category_flags, key=lambda t: _sort_key(t.date)):
            ws.cell(row=r, column=1, value=t.sheet)
            ws.cell(row=r, column=2, value=coerce_date(t.date))
            ws.cell(row=r, column=3, value=t.desc)
            ws.cell(row=r, column=4, value=t.kategori)
            ws.cell(row=r, column=5, value=t.objek)
            ws.cell(row=r, column=6, value=t.nominal)
            ws.cell(row=r, column=2).number_format = "dd/mm/yyyy"
            ws.cell(row=r, column=6).number_format = NUMBER_FORMAT
            for c in range(1, len(headers4) + 1):
                ws.cell(row=r, column=c).border = BORDER
                ws.cell(row=r, column=c).fill = MED_FILL
            r += 1
    else:
        ws.cell(row=r, column=1, value="Tidak ada transaksi berkategori baru bulan ini.")
        ws.cell(row=r, column=1).font = Font(italic=True, color="6B7280")
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="5. PENGELUARAN PRIBADI (perlu diverifikasi manual)")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    if personal_expense_flags:
        ws.cell(row=r, column=1, value=(
            "Dikeluarkan dari Laba Rugi (diperlakukan seperti prive/penarikan modal - mengurangi Ekuitas "
            "di Neraca, BUKAN beban bisnis), TAPI setiap transaksi berikut perlu diverifikasi manual satu-"
            "satu - pastikan memang benar pengeluaran pribadi owner, bukan salah kategori."
        ))
        ws.cell(row=r, column=1).font = Font(italic=True, size=9, color="6B7280")
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 40
        r += 1
        headers5 = ["Rekening", "Tanggal", "Keterangan", "Kategori Asli", "Objek", "Nominal"]
        hdr_row5 = r
        for i, h in enumerate(headers5, start=1):
            ws.cell(row=hdr_row5, column=i, value=h)
        style_header(ws, hdr_row5, len(headers5))
        r += 1
        for t in sorted(personal_expense_flags, key=lambda t: _sort_key(t.date)):
            ws.cell(row=r, column=1, value=t.sheet)
            ws.cell(row=r, column=2, value=coerce_date(t.date))
            ws.cell(row=r, column=3, value=t.desc)
            ws.cell(row=r, column=4, value=t.kategori)
            ws.cell(row=r, column=5, value=t.objek)
            ws.cell(row=r, column=6, value=t.nominal)
            ws.cell(row=r, column=2).number_format = "dd/mm/yyyy"
            ws.cell(row=r, column=6).number_format = NUMBER_FORMAT
            for c in range(1, len(headers5) + 1):
                ws.cell(row=r, column=c).border = BORDER
                ws.cell(row=r, column=c).fill = TRANSFER_MATCH_FILL
            r += 1
    else:
        ws.cell(row=r, column=1, value="Tidak ada transaksi Pengeluaran Pribadi bulan ini.")
        ws.cell(row=r, column=1).font = Font(italic=True, color="6B7280")
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="6. KEMUNGKINAN HUTANG BARU (perlu ditambahkan ke Buku Hutang)")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    if new_debt_flags:
        ws.cell(row=r, column=1, value=(
            "Transaksi berikut menyebut kata kunci pencairan hutang/pinjaman baru (bukan cicilan) - "
            "kategorinya SUDAH otomatis benar (Hutang Masuk, masuk Liabilitas di Neraca), TAPI "
            "reconcile.py bulanan TIDAK punya Buku Hutang sendiri (itu cuma ada di laporan kuartal/"
            "tahunan). Salin detail di bawah ini jadi baris baru di sheet 'Buku Hutang' pada laporan "
            "kuartal/tahunan berikutnya - Nilai Pinjaman diisi manual (bisa beda dari Nilai Diterima "
            "kalau ada potongan biaya di awal). Pemberi pinjaman diambil dari kolom Objek (konvensi bot "
            "konversi: Keterangan 'Hutang Baru - <Objek>', nama pemberi hutang ADA DI OBJEK, bukan di "
            "Keterangan langsung)."
        ))
        ws.cell(row=r, column=1).font = Font(italic=True, size=9, color="6B7280")
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 68
        r += 1
        headers6 = ["Rekening", "Tanggal Pinjam", "Keterangan", "Pemberi Pinjaman (dari Objek)", "Nilai Diterima", "Subjek"]
        hdr_row6 = r
        for i, h in enumerate(headers6, start=1):
            ws.cell(row=hdr_row6, column=i, value=h)
        style_header(ws, hdr_row6, len(headers6))
        r += 1
        for t in sorted(new_debt_flags, key=lambda t: _sort_key(t.date)):
            ws.cell(row=r, column=1, value=t.sheet)
            ws.cell(row=r, column=2, value=coerce_date(t.date))
            ws.cell(row=r, column=3, value=t.desc)
            ws.cell(row=r, column=4, value=t.objek)
            ws.cell(row=r, column=5, value=t.nominal)
            ws.cell(row=r, column=6, value=t.subjek)
            ws.cell(row=r, column=2).number_format = "dd/mm/yyyy"
            ws.cell(row=r, column=5).number_format = NUMBER_FORMAT
            for c in range(1, len(headers6) + 1):
                ws.cell(row=r, column=c).border = BORDER
                ws.cell(row=r, column=c).fill = TRANSFER_MATCH_FILL
            r += 1
    else:
        ws.cell(row=r, column=1, value="Tidak ada transaksi yang menyebut pencairan hutang/pinjaman baru bulan ini.")
        ws.cell(row=r, column=1).font = Font(italic=True, color="6B7280")
        r += 1

    widths = [22, 12, 26, 14, 22, 12, 26, 14, 10, 14, 14, 40, 12, 22, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A6"
    return ws, {"data_start": section1_data_start, "data_end": section1_last_row}


# ---------------------------------------------------------------------------
# Penambahan kolom bantu "Nominal Bersih" di tiap sheet rekening
# (rumus, dipakai sebagai basis SUMIF di laporan keuangan)
# ---------------------------------------------------------------------------

def add_helper_column(ws, last_row):
    """Tambah 3 kolom bantu berbasis rumus (bukan nilai mati):
    J = Nominal Bersih (dipakai basis SUMIF laporan keuangan)
    K = Saldo Kumulatif Rekonstruksi (dihitung ulang dari saldo awal +
        akumulasi nominal, sehingga selalu terisi walau kolom F/Saldo
        Kumulatif aslinya bolong-bolong di sebagian baris)
    L = Selisih vs Saldo Tercatat (alat bantu telusur/audit: harus 0 setiap
        kali kolom F terisi; kalau tidak 0, baris-baris sebelumnya di sheet
        ini kemungkinan TIDAK berurutan secara kronologis terhadap kolom
        Saldo Kumulatif aslinya - baris dengan selisih besar ditandai warna
        kuning otomatis, dan dirangkum di sheet Diagnostik Keseimbangan)
    """
    for col, title in ((10, "Nominal Bersih (Debit atau Kredit)"),
                       (11, "Saldo Kumulatif (Rekonstruksi)"),
                       (12, "Selisih vs Saldo Tercatat")):
        cell = ws.cell(row=1, column=col, value=title)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    for row in range(2, last_row + 1):
        # N() memperlakukan sel kosong sebagai 0 tanpa salah mengira nilai
        # 0 eksplisit sebagai "kosong" (beda dengan tes teks <>"")
        ws.cell(row=row, column=10, value=f"=N($D{row})+N($E{row})")
        if row == 2:
            ws.cell(row=row, column=11, value="=$F$2")
        else:
            ws.cell(row=row, column=11, value=f"=$K{row - 1}+$J{row}")
        ws.cell(row=row, column=12,
                value=f'=IF($F{row}<>"",$K{row}-$F{row},0)')

    for col, width in ((10, 26), (11, 28), (12, 22)):
        ws.column_dimensions[get_column_letter(col)].width = width

    # highlight visual: baris dengan penyimpangan signifikan (>Rp1.000)
    # antara saldo rekonstruksi dan saldo tercatat - penanda cepat untuk
    # menelusuri baris mana yang bikin data tidak berurutan/tidak konsisten
    from openpyxl.formatting.rule import CellIsRule
    rng = f"L2:L{last_row}"
    ws.conditional_formatting.add(
        rng, CellIsRule(operator="greaterThan", formula=["1000"], fill=MED_FILL)
    )


def add_effective_category_column(ws, txns):
    """Kolom bantu M = Kategori Efektif - Kategori ASLI (kolom C), KECUALI
    ada aturan override yang cocok (lihat Txn.category_override/
    CATEGORY_OVERRIDE_RULES) - ditulis sebagai NILAI (bukan rumus, karena
    logikanya melibatkan pencarian kata kunci yang jauh lebih mudah
    dilakukan di Python daripada rumus Excel murni). SEMUA rumus SUMIF/
    SUMIFS kategori di laporan keuangan (Laba Rugi, Neraca, Arus Kas, dst)
    merujuk ke kolom INI, bukan langsung ke kolom C, supaya override
    konsisten di semua laporan tanpa perlu duplikasi logika di tiap
    rumus. Kolom C tetap dibiarkan apa adanya (data asli, untuk audit)."""
    cell = ws.cell(row=1, column=13, value="Kategori Efektif (setelah override kata kunci - dipakai semua rumus)")
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    for t in txns:
        ws.cell(row=t.row, column=13, value=t.effective_kategori)
    ws.column_dimensions[get_column_letter(13)].width = 34


def _sender_receiver(t1, t2):
    """Tentukan mana PENGIRIM (nominal negatif/debit - uang keluar dari
    rekening ini) dan mana PENERIMA (nominal positif/kredit - uang
    masuk ke rekening ini) dari sepasang transaksi transfer yang sudah
    matched. PENTING: find_matches() TIDAK menjamin src=pengirim/
    dst=penerima - src/dst di sana cuma menandai "transaksi mana yang
    mulai dicari" vs "pasangannya yang ditemukan", BISA JADI src adalah
    sisi kredit (penerima) kalau itu yang lebih dulu diproses. Salah
    asumsi arah di sini akan menulis Subjek/Objek TERBALIK.

    Kalau tanda SAMA (same_sign - kasus tak lazim, mis. kedua sisi sama-
    sama tercatat kredit karena pembukuan ganda yang tidak standar),
    tidak ada dasar objektif untuk menentukan arah - kembalikan urutan
    asli (t1, t2) sebagai fallback, tidak dianggap error."""
    if t1.nominal < 0 and t2.nominal > 0:
        return t1, t2
    if t2.nominal < 0 and t1.nominal > 0:
        return t2, t1
    return t1, t2


BCA_MATCH_FILL = PatternFill("solid", fgColor="BDD7EE")  # biru
JAGO_MATCH_FILL = PatternFill("solid", fgColor="FCD9B6")  # orange
BRI_MATCH_FILL = PatternFill("solid", fgColor="FFF2A8")  # kuning
REKONLOKAL_UNMATCHED_FILL = PatternFill("solid", fgColor="FFC7CE")  # merah
REKONLOKAL_MEDIUM_CONFIDENCE_FILL = PatternFill("solid", fgColor="FFE5E8")  # merah shade lebih muda, untuk match confidence Medium
REKONLOKAL_SUSPECT_CATEGORY_FILL = PatternFill("solid", fgColor="8B0000")  # merah darah, mencolok

# Warna highlight yang PUNYA ARTI KHUSUS (arah transfer/belum direkon/
# kategori mencurigakan) - dipakai untuk cek "apakah baris ini SUDAH
# ada highlight yang lebih penting" sebelum menimpa dengan warna
# kelompok kategori yang sifatnya cuma visual. SENGAJA tidak mengecek
# "fill APAPUN yang bukan kosong" - banyak file sumber punya pewarnaan
# zebra-stripe/formatting bawaan yang TIDAK ADA ARTINYA untuk sistem
# ini, kalau itu ikut dianggap "sudah ada highlight", warna kelompok
# kategori jadi selang-seling tidak konsisten (kena di sebagian baris,
# tidak kena di baris lain kategori yang SAMA, cuma karena baris itu
# kebetulan sudah punya warna latar dari sumbernya).
_MEANINGFUL_HIGHLIGHT_HEXES = {
    "00BDD7EE", "00FCD9B6", "00FFF2A8", "00FFC7CE", "008B0000", "00FFE5E8",
}

# Highlight berdasarkan GRUP KATEGORI Layer 1 - dipakai di /rekonlokal
# MAUPUN flow rekonsiliasi utama (bukan cuma satu tempat), supaya
# konsisten dimanapun user melihat datanya. Beda dari highlight transfer
# (biru/orange/kuning - arah uang) dan highlight audit (merah/ungu -
# butuh perhatian) - ini murni pengelompokan VISUAL kategori P&L supaya
# gampang di-scan sekilas, tidak ada makna "butuh tindakan" apapun.
CATEGORY_GROUP_FILL_MAP = {
    "penjualan": PatternFill("solid", fgColor="C6EFCE"),  # hijau
    "penjualan shopeefood": PatternFill("solid", fgColor="C6EFCE"),
    "penjualan grabfood": PatternFill("solid", fgColor="C6EFCE"),
    "belanja bahan": PatternFill("solid", fgColor="E4C9A0"),  # coklat/tan
    "kemasan": PatternFill("solid", fgColor="E4C9A0"),
    "overhead": PatternFill("solid", fgColor="D3D3D3"),  # abu-abu
    "subscription": PatternFill("solid", fgColor="D3D3D3"),
    "belanja assets": PatternFill("solid", fgColor="FFD966"),  # emas
    "belanja utilitas": PatternFill("solid", fgColor="B2EBF2"),  # cyan/teal
    "tools dan equipments": PatternFill("solid", fgColor="B2EBF2"),
    "sewa dan maintenance bangunan": PatternFill("solid", fgColor="F4B6B0"),  # salmon
    "reparasi dan maintenance tools dan mesin": PatternFill("solid", fgColor="F4B6B0"),
    "pajak dan administrasi": PatternFill("solid", fgColor="D9D9A3"),  # olive/khaki
    "biaya admin bank": PatternFill("solid", fgColor="D9D9A3"),
    "marketing": PatternFill("solid", fgColor="F7C9DE"),  # pink
    "riset dan development": PatternFill("solid", fgColor="F7C9DE"),
    "konsumsi dan liburan": PatternFill("solid", fgColor="F7C9DE"),
    "pengeluaran pribadi": PatternFill("solid", fgColor="E8AB8C"),  # terracotta
}


def category_group_fill(kategori):
    """Warna highlight grup kategori untuk 'kategori' yang diberikan -
    None kalau kategori ini bukan bagian dari grup manapun (mis. Kas/
    Transaksi Internal/Modal/dst - tidak semua kategori didefinisikan
    perlu warna kelompok). Semua kategori 'Gaji*' (Gaji Bulan Ini/Gaji
    Accrual/pola apapun yang diawali 'gaji') dapat SATU warna yang sama,
    dicek terpisah dari peta di atas karena namanya dinamis (bukan
    string tetap)."""
    k = (kategori or "").strip().lower()
    if k.startswith("gaji"):
        return PatternFill("solid", fgColor="FFCC99")  # peach
    return CATEGORY_GROUP_FILL_MAP.get(k)


def _bank_group_fill(sheet_title):
    """Tentukan warna highlight berdasarkan grup bank tujuan - BCA (semua
    varian: BCA-887/BCA-417/BCA-292(Biz)) = biru, Jago = orange,
    BRI (BRI-507/BRI-567(Biz)) DAN BSI (semua varian) = kuning. Rekening
    lain (Kas-Buku, dst) -> None (tidak ada warna khusus, sesuai
    permintaan user cuma kelompok-kelompok ini)."""
    n = sheet_title.lower()
    if n.startswith("bca"):
        return BCA_MATCH_FILL
    if n.startswith("jago"):
        return JAGO_MATCH_FILL
    if n.startswith("bri") or n.startswith("bsi"):
        return BRI_MATCH_FILL
    return None


def _display_account_name(sheet_title):
    """Nama rekening untuk DITULIS ke sel Subjek/Objek - "Kas Buku"/
    "Kas/Buku" (nama rekening kas internal, hasil tebakan
    _infer_account_name atau filename_hint) diseragamkan jadi "Kas
    Kasir" untuk konsistensi penamaan (permintaan eksplisit user),
    rekening lain (BRI-507/BCA-887/dst) dipakai apa adanya. HANYA
    dipakai di titik PENULISAN ke sel - t.sheet sendiri (dipakai untuk
    matching/deduplikasi/key dict) TIDAK diubah, supaya tidak
    mengganggu logika pencocokan yang bergantung pada identitas asli."""
    if sheet_title.strip().lower().startswith("kas"):
        return "Kas Kasir"
    return sheet_title


def _settled_to_bank_label(sheet_title):
    """Label 'Settled to <bank>' untuk transaksi Penjualan - nama bank
    diekstrak dari identitas rekening (token pertama sebelum '-'/'(',
    di-uppercase, mis. 'BRI-507' -> 'BRI', 'BCA-887' -> 'BCA'); untuk
    rekening kas (bukan bank sungguhan) dipakai _display_account_name
    ('Kas Kasir') apa adanya, bukan 'Settled to Kas Kasir' yang janggal."""
    disp = _display_account_name(sheet_title)
    if disp == "Kas Kasir":
        return disp
    bank_prefix = sheet_title.split("-")[0].split("(")[0].strip().upper()
    return f"Settled to {bank_prefix}"


# Vendor Belanja Bahan/Kemasan yang keterangannya sering ditulis beda-
# beda di sumber (typo/singkatan/variasi ejaan) - diseragamkan di
# /rekonlokal jadi SATU nama baku per vendor, sekaligus dipastikan
# kategorinya benar. Urutan dalam grup: dari yang PALING SPESIFIK ke
# yang paling umum (dicek pakai batas kata/regex, bukan substring
# polos, supaya kata pendek seperti "SB"/"Pasar" tidak salah tangkap
# teks lain yang kebetulan mengandungnya).
_KAS_BUKU_VENDOR_RULES = [
    (["belanja sb", "sinar bahagia", "sb"], "Belanja Bahan Sinar Bahagia", "Belanja Bahan", "Sinar Bahagia"),
    (["belanja amanah", "amanah"], "Belanja Bahan Amanah", "Belanja Bahan", "Amanah"),
    (["belanja fadhilah", "fadhilah"], "Belanja Bahan Fadhilah", "Belanja Bahan", "Fadhilah"),
    (["belanja primer raya", "primer raya", "primer"], "Belanja Bahan Primer", "Belanja Bahan", "Primer"),
    (["belanja pasar", "pasar"], "Belanja Bahan Pasar", "Belanja Bahan", "Pasar"),
    (["belanja abadi", "abadi"], "Belanja Bahan Pasar", "Belanja Bahan", "Pasar"),
    (["dinda food", "dinda frozen", "belanja dinda"], "Belanja Bahan Dinda Food and Frozen", "Belanja Bahan", "Dinda Food and Frozen"),
    (["belanja mak opik", "mak opik", "mak opi"], "Belanja Bahan Mak Opik", "Belanja Bahan", "Mak Opik"),
    (["galon", "cleo"], "Belanja Bahan Air Mineral", "Belanja Bahan", "Air Mineral"),
    (["belanja arumi", "arumi"], "Belanja Bahan Arumi", "Kemasan", "Arumi"),
    (["masuya"], "Belanja Bahan UHT dan Pasta", "Belanja Bahan", "Masuya"),
    (["pembayaran briva ke tokopedia", "tokopedia"], "Belanja Tokopedia", "Belanja Bahan", "Tokopedia"),
    (["mira laundry"], "Mira Laundry", "Overhead", "Mira Laundry"),
    (["tomoro coffee", "tomoro"], "Rapat di Tomoro Coffee", "Belanja Bahan", "Tomoro Coffee"),
    (["sukanda jaya", "diamond fair", "sukanda"], "Produk Bahan Sukanda", "Belanja Bahan", "Sukanda"),
    (["shopeepay", "shopee pay", "shopee"], "Belanja Bahan Shopee", "Belanja Bahan", "Shopee"),
    (["iklan tiktok", "tiktok ads", "tiktok"], "TikTok", "Marketing", "TikTok"),
    (["iklan facebook", "facebook ads", "fb ads", "meta ads", "facebook"], "MetaAds", "Marketing", "MetaAds"),
    (["dompet anak bangsa"], "Gopay", "Overhead", "Gopay"),
]

# Vendor yang trigger-nya SPESIFIK dari kolom Objek (bukan Keterangan) -
# supplier/toko sering muncul di Objek, sementara Keterangan-nya sendiri
# masih generik ("Belanja Operasional"/"Belanja Bahan") dan tidak
# menyebut nama vendornya sama sekali. Format sama seperti
# _KAS_BUKU_VENDOR_RULES (keyword_objek, keterangan_baru, kategori_baru,
# objek_baru) - keyword dicek pakai batas kata terhadap Objek transaksi.
_OBJEK_VENDOR_RULES = [
    (["nanda audia agusti", "nanda audia agustin", "kliffer plastik"], "Kemasan Kliffer", "Kemasan", "Kliffer Plastik"),
    (["madam baha", "madam bahan kue", "toko madam"], "Bahan Kue", "Belanja Bahan", "Toko Madam"),
    (["dompet anak bangsa"], "Gopay", "Overhead", "Gopay"),
    (["yulia indah pratiwi", "yulia indah pratiw", "anugerah plastik"], "Kemasan Anugerah", "Kemasan", "Anugerah Plastik"),
    (["mira laundry"], "Mira Laundry", "Overhead", "Mira Laundry"),
    (["tomoro coffee", "tomoro"], "Rapat di Tomoro Coffee", "Belanja Bahan", "Tomoro Coffee"),
    (["sukanda jaya", "diamond fair", "sukanda"], "Produk Bahan Sukanda", "Belanja Bahan", "Sukanda"),
    (["shopeepay", "shopee pay"], "Belanja Bahan Shopee", "Belanja Bahan", "Shopee"),
    (["iklan tiktok", "tiktok ads", "tiktok"], "TikTok", "Marketing", "TikTok"),
    (["iklan facebook", "facebook ads", "fb ads", "meta ads", "facebook"], "MetaAds", "Marketing", "MetaAds"),
    (["rasbani"], "Konsumsi Internal Waroeng Rasbani", "Konsumsi dan Liburan", "Waroeng Rasbani"),
    (["biaya transfer keluar biaya"], "Biaya Transfer Keluar", "Biaya Admin Bank", "Biaya Admin Bank"),
    (["adobe"], "Adobe", "Subscription", "Adobe"),
    (["ace team hq i", "ace team"], "Rapat di Ace Team", "Pengeluaran Pribadi", "Ace Team"),
    (["kava coffee"], "Rapat di Kava Coffee", "Pengeluaran Pribadi", "Kava Coffee"),
    # ASUMSI (perlu dikonfirmasi user): kategori Doto & Gairah disamakan
    # dengan Ace Team/Kava Coffee ("Pengeluaran Pribadi") karena sama-sama
    # pola "Rapat di [tenant]" - belum ada kepastian kategori aslinya.
    (["doto"], "Rapat di Doto", "Pengeluaran Pribadi", "Doto"),
    (["gairah"], "Rapat di Gairah", "Pengeluaran Pribadi", "Gairah"),
    (["tokopedia"], "Belanja Tokopedia", "Belanja Bahan", "Tokopedia"),
    (["samsul padli"], "Mess Karyawan", "Overhead", "Samsul Padli"),
    (["muhammad zulfadli"], "Bahan Bangunan", "Sewa dan Maintenance Bangunan", "Muhammad Zulfadli"),
    (["muhammad umar al-khatib", "muhammad umar al khatib", "muhammad umar"], "Bahan Bangunan", "Sewa dan Maintenance Bangunan", "Muhammad Umar Al-Khatib"),
    (["angga eka"], "Furniture", "Belanja Assets", "Angga Eka"),
    (["asrul yusuf"], "Belanja Bahan Asrul Yusuf", "Belanja Bahan", "Asrul Yusuf"),
    (["dapoer ibu fenny", "dapur ibu fenny"], "Konsumsi Internal Dapoer Ibu Fenny", "Konsumsi dan Liburan", "Dapoer Ibu Fenny"),
    (["arafat bahaswen"], "Bahan Bangunan", "Sewa dan Maintenance Bangunan", "Arafat Bahaswen"),
    (["shopee"], "Belanja Bahan Shopee", "Belanja Bahan", "Shopee"),
    (["saddam"], "Bahan Bangunan", "Sewa dan Maintenance Bangunan", "Saddam"),
    (["pt yaoya berkat sejati", "yaoya berkat sejati", "yaoya"], "Bahan Kue", "Belanja Bahan", "Yaoya"),
    (["yudi haryono"], "Reparasi Tools Listrik", "Reparasi dan Maintenance Tools dan Mesin", "Yudi Haryono"),
    (["oriel chicken"], "Konsumsi Internal Oriel Chicken", "Konsumsi dan Liburan", "Oriel Chicken"),
    # ENTRY BARU (belum pernah ada sebelumnya) - dikonfirmasi user, sama
    # treatment-nya dengan vendor bahan bangunan lain di atas.
    (["mitra10", "mitra 10"], "Bahan Bangunan", "Sewa dan Maintenance Bangunan", "Mitra10 Bangunan"),
    (["depo bangunan"], "Bahan Bangunan", "Sewa dan Maintenance Bangunan", "Depo Bangunan"),
]


def _objek_vendor_info(t):
    """Sama seperti _kas_buku_vendor_info, tapi mengecek kolom Objek
    (bukan Keterangan) terhadap _OBJEK_VENDOR_RULES. Return
    (keterangan_baru, kategori_baru, objek_baru) atau None."""
    text = (t.objek or "").lower()
    for keywords, keterangan_baru, kategori_baru, objek_baru in _OBJEK_VENDOR_RULES:
        for kw in keywords:
            if re.search(r"\b" + re.escape(kw) + r"\b", text):
                return keterangan_baru, kategori_baru, objek_baru
    return None



def _kas_buku_vendor_info(t):
    """Kalau Keterangan/Objek transaksi ini mengandung salah satu kata
    kunci vendor yang dikenal (lihat _KAS_BUKU_VENDOR_RULES), kembalikan
    (keterangan_baru, kategori_baru, objek_baru_atau_None) - None kalau
    tidak ada yang cocok. objek_baru None berarti Objek TIDAK disentuh
    (kebanyakan vendor tidak perlu). Pencocokan pakai batas kata (regex
    \\b), bukan substring polos - penting untuk kata pendek seperti
    "SB"/"Pasar" yang berisiko salah tangkap kalau cuma dicek 'in'
    biasa."""
    text = (t.desc or "").lower()
    for keywords, keterangan_baru, kategori_baru, objek_baru in _KAS_BUKU_VENDOR_RULES:
        for kw in keywords:
            if re.search(r"\b" + re.escape(kw) + r"\b", text):
                return keterangan_baru, kategori_baru, objek_baru
    return None


def _gaji_rekon_lokal_info(t):
    """Kalau transaksi ini kemungkinan besar Gaji (Kategori mengandung
    kata 'gaji'), kembalikan (nama_depan, nama_bulan, tahun,
    is_bulan_ini) - None kalau bukan transaksi gaji atau tanggalnya
    tidak terbaca.

    Bulan gaji ditentukan (sesuai kontrak kategori v3 §3):
    1. Kalau ada teks eksplisit "Gaji <Bulan>" di Keterangan/Keterangan
       Tambahan (paling bisa diandalkan, ambil apa adanya dari sumber).
    2. Kalau tidak ada, pakai tanggal transaksi: tanggal <=15 (awal
       bulan) -> gaji bulan SEBELUMNYA (accrual, telat dibayar),
       tanggal >15 -> gaji bulan transaksi itu sendiri.

    Nama depan diambil dari kata pertama di Objek."""
    if "gaji" not in (t.kategori or "").lower():
        return None
    tgl = coerce_date(t.date)
    if tgl is None:
        return None
    objek = (t.objek or "").strip()
    objek_kata = objek.split()
    # Dua kata PERTAMA dari Objek (bukan cuma satu) - supaya kasus
    # seperti "Baiq Sabrina" vs "Baiq Widiani" otomatis kebeda dari
    # namanya sendiri, tanpa perlu nunggu ketemu kasus ambigu dulu satu
    # per satu. Kalau Objek CUMA satu kata (mis. "Baiq" doang, tanpa
    # nama belakang di data sumber sama sekali), tidak ada kata kedua
    # untuk diambil - baru di titik itu cek AMBIGUOUS_FIRST_NAMES.
    nama_depan = " ".join(w.capitalize() for w in objek_kata[:2]) if objek_kata else "Pegawai"
    if len(objek_kata) <= 1 and nama_depan.lower() in AMBIGUOUS_FIRST_NAMES:
        # Nama depan ini dipakai LEBIH DARI SATU pegawai, dan data
        # sumbernya (Objek) cuma satu kata ini saja - tidak ada nama
        # belakang/info lain untuk membedakan. Tandai eksplisit, jangan
        # dipaksakan ke salah satu pegawai (lihat AMBIGUOUS_FIRST_NAMES).
        nama_depan = f"{nama_depan} (?)"

    text = f"{t.desc or ''} {t.ket or ''}".lower()
    bulan_idx = None
    for i, nama_bulan in enumerate(MONTHS_ID):
        if i == 0:
            continue
        if re.search(rf"\bgaji\s+{nama_bulan.lower()}\b", text):
            bulan_idx = i
            break

    if bulan_idx is not None:
        # tahun: kalau bulan gaji ini "lebih besar" dari bulan transaksi
        # (mis. transaksi Januari tapi teks bilang "Gaji Desember"),
        # berarti tahun sebelumnya - selain itu tahun sama dgn transaksi.
        tahun = tgl.year - 1 if bulan_idx > tgl.month else tgl.year
    elif tgl.day <= 15:
        bulan_idx = tgl.month - 1 if tgl.month > 1 else 12
        tahun = tgl.year if tgl.month > 1 else tgl.year - 1
    else:
        bulan_idx = tgl.month
        tahun = tgl.year

    is_bulan_ini = (bulan_idx == tgl.month and tahun == tgl.year)
    return nama_depan, MONTHS_ID[bulan_idx], tahun, is_bulan_ini


def _looks_like_account_identifier(v):
    """True kalau v terlihat seperti identitas rekening/kode bank (ada
    digit, seperti 'BRI-507'/'BCA-887'/'BSI-288', atau menyebut kata
    kunci rekening umum seperti 'jago'/'kas'/'buku'), BUKAN nama
    vendor/pelanggan/pegawai biasa. Dipakai _infer_account_name untuk
    MEMPRIORITASKAN kandidat yang genuinely kemungkinan rekening,
    bukan cuma yang paling sering disebut - supaya nama vendor yang
    kebetulan sering muncul di transaksi (mis. 'SHOPEE' tempat belanja
    bahan online) tidak salah terpilih jadi identitas rekening, kalah
    dari rekening lawan transaksi yang genuinely lebih jarang disebut
    eksplisit di data."""
    if any(ch.isdigit() for ch in v):
        return True
    v_lower = v.lower()
    return any(tok in v_lower for tok in ("jago", "kas", "buku", "mandiri", "bni", "bri", "bca", "bsi"))


_MONTH_NAMES_LOWER = {m.lower() for m in ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                                            "Juli", "Agustus", "September", "Oktober", "November", "Desember"]}


def _account_name_from_filename(filename):
    """Ekstrak kemungkinan nama rekening dari nama file ASLI (mis.
    'BSI-288_Oktober_2024.xlsx' -> 'BSI-288', 'BSI-288_Oktober_2024
    (1).xlsx' atau 'BSI-288_Oktober_2024__1_.xlsx' -> 'BSI-288', buang
    embel-embel duplikat upload) - dipakai sebagai SINYAL TAMBAHAN
    untuk _infer_account_name kalau data transaksi file itu sendiri
    genuinely TIDAK PERNAH menyebut nama rekeningnya sendiri di kolom
    Subjek/Objek manapun (semua transaksi di file itu merujuk pihak
    LAIN - vendor/pelanggan/rekening lawan - bukan dirinya sendiri,
    kasus nyata: file BSI-288 yang isinya semua transaksi ke vendor
    luar, kata 'BSI-288' sendiri tidak pernah muncul sebagai Subjek/
    Objek di baris manapun). Return None kalau filename tidak
    mengikuti pola yang bisa diekstrak."""
    if not filename:
        return None
    stem = re.sub(r"\.xlsx?$", "", filename, flags=re.IGNORECASE)
    stem = re.sub(r"[\s_]*\(\d+\)$", "", stem)  # " (1)" gaya browser
    stem = re.sub(r"__\d+_$", "", stem)  # "__1_" gaya Telegram
    parts = re.split(r"[_\s]+", stem)
    kept = [p for p in parts
            if p and p.lower() not in _MONTH_NAMES_LOWER and not re.fullmatch(r"(19|20)\d{2}", p)]
    return " ".join(kept).strip() or None


def _infer_account_name(txns, filename_hint=None):
    """Tebak nama/kode rekening dari data transaksinya sendiri - ambil
    nilai Subjek/Objek (gabungan) yang PALING SERING muncul, kecuali
    placeholder ('-'), nama sheet generik ('Mutasi'), dan 'Tenant Lain'
    (penanda bot konversi kalau Objek/Subjek genuinely TIDAK DIKETAHUI,
    BUKAN nama rekening/entitas sungguhan - kalau ikut dihitung, bisa
    kebetulan jadi nilai PALING SERING muncul di file yang banyak
    transaksi tak dikenal objeknya, salah menggantikan nama rekening
    asli yang benar seperti 'BSI-288') - keduanya tidak merepresentasikan
    rekening apapun.

    Di antara kandidat yang tersisa, kandidat yang TERLIHAT SEPERTI
    identitas rekening (lihat _looks_like_account_identifier - ada
    digit atau kata kunci bank umum) DIPRIORITASKAN dari yang cuma
    nama vendor/pelanggan biasa - supaya nama vendor yang kebetulan
    sering disebut (mis. 'SHOPEE' tempat belanja online) tidak
    mengalahkan nama rekening lawan transaksi yang genuinely lebih
    jarang disebut eksplisit.

    filename_hint (opsional, dari nama file ASLI yang diupload user -
    lihat _account_name_from_filename): kalau ADA di antara kandidat
    (persis atau sebagai substring case-insensitive), MENANG mutlak -
    ini sinyal PALING KUAT karena nama file biasanya sengaja dinamai
    sesuai rekeningnya. Kalau TIDAK ada kandidat account-like SAMA
    SEKALI dari data transaksi (kasus nyata: file yang SEMUA
    transaksinya merujuk pihak luar, nama rekeningnya sendiri tidak
    pernah disebut sebagai Subjek/Objek di baris manapun),
    filename_hint dipakai LANGSUNG sebagai fallback terakhir sebelum
    None.

    Dipakai untuk file 'Rekon Lokal' berdiri sendiri yang sheet-nya
    sering dinamai generik ('Mutasi') di KEDUA file - nama sheet TIDAK
    BISA dipakai sebagai identitas rekening (selain tidak informatif,
    kalau kedua file kebetulan sheet-nya sama persis, itu bikin
    identitas keduanya tertukar total di sisi pencocokan/penulisan
    hasil).

    Return None kalau tidak ada kandidat jelas maupun filename_hint
    (fallback ke nama sheet apa adanya oleh pemanggil)."""
    counts = {}
    for t in txns:
        for v in (t.subjek, t.objek):
            v = (v or "").strip()
            if v and v.lower() not in ("-", "mutasi", "tenant lain"):
                counts[v] = counts.get(v, 0) + 1
    if filename_hint:
        # filename_hint SELALU menang kalau tersedia - nama file adalah
        # sinyal PALING kuat (rekening biasanya sengaja dinamai sesuai
        # file-nya), lebih dipercaya daripada kandidat apapun hasil
        # tebakan dari data transaksi, TERMASUK kandidat yang terlihat
        # account-like (mis. 'BCA-887' yang muncul di data cuma karena
        # itu REKENING LAWAN transaksi, bukan identitas file ini
        # sendiri). Cek dulu apakah ada kandidat data yang cocok (buat
        # konsistensi format, mis. 'BRI-567(Biz)' vs filename 'BRI-567')
        # - kalau tidak ada yang cocok, filename_hint tetap dipakai
        # LANGSUNG, bukan jatuh ke tebakan account-like.
        for cand in counts:
            if cand.lower() == filename_hint.lower() or filename_hint.lower() in cand.lower():
                return cand
        return filename_hint
    if not counts:
        return None
    account_like = {k: v for k, v in counts.items() if _looks_like_account_identifier(k)}
    if not account_like:
        return max(counts.items(), key=lambda kv: kv[1])[0]
    return max(account_like.items(), key=lambda kv: kv[1])[0]


def _find_style_reference_row(ws, row):
    """Cari baris TERDEKAT (coba ke atas dulu, baru ke bawah) yang aman
    dijadikan rujukan gaya (font/format) - BUKAN baris 'Saldo Awal
    Bulan' (baris ringkasan yang MEMANG sengaja beda gaya, biasanya
    bold + format General), BUKAN baris hasil split Fliptech lain yang
    SAMA-SAMA masih perlu dibetulkan (Kategori 'Biaya Admin Bank'/
    'Bunga Bank' DAN Subjek '-' - kalau beberapa baris begini berurutan,
    saling menjadikan satu sama lain sebagai rujukan bikin gaya rusak
    menyebar/muter, tidak pernah ketemu gaya yang genuinely benar), dan
    bukan baris footer (Kategori kosong, mis. 'Total Debit'/'Saldo
    Akhir').

    Return nomor baris rujukan, atau None kalau tidak ketemu."""
    def _aman(candidate):
        kat = str(ws.cell(row=candidate, column=3).value or "").strip().lower()
        if not kat or kat == "saldo awal bulan":
            return False
        if kat in ("biaya admin bank", "bunga bank"):
            subjek = str(ws.cell(row=candidate, column=7).value or "").strip()
            if subjek == "-":
                return False
        return True

    for candidate in range(row - 1, 1, -1):
        if _aman(candidate):
            return candidate
    for candidate in range(row + 1, ws.max_row + 1):
        if _aman(candidate):
            return candidate
    return None


def _cleanup_and_verify_sheet(ws):
    """3 pembersihan akhir untuk SATU sheet rekening sebelum file
    /rekonlokal disimpan:
    1. Hapus baris yang TIDAK PUNYA ANGKA sama sekali di Debit/Kredit/
       Saldo Kumulatif (D/E/F) - biasanya baris pemisah/artifak kosong,
       BUKAN baris Saldo Awal Bulan (F-nya SELALU terisi) atau baris
       footer (salah satu dari D/E/F selalu terisi) - keduanya aman
       tidak akan ikut terhapus oleh kriteria ini. Satu baris kosong
       pemisah kemudian disisipkan KEMBALI tepat sebelum tiap blok
       footer (Saldo Awal/Saldo Akhir/Total Debit/Total Kredit) supaya
       footer tidak menyatu jadi satu tabel data dengan transaksi -
       kalau menyatu, AutoFilter Excel pada kolom Kategori (kosong di
       baris footer) akan ikut menyembunyikan footer saat difilter.
    2. Format mata uang yang konsisten untuk kolom D/E/F - kalau ada
       sel bernilai angka tapi formatnya BUKAN format mata uang yang
       dominan dipakai kolom itu (mis. masih "General" karena sempat
       diedit manual), disamakan.
    3. Baris footer Saldo Awal/Total Debit/Total Kredit/Saldo Akhir
       (dikenali dari teks di kolom B) dihitung ULANG dari data
       transaksi yang SEBENARNYA (bukan dipercaya apa adanya - baris
       bisa saja sudah disisipkan/dihapus oleh proses /rekonlokal) dan
       ditulis sebagai NILAI STATIS (bukan formula) - supaya tetap jadi
       acuan tetap untuk verifikasi manual, tidak ikut berubah kalau
       user mengedit sel lain di dekatnya.
    4. Formula Saldo Kumulatif (F) di SEMUA baris transaksi ditulis
       ULANG dari nol (pola F{row} = D{row}+E{row}+F{row-1}) - openpyxl
       tidak otomatis menyesuaikan referensi formula saat insert_rows/
       delete_rows di langkah 1/1b di atas, jadi formula lama bisa
       SALAH REFERENSI baris (menyebabkan #VALUE! terutama saat user
       sorting/filter data)."""
    header = [ws.cell(row=1, column=c).value for c in range(1, 9)]
    if header != _STANDARD_HEADER:
        return  # bukan sheet rekening berformat standar, jangan diapa-apakan

    # 1. Hapus baris tanpa angka sama sekali di D/E/F
    rows_to_delete = []
    for r in range(2, ws.max_row + 1):
        d = ws.cell(row=r, column=4).value
        e = ws.cell(row=r, column=5).value
        f = ws.cell(row=r, column=6).value
        if d is None and e is None and f is None:
            rows_to_delete.append(r)
    for r in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(r)

    # 1b. Sisipkan KEMBALI satu baris kosong pemisah sebelum tiap
    # blok baris footer (Saldo Awal non-Bulan/Total Debit/Total
    # Kredit/Saldo Akhir) - tanpa jarak ini, footer jadi menyatu
    # LANGSUNG dengan data transaksi (khususnya kalau baris kosong
    # pemisah aslinya sempat terhapus di langkah 1 di atas, atau
    # memang tidak ada dari sumbernya) - AutoFilter Excel yang
    # diterapkan user pada kolom Kategori (kosong di baris footer)
    # jadi ikut MENYEMBUNYIKAN footer, karena footer dianggap BAGIAN
    # dari tabel data yang sama. Berlaku untuk SEMUA kemunculan blok
    # footer (beberapa sumber menulis footer ini lebih dari sekali di
    # tengah sheet, bukan cuma di baris paling akhir).
    def _is_footer_label(r):
        b = str(ws.cell(row=r, column=2).value or "").strip().lower()
        c = str(ws.cell(row=r, column=3).value or "").strip().lower()
        if c == "saldo awal bulan":
            return False  # baris pembuka, bukan footer
        return b in ("saldo awal", "saldo akhir") or b.startswith("total debit") or b.startswith("total kredit")

    footer_block_starts = []
    r = 2
    while r <= ws.max_row:
        if _is_footer_label(r):
            footer_block_starts.append(r)
            while r <= ws.max_row and _is_footer_label(r):
                r += 1
        else:
            r += 1
    for r in sorted(footer_block_starts, reverse=True):
        if r > 2:
            b_above = ws.cell(row=r - 1, column=2).value
            c_above = ws.cell(row=r - 1, column=3).value
            if b_above is not None or c_above is not None:
                ws.insert_rows(r)

    # 2. Format mata uang konsisten (setelah nomor baris stabil pasca hapus)
    fmt_counter = {col: Counter() for col in (4, 5, 6)}
    for r in range(2, ws.max_row + 1):
        for col in (4, 5, 6):
            cell = ws.cell(row=r, column=col)
            if isinstance(cell.value, (int, float)):
                fmt_counter[col][cell.number_format] += 1
    dominant_fmt = {col: c.most_common(1)[0][0] for col, c in fmt_counter.items() if c}
    for r in range(2, ws.max_row + 1):
        for col in (4, 5, 6):
            cell = ws.cell(row=r, column=col)
            if (isinstance(cell.value, (int, float)) and col in dominant_fmt
                    and cell.number_format != dominant_fmt[col]):
                cell.number_format = dominant_fmt[col]

    # 3. Cari Saldo Awal Bulan + baris footer, hitung ulang dari transaksi asli
    saldo_awal_value = 0.0
    footer_rows = {}
    total_debit = 0.0
    total_kredit = 0.0
    for r in range(2, ws.max_row + 1):
        b_lower = str(ws.cell(row=r, column=2).value or "").strip().lower()
        c_lower = str(ws.cell(row=r, column=3).value or "").strip().lower()
        d = ws.cell(row=r, column=4).value
        e = ws.cell(row=r, column=5).value
        f = ws.cell(row=r, column=6).value

        if c_lower == "saldo awal bulan":
            if isinstance(f, (int, float)):
                saldo_awal_value = f
            continue
        if b_lower == "saldo awal":
            footer_rows["saldo awal"] = r
            continue
        if b_lower.startswith("total debit"):
            footer_rows["total debit"] = r
            continue
        if b_lower.startswith("total kredit"):
            footer_rows["total kredit"] = r
            continue
        if b_lower == "saldo akhir":
            footer_rows["saldo akhir"] = r
            continue

        if isinstance(d, (int, float)):
            total_debit += d
        if isinstance(e, (int, float)):
            total_kredit += e

    saldo_akhir_value = saldo_awal_value + total_debit + total_kredit

    if "saldo awal" in footer_rows:
        ws.cell(row=footer_rows["saldo awal"], column=6, value=saldo_awal_value)
    if "total debit" in footer_rows:
        ws.cell(row=footer_rows["total debit"], column=4, value=total_debit)
    if "total kredit" in footer_rows:
        ws.cell(row=footer_rows["total kredit"], column=5, value=total_kredit)
    if "saldo akhir" in footer_rows:
        ws.cell(row=footer_rows["saldo akhir"], column=6, value=saldo_akhir_value)

    # 4. Perbaiki formula Saldo Kumulatif (F) di SEMUA baris transaksi -
    # openpyxl TIDAK otomatis menyesuaikan referensi formula saat baris
    # disisipkan/dihapus (beda dari Excel manual saat user insert/delete
    # row lewat UI) - formula LAMA (mis. "=D2+E2+F1") bisa jadi SALAH
    # REFERENSI baris setelah operasi insert_rows/delete_rows di langkah
    # 1/1b di atas, menyebabkan #VALUE! error kalau baris yang
    # direferensikan sekarang berisi teks (header/footer) bukan angka -
    # apalagi kalau user lalu SORTING/FILTER data, error itu ikut
    # bergeser dan makin membingungkan. Ditulis ULANG dari NOL mengikuti
    # pola standar F{row} = D{row}+E{row}+F{row-1} untuk SEMUA baris
    # transaksi (BUKAN baris Saldo Awal Bulan/footer, yang sudah
    # ditangani terpisah di atas sebagai nilai statis) - dijalankan
    # PALING AKHIR, setelah nomor baris benar-benar final.
    for r in range(2, ws.max_row + 1):
        b_lower = str(ws.cell(row=r, column=2).value or "").strip().lower()
        c_lower = str(ws.cell(row=r, column=3).value or "").strip().lower()
        if c_lower == "saldo awal bulan":
            continue
        if (b_lower in ("saldo awal", "saldo akhir")
                or b_lower.startswith("total debit") or b_lower.startswith("total kredit")):
            continue
        d = ws.cell(row=r, column=4).value
        e = ws.cell(row=r, column=5).value
        if d is None and e is None:
            continue  # baris kosong (seharusnya sudah tersaring di langkah 1, jaga-jaga)
        ws.cell(row=r, column=6, value=f"=D{r}+E{r}+F{r - 1}")


_OFFICIAL_LAYER1_CATEGORIES = [
    "Penjualan", "Penjualan Shopeefood", "Penjualan Grabfood",
    "Belanja Bahan", "Overhead", "Konsumsi dan Liburan", "Belanja Utilitas",
    "Tools dan Equipments", "Kemasan", "Subscription",
    "Sewa dan Maintenance Bangunan", "Reparasi dan Maintenance Tools dan Mesin",
    "Pajak dan Administrasi", "Belanja Assets", "Marketing",
    "Riset dan Development", "Biaya Admin Bank", "Tip/Minus/Lebih",
    "Hutang Masuk", "Pembayaran Hutang", "Pengeluaran Pribadi",
    "Modal & Setoran Pemilik", "Transaksi Internal",
]
_KATEGORI_STOPWORDS = {"dan", "&"}


def _find_closest_official_category(kategori_asli):
    """Kalau kategori_asli adalah versi TERPOTONG/tidak lengkap dari
    salah satu kategori resmi (semua kata yang ADA cocok, tinggal ada
    kata yang HILANG - mis. 'Reparasi Mesin'/'Reparasi Tools' vs
    'Reparasi dan Maintenance Tools dan Mesin'), kembalikan nama
    LENGKAP resmi. None kalau kategori_asli SUDAH persis salah satu
    kategori resmi, atau tidak ada kecocokan yang cukup jelas/spesifik
    (supaya tidak salah tangkap - kata umum sendirian seperti 'belanja'
    BUKAN sinyal cukup kuat, minimal 2 kata cocok, atau 1 kata yang
    sudah cukup panjang/spesifik)."""
    k = (kategori_asli or "").strip().lower()
    if not k:
        return None
    for official in _OFFICIAL_LAYER1_CATEGORIES:
        if k == official.lower():
            return None  # sudah persis benar
    k_words = set(k.split()) - _KATEGORI_STOPWORDS
    if not k_words:
        return None
    candidates = []
    for official in _OFFICIAL_LAYER1_CATEGORIES:
        official_words = set(official.lower().split()) - _KATEGORI_STOPWORDS
        if k_words <= official_words:
            candidates.append((len(k_words), len(official_words), official))
    if not candidates:
        return None
    # menangkan kandidat dengan PALING SEDIKIT kata (interpretasi paling
    # sederhana/langsung) kalau beberapa kategori resmi sama-sama cocok
    # dengan skor kata yang sama (mis. kata "tools" sendirian muncul di
    # "Tools dan Equipments" MAUPUN "Reparasi dan Maintenance Tools dan
    # Mesin").
    candidates.sort(key=lambda c: (-c[0], c[1]))
    score, _, best = candidates[0]
    if score >= 2:
        return best
    if score == 1:
        # untuk kecocokan SATU kata saja, pastikan kata itu tidak
        # sekadar prefiks generik yang dipakai BANYAK kategori resmi
        # sekaligus (mis. "belanja" muncul di "Belanja Bahan"/"Belanja
        # Assets"/"Belanja Utilitas") - kalau begitu genuinely ambigu,
        # jangan menebak salah satu meski katanya cukup panjang.
        satu_kata = next(iter(k_words))
        n_kategori_mengandung = sum(
            1 for official in _OFFICIAL_LAYER1_CATEGORIES
            if satu_kata in (set(official.lower().split()) - _KATEGORI_STOPWORDS)
        )
        if n_kategori_mengandung == 1 and len(satu_kata) >= 5:
            return best
    return None


_LEGACY_KATEGORI_RENAME = {
    "belanja operasional": "Overhead",
    "reparasi": "Reparasi dan Maintenance Tools dan Mesin",
    "reparasi dan maintenance": "Reparasi dan Maintenance Tools dan Mesin",
    "belanja konsumsi": "Konsumsi dan Liburan",
    "pajak daerah": "Pajak dan Administrasi",
    "biaya administrasi": "Pajak dan Administrasi",
    "administrasi": "Pajak dan Administrasi",
    "biaya renovasi atap": "Sewa dan Maintenance Bangunan",
    "renovasi bangunan": "Sewa dan Maintenance Bangunan",
    "renovasi": "Sewa dan Maintenance Bangunan",
    "tools": "Tools dan Equipments",
}


def run_rekon_bersih(path, output_path):
    """Fitur satu-file (input 1 file, output 1 file) - BUKAN pencocokan
    lintas file seperti /rekonlokal (yang butuh 2 file untuk mencari
    pasangan transfer). Tujuannya murni membersihkan format SATU file
    rekening:
    1. Subjek DAN Objek dilengkapi jadi nama lengkap Title Case kalau
       dikenali dari EMPLOYEE_ALIASES (pegawai+owner) - kalau tidak
       dikenali, ditulis ALL CAPS supaya jelas belum teridentifikasi
       (bukan cuma dibiarkan apa adanya).
    2. Baris yang tidak punya angka sama sekali di Debit/Kredit/Saldo
       Kumulatif dihapus (baris pemisah/artifak kosong) - kecuali satu
       baris kosong disisipkan kembali sebagai pemisah sebelum tiap
       blok footer (Saldo Awal/Saldo Akhir/Total Debit/Total Kredit),
       supaya AutoFilter Excel tidak ikut menyembunyikan footer.
    3. Format mata uang kolom Debit/Kredit/Saldo Kumulatif disamakan ke
       format dominan yang dipakai kolom itu.
    4. Footer Saldo Awal/Total Debit/Total Kredit/Saldo Akhir dihitung
       ULANG dari data transaksi sebenarnya dan ditulis sebagai nilai
       statis (bukan formula).
    5. Kategori yang TERPOTONG/tidak lengkap (typo/singkatan dari salah
       satu kategori resmi Layer 1, mis. "Reparasi Mesin"/"Reparasi
       Tools" -> "Reparasi dan Maintenance Tools dan Mesin") dilengkapi
       jadi nama resmi lengkap - lihat _LEGACY_KATEGORI_RENAME (kasus
       yang sudah diketahui pasti) dan _find_closest_official_category
       (pencocokan fuzzy berbasis kata untuk kasus lain, HANYA kalau
       cukup spesifik/tidak ambigu dengan kategori resmi lain).
    6. Keterangan Tambahan (I) diberi label status tenant (Objek) -
       tenant dikenal (EMPLOYEE_ALIASES/vendor terdaftar) -> "Paid Off
       to <tenant>", tenant terisi tapi tidak dikenal -> "Unrecognized
       Tenant", tenant kosong -> dibiarkan kosong.
    Tidak melakukan pencocokan transfer - fitur ini SENGAJA dibatasi
    cuma pada format+Subjek/Objek+Kategori sesuai permintaan eksplisit
    user, bukan rekategorisasi penuh berbasis kata kunci Keterangan
    seperti /rekonlokal."""
    wb = openpyxl.load_workbook(path)
    sheets = [s for s in wb.sheetnames if _looks_like_account_sheet(wb[s])]
    if not sheets:
        raise ValueError(f"File '{path}' tidak punya sheet rekening berformat standar yang dikenali.")

    # Nama vendor/merchant yang SUDAH dikenal benar (dari aturan vendor
    # /rekonlokal yang sudah ada) - kalau Objek/Subjek cocok (case-
    # insensitive) salah satu dari ini, ditulis bentuk Title Case yang
    # benar (BUKAN dipaksa ALL CAPS) - itu sudah "diketahui" dengan
    # benar, cuma bukan lewat EMPLOYEE_ALIASES (pegawai/owner) tapi
    # lewat pengenalan vendor.
    _known_vendor_names = {}  # lower() -> bentuk Title Case yang benar
    for _name in ("Grab Merchant", "Shopeefood Merchant"):
        _known_vendor_names[_name.lower()] = _name
    for _rules_list in (_KAS_BUKU_VENDOR_RULES, _OBJEK_VENDOR_RULES):
        for _keywords, _ket, _kat, _objek in _rules_list:
            _known_vendor_names[_ket.lower()] = _ket
            if _objek:
                _known_vendor_names[_objek.lower()] = _objek

    # Set gabungan (nilai TERISI, lower()) untuk pengecekan "tenant
    # dikenal" di label Keterangan Tambahan (I) - EMPLOYEE_ALIASES
    # (alias DAN nama lengkap resminya) plus vendor yang dikenal.
    _known_tenant_names_bersih = set(EMPLOYEE_ALIASES.keys()) | {v.lower() for v in EMPLOYEE_ALIASES.values()}
    _known_tenant_names_bersih |= set(_known_vendor_names.keys())

    n_subjek_objek_dilengkapi = 0
    n_kategori_dilengkapi = 0
    for sn in sheets:
        ws = wb[sn]
        split_fliptech_combined_rows(ws)
        txns, _ = read_account_sheet(ws)
        for t in txns:
            if t.is_opening:
                continue
            for col, val in ((7, t.subjek), (8, t.objek)):
                v = (val or "").strip()
                if not v or v == "-":
                    continue
                nama_lengkap = EMPLOYEE_ALIASES.get(v.lower()) or _known_vendor_names.get(v.lower())
                if nama_lengkap and v != nama_lengkap:
                    ws.cell(row=t.row, column=col, value=nama_lengkap)
                    n_subjek_objek_dilengkapi += 1
                elif not nama_lengkap and v != v.upper():
                    ws.cell(row=t.row, column=col, value=v.upper())
                    n_subjek_objek_dilengkapi += 1
            # Kategori yang TERPOTONG/tidak lengkap (typo/singkatan dari
            # salah satu kategori resmi, mis. "Reparasi Mesin"/"Reparasi
            # Tools" -> "Reparasi dan Maintenance Tools dan Mesin")
            # dilengkapi jadi nama resmi lengkap - dicek dulu daftar
            # legacy yang SUDAH DIKETAHUI PASTI (_LEGACY_KATEGORI_RENAME),
            # baru cocokkan fuzzy berbasis kata kalau belum ketemu di situ.
            kategori_asli = (t.kategori or "").strip()
            target = _LEGACY_KATEGORI_RENAME.get(kategori_asli.lower()) or _find_closest_official_category(kategori_asli)
            if not target and kategori_asli and not _is_recognized_category(kategori_asli):
                # Kategori aslinya genuinely tidak dikenal ("New
                # Kategori"/dst) DAN tidak cocok legacy-rename/fuzzy -
                # coba effective_kategori (aturan kata kunci berbasis
                # Keterangan/Objek/Subjek, BUKAN cuma teks Kategori itu
                # sendiri) - menangkap kasus seperti "Biaya Reparasi
                # (Bayar Tukang)" yang kata kuncinya ada di Keterangan,
                # bukan di kolom Kategori yang masih generik.
                hitung = t.effective_kategori
                if hitung and hitung != "Kategori Baru":
                    target = hitung
            if target and target != kategori_asli:
                ws.cell(row=t.row, column=3, value=target)
                n_kategori_dilengkapi += 1
            # Label Keterangan Tambahan (I) berdasarkan status tenant di
            # Objek (SAMA konsep dengan /rekonlokal, versi tanpa
            # pencocokan transfer karena /rekonbersih cuma 1 file):
            # tenant kosong -> dibiarkan kosong, tenant dikenal (dari
            # EMPLOYEE_ALIASES atau daftar vendor terdaftar) -> "Paid
            # Off to <tenant>", tenant terisi tapi tidak dikenal ->
            # "Unrecognized Tenant". Dicek dari Objek TERKINI (setelah
            # kemungkinan dilengkapi di pass Subjek/Objek di atas).
            objek_final = (ws.cell(row=t.row, column=8).value or "").strip()
            if objek_final and objek_final != "-":
                if objek_final.lower() in _known_tenant_names_bersih:
                    ws.cell(row=t.row, column=9, value=f"Paid Off to {objek_final}")
                else:
                    ws.cell(row=t.row, column=9, value="Unrecognized Tenant")
            # Highlight warna kelompok kategori - DISEGARKAN ulang sesuai
            # Kategori TERKINI (setelah kemungkinan dikoreksi di atas),
            # supaya kalau file ini sebelumnya sempat diwarnai versi
            # lama/salah (mis. dari /rekonlokal versi lama, atau
            # kategori yang baru saja dikoreksi barusan), warnanya ikut
            # diperbarui - bukan cuma dibiarkan warna basi yang sudah
            # tidak sesuai Kategori sekarang. Warna highlight yang
            # PUNYA ARTI KHUSUS (merah/ungu/biru/orange/kuning dari
            # /rekonlokal) TETAP dihormati/tidak ditimpa - fitur ini
            # tidak melakukan pencocokan transfer sendiri, jadi kalau
            # ada highlight seperti itu, itu peninggalan proses lain
            # yang lebih spesifik dan harus tetap terlihat.
            cell_b = ws.cell(row=t.row, column=2)
            current_fill = cell_b.fill.fgColor.rgb if cell_b.fill else None
            if current_fill not in _MEANINGFUL_HIGHLIGHT_HEXES:
                kategori_sekarang = ws.cell(row=t.row, column=3).value
                fill = category_group_fill(kategori_sekarang)
                target_fill_rgb = fill.fgColor.rgb if fill else None
                if fill is not None and current_fill != target_fill_rgb:
                    for c in range(1, 10):
                        ws.cell(row=t.row, column=c).fill = fill
        _cleanup_and_verify_sheet(ws)

    wb.save(output_path)
    return {"n_subjek_objek_dilengkapi": n_subjek_objek_dilengkapi, "n_kategori_dilengkapi": n_kategori_dilengkapi}


_KNOWN_I_LABEL_EXACT = {
    "unresolved", "medium unresolved", "low unresolved", "suspicious",
    "cookies kaola", "admin fee",
}
_KNOWN_I_LABEL_PREFIXES = ("solved ", "paid to ", "paid off to ", "sales via edc ", "settled to ")


def _is_known_i_label(text):
    """True kalau `text` (Keterangan Tambahan/kolom I) SUDAH cocok pola
    label sistem yang dikenal ('Unresolved', 'Solved X to Y', 'Paid to
    X', dst) - dipakai di run_rekon_lokal SEBAGAI TAMBAHAN atas cek
    "masih sama dengan t.ket" untuk menentukan baris "sudah dapat label
    bermakna". PENTING: cek "masih sama dengan t.ket" SENDIRIAN tidak
    cukup untuk file yang di-upload ULANG setelah pernah diproses
    /rekonlokal sebelumnya - baris yang TIDAK ketemu match lagi di run
    BARU ini akan punya t.ket = 'Unresolved'/'Solved ...' (dibaca dari
    file yang SUDAH berlabel itu), sehingga current_i == t.ket (SAMA
    PERSIS, karena TIDAK ADA pass baru yang menulis ulang) dan salah
    dikira "belum tersentuh". Fungsi ini menangkap kasus itu."""
    t = (text or "").strip().lower()
    if not t:
        return False
    if t in _KNOWN_I_LABEL_EXACT:
        return True
    return any(t.startswith(p) for p in _KNOWN_I_LABEL_PREFIXES)


def run_rekon_lokal(path1, path2, out1, out2, filename1=None, filename2=None):
    """Rekon Lokal - fitur MANUAL ringan: cocokkan HANYA transaksi
    'Transaksi Internal' antar 2 file rekening (bukan rekonsiliasi penuh
    - tidak ada Laba Rugi/Neraca/deteksi minus/dst). Untuk transfer yang
    match confidence High/Medium (nominal sama di tanggal sama, atau
    selisih wajar untuk settlement bank - lihat find_matches):
    1. Subjek/Objek di KEDUA file dikoreksi supaya saling menyebut
       rekening lawan yang benar - nama rekening diambil dari
       _infer_account_name (BUKAN nama sheet - lihat catatan di sana),
       supaya benar walau sheet kedua file kebetulan sama persis
       ('Mutasi').
    2. Highlight warna berdasarkan grup bank TUJUAN, SAMA seperti flow
       rekonsiliasi utama (biru=BCA, orange=Jago, kuning=BRI/BSI).
    3. Transaksi 'setoran tunai' yang self-referencing (Subjek==Objek==
       rekening sendiri, TIDAK ketemu pasangan cross-file - uang tunai
       masuk langsung tanpa lawan transaksi bank) - Keterangan
       diseragamkan jadi "Setoran <rekening>".
    4. Transaksi internal yang TIDAK ketemu pasangannya (selain pola
       setoran tunai di atas, yang memang wajar tidak match) dihighlight
       MERAH - supaya kelihatan jelas mana yang masih perlu ditelusuri
       manual.
    5. SEMUA transaksi (bukan cuma transfer internal) yang Kategori
       tersimpannya beda dari yang dihitung sistem berdasarkan
       Keterangan/Objek - dihighlight UNGU. HANYA menandai untuk audit
       manual, TIDAK PERNAH mengubah Kategori/Keterangan aslinya (beberapa
       mismatch bisa jadi false positive).
    Selain koreksi & highlight itu, KEDUA FILE DIKEMBALIKAN APA ADANYA -
    tidak ada sheet tambahan, tidak ada kategorisasi/perhitungan lain.
    Split/merge (combo_matches) SENGAJA di luar cakupan - fitur ini
    murni pasangan 1:1 sederhana.

    Return dict {"n_high": ..., "n_medium": ..., "n_belum_rekon": ...,
    "n_kategori_mencurigakan": ...} untuk caption bot."""
    wb1 = openpyxl.load_workbook(path1)
    wb2 = openpyxl.load_workbook(path2)
    sheets1 = [s for s in wb1.sheetnames if _looks_like_account_sheet(wb1[s])]
    sheets2 = [s for s in wb2.sheetnames if _looks_like_account_sheet(wb2[s])]
    if not sheets1:
        raise ValueError(f"File '{path1}' tidak punya sheet rekening berformat standar yang dikenali.")
    if not sheets2:
        raise ValueError(f"File '{path2}' tidak punya sheet rekening berformat standar yang dikenali.")

    all_txns = []
    filename_hints = {"1": _account_name_from_filename(filename1), "2": _account_name_from_filename(filename2)}
    # key = nama rekening HASIL TEBAKAN (unik per file, dipakai sebagai
    # t.sheet pengganti supaya find_matches tidak pernah menganggap 2
    # sheet dari file BEDA sebagai "sheet yang sama" cuma karena judul
    # sheet aslinya kebetulan identik ('Mutasi' di kedua file)
    name_to_real = {}  # nama rekening -> (workbook, nama sheet ASLI)
    for wb, sheets, tag in ((wb1, sheets1, "1"), (wb2, sheets2, "2")):
        for sn in sheets:
            split_fliptech_combined_rows(wb[sn])
            txns, _ = read_account_sheet(wb[sn])
            akun = _infer_account_name(txns, filename_hint=filename_hints[tag]) or f"{sn} ({tag})"
            # Tabrakan nama rekening ANTAR FILE - bisa terjadi walau
            # _infer_account_name jalan benar, kalau KEDUA file memang
            # jenis yang sama (mis. dua file Kas Buku, keduanya sama-
            # sama paling sering menyebut diri sendiri "Kas/Buku").
            # Tanpa disambiguasi ini, entry file PERTAMA di name_to_real
            # TERTIMPA file KEDUA - koreksi yang seharusnya ditulis ke
            # file pertama malah ketulis ke file kedua (row number sama,
            # tapi workbook/sheet beda -> data campur aduk).
            if akun in name_to_real:
                akun = f"{akun} ({tag})"
            for t in txns:
                t.sheet = akun
            all_txns.extend(txns)
            name_to_real[akun] = (wb, sn)

    matches, _combo_matches = find_matches(all_txns, list(name_to_real.keys()))

    # Lengkapi Objek jadi nama lengkap (Title Case) kalau dikenali dari
    # EMPLOYEE_ALIASES (pegawai+owner) - kalau TIDAK dikenali, tulis
    # ALL CAPS supaya jelas kelihatan ini belum teridentifikasi (bukan
    # cuma dibiarkan apa adanya, yang bisa ambigu antara "sudah diperiksa
    # dan memang begitu" vs "belum diperiksa"). Dijalankan PALING AWAL di
    # antara pass-pass koreksi lain (sebelum transfer-matching/vendor/
    # Gaji/Grab-Shopeefood) - pass-pass itu MENIMPA Objek dengan nilai
    # spesifik mereka sendiri (nama rekening lawan, nama vendor, dst)
    # kalau memang berlaku, jadi tidak masalah kalau pass ini sempat
    # menulis sesuatu di situ dulu.
    for t in all_txns:
        if t.is_opening:
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        # "Kas Buku"/"Kas/Buku" (rekening kas internal, dalam bentuk
        # apapun ia ditulis - beda kapitalisasi/tanda baca) diseragamkan
        # jadi "Kas Kasir" di Subjek MAUPUN Objek untuk konsistensi
        # penamaan (permintaan eksplisit user) - dicek LEBIH DULU
        # sebelum EMPLOYEE_ALIASES, supaya tidak keduluan aturan lain.
        objek_asli = (t.objek or "").strip()
        subjek_asli = (t.subjek or "").strip()
        _kas_buku_re = re.compile(r"^kas\s*/?\s*buku$", re.IGNORECASE)
        if subjek_asli and _kas_buku_re.match(subjek_asli):
            wb_t[sheet_t].cell(row=t.row, column=7, value="Kas Kasir")
        if objek_asli and _kas_buku_re.match(objek_asli):
            wb_t[sheet_t].cell(row=t.row, column=8, value="Kas Kasir")
            continue  # sudah ditangani, lewati pass EMPLOYEE_ALIASES di bawah untuk Objek ini
        if not objek_asli or objek_asli == "-":
            continue
        nama_lengkap = EMPLOYEE_ALIASES.get(objek_asli.lower())
        if nama_lengkap:
            wb_t[sheet_t].cell(row=t.row, column=8, value=nama_lengkap)
        elif objek_asli != objek_asli.upper():
            wb_t[sheet_t].cell(row=t.row, column=8, value=objek_asli.upper())

    n_high = 0
    n_medium = 0
    n_low = 0
    matched_ids = set()
    for m in matches:
        # Confidence "Low" DULU jatuh ke celah - tidak ditangani loop
        # matching ini (excluded karena confidence bukan High/Medium)
        # MAUPUN loop "belum direkon" di bawah (excluded karena m.dst
        # SUDAH terisi, technically "matched" walau confidence rendah) -
        # akibatnya baris begini tidak pernah dapat label "Solved"/
        # "Medium Unresolved"/"Unresolved" SAMA SEKALI, dan pass label
        # tenant paling akhir salah mengiranya "belum tersentuh" lalu
        # menandainya "Unrecognized Tenant" - padahal ini genuinely
        # transfer internal, bukan soal tenant sama sekali. Confidence
        # Low SEKARANG ditangani DI SINI juga (sejajar dengan Medium).
        if m.dst is None or m.confidence not in ("High", "Medium", "Low"):
            continue
        if m.confidence == "High":
            n_high += 1
        elif m.confidence == "Medium":
            n_medium += 1
        else:
            n_low += 1
        pengirim, penerima = _sender_receiver(m.src, m.dst)
        matched_ids.add(id(pengirim))
        matched_ids.add(id(penerima))
        wb_p, sheet_p = name_to_real[pengirim.sheet]
        wb_r, sheet_r = name_to_real[penerima.sheet]
        ws_p = wb_p[sheet_p]
        ws_r = wb_r[sheet_r]
        ws_p.cell(row=pengirim.row, column=8, value=_display_account_name(penerima.sheet))
        ws_r.cell(row=penerima.row, column=7, value=_display_account_name(pengirim.sheet))
        # Kategori (C) dipastikan benar - transaksi ini SUDAH terbukti
        # transfer internal matched, terlepas dari kategori asalnya
        # (mis. 'Transfer Lainnya'/'New Kategori' sebelum ketemu
        # pasangan). Keterangan (B) diseragamkan jadi 'Transfer ke/dari
        # <rekening>'. Keterangan Tambahan (I) ditulis 'Solved <Subjek>
        # to <Objek>' - penanda EKSPLISIT bahwa baris ini SUDAH selesai
        # direkon, supaya kalau /rekonlokal dijalankan LAGI nanti dengan
        # file pasangan yang BEDA (baris ini otomatis tidak ketemu match
        # di run itu, karena rekening lawannya tidak ada di file yang
        # dibandingkan), baris ini TIDAK dihighlight merah lagi -
        # sudah terbukti selesai dari run sebelumnya, bukan genuinely
        # belum direkon.
        ws_p.cell(row=pengirim.row, column=3, value="Transaksi Internal")
        ws_r.cell(row=penerima.row, column=3, value="Transaksi Internal")
        # Confidence Medium/Low: cocok tapi tidak 100% pasti (selisih
        # nominal/tanggal masih dalam toleransi, Low = toleransi lebih
        # longgar dari Medium) - Keterangan Tambahan ditulis "Medium/Low
        # Unresolved" (BUKAN "Solved X to Y" seperti confidence High)
        # supaya jelas kelihatan ini masih perlu verifikasi manual.
        if m.confidence == "Medium":
            note = "Medium Unresolved"
        elif m.confidence == "Low":
            note = "Low Unresolved"
        else:
            note = f"Solved {_display_account_name(pengirim.sheet)} to {_display_account_name(penerima.sheet)}"
        ws_p.cell(row=pengirim.row, column=9, value=note)
        ws_r.cell(row=penerima.row, column=9, value=note)
        ws_p.cell(row=pengirim.row, column=2, value=f"Transfer ke {_display_account_name(penerima.sheet)}")
        ws_r.cell(row=penerima.row, column=2, value=f"Transfer dari {_display_account_name(pengirim.sheet)}")
        # Confidence Medium/Low (cocok tapi tidak 100% pasti - selisih
        # nominal/tanggal masih dalam toleransi, bukan match persis) -
        # dihighlight merah shade LEBIH MUDA daripada highlight "belum
        # direkon" (FFC7CE), MENGGANTIKAN warna grup bank tujuan biasa -
        # supaya baris ini tetap kelihatan beda/perlu perhatian ekstra
        # dibanding match High yang sudah pasti, tapi TIDAK disamakan
        # semencolok "belum direkon" yang genuinely belum ketemu sama
        # sekali. Low dan Medium pakai warna sama (beda ditandai lewat
        # teks "Low Unresolved" vs "Medium Unresolved" di kolom I).
        if m.confidence in ("Medium", "Low"):
            fill = REKONLOKAL_MEDIUM_CONFIDENCE_FILL
        else:
            fill = _bank_group_fill(penerima.sheet)
        if fill is not None:
            for c in range(1, 10):
                ws_p.cell(row=pengirim.row, column=c).fill = fill
                ws_r.cell(row=penerima.row, column=c).fill = fill

    # Transaksi internal yang TIDAK ketemu pasangannya - dihighlight
    # MERAH supaya kelihatan jelas mana yang masih perlu ditelusuri
    # manual. Transaksi self-referencing (Subjek==Objek==rekening
    # sendiri, pola setoran tunai via CDM/ATM) DIKECUALIKAN dari sini -
    # itu MEMANG tidak akan pernah ketemu pasangan lintas file (bukan
    # transfer antar rekening, cuma uang tunai masuk langsung), sudah
    # ditangani wajar lewat normalisasi "Setoran <rekening>" di bawah,
    # bukan kasus 'belum direkon' yang perlu ditandai.
    n_belum_rekon = 0
    for m in matches:
        if m.dst is not None:
            continue
        src = m.src
        if (src.subjek or "").strip() == (src.objek or "").strip():
            # Self-referencing (Subjek==Objek) BIASANYA pola setoran
            # tunai yang wajar (lihat komentar di bawah) - TAPI kalau
            # Keterangan Tambahan-nya sudah "Solved <X> to <X>" (rekening
            # SAMA di kedua sisi), ini bukan setoran tunai genuine -
            # ini SISA DATA RUSAK dari bug find_matches yang sudah
            # diperbaiki (transaksi sempat salah matched dengan transaksi
            # LAIN DI REKENING YANG SAMA, ditandai 'Solved' padahal
            # rekening lawannya harusnya BEDA). Highlight merah supaya
            # user sadar perlu ditelusuri manual - jangan dianggap wajar.
            ket = (src.ket or "").strip().lower()
            subjek = (src.subjek or "").strip().lower()
            if ket.startswith("solved ") and subjek and subjek in ket.split(" to "):
                pass  # lanjut ke bawah, JANGAN di-skip - tandai merah
            else:
                continue
        elif (src.ket or "").strip().lower().startswith("solved "):
            # Sudah pernah "Solved <X> to <Y>" (X != Y, valid) dari run
            # /rekonlokal SEBELUMNYA (dengan file pasangan yang berbeda)
            # - baris ini SUDAH selesai direkon, cuma kebetulan rekening
            # lawannya tidak ada di file yang dibandingkan pada run kali
            # ini. Bukan genuinely belum direkon - jangan dihighlight
            # merah lagi.
            continue
        elif any(kw in (src.kategori or "").lower() for kw in CAPITAL_KEYWORDS):
            # Kategori ASLI-nya Modal & Setoran Pemilik (jadi kandidat
            # transfer karena melibatkan owner - lihat Txn.is_transfer),
            # tapi TIDAK ketemu pasangan valid. Ini BUKAN kegagalan -
            # "Setoran Pemilik" TETAP kategori yang sah kalau memang
            # tidak ada transaksi berlawanan (user menegaskan: itu cuma
            # prioritas KEDUA, dipakai kalau tidak ketemu pasangan
            # transfer valid). Jangan dihighlight merah - biarkan tetap
            # Modal & Setoran Pemilik apa adanya.
            continue
        n_belum_rekon += 1
        wb_t, sheet_t = name_to_real[src.sheet]
        ws_t = wb_t[sheet_t]
        ws_t.cell(row=src.row, column=9, value="Unresolved")
        for c in range(1, 10):
            ws_t.cell(row=src.row, column=c).fill = REKONLOKAL_UNMATCHED_FILL

    # Setoran tunai (mis. "SETORAN VIA CDM") - transaksi yang self-
    # referencing (Subjek==Objek==rekening sendiri, uang tunai masuk
    # langsung ke rekening tanpa lawan transaksi bank untuk dicocokkan)
    # DAN tidak ketemu pasangan cross-file (bukan bagian dari matches di
    # atas) - Keterangan diseragamkan jadi "Setoran <rekening>", supaya
    # konsisten apapun istilah aslinya ("SETORAN VIA CDM", dst).
    for t in all_txns:
        if id(t) in matched_ids:
            continue
        if not t.is_transfer or t.nominal <= 0:
            continue
        if (t.subjek or "").strip() != (t.objek or "").strip():
            continue
        if "setoran" not in (t.desc or "").lower():
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        wb_t[sheet_t].cell(row=t.row, column=2, value=f"Setoran {_display_account_name(t.sheet)}")

    # Baris hasil pemisahan Fliptech (split_fliptech_combined_rows) -
    # dikenali dari CIRI KHASNYA (Kategori persis 'Biaya Admin Bank'/
    # 'Bunga Bank' DAN Subjek '-', signature yang ditulis fungsi itu),
    # BUKAN dari teks 'Bagian dari transaksi Fliptech...' - teks itu
    # bisa SUDAH HILANG kalau file ini pernah diproses /rekonlokal
    # sebelumnya (kategori/keterangan sudah dibersihkan duluan), tapi
    # font/number_format-nya BISA SAJA belum sempat ikut dibetulkan
    # (versi /rekonlokal sebelum ini cuma benerin teks, font di-
    # hardcode salah). Jadi di sini SELALU dicek & dibetulkan ulang,
    # idempotent - aman dijalankan berkali-kali di file yang sama.
    for t in all_txns:
        kat = (t.kategori or "").strip().lower()
        if kat not in ("biaya admin bank", "bunga bank"):
            continue
        if (t.subjek or "").strip() != "-":
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        if "bagian dari transaksi fliptech" in (t.desc or "").lower():
            ws_t.cell(row=t.row, column=9, value=t.desc)
            ws_t.cell(row=t.row, column=2, value=t.kategori)
        ref_row = _find_style_reference_row(ws_t, t.row)
        if ref_row is not None:
            for col in range(1, 10):
                ref_cell = ws_t.cell(row=ref_row, column=col)
                this_cell = ws_t.cell(row=t.row, column=col)
                this_cell.font = copy.copy(ref_cell.font)
                this_cell.number_format = ref_cell.number_format
                this_cell.alignment = copy.copy(ref_cell.alignment)
                this_cell.border = copy.copy(ref_cell.border)

    # Transaksi terindikasi Gaji - Keterangan (B) diseragamkan jadi
    # "Gaji <Nama Depan> <Bulan> <Tahun>" (nama depan dari Objek), dan
    # Kategori (C) dipastikan "Gaji Bulan Ini" atau "Gaji Accrual"
    # (bulan sebelumnya, dibayar telat/awal bulan) - lihat
    # _gaji_rekon_lokal_info untuk logika penentuan bulan gajinya.
    gaji_ambiguous_ids = set()
    for t in all_txns:
        info = _gaji_rekon_lokal_info(t)
        if info is None:
            continue
        nama_depan, bulan_nama, tahun, is_bulan_ini = info
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        # Objek diseragamkan jadi Title Case (huruf awal tiap kata
        # kapital, sisanya kecil) - data sumber sering ALL CAPS
        # ("ADINDA NURUSSHAFWA"), tidak enak dibaca dan tidak konsisten
        # dengan gaya penulisan nama di tempat lain. Dilakukan DULU di
        # sini (sebelum menulis Keterangan Tambahan) supaya "Paid to
        # <nama>" di bawah memakai nama yang sudah rapi.
        objek_asli = (ws_t.cell(row=t.row, column=8).value or "").strip()
        objek_title = " ".join(w.capitalize() for w in objek_asli.split()) if objek_asli else objek_asli
        if objek_asli:
            ws_t.cell(row=t.row, column=8, value=objek_title)
        ws_t.cell(row=t.row, column=2, value=f"Gaji {nama_depan} {bulan_nama} {tahun}")
        ws_t.cell(row=t.row, column=3, value="Gaji Bulan Ini" if is_bulan_ini else "Gaji Accrual")
        # Keterangan Tambahan (I) untuk SEMUA transaksi Gaji ditulis
        # "Paid to <nama lengkap>" (pakai Objek penuh yang sudah Title
        # Case, BUKAN cuma 2 kata pertama yang dipakai di Keterangan) -
        # override permintaan eksplisit user, menggantikan pendekatan
        # arsip Keterangan lama yang dipakai sebelumnya untuk kategori
        # Gaji secara khusus.
        ws_t.cell(row=t.row, column=9, value=f"Paid to {objek_title}" if objek_title else "-")
        if "(?)" in nama_depan:
            # Nama depan ambigu (dipakai >1 pegawai, data sumber tidak
            # cukup buat membedakan) - highlight ungu yang sama dengan
            # 'Kategori mencurigakan', sama-sama butuh verifikasi manual.
            # Dicatat di gaji_ambiguous_ids supaya pass "Kategori
            # mencurigakan" di bawah TIDAK menganggap ungu ini "stale"
            # dan menghapusnya lagi - ini ungu SEGAR dari run ini juga.
            gaji_ambiguous_ids.add(id(t))
            for c in range(1, 10):
                ws_t.cell(row=t.row, column=c).fill = REKONLOKAL_SUSPECT_CATEGORY_FILL
                ws_t.cell(row=t.row, column=c).font = Font(color="FFFFFF")

    # Pembayaran Hutang - Keterangan Tambahan (I) ditulis "Paid to
    # <penerima>" (pakai Objek, Title Case) - sama pola dengan Gaji di
    # atas, cuma untuk kategori Pembayaran Hutang.
    for t in all_txns:
        if (t.effective_kategori or "").strip().lower() != "pembayaran hutang":
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        objek_asli = (ws_t.cell(row=t.row, column=8).value or "").strip()
        objek_title = " ".join(w.capitalize() for w in objek_asli.split()) if objek_asli else objek_asli
        if objek_asli:
            ws_t.cell(row=t.row, column=8, value=objek_title)
        ws_t.cell(row=t.row, column=9, value=f"Paid to {objek_title}" if objek_title else "-")

    # Keterangan Tambahan (I) untuk transaksi Penjualan/Shopeefood/
    # Grabfood sering berisi kode referensi teknis mentah dari mesin
    # EDC/QRIS (timestamp+kode transaksi+Teller ID, atau MID/CBG/QR/DDR)
    # - tidak menambah informasi berguna begitu transaksinya SENDIRI
    # sudah benar terklasifikasi sebagai Penjualan. Dibersihkan jadi
    # "-" KHUSUS untuk transaksi yang Kategori-nya SUDAH salah satu dari
    # 3 kategori Penjualan ini - tidak menyentuh transaksi lain (baris
    # yang belum jelas kategorinya, catatan referensi masih penting
    # untuk audit).
    _qris_teller_pattern = re.compile(
        r"^jam \d{2}:\d{2}:\d{2};.*teller/user id:\s*\S+$", re.IGNORECASE)
    _qris_mid_pattern = re.compile(
        r"^mid:\s*\d+;\s*cbg:\s*\d+;\s*qr\s*:\s*[\d.]+;\s*ddr:\s*[\d.]+(;.*)?$", re.IGNORECASE)
    # Kode referensi settlement BCA (mis. "0309/FTSCY/WS95051;
    # 0561864887; 24930FFV02186485", kadang ada embel-embel tambahan di
    # ujung seperti "; VISIONET INTERNASI") - format "<4 digit>/FTSCY/
    # <kode>; <nomor rekening>; <kode>", boleh diikuti keterangan
    # tambahan opsional setelahnya.
    _bca_ftscy_pattern = re.compile(
        r"^\d{4}/ftscy/\S+;\s*\d+;\s*\S+(;.*)?$", re.IGNORECASE)
    _PENJUALAN_KATEGORI = {"penjualan", "penjualan shopeefood", "penjualan grabfood"}
    for t in all_txns:
        if (t.kategori or "").strip().lower() not in _PENJUALAN_KATEGORI:
            continue
        ket_text = (t.ket or "").strip()
        if not ket_text or ket_text == "-":
            continue
        if (_qris_teller_pattern.match(ket_text) or _qris_mid_pattern.match(ket_text)
                or _bca_ftscy_pattern.match(ket_text)):
            wb_t, sheet_t = name_to_real[t.sheet]
            wb_t[sheet_t].cell(row=t.row, column=9, value="-")

    # Payment gateway settlement (Grabfood via Visionet, Shopeefood via
    # Airpay International) - Objek diseragamkan jadi nama merchant
    # yang jelas (bukan nama payment gateway teknis), dan Keterangan
    # Tambahan dikosongkan (tidak ada info tambahan yang perlu
    # dipertahankan untuk settlement rutin seperti ini). Juga berlaku
    # untuk 'Biaya Admin Bank' - Keterangan Tambahan-nya dikosongkan
    # tanpa syarat pola tertentu (lebih luas dari pembersihan QRIS di
    # atas yang cuma untuk pola spesifik).
    penjualan_fixed_ids = set()
    for t in all_txns:
        kat = (t.kategori or "").strip().lower()
        subjek_k = (t.subjek or "").strip().lower()
        objek_k = (t.objek or "").strip().lower()
        desc_k = (t.desc or "").strip().lower()
        ket_k = (t.ket or "").strip().lower()
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        # Selain cek Kategori/Subjek/Objek, JUGA cek Keterangan (B) -
        # sumber kadang sudah menulis "Penjualan Grabfood"/"Penjualan
        # Shopeefood" di Keterangan tapi Kategori-nya masih salah cuma
        # "Penjualan" generik (belum spesifik) - kalau ketemu lewat
        # Keterangan begini, Kategori JUGA dibetulkan, tidak cuma
        # Objek/Keterangan Tambahan.
        #
        # Arah Subjek/Objek untuk transaksi Penjualan: uang MENGALIR
        # DARI merchant/pembeli KE rekening penerima - jadi Subjek =
        # sumber uang (merchant), Objek = rekening tujuan (rekening
        # bank ini sendiri, t.sheet) - BUKAN sebaliknya seperti versi
        # lama (Subjek dibiarkan apa adanya, Objek ditulis nama
        # merchant, yang justru menggambarkan arah TERBALIK).
        if (kat == "penjualan grabfood" or subjek_k == "visionet" or objek_k == "visionet"
                or "grabfood" in desc_k or "visionet" in ket_k or "grabfood" in ket_k):
            ws_t.cell(row=t.row, column=3, value="Penjualan Grabfood")
            ws_t.cell(row=t.row, column=7, value="Grab Merchant")
            ws_t.cell(row=t.row, column=8, value=_display_account_name(t.sheet))
            ws_t.cell(row=t.row, column=9, value="-")
            penjualan_fixed_ids.add(id(t))
        elif (kat == "penjualan shopeefood" or subjek_k == "airpay" or objek_k == "airpay"
                or "shopeefood" in desc_k or "airpay" in ket_k or "shopeefood" in ket_k):
            ws_t.cell(row=t.row, column=3, value="Penjualan Shopeefood")
            ws_t.cell(row=t.row, column=7, value="Shopeefood Merchant")
            ws_t.cell(row=t.row, column=8, value=_display_account_name(t.sheet))
            ws_t.cell(row=t.row, column=9, value="-")
            penjualan_fixed_ids.add(id(t))
        elif "diva nispi yolanda" in objek_k:
            # "Diva Nispi Yolanda" di Objek = penanda transaksi ke
            # vendor "Kaola Cookies" - langsung ditetapkan Konsumsi dan
            # Liburan, Keterangan Tambahan (I) ditulis "Cookies Kaola"
            # (BUKAN Keterangan/B - beda dari pola vendor standar
            # lainnya, sesuai permintaan eksplisit user).
            ws_t.cell(row=t.row, column=3, value="Konsumsi dan Liburan")
            ws_t.cell(row=t.row, column=9, value="Cookies Kaola")
            penjualan_fixed_ids.add(id(t))
        elif "stoa space" in objek_k or "stoa space" in desc_k or "stoa space" in subjek_k:
            # "STOA SPACE HO" di Objek - penanda penjualan masuk via
            # QRIS/EDC BCA (nama tenant sendiri muncul di catatan bank
            # sebagai identitas terminal, BUKAN pihak lawan transaksi
            # sungguhan) - Kategori dipastikan Penjualan, Keterangan
            # diseragamkan, Subjek/Objek ikut arah Penjualan standar
            # (Subjek = sumber/label pembayar, Objek = rekening
            # penerima ini sendiri). Keterangan Tambahan (I) JUGA
            # ditulis label yang sama - tanpa ini, Objek (rekening
            # sendiri, mis. "BCA-887") salah kena "Unrecognized Tenant"
            # dari pass label akhir kalau file ini diproses ulang.
            ws_t.cell(row=t.row, column=3, value="Penjualan")
            ws_t.cell(row=t.row, column=2, value="Sales via EDC BCA")
            ws_t.cell(row=t.row, column=7, value="Sales via EDC BCA")
            ws_t.cell(row=t.row, column=8, value=_display_account_name(t.sheet))
            ws_t.cell(row=t.row, column=9, value=_settled_to_bank_label(t.sheet))
            penjualan_fixed_ids.add(id(t))
        elif kat == "penjualan" and t.sheet.strip().lower().startswith("bca"):
            # BCA juga menerima penjualan via EDC BCA sendiri (di luar
            # Grabfood/Shopeefood yang sudah ditangani cabang khusus di
            # atas, dan di luar pola "Stoa Space" yang menangkap kasus
            # Kategori BELUM Penjualan) - aturan generik sejajar dengan
            # BRI di bawah, supaya SEMUA transaksi Penjualan di BCA
            # konsisten dapat label "Sales via EDC BCA" terlepas ada
            # tidaknya teks "Stoa Space" spesifik.
            ws_t.cell(row=t.row, column=2, value="Sales via EDC BCA")
            ws_t.cell(row=t.row, column=7, value="Sales via EDC BCA")
            ws_t.cell(row=t.row, column=8, value=_display_account_name(t.sheet))
            ws_t.cell(row=t.row, column=9, value=_settled_to_bank_label(t.sheet))
            penjualan_fixed_ids.add(id(t))
        elif kat == "penjualan" and t.sheet.strip().lower().startswith("bri"):
            # BRI HANYA menerima penjualan via EDC BRI (BEDA dari BCA
            # yang JUGA menerima Shopeefood/Grabfood via payment
            # gateway, sudah ditangani cabang Grab/Shopeefood di atas
            # yang dicek LEBIH DULU) - user menegaskan SEMUA transaksi
            # Penjualan yang masuk ke rekening BRI genuinely EDC BRI,
            # jadi TIDAK perlu kata kunci spesifik seperti "Stoa Space"
            # untuk BCA - langsung diseragamkan begitu Kategori-nya
            # Penjualan dan rekeningnya BRI, konsisten dengan pola BCA.
            ws_t.cell(row=t.row, column=2, value="Sales via EDC BRI")
            ws_t.cell(row=t.row, column=7, value="Sales via EDC BRI")
            ws_t.cell(row=t.row, column=8, value=_display_account_name(t.sheet))
            ws_t.cell(row=t.row, column=9, value=_settled_to_bank_label(t.sheet))
            penjualan_fixed_ids.add(id(t))
        elif kat == "penjualan" and t.sheet.strip().lower().startswith("kas"):
            # Penjualan yang tercatat DI buku kas sendiri (bukan
            # settlement payment gateway seperti Grabfood/Shopeefood) -
            # ini penjualan CASH langsung, uangnya mengalir dari
            # penjualan tunai KE kas - Subjek = "Penjualan Cash", Objek
            # = buku kas ini sendiri (t.sheet, mis. "Kas/Buku").
            ws_t.cell(row=t.row, column=7, value="Penjualan Cash")
            ws_t.cell(row=t.row, column=8, value=_display_account_name(t.sheet))
        elif kat == "biaya admin bank":
            ws_t.cell(row=t.row, column=9, value="Admin Fee")
            penjualan_fixed_ids.add(id(t))

    # Rename kategori LEGACY (nama lama/pendek) ke nama resmi kontrak
    # kategori terbaru - transformasi yang MEMANG disengaja, konsisten
    # dengan Kategori Layer 1 resmi. HANYA menulis ulang kolom Kategori
    # (C) - Keterangan (B) dibiarkan apa adanya (sudah cukup deskriptif,
    # cuma label Kategori-nya yang ketinggalan zaman).
    legacy_renamed_ids = set()
    for t in all_txns:
        if t.is_opening:
            continue
        asli = (t.kategori or "").strip().lower()
        target = _LEGACY_KATEGORI_RENAME.get(asli)
        if target is None or target == (t.kategori or "").strip():
            continue
        legacy_renamed_ids.add(id(t))
        wb_t, sheet_t = name_to_real[t.sheet]
        wb_t[sheet_t].cell(row=t.row, column=3, value=target)

    # Objek diinferensi dari pola di Keterangan Tambahan (I) - beberapa
    # bank menulis catatan referensi yang menyebut nama penerima/pihak
    # lain di sana, tapi Objek transaksinya sendiri kosong/generik.
    #   "GoPay <nomor>" -> Objek "Gopay Owner"
    #   "BCA <nomor rekening>" -> Objek "BCA-<3 digit terakhir>"
    _gopay_ket_pattern = re.compile(r"\bgopay\s+0?\d{6,}\b", re.IGNORECASE)
    # Nomor rekening BCA GENUINE biasanya 10 digit (kadang 9-11
    # tergantung jenis rekening) - dibatasi rentang ini (BUKAN "6 atau
    # lebih" seperti sebelumnya) supaya tidak salah menangkap kode
    # referensi lain yang kebetulan diawali "BCA" tapi angkanya JAUH
    # lebih panjang (mis. kode QRIS/merchant 19 digit seperti
    # "9360001430017297828" - itu BUKAN indikasi rekening BCA sungguhan,
    # cuma kebetulan format teksnya mirip).
    _bca_ket_pattern = re.compile(r"\bbca\s+(\d{9,11})\b", re.IGNORECASE)
    _bca_label_pattern = re.compile(r"^bca-\d{3}$", re.IGNORECASE)
    for t in all_txns:
        ket_text = t.ket or ""
        objek_baru = None
        if _gopay_ket_pattern.search(ket_text):
            objek_baru = "Gopay Owner"
        else:
            m = _bca_ket_pattern.search(ket_text)
            if m:
                objek_baru = f"BCA-{m.group(1)[-3:]}"
            elif _bca_label_pattern.match((t.objek or "").strip()):
                # Objek SUDAH menyandang label "BCA-XXX" (dari run
                # sebelumnya, sebelum pembatasan 9-11 digit ada), tapi
                # Keterangan Tambahan SAAT INI tidak mendukung label itu
                # sebagai indikasi rekening BCA genuine (baik karena
                # angkanya di luar rentang genuine, ATAU tidak ada pola
                # "BCA <angka>" sama sekali di teksnya) - batalkan label
                # basi itu, jangan dibiarkan tersandang tanpa dasar.
                wb_t, sheet_t = name_to_real[t.sheet]
                wb_t[sheet_t].cell(row=t.row, column=8, value="-")
                continue
        if objek_baru is None or (t.objek or "").strip() == objek_baru:
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        wb_t[sheet_t].cell(row=t.row, column=8, value=objek_baru)

    # Keterangan berupa KODE ANGKA PANJANG (mis. nomor rekening pengirim
    # diulang - format umum di BRI) - diseragamkan jadi "Setoran
    # <rekening>", konsisten dengan pola setoran tunai lain, BUKAN
    # dibiarkan sebagai deretan angka yang tidak informatif. Hanya untuk
    # transaksi KREDIT (uang masuk) - konsisten dengan
    # _looks_like_long_numeric_code() yang sudah dipakai di
    # effective_kategori untuk kasus yang sama.
    for t in all_txns:
        if t.nominal <= 0 or not _looks_like_long_numeric_code(t.desc):
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        ws_t.cell(row=t.row, column=9, value=ws_t.cell(row=t.row, column=2).value)
        ws_t.cell(row=t.row, column=2, value=f"Setoran {_display_account_name(t.sheet)}")

    # Vendor Belanja Bahan/Kemasan yang sering ditulis beda-beda di
    # sumber - diseragamkan jadi satu nama baku, Kategori dipastikan
    # benar. Keterangan lama diarsip ke Keterangan Tambahan dulu.
    vendor_fixed_ids = set()
    for t in all_txns:
        if t.is_opening:
            continue
        info = _kas_buku_vendor_info(t)
        if info is None:
            continue
        keterangan_baru, kategori_baru, objek_baru = info
        sudah_benar = (t.desc == keterangan_baru and (t.kategori or "").strip() == kategori_baru
                       and (objek_baru is None or (t.objek or "").strip() == objek_baru))
        if sudah_benar:
            continue  # sudah benar, tidak perlu apa-apa
        vendor_fixed_ids.add(id(t))
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        # Vendor sudah "confirmed" (dikenali pasti dari daftar) -
        # Keterangan Tambahan ditulis "Paid Off to <tenant>" sesuai
        # permintaan user, BUKAN diarsip (vendor yang sudah pasti
        # dikenali dari daftar tidak perlu menyimpan catatan referensi
        # lama).
        tenant_label = objek_baru or keterangan_baru
        ws_t.cell(row=t.row, column=9, value=f"Paid Off to {tenant_label}")
        ws_t.cell(row=t.row, column=2, value=keterangan_baru)
        ws_t.cell(row=t.row, column=3, value=kategori_baru)
        if objek_baru is not None:
            ws_t.cell(row=t.row, column=8, value=objek_baru)

    # Vendor yang trigger-nya dari kolom Objek (bukan Keterangan) - lihat
    # _OBJEK_VENDOR_RULES/_objek_vendor_info.
    for t in all_txns:
        if t.is_opening:
            continue
        info = _objek_vendor_info(t)
        if info is None:
            continue
        keterangan_baru, kategori_baru, objek_baru = info
        sudah_benar = (t.desc == keterangan_baru and (t.kategori or "").strip() == kategori_baru
                       and (t.objek or "").strip() == objek_baru)
        if sudah_benar:
            continue
        vendor_fixed_ids.add(id(t))
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        # Sama seperti pass vendor Keterangan - vendor sudah "confirmed"
        # dari Objek, Keterangan Tambahan ditulis "Paid Off to <tenant>".
        ws_t.cell(row=t.row, column=9, value=f"Paid Off to {objek_baru}")
        ws_t.cell(row=t.row, column=2, value=keterangan_baru)
        ws_t.cell(row=t.row, column=3, value=kategori_baru)
        ws_t.cell(row=t.row, column=8, value=objek_baru)

    # Tokopedia: default Kolom B "Belanja Tokopedia", KECUALI nominal
    # besar (asumsi ambang Rp2.000.000, item mahal cenderung Assets/
    # peralatan, bukan bahan habis pakai) -> Kolom B "Belanja Assets".
    # HANYA Kolom B yang disentuh - Kategori (Kolom C, tetap "Belanja
    # Bahan" dari rule di atas) dan Objek TIDAK diubah oleh pass ini.
    # ASUMSI ambang nominal ini perlu dikonfirmasi/disesuaikan user.
    _TOKOPEDIA_ASSET_THRESHOLD = 2_000_000
    for t in all_txns:
        if t.is_opening:
            continue
        if (t.objek or "").strip().lower() != "tokopedia":
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        if abs(t.nominal) >= _TOKOPEDIA_ASSET_THRESHOLD:
            ws_t.cell(row=t.row, column=2, value="Belanja Assets")

    # OVO: dipakai khusus untuk bayar Grabfood - Kolom B "Konsumsi
    # Internal via Grabfood", KECUALI baris ini sudah di-override manual
    # jadi "Pengeluaran Pribadi" (Kategori Kolom C), sesuai penegasan
    # user. HANYA Kolom B yang disentuh.
    for t in all_txns:
        if t.is_opening:
            continue
        if (t.effective_kategori or "").strip().lower() == "pengeluaran pribadi":
            continue
        teks_asli = f"{t.desc or ''} {t.objek or ''} {t.subjek or ''}".lower()
        if "ovo" not in teks_asli:
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        ws_t.cell(row=t.row, column=2, value="Konsumsi Internal via Grabfood")

    # Rekening Keluarga/Owner: dicurigai transaksi internal terlebih
    # dahulu (dua arah, masuk maupun keluar) - Kolom B "dari Rekening
    # Keluarga". HANYA Kolom B yang disentuh, Kategori/Subjek/Objek asli
    # dibiarkan apa adanya untuk diverifikasi manual.
    for t in all_txns:
        if t.is_opening:
            continue
        subjek_lc = (t.subjek or "").strip().lower()
        objek_lc = (t.objek or "").strip().lower()
        if not (subjek_lc.startswith("rekening keluarga") or objek_lc.startswith("rekening keluarga")):
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        ws_t.cell(row=t.row, column=2, value="dari Rekening Keluarga")

    # Kategori mencurigakan - Kategori TERSIMPAN beda dari yang
    # DIHITUNG sistem berdasarkan Keterangan/Objek (effective_kategori,
    # logika sama persis dipakai flow rekonsiliasi utama). HANYA
    # menandai (highlight ungu), TIDAK PERNAH mengubah Kategori/
    # Keterangan aslinya - ini murni sinyal untuk audit manual, bukan
    # auto-koreksi (beberapa mismatch bisa jadi false positive, aturan
    # kata kunci tidak selalu sempurna menangkap konteks).
    #
    # 'Belanja Operasional' -> 'Overhead', 'Reparasi' -> 'Reparasi dan
    # Maintenance Tools dan Mesin', dst DIKECUALIKAN dari sini - sudah
    # AKTIF dikoreksi di pass _legacy_kategori_rename di atas (lihat
    # legacy_renamed_ids), bukan kesalahan kategorisasi genuine yang
    # perlu ditandai untuk audit manual lagi.
    #
    # Kasus KHUSUS: kalau Kategori ASLI-nya literal "New Kategori"/
    # "Kategori Baru" (placeholder GENERIK "belum diketahui", bukan
    # kategori spesifik yang salah) DAN sistem SUDAH bisa menentukan
    # kategori spesifik dengan yakin dari Keterangan/Objek/Subjek
    # (effective_kategori bukan lagi 'Kategori Baru') - AKTIF dikoreksi
    # juga, BUKAN cuma diflag ungu. Beda dari kasus Kategori SPESIFIK
    # yang salah (mis. Kategori sudah 'Overhead' tapi harusnya
    # 'Reparasi dan Maintenance Tools dan Mesin' - itu tetap cuma
    # diflag, karena ada KEMUNGKINAN kategori spesifik yang tersimpan
    # itu sengaja/benar dan aturan kata kunci-lah yang salah tangkap -
    # sedangkan "New Kategori" jelas-jelas "belum diketahui", tidak ada
    # nilai tersimpan yang perlu dipertahankan.
    new_kategori_fixed_ids = set()
    for t in all_txns:
        if t.is_opening:
            continue
        # Baris yang SUDAH dikoreksi aktif oleh pass vendor/legacy-
        # rename/Grabfood-Shopeefood di atas dilewati - vendor pass bisa
        # menulis Kategori yang LEBIH SPESIFIK/BENAR (mis. 'Belanja
        # Bahan' untuk vendor Shopee tertentu) daripada apa yang
        # dihitung effective_kategori dari aturan kata kunci GENERIK
        # (mis. 'shopee' -> 'Overhead' di CATEGORY_OVERRIDE_RULES, yang
        # tidak tahu-menahu soal vendor spesifik ini) - t.kategori
        # sendiri TIDAK berubah (Txn tidak dimutasi saat menulis ke
        # sel), jadi tanpa pengecualian ini, pass ini akan salah
        # menimpa BALIK hasil vendor yang sudah benar dengan hasil
        # effective_kategori yang lebih generik.
        if id(t) in vendor_fixed_ids or id(t) in legacy_renamed_ids or id(t) in penjualan_fixed_ids:
            continue
        asli = (t.kategori or "").strip().lower()
        if asli not in ("new kategori", "kategori baru"):
            continue
        hitung = t.effective_kategori
        if not hitung or hitung.strip().lower() == "kategori baru":
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        wb_t[sheet_t].cell(row=t.row, column=3, value=hitung)
        new_kategori_fixed_ids.add(id(t))

    # Transaksi MASUK (kredit) yang Subjek/pengirimnya adalah nama pegawai
    # dari roster gaji (EMPLOYEE_ALIASES, BUKAN owner) - kemungkinan besar
    # setoran tunai hasil penjualan yang disetor pegawai ke rekening bank,
    # BUKAN transfer dari pihak luar. Dipaksa jadi "Setoran Tunai"
    # (Kolom B) + "Transaksi Internal" (Kolom C), SEKALIGUS ditandai
    # Suspicious (perlu verifikasi manual - kenapa nama pegawai yang
    # muncul sebagai pengirim, bukan nama rekening bank/kas sendiri).
    _employee_only_keywords = sorted(
        {k for k, v in EMPLOYEE_ALIASES.items() if v != "Ahmad Roziyan Hidayat"},
        key=len, reverse=True,
    )
    n_kategori_mencurigakan = 0
    for t in all_txns:
        if t.is_opening or t.nominal <= 0:
            continue
        subjek_lc = (t.subjek or "").strip().lower()
        if not subjek_lc:
            continue
        if not any(re.search(r"\b" + re.escape(kw) + r"\b", subjek_lc) for kw in _employee_only_keywords):
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        ws_t.cell(row=t.row, column=2, value="Setoran Tunai")
        ws_t.cell(row=t.row, column=3, value="Transaksi Internal")
        ws_t.cell(row=t.row, column=9, value="Suspicious")
        for c in range(1, 10):
            ws_t.cell(row=t.row, column=c).fill = REKONLOKAL_SUSPECT_CATEGORY_FILL
            ws_t.cell(row=t.row, column=c).font = Font(color="FFFFFF")
        vendor_fixed_ids.add(id(t))  # cegah pass "Kategori mencurigakan" di bawah menimpa/bersihkan balik
        n_kategori_mencurigakan += 1

    # Transaksi MASUK (kredit) yang Keterangan-nya masih "Belanja ..."
    # (kategori pengeluaran) - JANGGAL kalau arahnya masuk, kemungkinan
    # besar ini sebenarnya setoran tunai yang salah tercatat. HANYA
    # berlaku di rekening BANK (BRI/BCA/BSI/Jago/dst) - sheet Kas Buku/
    # Kas Kasir DILINDUNGI dan TIDAK PERNAH kena override ini (setoran/
    # penarikan kas fisik di buku kas memang wajar, bukan anomali),
    # sesuai penegasan user.
    for t in all_txns:
        if t.is_opening or t.nominal <= 0:
            continue
        if id(t) in vendor_fixed_ids:
            continue  # sudah ditangani pass lain (mis. setoran tunai pegawai di atas)
        if (t.sheet or "").strip().lower().startswith("kas"):
            continue  # lindungi Kas Buku/Kas Kasir - tidak pernah disentuh pass ini
        desc_lc = (t.desc or "").strip().lower()
        if not desc_lc.startswith("belanja"):
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        ws_t.cell(row=t.row, column=2, value="Setoran Tunai")
        ws_t.cell(row=t.row, column=3, value="Transaksi Internal")
        ws_t.cell(row=t.row, column=9, value="Suspicious")
        for c in range(1, 10):
            ws_t.cell(row=t.row, column=c).fill = REKONLOKAL_SUSPECT_CATEGORY_FILL
            ws_t.cell(row=t.row, column=c).font = Font(color="FFFFFF")
        vendor_fixed_ids.add(id(t))
        n_kategori_mencurigakan += 1

    for t in all_txns:
        if t.is_opening:
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        # baru saja dibetulkan AKTIF oleh pass vendor/legacy-rename/
        # Grabfood-Shopeefood di atas - diperlakukan SAMA seperti "tidak
        # (lagi) mencurigakan" (bukan di-skip total), supaya baris ini
        # TETAP dapat kesempatan dibersihkan dari ungu BASI kalau
        # kebetulan file ini sudah pernah diproses SEBELUM koreksi aktif
        # ini ada - kalau di-skip total, ungu basi dari run lama tidak
        # akan pernah terhapus untuk baris yang SEKARANG sudah aktif
        # dikoreksi.
        sudah_dikoreksi_aktif = (id(t) in vendor_fixed_ids or id(t) in legacy_renamed_ids
                                  or id(t) in penjualan_fixed_ids or id(t) in new_kategori_fixed_ids)
        asli = (t.kategori or "").strip().lower()
        hitung = (t.effective_kategori or "").strip().lower()
        if sudah_dikoreksi_aktif or not asli or asli == hitung or hitung == "kategori baru":
            # TIDAK/tidak lagi mencurigakan - tapi kalau baris ini masih
            # bertahan warna ungu dari RUN /rekonlokal SEBELUMNYA (mis.
            # kategori sekarang sudah konsisten setelah perbaikan aturan,
            # padahal saat run lalu masih dianggap mismatch), bersihkan
            # supaya tidak menyesatkan seolah masih perlu diaudit -
            # highlight grup kategori di bawah akan mewarnai ulang baris
            # ini dengan benar kalau memang termasuk salah satu grup.
            cell_b = ws_t.cell(row=t.row, column=2)
            current_fill = cell_b.fill.fgColor.rgb if cell_b.fill else None
            if current_fill == "008B0000" and id(t) not in gaji_ambiguous_ids:
                for c in range(1, 10):
                    ws_t.cell(row=t.row, column=c).fill = PatternFill(fill_type=None)
                    ws_t.cell(row=t.row, column=c).font = Font(color="FF000000")
            continue
        n_kategori_mencurigakan += 1
        ws_t.cell(row=t.row, column=9, value="Suspicious")
        for c in range(1, 10):
            ws_t.cell(row=t.row, column=c).fill = REKONLOKAL_SUSPECT_CATEGORY_FILL
            ws_t.cell(row=t.row, column=c).font = Font(color="FFFFFF")

    # Highlight grup kategori (Penjualan/Belanja Bahan+Kemasan/dst) -
    # HANYA diterapkan pada baris yang BELUM punya highlight dari pass
    # lain (transfer biru/orange/kuning, belum-direkon merah, kategori
    # mencurigakan ungu) - supaya highlight yang lebih spesifik/penting
    # itu tidak tertimpa oleh pewarnaan kelompok yang sifatnya cuma
    # visual, bukan penanda audit.
    for t in all_txns:
        if t.is_opening:
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        cell_b = ws_t.cell(row=t.row, column=2)
        current_fill = cell_b.fill.fgColor.rgb if cell_b.fill else None
        if current_fill in _MEANINGFUL_HIGHLIGHT_HEXES:
            continue  # sudah ada highlight lain yang lebih penting, jangan ditimpa
        kategori_sekarang = ws_t.cell(row=t.row, column=3).value
        fill = category_group_fill(kategori_sekarang)
        if fill is None:
            continue
        for c in range(1, 10):
            ws_t.cell(row=t.row, column=c).fill = fill

    # Label akhir untuk baris yang BELUM disentuh pass manapun di atas -
    # dicek dari Keterangan Tambahan (I) yang MASIH SAMA PERSIS dengan
    # nilai mentah aslinya (t.ket) - kalau sudah beda, berarti SUDAH
    # dapat label bermakna dari salah satu pass sebelumnya (Solved,
    # Medium Unresolved, Unresolved, Suspicious, Paid Off to, Paid to,
    # -, dll), jadi TIDAK disentuh lagi di sini.
    # - Objek KOSONG ('-'/kosong) -> dibiarkan apa adanya, TIDAK ditulis
    #   apa-apa (sesuai permintaan eksplisit user).
    # - Objek TERISI dan DIKENALI (cocok EMPLOYEE_ALIASES atau nama
    #   vendor yang sudah terdaftar di sistem) -> "Paid Off to <Objek>".
    # - Objek TERISI tapi TIDAK dikenali sama sekali -> "Unrecognized
    #   Tenant".
    _known_tenant_names = {v.lower() for v in EMPLOYEE_ALIASES.values()}
    for _rules_list in (_KAS_BUKU_VENDOR_RULES, _OBJEK_VENDOR_RULES):
        for _keywords, _ket, _kat, _objek in _rules_list:
            _known_tenant_names.add(_ket.lower())
            if _objek:
                _known_tenant_names.add(_objek.lower())
    _known_tenant_names.update({"grab merchant", "shopeefood merchant", "gopay owner"})
    for t in all_txns:
        if t.is_opening:
            continue
        wb_t, sheet_t = name_to_real[t.sheet]
        ws_t = wb_t[sheet_t]
        current_i = ws_t.cell(row=t.row, column=9).value
        if _is_known_i_label(current_i) or (current_i or "") != (t.ket or ""):
            continue  # sudah dapat label dari pass lain (run ini ATAU run sebelumnya), jangan disentuh
        objek_sekarang = (ws_t.cell(row=t.row, column=8).value or "").strip()
        if not objek_sekarang or objek_sekarang == "-":
            continue  # tenant kosong - biarkan kosong
        if objek_sekarang.lower() in _known_tenant_names:
            ws_t.cell(row=t.row, column=9, value=f"Paid Off to {objek_sekarang}")
        elif (t.effective_kategori or "").strip().lower() in (
                "penjualan", "penjualan shopeefood", "penjualan grabfood"):
            # Fallback: transaksi Penjualan APAPUN yang belum tertangani
            # pass spesifik manapun di atas (mis. Grabfood/Shopeefood
            # settlement, atau bank selain BRI/BCA) - TETAP TIDAK boleh
            # dianggap soal "tenant" (Objek di sini adalah rekening
            # PENERIMA, bukan tenant/vendor) - label "Settled to <bank>"
            # sesuai permintaan user, bukan "Unrecognized Tenant".
            ws_t.cell(row=t.row, column=9, value=_settled_to_bank_label(t.sheet))
        else:
            ws_t.cell(row=t.row, column=9, value="Unrecognized Tenant")

    # Pembersihan akhir per sheet - hapus baris tanpa angka, samakan
    # format mata uang, hitung ulang & patenkan (nilai statis, bukan
    # formula) Saldo Awal/Total Debit/Total Kredit/Saldo Akhir sebagai
    # acuan tetap untuk verifikasi manual user. Dijalankan PALING AKHIR,
    # setelah semua koreksi lain selesai (row number sudah final).
    for sn in sheets1:
        _cleanup_and_verify_sheet(wb1[sn])
    for sn in sheets2:
        _cleanup_and_verify_sheet(wb2[sn])

    wb1.save(out1)
    wb2.save(out2)
    return {
        "n_high": n_high, "n_medium": n_medium, "n_low": n_low, "n_belum_rekon": n_belum_rekon,
        "n_kategori_mencurigakan": n_kategori_mencurigakan,
    }


def correct_and_highlight_matched_transfers(wb, matches, combo_matches):
    """Untuk tiap transfer yang SUDAH ketemu pasangannya (matches dengan
    dst terisi, confidence High/Medium/Low, atau combo_matches):
    1. KOREKSI Subjek/Objek di kedua sisi (src & dst) supaya benar-benar
       menyebut rekening lawan transaksinya - banyak data sumber
       menulis Objek = nama rekening SENDIRI (tidak berguna untuk audit,
       mis. sheet BCA-292(Biz) isi Objek-nya 'BCA-292(Biz)' juga),
       padahal sistem SUDAH TAHU pasangan sebenarnya dari hasil
       pencocokan tanggal+nominal - jadi ditulis ulang jadi akurat.
    2. HIGHLIGHT warna berdasarkan grup bank TUJUAN uang (bukan asal) -
       biru=BCA, orange=Jago, kuning=BRI - diterapkan di KEDUA sisi
       (baris pengirim maupun penerima), supaya audit visual langsung
       kelihatan kemana uang itu benar-benar mengalir tanpa perlu buka
       sheet Rekonsiliasi."""
    def _base_name(sheet_title):
        return sheet_title.rsplit(" ", 2)[0]

    def _apply(t1, t2):
        if t1 is None or t2 is None:
            return
        if t1.sheet not in wb.sheetnames or t2.sheet not in wb.sheetnames:
            return
        pengirim, penerima = _sender_receiver(t1, t2)
        ws_pengirim = wb[pengirim.sheet]
        ws_penerima = wb[penerima.sheet]
        ws_pengirim.cell(row=pengirim.row, column=8, value=_base_name(penerima.sheet))  # Objek pengirim = penerima
        ws_penerima.cell(row=penerima.row, column=7, value=_base_name(pengirim.sheet))  # Subjek penerima = pengirim
        fill = _bank_group_fill(penerima.sheet)
        if fill is not None:
            for c in range(1, 10):
                ws_pengirim.cell(row=pengirim.row, column=c).fill = fill
                ws_penerima.cell(row=penerima.row, column=c).fill = fill

    for m in matches:
        if m.dst is not None and m.confidence in ("High", "Medium", "Low"):
            _apply(m.src, m.dst)
    for combo in combo_matches:
        # split/merge: src (satu sisi) vs 2 parts (sisi lain, di rekening
        # yang sama) - arah pengirim/penerima ditentukan otomatis di
        # dalam _apply() lewat tanda nominal, urutan argumen di sini
        # tidak lagi krusial.
        src = combo["src"]
        for part in combo["parts"]:
            _apply(src, part)


# ---------------------------------------------------------------------------
# Laporan Laba Rugi
# ---------------------------------------------------------------------------

INCOME_CATEGORIES_REVENUE = ["Penjualan", "Penjualan Shopeefood", "Penjualan Grabfood"]

# Kategori beban yang dicocokkan persis apa adanya (SUMIF biasa) - LAYER
# 1 (bot konversi), tidak boleh ada istilah Layer 2 (COGS/OpEx/CapEx) di
# sini, itu murni hasil roll-up di bagian "RINGKASAN LAYER 2" Laba Rugi.
INCOME_CATEGORIES_EXPENSE = [
    "Belanja Bahan",
    "Overhead",
    "Konsumsi dan Liburan",
    "Belanja Utilitas",
    "Tools dan Equipments",
    "Kemasan",
    "Subscription",
    "Sewa dan Maintenance Bangunan",
    "Reparasi dan Maintenance Tools dan Mesin",
    "Pajak dan Administrasi",
    "Belanja Assets",
]

# Marketing dan Riset dan Pengembangan (RnD) digabung jadi satu baris -
# dulu "Riset dan Pengembangan" (kategori baru, dipicu keyword "Pelatihan")
# tidak terdaftar sama sekali di INCOME_CATEGORIES_EXPENSE, jadi uangnya
# hilang dari Laba Rugi (sumber selisih Neraca di BCA).
MARKETING_RND_CATEGORY_TEXTS = ["Marketing", "Riset dan Development"]

# Gaji: dulu satu baris per "Gaji <Bulan> <Tahun>" (mis. "Gaji Desember 2024")
# Gaji: dulu satu baris per "Gaji <Bulan> <Tahun>" (mis. "Gaji Desember 2024")
# yang berarti daftar kategori harus terus ditambah tiap tahun, dan asumsi
# lama (info bulan ada di kolom Kategori) ternyata tidak berlaku di semua
# parser - versi bank ("preformatted") pakai Kategori tetap "Gaji Pegawai"
# untuk SEMUA gaji, info bulannya cuma ada di Keterangan (mis. "Gaji
# Latifatul Husna Januari"). Makanya deteksi bulan ini/lalu sekarang
# dicocokkan ke kolom KETERANGAN (bukan Kategori), dengan Kategori cuma
# dipakai untuk memastikan barisnya memang tentang gaji (wildcard "Gaji*").
def gaji_category_patterns(period_month):
    """Return (nama_bulan_ini, nama_bulan_lalu) berdasarkan nomor bulan
    periode (1-12). Kalau bulan tidak terdeteksi, return (None, None)."""
    if not period_month:
        return None, None
    prev_month = 12 if period_month == 1 else period_month - 1
    return MONTHS_ID[period_month], MONTHS_ID[prev_month]


def sumif_gaji_bulan_formula(sheet, last_row, nama_bulan):
    """SUMIFS: baris berkategori 'Gaji*' DAN keterangannya menyebut nama
    bulan tertentu (mis. '*Januari*') - menangani baik gaya lama (bulan ada
    di Kategori) maupun gaya baru/bank (Kategori tetap 'Gaji Pegawai', bulan
    cuma disebut di Keterangan)."""
    rng_c = f"'{sheet}'!$M$2:$M${last_row}"
    rng_b = f"'{sheet}'!$B$2:$B${last_row}"
    rng_j = f"'{sheet}'!$J$2:$J${last_row}"
    if not nama_bulan:
        return f"=SUMIF({rng_c},\"Gaji*\",{rng_j})"
    return f"=SUMIFS({rng_j},{rng_c},\"Gaji*\",{rng_b},\"*{nama_bulan}*\")"


# Biaya admin, biaya admin transfer (Fliptech, auto dari Rekonsiliasi), bunga,
# dan pajak bank digabung jadi SATU baris "Biaya Admin Bank" - dulu terpecah
# karena bank/versi parser beda pakai istilah berbeda (Biaya Admin & Pajak
# Bank / Biaya Admin dan Bunga Bank / Bunga dan Admin Bank / kini disatukan
# jadi "Biaya Admin Bank" di parser terbaru), padahal secara ekonomi
# sama-sama biaya jasa perbankan. Daftar lama tetap disertakan supaya file
# historis yang masih pakai istilah lama tetap tertangkap.
BANK_FEE_CATEGORY_TEXTS = [
    "Biaya Admin Bank",
    "Biaya Admin & Pajak Bank",
    "Biaya Admin dan Bunga Bank",
    "Bunga dan Admin Bank",
]

OTHER_CATEGORIES = ["Tip/Minus/Lebih"]
# 'Penarikan'/'Penerimaan' SENGAJA DIHAPUS dari daftar kategori resmi -
# terlalu ambigu (staff cash withdrawal vs uang masuk tak terklasifikasi)
# untuk otomatis diterima begitu saja. Kalau ketemu Kategori persis ini
# di data sumber, sekarang JATUH ke 'Kategori Baru' (perlu diaudit
# manual), bukan diam-diam dianggap sudah benar.
# 'Pembayaran Hutang' SENGAJA dikeluarkan dari sini - pelunasan pokok
# hutang BUKAN beban bisnis (tidak boleh mengurangi Laba Rugi), itu
# pengurang LIABILITAS. Ditangani di bagian Liabilitas Neraca bareng
# 'Hutang Masuk', bukan di Laba Rugi - lihat write_balance_sheet.


def _is_recognized_category(kategori):
    """True kalau `kategori` (Kategori ASLI dari sumber, sebelum
    override) sudah salah satu kategori yang dikenal sistem - dicek
    lewat effective_kategori's fallback chain SEBELUM jatuh ke 'Kategori
    Baru'. Mencakup: kategori protected (modal, saldo awal, dst), gaji*,
    kategori transfer-like, kategori modal, dan exact match ke semua
    kategori Laba Rugi/Lain-lain yang dikenal."""
    k = (kategori or "").strip().lower()
    if not k:
        return False
    if k in _PROTECTED_FROM_CATEGORY_OVERRIDE or k.startswith("gaji") or k.startswith("modal"):
        return True
    if any(kw in k for kw in TRANSFER_KEYWORDS):
        return True
    if any(kw in k for kw in CAPITAL_KEYWORDS):
        return True
    if any(kw in k for kw in PERSONAL_EXPENSE_KEYWORDS):
        return True
    known_exact = {
        c.lower() for c in (
            INCOME_CATEGORIES_EXPENSE + INCOME_CATEGORIES_REVENUE +
            MARKETING_RND_CATEGORY_TEXTS + BANK_FEE_CATEGORY_TEXTS + OTHER_CATEGORIES +
            ["Hutang Masuk", "Pembayaran Hutang"]
        )
    }
    return k in known_exact



# ---------------------------------------------------------------------------
# Helper pivot: setiap laporan keuangan ditulis per-rekening (kolom),
# dengan kolom paling kanan = TOTAL keseluruhan. Ini supaya selisih di
# Neraca/Diagnostik bisa langsung ditelusuri ke rekening mana penyebabnya,
# tanpa harus buka satu-satu.
# ---------------------------------------------------------------------------

def col_letter(i):
    return get_column_letter(i)


def pivot_total_col(sheets):
    return 2 + len(sheets)


def write_pivot_header(ws, row, sheets, label=""):
    ws.cell(row=row, column=1, value=label)
    for i, sheet in enumerate(sheets):
        c = 2 + i
        cell = ws.cell(row=row, column=c, value=sheet)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    total_col = pivot_total_col(sheets)
    ws.cell(row=row, column=total_col, value="TOTAL")
    style_header(ws, row, total_col)
    ws.row_dimensions[row].height = 32


def write_pivot_section(ws, row, label, sheets):
    """Baris judul section (mis. 'PENDAPATAN'), fill/bold di seluruh lebar tabel."""
    total_col = pivot_total_col(sheets)
    ws.cell(row=row, column=1, value=label)
    for c in range(1, total_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = SECTION_FONT
        cell.fill = SECTION_FILL


def write_pivot_data_row(ws, row, label, sheets, formula_fn, bold=False):
    """formula_fn(sheet) -> formula string (SUMIF per sheet, dll). Kolom
    TOTAL diisi SUM dari sel-sel per rekening di baris yang sama (bukan
    dihitung ulang terpisah), supaya konsisten dan gampang dicek manual."""
    ws.cell(row=row, column=1, value=label)
    n = len(sheets)
    for i, sheet in enumerate(sheets):
        c = 2 + i
        cell = ws.cell(row=row, column=c, value=formula_fn(sheet))
        cell.number_format = NUMBER_FORMAT
    total_col = 2 + n
    ws.cell(row=row, column=total_col,
            value=f"=SUM({col_letter(2)}{row}:{col_letter(1 + n)}{row})")
    ws.cell(row=row, column=total_col).number_format = NUMBER_FORMAT
    if bold:
        for c in range(1, total_col + 1):
            ws.cell(row=row, column=c).font = Font(bold=True)


def write_pivot_subtotal_row(ws, row, label, sheets, ref_rows, bold=True):
    """Subtotal per kolom = SUM baris-baris ref_rows di kolom yang sama.
    Kolom TOTAL = SUM sel-sel per rekening di baris subtotal itu sendiri."""
    ws.cell(row=row, column=1, value=label)
    n = len(sheets)
    for i in range(n):
        c = 2 + i
        cl = col_letter(c)
        ws.cell(row=row, column=c, value=f"=SUM({cl}{ref_rows[0]}:{cl}{ref_rows[-1]})")
        ws.cell(row=row, column=c).number_format = NUMBER_FORMAT
    total_col = 2 + n
    ws.cell(row=row, column=total_col,
            value=f"=SUM({col_letter(2)}{row}:{col_letter(1 + n)}{row})")
    ws.cell(row=row, column=total_col).number_format = NUMBER_FORMAT
    if bold:
        for c in range(1, total_col + 1):
            ws.cell(row=row, column=c).font = Font(bold=True)


def write_pivot_formula_row(ws, row, label, sheets, per_col_formula_fn, bold=False):
    """Baris hasil kombinasi rumus antar-baris (mis. Laba = Pendapatan+Beban),
    per_col_formula_fn(col_letter) -> formula string, dipakai sama persis
    untuk tiap kolom rekening MAUPUN kolom TOTAL (referensi sel berbeda,
    logika sama)."""
    ws.cell(row=row, column=1, value=label)
    n = len(sheets)
    for i in range(n):
        c = 2 + i
        cl = col_letter(c)
        ws.cell(row=row, column=c, value=per_col_formula_fn(cl))
        ws.cell(row=row, column=c).number_format = NUMBER_FORMAT
    total_col = 2 + n
    ws.cell(row=row, column=total_col, value=per_col_formula_fn(col_letter(total_col)))
    ws.cell(row=row, column=total_col).number_format = NUMBER_FORMAT
    if bold:
        for c in range(1, total_col + 1):
            ws.cell(row=row, column=c).font = Font(bold=True)


def sumif_one_sheet(sheet, last_row, category):
    return (f"=SUMIF('{sheet}'!$M$2:$M${last_row},\"{category}\","
            f"'{sheet}'!$J$2:$J${last_row})")


def sumif_modal_one_sheet(sheet, last_row):
    """Modal & Setoran Pemilik + kasus 'Transfer Masuk ... dari rekening
    sendiri' (uang milik owner sendiri dipindah antar rekening, mis.
    pencairan investasi pribadi) - lihat catatan di CAPITAL_SELF_TRANSFER_KEYWORDS.
    Ditambah 'Laba Ditahan Bulanan' (laba bulan berjalan yang disimpan,
    biasanya dimasukkan sebagai modal baru bulan berikutnya atau dana
    darurat) - user menegaskan ini dianggap kategori modal dari owner."""
    rng_c = f"'{sheet}'!$M$2:$M${last_row}"
    rng_b = f"'{sheet}'!$B$2:$B${last_row}"
    rng_j = f"'{sheet}'!$J$2:$J${last_row}"
    return (f"=SUMIF({rng_c},\"Modal*\",{rng_j})"
            f"+SUMIF({rng_c},\"Laba Ditahan Bulanan\",{rng_j})"
            f"+SUMIFS({rng_j},{rng_c},\"Transfer Masuk\",{rng_b},\"*rekening sendiri*\")")


def sumif_tip_minus_one_sheet(sheet, last_row):
    rng_c = f"'{sheet}'!$M$2:$M${last_row}"
    rng_j = f"'{sheet}'!$J$2:$J${last_row}"
    return f"=SUMIF({rng_c},\"Tip/Minus/Lebih\",{rng_j})"


def sumif_gaji_lainnya_formula(sheet, last_row, nama_bulan_list):
    """Jaring pengaman: tangkap semua baris berkategori 'Gaji*' TAPI
    keterangannya tidak menyebut bulan ini/lalu (mis. gaji utuh dari bulan
    lain, atau baris gaji tanpa nama bulan sama sekali) - supaya tidak ada
    beban gaji yang diam-diam hilang dari Laba Rugi."""
    rng_c = f"'{sheet}'!$M$2:$M${last_row}"
    rng_b = f"'{sheet}'!$B$2:$B${last_row}"
    rng_j = f"'{sheet}'!$J$2:$J${last_row}"
    parts = [f"SUMIF({rng_c},\"Gaji*\",{rng_j})"]
    for nama_bulan in nama_bulan_list:
        if nama_bulan:
            parts.append(f"SUMIFS({rng_j},{rng_c},\"Gaji*\",{rng_b},\"*{nama_bulan}*\")")
    return "=" + parts[0] + "".join(f"-{p}" for p in parts[1:])


def write_income_statement(wb, sheets_last_row, period_label, period_month, recon_range):
    name = "Laporan Laba Rugi"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    sheets = list(sheets_last_row.keys())
    ws["A1"] = f"LAPORAN LABA RUGI - {period_label.upper()}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Pivot per rekening (rumus SUMIF beralamat absolut), kolom TOTAL paling kanan "
                "= jumlah keseluruhan. Bandingkan antar kolom untuk menelusuri selisih per rekening.")
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    r = 4
    write_pivot_header(ws, r, sheets)
    r += 1

    write_pivot_section(ws, r, "PENDAPATAN", sheets)
    r += 1
    rev_rows = []
    for cat in INCOME_CATEGORIES_REVENUE:
        write_pivot_data_row(ws, r, cat, sheets,
                              lambda sheet, cat=cat: sumif_one_sheet(sheet, sheets_last_row[sheet], cat))
        rev_rows.append(r)
        r += 1
    total_rev_row = r
    write_pivot_subtotal_row(ws, r, "Total Pendapatan", sheets, rev_rows)
    r += 2

    write_pivot_section(ws, r, "BEBAN", sheets)
    r += 1
    exp_rows = []
    exp_row_by_cat = {}
    for cat in INCOME_CATEGORIES_EXPENSE:
        write_pivot_data_row(ws, r, cat, sheets,
                              lambda sheet, cat=cat: sumif_one_sheet(sheet, sheets_last_row[sheet], cat))
        exp_rows.append(r)
        exp_row_by_cat[cat] = r
        r += 1
    write_pivot_data_row(
        ws, r, "Marketing & RnD", sheets,
        lambda sheet: sumif_multi_one_sheet(sheet, sheets_last_row[sheet], MARKETING_RND_CATEGORY_TEXTS),
    )
    exp_rows.append(r)
    marketing_rnd_row = r
    r += 1
    # Gaji: dicocokkan lewat KETERANGAN (bukan Kategori) supaya konsisten
    # baik format lama (bulan ada di Kategori) maupun format bank/preformatted
    # (Kategori tetap "Gaji Pegawai", bulan cuma disebut di Keterangan, mis.
    # "Gaji Latifatul Husna Januari"). "Gaji <bulan ini>" = beban berjalan,
    # "Gaji <bulan lalu>" = accrual, "Gaji Lainnya" = jaring pengaman.
    nama_ini, nama_lalu = gaji_category_patterns(period_month)
    write_pivot_data_row(ws, r, f"Gaji Bulan Ini (Gaji {nama_ini or 'Bulan Ini'})", sheets,
                          lambda sheet: sumif_gaji_bulan_formula(sheet, sheets_last_row[sheet], nama_ini))
    exp_rows.append(r)
    gaji_ini_row = r
    r += 1
    write_pivot_data_row(ws, r, f"Gaji Accrual (Gaji {nama_lalu or 'Accrual'})", sheets,
                          lambda sheet: sumif_gaji_bulan_formula(sheet, sheets_last_row[sheet], nama_lalu))
    exp_rows.append(r)
    gaji_accrual_row = r
    r += 1
    write_pivot_data_row(
        ws, r, "Gaji Lainnya (bulan lain/tidak disebutkan)", sheets,
        lambda sheet: sumif_gaji_lainnya_formula(sheet, sheets_last_row[sheet], [nama_ini, nama_lalu]),
    )
    exp_rows.append(r)
    gaji_lainnya_row = r
    r += 1
    # Biaya Admin Bank: gabungan biaya admin bank, biaya admin
    # transfer (mis. via Fliptech, auto-terdeteksi dari Rekonsiliasi kolom
    # N/O), bunga, dan pajak bank - dulu terpecah jadi beberapa baris karena
    # tiap bank/versi parser pakai istilah beda, sekarang satu baris saja
    fee_row = r
    write_pivot_data_row(
        ws, r, "Biaya Admin Bank (termasuk biaya transfer Fliptech)", sheets,
        lambda sheet: (
            f"={sumif_multi_one_sheet(sheet, sheets_last_row[sheet], BANK_FEE_CATEGORY_TEXTS)[1:]}"
            f"-SUMIFS('Rekonsiliasi'!$O${recon_range['data_start']}:$O${recon_range['data_end']},"
            f"'Rekonsiliasi'!$N${recon_range['data_start']}:$N${recon_range['data_end']},\"{sheet}\")"
        ),
    )
    exp_rows.append(r)
    r += 1
    total_exp_row = r
    write_pivot_subtotal_row(ws, r, "Total Beban", sheets, exp_rows)
    r += 2

    # RINGKASAN LAYER 2 (COGS/OpEx/CapEx) - roll-up dari kategori Layer 1
    # di atas, sesuai "Kontrak Kategori: Bot Konversi -> Bot Rekonsiliasi"
    # yang disepakati dengan tim bot konversi:
    #   COGS  = Belanja Bahan (Belanja Konsumsi sudah digabung ke
    #           Konsumsi dan Liburan yang sekarang OpEx, bukan COGS lagi)
    #   OpEx  = Overhead (sudah menyerap Belanja Operasional) +
    #           Konsumsi dan Liburan + Belanja Utilitas + Tools dan
    #           Equipments + Kemasan + Subscription + Sewa dan
    #           Maintenance Bangunan (sudah menyerap Biaya Renovasi Atap)
    #           + Reparasi dan Maintenance Tools dan Mesin + Pajak dan
    #           Administrasi + Marketing & RnD + Gaji* + Biaya Admin Bank
    #   CapEx = Belanja Assets
    # Baris Layer 1 di atas TETAP ditampilkan lengkap untuk audit detail -
    # ini cuma ringkasan tambahan, bukan pengganti.
    write_pivot_section(ws, r, "RINGKASAN LAYER 2 (roll-up COGS/OpEx/CapEx)", sheets)
    r += 1
    cogs_ref_rows = [exp_row_by_cat["Belanja Bahan"]]
    write_pivot_formula_row(
        ws, r, "COGS (Belanja Bahan)", sheets,
        lambda cl: "=" + "+".join(f"{cl}{rr}" for rr in cogs_ref_rows),
        bold=True,
    )
    r += 1
    opex_ref_rows = [
        exp_row_by_cat["Overhead"],
        exp_row_by_cat["Konsumsi dan Liburan"], exp_row_by_cat["Belanja Utilitas"],
        exp_row_by_cat["Tools dan Equipments"], exp_row_by_cat["Kemasan"],
        exp_row_by_cat["Subscription"], exp_row_by_cat["Sewa dan Maintenance Bangunan"],
        exp_row_by_cat["Reparasi dan Maintenance Tools dan Mesin"], exp_row_by_cat["Pajak dan Administrasi"],
        marketing_rnd_row,
        gaji_ini_row, gaji_accrual_row, gaji_lainnya_row, fee_row,
    ]
    write_pivot_formula_row(
        ws, r, "OpEx (Overhead+Konsumsi&Liburan+Utilitas+Tools&Equip+Kemasan+Subscription+Sewa&Maintenance Bangunan+Reparasi&Maintenance Tools&Mesin+Pajak&Administrasi+Marketing&RnD+Gaji*+Biaya Admin Bank)", sheets,
        lambda cl: "=" + "+".join(f"{cl}{rr}" for rr in opex_ref_rows),
        bold=True,
    )
    r += 1
    write_pivot_formula_row(
        ws, r, "CapEx (Belanja Assets)", sheets,
        lambda cl: f"={cl}{exp_row_by_cat['Belanja Assets']}",
        bold=True,
    )
    r += 2

    write_pivot_section(ws, r, "LAIN-LAIN (perlu verifikasi manual)", sheets)
    r += 1
    other_rows = []
    for cat in OTHER_CATEGORIES:
        if cat == "Tip/Minus/Lebih":
            write_pivot_data_row(ws, r, cat, sheets,
                                  lambda sheet: sumif_tip_minus_one_sheet(sheet, sheets_last_row[sheet]))
        else:
            write_pivot_data_row(ws, r, cat, sheets,
                                  lambda sheet, cat=cat: sumif_one_sheet(sheet, sheets_last_row[sheet], cat))
        other_rows.append(r)
        r += 1
    total_other_row = r
    write_pivot_subtotal_row(ws, r, "Total Lain-lain", sheets, other_rows)
    r += 2

    net_row = r
    write_pivot_formula_row(
        ws, r, "LABA / RUGI BERSIH", sheets,
        lambda cl: f"={cl}{total_rev_row}+{cl}{total_exp_row}+{cl}{total_other_row}",
        bold=True,
    )
    for c in range(1, pivot_total_col(sheets) + 1):
        ws.cell(row=r, column=c).font = Font(bold=True, size=12)

    ws.column_dimensions["A"].width = 34
    for i in range(len(sheets)):
        ws.column_dimensions[col_letter(2 + i)].width = 16
    ws.column_dimensions[col_letter(pivot_total_col(sheets))].width = 18
    ws.freeze_panes = "B5"
    return ws, {"total_rev": total_rev_row, "total_exp": total_exp_row,
                "total_other": total_other_row, "net": net_row, "sheet": name,
                "sheets": sheets, "total_col": pivot_total_col(sheets)}


# ---------------------------------------------------------------------------
# Neraca (Balance Sheet)
# ---------------------------------------------------------------------------

TRANSFER_CATEGORY_TEXTS = [
    "Pindah Rekening Internal",
    "Pindang Rekening Internal",
    "Transfer Internal",
    "Transfer Lainnya",
    "Transaksi Internal",
]


def _validate_category_override_targets():
    """Cegah kelas bug yang pernah kejadian: kategori tujuan di
    CATEGORY_OVERRIDE_RULES harus PERSIS salah satu string yang benar-benar
    dicek rumus SUMIF/SUMIFS laporan keuangan (INCOME_CATEGORIES_EXPENSE/
    REVENUE, MARKETING_RND_CATEGORY_TEXTS, BANK_FEE_CATEGORY_TEXTS,
    OTHER_CATEGORIES, atau 'Modal & Setoran Pemilik') - BUKAN cuma label
    baris yang enak dibaca (mis. 'Marketing & RnD' pernah kepakai padahal
    yang benar-benar dicek SUMIF adalah 'Marketing' saja, uangnya jadi
    tidak ketangkap di manapun dan bikin Neraca selisih diam-diam)."""
    known = set(
        INCOME_CATEGORIES_EXPENSE + INCOME_CATEGORIES_REVENUE +
        MARKETING_RND_CATEGORY_TEXTS + BANK_FEE_CATEGORY_TEXTS +
        OTHER_CATEGORIES + TRANSFER_CATEGORY_TEXTS +
        ["Modal & Setoran Pemilik", "Pengeluaran Pribadi", "Hutang Masuk", "Pembayaran Hutang"]
    )
    # 'Kategori Baru' SENGAJA dikecualikan dari 'known' - ini bukan bug,
    # ini SATU-SATUNYA target yang memang sengaja TIDAK dihitung SUMIF
    # manapun (dikeluarkan total dari Laba Rugi/Neraca sesuai penegasan
    # user, supaya Neraca genuinely tidak balance selama masih ada
    # transaksi yang belum diaudit - lihat compute_balance_status).
    allowed_special = known | {"Kategori Baru"}
    for rule in CATEGORY_OVERRIDE_RULES:
        target = rule["category"]
        if target not in allowed_special:
            raise AssertionError(
                f"CATEGORY_OVERRIDE_RULES: kategori tujuan {target!r} bukan salah satu kategori "
                "yang dikenali rumus SUMIF/SUMIFS laporan keuangan - uang yang di-override ke sini "
                "tidak akan ketangkap di manapun. Perbaiki jadi salah satu dari: "
                f"{sorted(known)}"
            )


_validate_category_override_targets()


def sumif_multi_one_sheet(sheet, last_row, categories):
    parts = [
        f"SUMIF('{sheet}'!$M$2:$M${last_row},\"{cat}\",'{sheet}'!$J$2:$J${last_row})"
        for cat in categories
    ]
    return "=" + "+".join(parts)


def write_balance_sheet(wb, sheets_last_row, opening_rows, income_ref, period_end_label, recon_range):
    name = "Neraca"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    sheets = list(sheets_last_row.keys())
    ws["A1"] = f"NERACA - PER {period_end_label.upper()}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Pivot per rekening, kolom TOTAL paling kanan = keseluruhan. Baris 'CEK KESEIMBANGAN' "
                "dan 'Selisih Belum Terjelaskan' per kolom langsung menunjukkan rekening mana yang "
                "selisih - lihat juga sheet Diagnostik Keseimbangan.")
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    r = 4
    write_pivot_header(ws, r, sheets)
    r += 1

    write_pivot_section(ws, r, "ASET (KAS & SETARA KAS)", sheets)
    r += 1
    kas_row = r
    write_pivot_data_row(ws, r, "Kas & Setara Kas (Saldo Akhir)", sheets,
                          lambda sheet: f"='{sheet}'!$K${sheets_last_row[sheet]}", bold=True)
    total_asset_row = kas_row
    r += 2

    write_pivot_section(ws, r, "EKUITAS", sheets)
    r += 1
    saldo_awal_row = r
    write_pivot_data_row(ws, r, "Saldo Awal Bulan", sheets,
                          lambda sheet: f"='{sheet}'!$K${opening_rows[sheet]}")
    r += 1
    modal_row = r
    write_pivot_data_row(ws, r, "Modal & Setoran Pemilik (+ Laba Ditahan Bulanan) (bulan ini)", sheets,
                          lambda sheet: sumif_modal_one_sheet(sheet, sheets_last_row[sheet]))
    r += 1
    pengeluaran_pribadi_row = r
    write_pivot_data_row(ws, r, "Pengeluaran Pribadi (mengurangi ekuitas, bukan beban bisnis)", sheets,
                          lambda sheet: sumif_one_sheet(sheet, sheets_last_row[sheet], "Pengeluaran Pribadi"))
    r += 1
    laba_row = r
    # kolom rekening di Neraca urutannya sama dengan di Laporan Laba Rugi
    # (keduanya dari sheets_last_row.keys() yang sama), jadi tinggal pakai
    # huruf kolom yang sama untuk menautkan baris LABA/RUGI BERSIH per rekening
    write_pivot_formula_row(
        ws, r, "Laba Bersih Bulan Ini", sheets,
        lambda cl: f"='{income_ref['sheet']}'!{cl}{income_ref['net']}",
    )
    r += 1
    total_equity_row = r
    write_pivot_subtotal_row(ws, r, "Total Ekuitas", sheets, [saldo_awal_row, laba_row])
    r += 2

    write_pivot_section(ws, r, "LIABILITAS (Hutang)", sheets)
    r += 1
    hutang_masuk_row = r
    write_pivot_data_row(ws, r, "Hutang Masuk (bulan ini)", sheets,
                          lambda sheet: sumif_one_sheet(sheet, sheets_last_row[sheet], "Hutang Masuk"))
    r += 1
    pembayaran_hutang_row = r
    write_pivot_data_row(ws, r, "Pembayaran Hutang (bulan ini, mengurangi liabilitas)", sheets,
                          lambda sheet: sumif_one_sheet(sheet, sheets_last_row[sheet], "Pembayaran Hutang"))
    r += 1
    total_liability_row = r
    write_pivot_subtotal_row(ws, r, "Total Liabilitas (Hutang)", sheets, [hutang_masuk_row, pembayaran_hutang_row])
    r += 2

    balance_check_row = r
    write_pivot_formula_row(
        ws, r, "CEK KESEIMBANGAN (Aset - Ekuitas - Liabilitas)", sheets,
        lambda cl: f"={cl}{total_asset_row}-{cl}{total_equity_row}-{cl}{total_liability_row}",
        bold=True,
    )
    r += 1
    transfer_row = r
    # Transfer Bersih murni (SUMIF kategori transfer) DITAMBAH biaya admin
    # yang sudah "dipisahkan" jadi beban riil di Laba Rugi (baris Biaya Admin
    # Transfer) - supaya baris ini hanya berisi porsi transfer yang benar
    # dua sisinya matched, bukan lagi bercampur dengan biaya admin
    write_pivot_data_row(
        ws, r, "Transfer Bersih (rekening ini, info)", sheets,
        lambda sheet: (
            f"={sumif_multi_one_sheet(sheet, sheets_last_row[sheet], TRANSFER_CATEGORY_TEXTS)[1:]}"
            f"+SUMIFS('Rekonsiliasi'!$O${recon_range['data_start']}:$O${recon_range['data_end']},"
            f"'Rekonsiliasi'!$N${recon_range['data_start']}:$N${recon_range['data_end']},\"{sheet}\")"
        ),
    )
    r += 1
    residual_row = r
    write_pivot_formula_row(
        ws, r, "Selisih Belum Terjelaskan (Selisih - Transfer Bersih)", sheets,
        lambda cl: f"={cl}{balance_check_row}-{cl}{transfer_row}",
        bold=True,
    )
    r += 1
    ws.cell(row=r, column=1,
            value=("Transfer Bersih seharusnya saling menutup ~0 di kolom TOTAL (lihat sheet "
                   "Rekonsiliasi kalau tidak). Per rekening wajar tidak 0 (rekening itu bisa jadi "
                   "pengirim/penerima bersih bulan ini). Yang perlu ditelusuri adalah 'Selisih Belum "
                   "Terjelaskan' - kalau besar di satu rekening, itu tandanya data di sheet rekening "
                   "itu (lihat kolom L) yang bermasalah, bukan soal transfer."))
    ws.cell(row=r, column=1).font = Font(italic=True, size=9, color="6B7280")
    ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=pivot_total_col(sheets))
    ws.row_dimensions[r].height = 48

    ws.column_dimensions["A"].width = 38
    for i in range(len(sheets)):
        ws.column_dimensions[col_letter(2 + i)].width = 16
    ws.column_dimensions[col_letter(pivot_total_col(sheets))].width = 18
    ws.freeze_panes = "B5"
    return ws, {"total_asset": total_asset_row, "total_equity": total_equity_row,
                "total_liability": total_liability_row,
                "saldo_awal": saldo_awal_row, "balance_check": balance_check_row,
                "transfer_row": transfer_row, "residual_row": residual_row,
                "sheet": name, "sheets": sheets, "total_col": pivot_total_col(sheets)}


# ---------------------------------------------------------------------------
# Laporan Arus Kas (Cash Flow Statement)
# ---------------------------------------------------------------------------

def write_cash_flow(wb, sheets_last_row, income_ref, balance_ref, period_label):
    name = "Laporan Arus Kas"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    sheets = list(sheets_last_row.keys())
    ws["A1"] = f"LAPORAN ARUS KAS - {period_label.upper()}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Metode langsung, pivot per rekening. Transfer antar rekening sendiri sengaja tidak "
                "dimasukkan karena saling menutup nol (lihat sheet Rekonsiliasi).")
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    r = 4
    write_pivot_header(ws, r, sheets)
    r += 1

    write_pivot_section(ws, r, "ARUS KAS DARI AKTIVITAS OPERASI", sheets)
    r += 1
    op_row = r
    write_pivot_formula_row(
        ws, r, "Laba Bersih Bulan Ini (basis kas)", sheets,
        lambda cl: f"='{income_ref['sheet']}'!{cl}{income_ref['net']}",
    )
    r += 1
    total_op_row = r
    write_pivot_subtotal_row(ws, r, "Kas Bersih dari Operasi", sheets, [op_row, op_row])
    r += 2

    write_pivot_section(ws, r, "ARUS KAS DARI AKTIVITAS PENDANAAN", sheets)
    r += 1
    fin_row = r
    write_pivot_data_row(ws, r, "Modal & Setoran Pemilik (+ Laba Ditahan Bulanan)", sheets,
                          lambda sheet: sumif_modal_one_sheet(sheet, sheets_last_row[sheet]))
    r += 1
    write_pivot_data_row(ws, r, "Pengeluaran Pribadi", sheets,
                          lambda sheet: sumif_one_sheet(sheet, sheets_last_row[sheet], "Pengeluaran Pribadi"))
    r += 1
    total_fin_row = r
    write_pivot_subtotal_row(ws, r, "Kas Bersih dari Pendanaan", sheets, [fin_row, total_fin_row - 1])
    r += 2

    net_change_row = r
    write_pivot_formula_row(
        ws, r, "KENAIKAN (PENURUNAN) KAS BERSIH", sheets,
        lambda cl: f"={cl}{total_op_row}+{cl}{total_fin_row}",
        bold=True,
    )
    r += 1
    saldo_awal_row = r
    write_pivot_formula_row(
        ws, r, "Saldo Kas Awal Bulan", sheets,
        lambda cl: f"='{balance_ref['sheet']}'!{cl}{balance_ref['saldo_awal']}",
    )
    r += 1
    saldo_akhir_row = r
    write_pivot_formula_row(
        ws, r, "Saldo Kas Akhir Bulan", sheets,
        lambda cl: f"={cl}{net_change_row}+{cl}{saldo_awal_row}",
        bold=True,
    )
    r += 1
    write_pivot_formula_row(
        ws, r, "Cek vs Total Aset di Neraca", sheets,
        lambda cl: f"={cl}{saldo_akhir_row}-'{balance_ref['sheet']}'!{cl}{balance_ref['total_asset']}",
    )

    ws.column_dimensions["A"].width = 38
    for i in range(len(sheets)):
        ws.column_dimensions[col_letter(2 + i)].width = 16
    ws.column_dimensions[col_letter(pivot_total_col(sheets))].width = 18
    ws.freeze_panes = "B5"
    return ws


# ---------------------------------------------------------------------------
# Diagnostik Keseimbangan - alat telusur kalau Neraca/Arus Kas selisih
# ---------------------------------------------------------------------------

def write_diagnostic_sheet(wb, sheets_last_row, balance_ref, closing_info_by_sheet):
    """Sheet khusus buat menelusuri KENAPA CEK KESEIMBANGAN di Neraca tidak
    nol. Beberapa kemungkinan penyebab yang paling sering terjadi:
    1. Transfer antar rekening yang belum matched (lihat sheet Rekonsiliasi
       bagian 1) - nilainya tidak ikut dihitung di Laba Rugi/Ekuitas, tapi
       tetap mempengaruhi saldo kas riil.
    2. Baris-baris di sheet rekening TIDAK berurutan kronologis terhadap
       kolom Saldo Kumulatif aslinya (kolom F), sehingga saldo hasil
       rekonstruksi (kolom K) menyimpang dari saldo tercatat. Bagian ini
       menunjukkan tepat di rekening mana dan seberapa besar penyimpangan
       itu terjadi, lewat kolom L (Selisih vs Saldo Tercatat) di tiap sheet
       rekening.
    3. Kalau sheet rekening punya blok rekap penutup (Saldo Akhir/Total
       Debit/Total Kredit di baris-baris akhir), nilai itu SUDAH dikeluarkan
       dari rekonstruksi kolom J/K/L (supaya tidak dobel/merusak saldo
       berjalan) dan dipakai di sini sebagai acuan pembanding independen."""
    name = "Diagnostik Keseimbangan"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws["A1"] = "DIAGNOSTIK KESEIMBANGAN - ALAT TELUSUR SELISIH"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (
        "Dipakai kalau baris 'CEK KESEIMBANGAN' di Neraca tidak nol. "
        "Cek bagian di bawah: transfer yang belum matched (sheet "
        "Rekonsiliasi bagian 1), penyimpangan urutan data per rekening, "
        "dan cross-check terhadap blok rekap penutup sheet (kalau ada)."
    )
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    r = 4
    ws.cell(row=r, column=1, value="1. RINGKASAN KESEIMBANGAN")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    total_col_letter = col_letter(balance_ref["total_col"])
    ws.cell(row=r, column=1, value="Total Aset (Neraca, kolom TOTAL)")
    ws.cell(row=r, column=2, value=f"='{balance_ref['sheet']}'!{total_col_letter}{balance_ref['total_asset']}")
    r += 1
    ws.cell(row=r, column=1, value="Total Ekuitas (Neraca, kolom TOTAL)")
    ws.cell(row=r, column=2, value=f"='{balance_ref['sheet']}'!{total_col_letter}{balance_ref['total_equity']}")
    r += 1
    selisih_row = r
    ws.cell(row=r, column=1, value="Selisih (Aset - Ekuitas)")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"='{balance_ref['sheet']}'!{total_col_letter}{balance_ref['balance_check']}")
    ws.cell(row=r, column=2).font = Font(bold=True)
    r += 1
    ws.cell(row=r, column=1, value="Kemungkinan sumber #1: transfer belum matched (lihat sheet Rekonsiliasi bagian 1)")
    ws.cell(row=r, column=2, value="=COUNTIF(Rekonsiliasi!$K$6:$K$300,\"Needs manual verification\")")
    ws.cell(row=r, column=3, value="baris - buka sheet Rekonsiliasi, cari warna merah")
    r += 2

    ws.cell(row=r, column=1,
            value="1b. SELISIH BELUM TERJELASKAN PER REKENING (dari sheet Neraca)")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    headers1b = ["Rekening", "Selisih Belum Terjelaskan (Rp)"]
    hdr_row1b = r
    for i, h in enumerate(headers1b, start=1):
        ws.cell(row=hdr_row1b, column=i, value=h)
    style_header(ws, hdr_row1b, len(headers1b))
    r += 1
    sheets_list = list(sheets_last_row.keys())
    for i, sheet in enumerate(sheets_list):
        cl = col_letter(2 + i)
        ws.cell(row=r, column=1, value=sheet)
        ws.cell(row=r, column=2, value=f"='{balance_ref['sheet']}'!{cl}{balance_ref['residual_row']}")
        ws.cell(row=r, column=2).number_format = NUMBER_FORMAT
        r += 1
    ws.cell(row=r, column=1, value="TOTAL")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"='{balance_ref['sheet']}'!{total_col_letter}{balance_ref['residual_row']}")
    ws.cell(row=r, column=2).font = Font(bold=True)
    ws.cell(row=r, column=2).number_format = NUMBER_FORMAT
    r += 2

    ws.cell(row=r, column=1,
            value="1c. CROSS-CHECK SALDO AKHIR RESMI (dari blok rekap penutup sheet, kalau ada)")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    headers1c = ["Rekening", "Saldo Akhir Resmi (dari blok penutup)", "Saldo Akhir Rekonstruksi (Kolom K)", "Selisih"]
    hdr_row1c = r
    for i, h in enumerate(headers1c, start=1):
        ws.cell(row=hdr_row1c, column=i, value=h)
    style_header(ws, hdr_row1c, len(headers1c))
    r += 1
    any_closing_found = False
    for sheet, last_row in sheets_last_row.items():
        info = closing_info_by_sheet.get(sheet, {})
        saldo_akhir_resmi = info.get("saldo_akhir")
        ws.cell(row=r, column=1, value=sheet)
        if saldo_akhir_resmi is not None:
            any_closing_found = True
            ws.cell(row=r, column=2, value=saldo_akhir_resmi)
            ws.cell(row=r, column=2).number_format = NUMBER_FORMAT
            ws.cell(row=r, column=3, value=f"='{sheet}'!$K${last_row}")
            ws.cell(row=r, column=3).number_format = NUMBER_FORMAT
            ws.cell(row=r, column=4, value=f"=B{r}-C{r}")
            ws.cell(row=r, column=4).number_format = NUMBER_FORMAT
        else:
            ws.cell(row=r, column=2, value="(tidak ada blok penutup di sheet ini)")
            ws.cell(row=r, column=2).font = Font(italic=True, color="6B7280")
        r += 1
    if not any_closing_found:
        ws.cell(row=r, column=1,
                value="Tidak ada sheet dengan blok rekap penutup (Saldo Akhir/Total Debit/Kredit) terdeteksi.")
        ws.cell(row=r, column=1).font = Font(italic=True, color="6B7280")
        r += 1
    r += 1

    ws.cell(row=r, column=1, value="2. PENYIMPANGAN URUTAN DATA PER REKENING (Kolom K vs Kolom F)")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    headers = [
        "Rekening", "Jumlah Baris Menyimpang (>Rp1rb)", "Selisih Maksimum (Rp)",
        "Selisih di Baris Terakhir (Rp)", "Baris Pertama Menyimpang", "Keterangan Baris Itu",
    ]
    hdr_row = r
    for i, h in enumerate(headers, start=1):
        ws.cell(row=hdr_row, column=i, value=h)
    style_header(ws, hdr_row, len(headers))
    r += 1
    for sheet, last_row in sheets_last_row.items():
        rng_l = f"'{sheet}'!$L$2:$L${last_row}"
        rng_a = f"'{sheet}'!$A$2:$A${last_row}"
        rng_c = f"'{sheet}'!$B$2:$B${last_row}"
        ws.cell(row=r, column=1, value=sheet)
        ws.cell(row=r, column=2, value=f"=COUNTIF({rng_l},\">1000\")+COUNTIF({rng_l},\"<-1000\")")
        ws.cell(row=r, column=3, value=f"=SUMPRODUCT(MAX(ABS({rng_l})))")
        ws.cell(row=r, column=3).number_format = NUMBER_FORMAT
        ws.cell(row=r, column=4, value=f"='{sheet}'!$L${last_row}")
        ws.cell(row=r, column=4).number_format = NUMBER_FORMAT
        ws.cell(row=r, column=5,
                value=(f'=IFERROR(INDEX({rng_a},MATCH(TRUE,INDEX(ABS({rng_l})>1000,0),0)),'
                       f'"Tidak ada penyimpangan signifikan")'))
        ws.cell(row=r, column=5).number_format = DATE_FORMAT
        ws.cell(row=r, column=6,
                value=(f'=IFERROR(INDEX({rng_c},MATCH(TRUE,INDEX(ABS({rng_l})>1000,0),0)),"-")'))
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(c == 6))
        r += 1

    r += 1
    ws.cell(row=r, column=1,
            value=("Catatan: penyimpangan berarti baris-baris SEBELUM titik itu di sheet rekening "
                   "tidak tersusun berurutan sesuai kolom Saldo Kumulatif aslinya (kemungkinan input "
                   "manual tidak kronologis, atau digabung per sesi/hari alih-alih per transaksi). "
                   "Saldo akhir bulan (Total Aset di Neraca) tetap dihitung dari hasil rekonstruksi "
                   "(kolom K), bukan dari kolom F yang bolong urutannya - tapi 'Saldo Awal' yang "
                   "dipakai di Neraca mengasumsikan baris pertama tiap sheet adalah titik awal yang "
                   "valid. Kalau baris pertama BUKAN baris Saldo Awal (lihat sheet Rekonsiliasi atau "
                   "cek manual), kemungkinan itu sumber selisihnya.")
            )
    ws.cell(row=r, column=1).font = Font(italic=True, size=9, color="6B7280")
    ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 60

    widths = [30, 24, 20, 22, 20, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


MONTHS_ID = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]


def coerce_date(value):
    """Ubah nilai tanggal dari sumber manapun (datetime.datetime,
    datetime.date, atau teks format umum) jadi datetime.date. Dipakai
    supaya file lama/preformatted yang kolom tanggalnya bukan objek
    datetime asli (mis. tersimpan sebagai teks) tetap bisa dideteksi
    periodenya. Return None kalau tidak bisa diparse."""
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, str):
        s = value.strip()
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y",
                    "%d-%b-%y", "%d-%b-%Y", "%d %B %Y", "%m/%d/%Y"):
            try:
                return datetime.datetime.strptime(s, fmt).date()
            except ValueError:
                continue
    return None


def detect_period(all_txns):
    """Tebak (tahun, bulan) laporan dari tanggal transaksi yang paling
    sering muncul. Return (None, None) kalau tidak ada tanggal valid."""
    from collections import Counter

    counts = Counter()
    for t in all_txns:
        d = coerce_date(t.date)
        if d is not None:
            counts[(d.year, d.month)] += 1
    if not counts:
        return None, None
    (year, month), _ = counts.most_common(1)[0]
    return year, month


def detect_period_label(all_txns):
    """Label teks 'Bulan Tahun' dari (tahun, bulan) yang terdeteksi, dipakai
    di judul-judul laporan."""
    year, month = detect_period(all_txns)
    if year is None:
        return "PERIODE TIDAK TERDETEKSI"
    return f"{MONTHS_ID[month]} {year}"


def detect_period_end_date(all_txns, year_month_label):
    """Tanggal terakhir transaksi pada bulan yang terdeteksi, dipakai untuk
    judul Neraca ('per tanggal X')."""
    dated = [coerce_date(t.date) for t in all_txns]
    dated = [d for d in dated if d is not None]
    if not dated:
        return ""
    last = max(dated)
    return f"{last.day} {MONTHS_ID[last.month]} {last.year}"


# ---------------------------------------------------------------------------
# Orkestrasi utama
# ---------------------------------------------------------------------------

def reload_shared_rules():
    """Baca ulang shared_rules (Postgres/JSON) dan timpa variabel modul
    yang relevan - dipanggil di awal run_reconciliation() supaya bot yang
    sudah lama jalan (proses long-running di Railway) tetap pakai aturan
    TERBARU dari Postgres tiap file baru diproses, bukan cuma versi yang
    kebetulan aktif saat bot pertama kali start."""
    global CATEGORY_OVERRIDE_RULES, TRANSFER_KEYWORDS, CAPITAL_KEYWORDS, _STATIC_CATEGORY_OVERRIDE_RULES
    global DESC_TRANSFER_KEYWORDS, CAPITAL_SELF_TRANSFER_KEYWORDS
    global _PROTECTED_FROM_CATEGORY_OVERRIDE, TIP_MINUS_THRESHOLD, FLIPTECH_FEE_THRESHOLD
    shared_rules._cache = None
    CATEGORY_OVERRIDE_RULES = shared_rules.get("category_override_rules", _DEFAULT_CATEGORY_OVERRIDE_RULES)
    _STATIC_CATEGORY_OVERRIDE_RULES = CATEGORY_OVERRIDE_RULES
    CATEGORY_OVERRIDE_RULES = CATEGORY_OVERRIDE_RULES + _build_transfer_masuk_rules()
    TRANSFER_KEYWORDS = shared_rules.get("transfer_keywords", TRANSFER_KEYWORDS)
    CAPITAL_KEYWORDS = shared_rules.get("capital_keywords", CAPITAL_KEYWORDS)
    DESC_TRANSFER_KEYWORDS = shared_rules.get("desc_transfer_keywords", DESC_TRANSFER_KEYWORDS)
    CAPITAL_SELF_TRANSFER_KEYWORDS = shared_rules.get("capital_self_transfer_keywords", CAPITAL_SELF_TRANSFER_KEYWORDS)
    _PROTECTED_FROM_CATEGORY_OVERRIDE = set(shared_rules.get(
        "protected_from_category_override", sorted(_PROTECTED_FROM_CATEGORY_OVERRIDE)
    ))
    TIP_MINUS_THRESHOLD = shared_rules.get("tip_minus_threshold", TIP_MINUS_THRESHOLD)
    FLIPTECH_FEE_THRESHOLD = shared_rules.get("fliptech_fee_threshold", FLIPTECH_FEE_THRESHOLD)
    _validate_category_override_targets()


def run_reconciliation(input_path, output_path, with_statements=None):
    """Fungsi utama: rekonsiliasi antar rekening (pencocokan transfer,
    deteksi split/merge, indikasi minus/selisih). Ini yang jalan secara
    default setiap ada file masuk.

    with_statements: None (default) = OTOMATIS - laporan keuangan (Laba
    Rugi/Neraca/Arus Kas/Diagnostik) HANYA disertakan kalau rekonsiliasi
    ini benar-benar tidak ada isu (no_issues=True: semua transfer sudah
    matched, tidak ada indikasi minus, Neraca balanced, tidak ada
    Kategori Baru). Kalau masih ada isu, HANYA sheet rekening + Rekonsiliasi
    yang dibuat - user harus beresin isu itu dulu (upload ulang) sebelum
    laporan keuangan lengkap ikut dibuat, supaya tidak menganalisis angka
    yang masih berpotensi salah/belum lengkap.

    True/False eksplisit = PAKSA (abaikan status no_issues) - dipakai
    /laporan kalau user memang mau lihat laporan keuangan meski masih ada
    isu yang belum beres.
    """
    reload_shared_rules()
    wb = openpyxl.load_workbook(input_path)
    REPORT_SHEET_NAMES = {
        "Rekonsiliasi", "Laporan Laba Rugi", "Neraca", "Laporan Arus Kas",
        "Diagnostik Keseimbangan", "Roster Gaji Bulan Ini",
    }
    account_sheets = [s for s in wb.sheetnames if s not in REPORT_SHEET_NAMES]

    all_txns = []
    all_txns_by_sheet = {}
    sheets_last_row = {}
    opening_rows = {}
    closing_info_by_sheet = {}
    for sname in account_sheets:
        ws = wb[sname]
        split_fliptech_combined_rows(ws)
        lr = last_data_row(ws)
        sheets_last_row[sname] = lr
        add_helper_column(ws, lr)
        txns, closing_info = read_account_sheet(ws)
        add_effective_category_column(ws, txns)
        all_txns.extend(txns)
        all_txns_by_sheet[sname] = txns
        closing_info_by_sheet[sname] = closing_info
        opening = next((t for t in txns if t.is_opening), None)
        opening_rows[sname] = opening.row if opening else 2

    matches, combo_matches = find_matches(all_txns, account_sheets)
    correct_and_highlight_matched_transfers(wb, matches, combo_matches)

    # Highlight grup kategori (Penjualan/Belanja Bahan+Kemasan/dst) -
    # SAMA seperti di /rekonlokal - HANYA diterapkan pada baris yang
    # BELUM punya highlight dari transfer matching di atas (biru/orange/
    # kuning), supaya highlight arah-uang yang lebih spesifik itu tidak
    # tertimpa oleh pewarnaan kelompok yang sifatnya cuma visual.
    for sname in account_sheets:
        ws = wb[sname]
        for t in all_txns_by_sheet[sname]:
            if t.is_opening:
                continue
            cell_b = ws.cell(row=t.row, column=2)
            current_fill = cell_b.fill.fgColor.rgb if cell_b.fill else None
            if current_fill in _MEANINGFUL_HIGHLIGHT_HEXES:
                continue
            fill = category_group_fill(ws.cell(row=t.row, column=3).value)
            if fill is None:
                continue
            for c in range(1, 10):
                ws.cell(row=t.row, column=c).fill = fill

    minus_flags = find_minus_flags(all_txns_by_sheet)
    balance_status = compute_balance_status(all_txns_by_sheet)
    new_category_flags = find_new_category_flags(all_txns_by_sheet)
    personal_expense_flags = find_personal_expense_flags(all_txns_by_sheet)
    new_debt_flags = find_new_debt_flags(all_txns_by_sheet)

    # transaksi yang sudah terjelaskan lewat split/merge tidak perlu lagi
    # tampil sebagai "Needs manual verification" biasa di bagian 1
    combo_covered_ids = set()
    for cm in combo_matches:
        combo_covered_ids.add(id(cm["src"]))
        combo_covered_ids.add(id(cm["parts"][0]))
        combo_covered_ids.add(id(cm["parts"][1]))
    matches_section1 = [
        m for m in matches
        if m.dst is not None or id(m.src) not in combo_covered_ids
    ]

    ws_recon, recon_range = write_rekonsiliasi_sheet(wb, matches_section1, combo_matches, minus_flags, balance_status, new_category_flags, personal_expense_flags, new_debt_flags)

    order = list(account_sheets) + ["Rekonsiliasi"]

    period_year, period_month = detect_period(all_txns)
    period_label = detect_period_label(all_txns)
    period_end_label = detect_period_end_date(all_txns, period_label)

    # match dengan confidence "Not applicable" (teridentifikasi sebagai
    # cicilan pinjaman via Fliptech, BUKAN transfer internal yang genuinely
    # belum ketemu pasangannya) tidak dihitung sebagai unmatched - sudah
    # ada penjelasannya, tidak perlu verifikasi manual lagi
    n_transfer_unmatched = sum(
        1 for m in matches_section1 if m.dst is None and m.confidence != "Not applicable (bukan transfer internal)"
    )
    n_transfer_not_applicable = sum(
        1 for m in matches_section1 if m.confidence == "Not applicable (bukan transfer internal)"
    )
    n_balance_issues = sum(1 for s in balance_status.values() if abs(s["selisih"]) >= 1)
    no_issues = n_transfer_unmatched == 0 and len(minus_flags) == 0 and n_balance_issues == 0

    # with_statements=None (default) -> otomatis ikuti status no_issues:
    # laporan keuangan lengkap CUMA dibuat kalau rekonsiliasi ini benar-
    # benar bersih. True/False eksplisit (mis. dari /laporan) memaksa
    # abaikan status ini.
    actually_write_statements = no_issues if with_statements is None else with_statements

    if actually_write_statements:
        income_ws, income_ref = write_income_statement(wb, sheets_last_row, period_label, period_month, recon_range)
        balance_ws, balance_ref = write_balance_sheet(wb, sheets_last_row, opening_rows, income_ref, period_end_label, recon_range)
        write_cash_flow(wb, sheets_last_row, income_ref, balance_ref, period_label)
        write_diagnostic_sheet(wb, sheets_last_row, balance_ref, closing_info_by_sheet)
        order += [income_ref["sheet"], balance_ref["sheet"], "Laporan Arus Kas", "Diagnostik Keseimbangan"]

    # urutan sheet: rekening dulu, lalu laporan
    wb._sheets = [wb[s] for s in order]

    wb.save(output_path)

    summary = {
        "n_transfer_high": sum(1 for m in matches if m.confidence == "High"),
        "n_transfer_medium": sum(1 for m in matches if m.confidence == "Medium"),
        "n_transfer_low": sum(1 for m in matches if m.confidence == "Low"),
        "n_transfer_split_merge": len(combo_matches),
        "n_transfer_unmatched": n_transfer_unmatched,
        "n_transfer_not_applicable": n_transfer_not_applicable,
        "n_minus_flags": len(minus_flags),
        "n_balance_issues": n_balance_issues,
        "n_new_category": len(new_category_flags),
        "n_personal_expense": len(personal_expense_flags),
        "n_new_debt": len(new_debt_flags),
        "with_statements": actually_write_statements,
        "period_label": period_label,
        "no_issues": no_issues,
    }
    return summary


if __name__ == "__main__":
    import sys
    inp = sys.argv[1] if len(sys.argv) > 1 else "Recon_Januari_2025.xlsx"
    out = sys.argv[2] if len(sys.argv) > 2 else "Recon_Januari_2025_HASIL.xlsx"
    if "--recon-only" in sys.argv:
        with_stmt = False
    elif "--laporan" in sys.argv:
        with_stmt = True
    else:
        with_stmt = None  # otomatis: ikuti status no_issues
    s = run_reconciliation(inp, out, with_statements=with_stmt)
    print(s)
