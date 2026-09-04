"""Audit silang mesin kasir (POS) vs rekap keuangan bank/kas - FITUR
MANUAL, tidak jalan otomatis, harus di-trigger via /auditkasir (bot.py)
atau dipanggil langsung sebagai skrip.

KENAPA AGREGASI BULANAN, BUKAN PENCOCOKAN PER TRANSAKSI/HARIAN:
1. Data bank (Rekap) tidak menyimpan referensi transaksi POS (No
   Transaksi kasir) - cuma "Jam ...; Teller/User ID" dan kadang kode
   settlement (QRISOnUs/QRISOffUs/MID). Tidak ada kunci unik untuk
   mencocokkan satu transaksi kasir ke satu baris bank secara presisi.
2. QRIS/kartu settle ke rekening H+1 atau lebih lambat, dan TIDAK SELALU
   HARIAN (bisa beberapa hari terkumpul jadi satu setoran) - dikonfirmasi
   user. Pencocokan harian akan penuh false-positive selisih di
   perbatasan settlement yang sebenarnya normal.
3. Yang bisa diaudit secara andal: TOTAL BULANAN per kategori metode
   bayar (Cash/QRIS BRI/QRIS BCA/Kartu) - kalau selisihnya kecil (relatif
   terhadap volume), settlement timing yang menjelaskan. Kalau selisihnya
   besar, itu sinyal kuat ada salah kategori/metode/refund yang tidak
   tercatat rapi/uang hilang, perlu ditelusuri manual.

KETERBATASAN yang harus disadari user:
- Card (Kartu Debit/Kredit - BRI) dan sebagian QRIS BRI SAMA-SAMA settle
  ke rekening BRI-507 tanpa penanda konsisten yang membedakan keduanya di
  sisi bank (kadang ada teks 'QRISOnUs/QRISOffUs', kadang tidak) - jadi
  audit ini membandingkan GABUNGAN (QRIS BRI + Kartu BRI dari POS) vs
  TOTAL 'Penjualan' di BRI-507 (bank), bukan dipisah per metode.
- QRIS BCA vs 'Penjualan' di BCA-887 (bank) - dari data yang diperiksa,
  semua 'Penjualan' BCA-887 memang muncul terkait QRIS (pola 'MID:...QR:'
  atau kode settlement harian), jadi perbandingan 1:1 masuk akal.
- Beberapa hari di awal/akhir bulan WAJAR meleset ke bulan
  sebelum/sesudahnya karena jeda settlement - laporan ini menandai hari
  pertama/terakhir tiap bulan sebagai konteks, bukan tuduhan kesalahan.
"""
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime

import reconcile as rc

HEADER_FILL = "1F2937"
ALT_FILL = "F3F4F6"
GOOD_FILL = "DCFCE7"
WARN_FILL = "FEF3C7"
BAD_FILL = "FEE2E2"
BORDER = Border(*[Side(style="thin", color="D1D5DB")] * 4)

METODE_KATEGORI = {
    "cash": "Cash",
    "bank transfer - qris bri": "QRIS BRI",
    "bank transfer - qris bca": "QRIS BCA",
    "kartu debit/kredit - bri": "Kartu BRI",
}


class KasirAuditError(Exception):
    """Dilempar kalau file input tidak sesuai format yang diharapkan -
    pesannya dirancang untuk langsung ditunjukkan ke user di bot."""


def normalize_metode(metode_raw):
    """Petakan teks 'Metode Pembayaran' dari POS ke kategori standar.
    Metode yang tidak dikenali (mis. baru dari POS) dikembalikan apa
    adanya dengan prefix 'Lainnya:' supaya tetap kelihatan di laporan,
    bukan diam-diam diabaikan."""
    key = str(metode_raw or "").strip().lower()
    if key in ("", "-"):
        return "Belum Bayar"
    return METODE_KATEGORI.get(key, f"Lainnya: {metode_raw}")


def _parse_pos_datetime(s):
    s = str(s or "").strip()
    if not s:
        return None
    for fmt in ("%d-%m-%Y %H:%M:%S", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def parse_pos_sales(path):
    """Baca file 'Detail Penjualan' (export mesin kasir/POS). Header
    kolom ada di baris 13 (baris 1-12 metadata ringkasan). Return list of
    dict per transaksi. Melempar KasirAuditError kalau header tidak
    cocok format yang diharapkan (supaya gagal jelas, bukan salah baca
    diam-diam)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row = None
    for r in range(1, min(ws.max_row, 20) + 1):
        if str(ws.cell(row=r, column=1).value or "").strip() == "No Transaksi":
            header_row = r
            break
    if header_row is None:
        raise KasirAuditError(
            f"File '{path}' tidak dikenali sebagai export Detail Penjualan POS - "
            "tidak ketemu baris header 'No Transaksi' di 20 baris pertama."
        )
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        no_transaksi = ws.cell(row=r, column=1).value
        if not no_transaksi:
            continue
        waktu_bayar = ws.cell(row=r, column=3).value
        tgl = _parse_pos_datetime(waktu_bayar) or _parse_pos_datetime(ws.cell(row=r, column=2).value)
        rows.append({
            "no_transaksi": no_transaksi,
            "waktu_order": ws.cell(row=r, column=2).value,
            "waktu_bayar": waktu_bayar,
            "tanggal": tgl,
            "total": ws.cell(row=r, column=6).value or 0,
            "metode_raw": ws.cell(row=r, column=7).value,
            "metode": normalize_metode(ws.cell(row=r, column=7).value),
            "status": ws.cell(row=r, column=9).value,
            "kasir": ws.cell(row=r, column=11).value,
            "refund_tanggal": ws.cell(row=r, column=12).value,
            "refund_jumlah": ws.cell(row=r, column=13).value or 0,
            "refund_metode": ws.cell(row=r, column=14).value,
            "refund_alasan": ws.cell(row=r, column=15).value,
        })
    if not rows:
        raise KasirAuditError(f"File '{path}': tidak ada baris transaksi ditemukan setelah header.")
    return rows


def _bulan_label_dari_tanggal(tgl):
    if tgl is None:
        return None
    return f"{rc.MONTHS_ID[tgl.month]} {tgl.year}"


def aggregate_pos_by_month(all_pos_txns):
    """Kelompokkan transaksi POS per (bulan, kategori metode) - hitung
    total kotor, total refund, bersih (kotor - refund), jumlah transaksi.
    Transaksi 'Belum Bayar' (status Belum Lunas) DIKECUALIKAN dari total
    (uangnya belum benar-benar berpindah tangan)."""
    agg = {}
    for t in all_pos_txns:
        bulan = _bulan_label_dari_tanggal(t["tanggal"])
        if bulan is None or t["status"] == "Belum Lunas":
            continue
        key = (bulan, t["metode"])
        a = agg.setdefault(key, {"kotor": 0, "refund": 0, "n": 0})
        a["kotor"] += t["total"]
        a["refund"] += t["refund_jumlah"]
        a["n"] += 1
    return agg


def aggregate_pos_by_day(all_pos_txns):
    """Kelompokkan transaksi POS per (tanggal, kategori metode) - dipakai
    untuk rekonsiliasi harian dengan asumsi H+1 (QRIS/kartu baru masuk
    rekening hari berikutnya). Refund diperlakukan sama seperti agregasi
    bulanan: dikurangi dari total kotor tanggal transaksi ASLI (bukan
    tanggal refund) - refund adalah koreksi atas penjualan hari itu."""
    agg = {}
    for t in all_pos_txns:
        tgl = t["tanggal"]
        if tgl is None or t["status"] == "Belum Lunas":
            continue
        key = (tgl.date(), t["metode"])
        a = agg.setdefault(key, {"kotor": 0, "refund": 0, "n": 0})
        a["kotor"] += t["total"]
        a["refund"] += t["refund_jumlah"]
        a["n"] += 1
    return agg


_EXPECTED_HEADER = ["Tanggal", "Keterangan Transaksi", "Kategori Transaksi", "Debit", "Kredit",
                    "Saldo Kumulatif", "Subjek Transaksi", "Objek Transaksi"]


def _looks_like_account_sheet(ws):
    header = [ws.cell(row=1, column=c).value for c in range(1, len(_EXPECTED_HEADER) + 1)]
    return header == _EXPECTED_HEADER


def aggregate_bank_penjualan(rekap_paths):
    """Untuk tiap file Rekap (rekonsiliasi bulanan biasa), jumlahkan
    transaksi berkategori efektif 'Penjualan' per (bulan, nama dasar
    rekening). Pakai reconcile.py langsung supaya konsisten dengan
    logika kategorisasi yang sama dipakai rekonsiliasi normal."""
    agg = {}
    daily = {}  # (bulan, rekening, tanggal) -> total, buat cek settlement lintas bulan
    for path in rekap_paths:
        wb = openpyxl.load_workbook(path)
        account_sheets = [s for s in wb.sheetnames if _looks_like_account_sheet(wb[s])]
        if not account_sheets:
            raise KasirAuditError(
                f"File '{path}' tidak punya sheet rekening yang dikenali - pastikan ini file "
                "rekonsiliasi bulanan (hasil /start atau upload biasa ke bot), bukan laporan "
                "kuartalan/tahunan gabungan."
            )
        for sn in account_sheets:
            rc.split_fliptech_combined_rows(wb[sn])
            txns, _ = rc.read_account_sheet(wb[sn])
            base = sn.rsplit(" ", 2)[0]
            for t in txns:
                if t.effective_kategori != "Penjualan":
                    continue
                tgl = rc.coerce_date(t.date)
                bulan = _bulan_label_dari_tanggal(tgl) if tgl else None
                if bulan is None:
                    continue
                key = (bulan, base)
                agg[key] = agg.get(key, 0) + t.nominal
                dkey = (base, tgl)
                daily[dkey] = daily.get(dkey, 0) + t.nominal
    return agg, daily


# Rekening bank tujuan settlement per kategori metode POS non-tunai -
# lihat catatan keterbatasan di docstring modul ini.
_SETTLEMENT_TARGET = {
    "QRIS BRI": "BRI-507",
    "Kartu BRI": "BRI-507",
    "QRIS BCA": "BCA-887",
}


def _write_harian_sheet(wb, title, pos_daily, bank_daily, bank_rekening, all_dates, shift_hari,
                          catatan_header):
    """Satu sheet rekonsiliasi harian untuk SATU kategori metode -
    bandingkan total POS tanggal T dengan total Bank tanggal T+shift_hari
    (shift_hari=0 untuk Cash/tanpa penundaan, 1 untuk asumsi H+1 QRIS/
    kartu). Baris dengan selisih signifikan (>Rp5.000 ATAU >5% dari nilai
    POS, mana yang lebih longgar) ditandai merah - ini yang perlu
    ditelusuri (kemungkinan settlement tidak terjadi hari itu/terlambat
    lebih dari H+1, refund, atau salah metode)."""
    from datetime import timedelta
    ws = wb.create_sheet(title)
    ws["A1"] = title.upper()
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = catatan_header
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 40

    r = 4
    headers = ["Tanggal POS", "POS - Kotor", "POS - Refund", "POS - Bersih",
               f"Tanggal Bank (+{shift_hari} hari)", f"Bank ({bank_rekening})", "Selisih", "Status"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEADER_FILL)
        c.border = BORDER
    r += 1

    running_pos = 0
    running_bank = 0
    n_selisih = 0
    for tgl in all_dates:
        pos = pos_daily.get(tgl)
        tgl_bank = tgl + timedelta(days=shift_hari)
        bank_val = bank_daily.get((bank_rekening, tgl_bank))
        if pos is None and bank_val is None:
            continue
        kotor = pos["kotor"] if pos else 0
        refund = pos["refund"] if pos else 0
        bersih = kotor - refund
        bank_val = bank_val or 0
        selisih = bersih - bank_val
        running_pos += bersih
        running_bank += bank_val
        toleransi = max(5000, bersih * 0.05)
        if abs(selisih) <= toleransi:
            status, fill = "OK", GOOD_FILL
        else:
            status, fill = "SELISIH - cek manual", BAD_FILL
            n_selisih += 1
        ws.cell(row=r, column=1, value=tgl.strftime("%d/%m/%Y"))
        ws.cell(row=r, column=2, value=kotor)
        ws.cell(row=r, column=3, value=refund)
        ws.cell(row=r, column=4, value=bersih)
        ws.cell(row=r, column=5, value=tgl_bank.strftime("%d/%m/%Y"))
        ws.cell(row=r, column=6, value=bank_val)
        ws.cell(row=r, column=7, value=selisih)
        ws.cell(row=r, column=8, value=status)
        for c in (2, 3, 4, 6, 7):
            ws.cell(row=r, column=c).number_format = rc.NUMBER_FORMAT
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=8).fill = PatternFill("solid", fgColor=fill)
        ws.cell(row=r, column=8).font = Font(bold=True)
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="TOTAL PERIODE")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=4, value=running_pos)
    ws.cell(row=r, column=6, value=running_bank)
    ws.cell(row=r, column=7, value=running_pos - running_bank)
    for c in (4, 6, 7):
        ws.cell(row=r, column=c).number_format = rc.NUMBER_FORMAT
        ws.cell(row=r, column=c).font = Font(bold=True)
    r += 1
    selisih_total = running_pos - running_bank
    if abs(selisih_total) <= max(10000, running_pos * 0.02):
        ws.cell(row=r, column=1, value=(
            "Total periode NYAMBUNG (selisih kecil) - artinya hari-hari SELISIH di atas kemungkinan "
            "besar cuma jeda settlement yang lebih lambat dari H+1 (menumpuk lalu masuk sekaligus di "
            "hari lain), bukan uang yang benar-benar hilang."
        ))
        ws.cell(row=r, column=1).font = Font(italic=True, color="15803D")
    else:
        rp_selisih = f"{abs(selisih_total):,.0f}".replace(",", ".")
        ws.cell(row=r, column=1, value=(
            f"Total periode TIDAK NYAMBUNG (selisih Rp{rp_selisih}) - ini BUKAN cuma "
            "soal jeda settlement, ada uang yang genuinely tidak tercatat di salah satu sisi. Perlu "
            "ditelusuri manual."
        ))
        ws.cell(row=r, column=1).font = Font(italic=True, bold=True, color="B91C1C")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.row_dimensions[r].height = 30

    widths = [14, 15, 13, 15, 18, 16, 14, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    return n_selisih


def run_kasir_audit(pos_paths, rekap_paths, output_path):
    all_pos = []
    for p in pos_paths:
        all_pos.extend(parse_pos_sales(p))
    pos_agg = aggregate_pos_by_month(all_pos)
    pos_daily_raw = aggregate_pos_by_day(all_pos)
    bank_agg, bank_daily = aggregate_bank_penjualan(rekap_paths)

    bulan_list = sorted({b for b, _ in pos_agg} | {b for b, _ in bank_agg},
                         key=lambda lbl: (lbl.split()[1], rc.MONTHS_ID.index(lbl.split()[0])))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Ringkasan Audit Kasir")
    ws["A1"] = "AUDIT SILANG MESIN KASIR vs REKAP KEUANGAN"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = (
        "Perbandingan TOTAL BULANAN per metode bayar (bukan per transaksi/harian) - QRIS/kartu settle "
        "ke rekening beberapa hari kemudian dan tidak selalu harian, jadi pencocokan bulanan lebih andal "
        "daripada harian. Selisih kecil relatif terhadap volume bulan itu WAJAR (jeda settlement lintas "
        "bulan) - selisih besar perlu ditelusuri manual (kemungkinan salah kategori/metode, refund yang "
        "tidak konsisten, atau setoran yang belum masuk)."
    )
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 55

    r = 4
    headers = ["Bulan", "Metode", "POS - Kotor", "POS - Refund", "POS - Bersih",
               "Bank Tercatat (Penjualan)", "Selisih", "Selisih (%)", "Catatan"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEADER_FILL)
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center")
    r += 1

    metode_order = ["Cash", "QRIS BCA"]
    row_i = 0
    for bulan in bulan_list:
        for metode in metode_order:
            pos = pos_agg.get((bulan, metode))
            kotor = pos["kotor"] if pos else 0
            refund = pos["refund"] if pos else 0
            bersih = kotor - refund
            n = pos["n"] if pos else 0
            if metode == "Cash":
                bank_val = bank_agg.get((bulan, "Kas-Buku"))
                catatan_target = "Kas-Buku"
            else:
                target = _SETTLEMENT_TARGET.get(metode)
                bank_val = bank_agg.get((bulan, target)) if target else None
                catatan_target = target or "-"
            if pos is None and bank_val is None:
                continue  # metode ini tidak muncul sama sekali bulan ini di kedua sisi
            selisih = (bersih - bank_val) if bank_val is not None else None
            pct = (abs(selisih) / bersih * 100) if (selisih is not None and bersih) else None
            if bank_val is None:
                catatan = f"Tidak ada data Rekap untuk {catatan_target} bulan ini."
                fill = WARN_FILL
            elif pct is None:
                catatan = "-"
                fill = ALT_FILL if row_i % 2 else None
            elif pct < 5:
                catatan = "Selisih kecil, kemungkinan besar jeda settlement normal."
                fill = GOOD_FILL
            elif pct < 15:
                catatan = "Selisih cukup besar - cek transaksi awal/akhir bulan (kemungkinan settlement lintas bulan) atau refund."
                fill = WARN_FILL
            else:
                catatan = "SELISIH BESAR - kemungkinan salah kategori/metode, refund tidak tercatat rapi, atau setoran belum masuk. Perlu ditelusuri manual."
                fill = BAD_FILL
            ws.cell(row=r, column=1, value=bulan)
            ws.cell(row=r, column=2, value=f"{metode} ({n} transaksi)")
            ws.cell(row=r, column=3, value=kotor)
            ws.cell(row=r, column=4, value=refund)
            ws.cell(row=r, column=5, value=bersih)
            ws.cell(row=r, column=6, value=bank_val if bank_val is not None else "-")
            ws.cell(row=r, column=7, value=selisih if selisih is not None else "-")
            ws.cell(row=r, column=8, value=f"{pct:.1f}%" if pct is not None else "-")
            ws.cell(row=r, column=9, value=catatan)
            for c in (3, 4, 5, 6, 7):
                ws.cell(row=r, column=c).number_format = rc.NUMBER_FORMAT
            for c in range(1, 10):
                ws.cell(row=r, column=c).border = BORDER
                if fill:
                    ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=fill)
            row_i += 1
            r += 1

    # QRIS BRI + Kartu BRI digabung (lihat keterbatasan di docstring) -
    # baris terpisah, dibangun ulang bersih di sini
    r += 1
    ws.cell(row=r, column=1, value="Rincian gabungan non-tunai BRI (QRIS BRI + Kartu BRI vs total Penjualan BRI-507):")
    ws.cell(row=r, column=1).font = Font(bold=True, italic=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    r += 1
    for bulan in bulan_list:
        qris = pos_agg.get((bulan, "QRIS BRI"), {"kotor": 0, "refund": 0, "n": 0})
        kartu = pos_agg.get((bulan, "Kartu BRI"), {"kotor": 0, "refund": 0, "n": 0})
        kotor = qris["kotor"] + kartu["kotor"]
        refund = qris["refund"] + kartu["refund"]
        bersih = kotor - refund
        n = qris["n"] + kartu["n"]
        bank_val = bank_agg.get((bulan, "BRI-507"))
        if bank_val is None and n == 0:
            continue
        selisih = (bersih - bank_val) if bank_val is not None else None
        pct = (abs(selisih) / bersih * 100) if (selisih is not None and bersih) else None
        if bank_val is None:
            catatan, fill = f"Tidak ada data Rekap untuk BRI-507 bulan ini.", WARN_FILL
        elif pct is None:
            catatan, fill = "-", None
        elif pct < 5:
            catatan, fill = "Selisih kecil, kemungkinan besar jeda settlement normal.", GOOD_FILL
        elif pct < 15:
            catatan, fill = "Selisih cukup besar - cek transaksi awal/akhir bulan atau refund.", WARN_FILL
        else:
            catatan, fill = "SELISIH BESAR - perlu ditelusuri manual.", BAD_FILL
        ws.cell(row=r, column=1, value=bulan)
        ws.cell(row=r, column=2, value=f"QRIS BRI + Kartu BRI ({n} transaksi)")
        ws.cell(row=r, column=3, value=kotor)
        ws.cell(row=r, column=4, value=refund)
        ws.cell(row=r, column=5, value=bersih)
        ws.cell(row=r, column=6, value=bank_val if bank_val is not None else "-")
        ws.cell(row=r, column=7, value=selisih if selisih is not None else "-")
        ws.cell(row=r, column=8, value=f"{pct:.1f}%" if pct is not None else "-")
        ws.cell(row=r, column=9, value=catatan)
        for c in (3, 4, 5, 6, 7):
            ws.cell(row=r, column=c).number_format = rc.NUMBER_FORMAT
        for c in range(1, 10):
            ws.cell(row=r, column=c).border = BORDER
            if fill:
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=fill)
        r += 1

    widths = [16, 26, 15, 14, 15, 18, 14, 12, 45]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A5"

    # Sheet detail refund
    ws2 = wb.create_sheet("Detail Refund")
    ws2["A1"] = "DETAIL TRANSAKSI REFUND"
    ws2["A1"].font = Font(bold=True, size=13)
    r2 = 3
    h2 = ["No Transaksi", "Tanggal Bayar", "Metode", "Total", "Tanggal Refund", "Jumlah Refund", "Metode Refund", "Alasan", "Kasir"]
    for i, h in enumerate(h2, start=1):
        c = ws2.cell(row=r2, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEADER_FILL)
        c.border = BORDER
    r2 += 1
    refund_txns = sorted((t for t in all_pos if t["status"] == "Refund"), key=lambda t: t["tanggal"] or datetime.min)
    for t in refund_txns:
        ws2.cell(row=r2, column=1, value=t["no_transaksi"])
        ws2.cell(row=r2, column=2, value=t["waktu_bayar"])
        ws2.cell(row=r2, column=3, value=t["metode"])
        ws2.cell(row=r2, column=4, value=t["total"])
        ws2.cell(row=r2, column=5, value=t["refund_tanggal"])
        ws2.cell(row=r2, column=6, value=t["refund_jumlah"])
        ws2.cell(row=r2, column=7, value=t["refund_metode"])
        ws2.cell(row=r2, column=8, value=t["refund_alasan"])
        ws2.cell(row=r2, column=9, value=t["kasir"])
        for c in (4, 6):
            ws2.cell(row=r2, column=c).number_format = rc.NUMBER_FORMAT
        for c in range(1, 10):
            ws2.cell(row=r2, column=c).border = BORDER
        r2 += 1
    if not refund_txns:
        ws2.cell(row=r2, column=1, value="(tidak ada transaksi refund di periode ini)")
        ws2.cell(row=r2, column=1).font = Font(italic=True, color="6B7280")
    widths2 = [20, 18, 22, 14, 18, 14, 16, 18, 14]
    for i, w in enumerate(widths2, start=1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # Sheet transaksi belum lunas (informational)
    ws3 = wb.create_sheet("Belum Lunas (dikecualikan)")
    ws3["A1"] = "TRANSAKSI 'BELUM LUNAS' - DIKECUALIKAN DARI AUDIT (uang belum berpindah tangan)"
    ws3["A1"].font = Font(bold=True, size=13)
    r3 = 3
    for i, h in enumerate(["No Transaksi", "Waktu Order", "Total", "Kasir"], start=1):
        c = ws3.cell(row=r3, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=HEADER_FILL)
        c.border = BORDER
    r3 += 1
    belum_lunas = [t for t in all_pos if t["status"] == "Belum Lunas"]
    for t in belum_lunas:
        ws3.cell(row=r3, column=1, value=t["no_transaksi"])
        ws3.cell(row=r3, column=2, value=t["waktu_order"])
        ws3.cell(row=r3, column=3, value=t["total"])
        ws3.cell(row=r3, column=4, value=t["kasir"])
        ws3.cell(row=r3, column=3).number_format = rc.NUMBER_FORMAT
        for c in range(1, 5):
            ws3.cell(row=r3, column=c).border = BORDER
        r3 += 1
    if not belum_lunas:
        ws3.cell(row=r3, column=1, value="(tidak ada transaksi belum lunas di periode ini)")
        ws3.cell(row=r3, column=1).font = Font(italic=True, color="6B7280")
    for i, w in enumerate([20, 20, 14, 14], start=1):
        ws3.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # Rekonsiliasi harian per kategori metode, asumsi H+1 untuk QRIS/kartu
    all_dates_set = {tgl for (tgl, _) in pos_daily_raw} | {tgl for (_, tgl) in bank_daily}
    all_dates_sorted = sorted(all_dates_set)

    cash_daily = {tgl: v for (tgl, m), v in pos_daily_raw.items() if m == "Cash"}
    qris_bri_kartu_daily = {}
    for (tgl, m), v in pos_daily_raw.items():
        if m in ("QRIS BRI", "Kartu BRI"):
            a = qris_bri_kartu_daily.setdefault(tgl, {"kotor": 0, "refund": 0, "n": 0})
            a["kotor"] += v["kotor"]
            a["refund"] += v["refund"]
            a["n"] += v["n"]
    qris_bca_daily = {tgl: v for (tgl, m), v in pos_daily_raw.items() if m == "QRIS BCA"}

    n_selisih_cash = _write_harian_sheet(
        wb, "Harian - Cash", cash_daily, bank_daily, "Kas-Buku", all_dates_sorted, 0,
        "Cash tidak ada penundaan settlement - dibandingkan tanggal yang SAMA persis."
    )
    n_selisih_bri = _write_harian_sheet(
        wb, "Harian - QRIS BRI+Kartu (H+1)", qris_bri_kartu_daily, bank_daily, "BRI-507", all_dates_sorted, 1,
        "Asumsi QRIS BRI dan Kartu BRI baru masuk rekening BRI-507 SEHARI SETELAH tanggal transaksi POS "
        "(H+1). QRIS BRI + Kartu BRI digabung karena sama-sama settle ke BRI-507 tanpa penanda konsisten "
        "yang membedakan keduanya di sisi bank."
    )
    n_selisih_bca = _write_harian_sheet(
        wb, "Harian - QRIS BCA (H+1)", qris_bca_daily, bank_daily, "BCA-887", all_dates_sorted, 1,
        "Asumsi QRIS BCA baru masuk rekening BCA-887 SEHARI SETELAH tanggal transaksi POS (H+1)."
    )

    wb.save(output_path)
    n_refund_total = sum(t["refund_jumlah"] for t in all_pos if t["status"] == "Refund")
    return {
        "n_bulan": len(bulan_list),
        "n_transaksi_pos": len(all_pos),
        "n_refund": len(refund_txns),
        "total_refund": n_refund_total,
        "n_belum_lunas": len(belum_lunas),
        "n_hari_selisih_cash": n_selisih_cash,
        "n_hari_selisih_qris_bri": n_selisih_bri,
        "n_hari_selisih_qris_bca": n_selisih_bca,
    }


if __name__ == "__main__":
    import sys
    args = sys.argv[1:]
    pos_paths = [a for a in args if "detail" in a.lower() or "sales" in a.lower()]
    rekap_paths = [a for a in args if a not in pos_paths and not a.endswith("_out.xlsx")]
    out = "kasir_audit_output.xlsx"
    print(run_kasir_audit(pos_paths, rekap_paths, out))
