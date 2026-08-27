"""
quarterly.py
Laporan keuangan kuartalan: Laba Rugi, Neraca, Arus Kas, dan Roster Gaji 3
Bulan - dibangun dari 3 file hasil rekonsiliasi bulanan yang SUDAH
completed (bulan harus berurutan).

Beda dengan reconcile.py (yang MELAKUKAN rekonsiliasi), modul ini murni
MENGGABUNGKAN 3 bulan yang datanya sudah bersih. Tidak ada pencocokan
transfer baru di sini - itu tanggung jawab reconcile.py di masing-masing
file bulanan sebelum digabung.

Semua angka tetap rumus Excel beralamat absolut, merujuk ke sheet rekening
yang disalin apa adanya dari tiap file bulanan ke satu workbook kuartalan.
"""

import datetime
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy as copy_style

import reconcile as rc

REPORT_SHEET_NAMES = {
    "Rekonsiliasi", "Laporan Laba Rugi", "Neraca", "Laporan Arus Kas",
    "Diagnostik Keseimbangan",
}


class QuarterlyInputError(Exception):
    """Dilempar kalau input tidak valid (bulan tidak berurutan, file rusak,
    dll) - pesannya dirancang untuk langsung ditunjukkan ke user di bot."""


# ---------------------------------------------------------------------------
# Baca & validasi 3 file bulanan
# ---------------------------------------------------------------------------

def load_month(path):
    wb = openpyxl.load_workbook(path)
    account_sheets = [s for s in wb.sheetnames if s not in REPORT_SHEET_NAMES]
    if not account_sheets:
        raise QuarterlyInputError(
            f"File '{path}' tidak punya sheet rekening yang dikenali (cuma ada: {wb.sheetnames})."
        )
    all_txns = []
    sheets_last_row = {}
    opening_rows = {}
    txns_by_sheet = {}
    for sname in account_sheets:
        ws = wb[sname]
        lr = rc.last_data_row(ws)
        sheets_last_row[sname] = lr
        txns, _ = rc.read_account_sheet(ws)
        all_txns.extend(txns)
        txns_by_sheet[sname] = txns
        opening = next((t for t in txns if t.is_opening), None)
        opening_rows[sname] = opening.row if opening else 2
    year, month = rc.detect_period(all_txns)
    if year is None:
        raise QuarterlyInputError(f"Periode tidak terdeteksi di file '{path}'.")
    return {
        "wb": wb,
        "year": year,
        "month": month,
        "sheets_last_row": sheets_last_row,
        "opening_rows": opening_rows,
        "account_sheets": account_sheets,
        "txns_by_sheet": txns_by_sheet,
    }


def validate_consecutive(months):
    """months: list of dict dari load_month(), URUTAN SESUAI UPLOAD (belum
    tentu urut kronologis). Return list yang sama tapi sudah disortir
    kronologis. Raise QuarterlyInputError kalau bulannya bukan 3 bulan
    berturut-turut."""
    if len(months) != 3:
        raise QuarterlyInputError(f"Butuh tepat 3 file, yang diterima {len(months)}.")
    ordered = sorted(months, key=lambda m: (m["year"], m["month"]))
    labels = [f"{rc.MONTHS_ID[m['month']]} {m['year']}" for m in ordered]
    for i in range(1, 3):
        y0, m0 = ordered[i - 1]["year"], ordered[i - 1]["month"]
        y1, m1 = ordered[i]["year"], ordered[i]["month"]
        exp_m = 1 if m0 == 12 else m0 + 1
        exp_y = y0 + 1 if m0 == 12 else y0
        if (y1, m1) != (exp_y, exp_m):
            raise QuarterlyInputError(
                "Bulan yang diupload tidak berurutan: " + ", ".join(labels) +
                f". Setelah {rc.MONTHS_ID[m0]} {y0} seharusnya "
                f"{rc.MONTHS_ID[exp_m]} {exp_y}, bukan {rc.MONTHS_ID[m1]} {y1}."
            )
    return ordered


# ---------------------------------------------------------------------------
# Salin sheet rekening apa adanya ke workbook kuartalan
# ---------------------------------------------------------------------------

def copy_sheet_into(src_ws, dest_wb, new_title):
    dest_ws = dest_wb.create_sheet(new_title)
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = dest_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy_style(cell.font)
                new_cell.fill = copy_style(cell.fill)
                new_cell.border = copy_style(cell.border)
                new_cell.alignment = copy_style(cell.alignment)
                new_cell.number_format = cell.number_format
    for col_letter, dim in src_ws.column_dimensions.items():
        dest_ws.column_dimensions[col_letter].width = dim.width
    dest_ws.freeze_panes = src_ws.freeze_panes
    return dest_ws


# ---------------------------------------------------------------------------
# Helper rumus lintas-sheet-dalam-satu-bulan
# ---------------------------------------------------------------------------

def sumif_across_sheets(sheets_info, category):
    parts = [
        f"SUMIF('{sheet}'!$C$2:$C${last_row},\"{category}\",'{sheet}'!$J$2:$J${last_row})"
        for sheet, last_row in sheets_info
    ]
    return "=" + "+".join(parts) if parts else "=0"


def sumif_multi_across_sheets(sheets_info, categories):
    parts = []
    for sheet, last_row in sheets_info:
        for cat in categories:
            parts.append(f"SUMIF('{sheet}'!$C$2:$C${last_row},\"{cat}\",'{sheet}'!$J$2:$J${last_row})")
    return "=" + "+".join(parts) if parts else "=0"


# ---------------------------------------------------------------------------
# Laba Rugi Kuartal (kolom = 3 bulan + TOTAL)
# ---------------------------------------------------------------------------

def write_quarterly_income_statement(wb, month_sheet_map):
    name = "Laba Rugi Kuartal"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    labels = [m["label"] for m in month_sheet_map]
    period_text = f"{labels[0]} - {labels[-1]}"
    ws["A1"] = f"LAPORAN LABA RUGI KUARTALAN - {period_text.upper()}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Kolom = tiap bulan (rumus SUMIF ke sheet rekening bulan itu), kolom TOTAL "
                "= jumlah kuartal. Roster gaji per pegawai ada di sheet terpisah.")
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    r = 4
    rc.write_pivot_header(ws, r, labels)
    r += 1

    rc.write_pivot_section(ws, r, "PENDAPATAN", labels)
    r += 1
    rev_rows = []
    for cat in rc.INCOME_CATEGORIES_REVENUE:
        rc.write_pivot_data_row(
            ws, r, cat, labels,
            lambda label, cat=cat: sumif_across_sheets(_sheets_for_label(month_sheet_map, label), cat),
        )
        rev_rows.append(r)
        r += 1
    total_rev_row = r
    rc.write_pivot_subtotal_row(ws, r, "Total Pendapatan", labels, rev_rows)
    r += 2

    rc.write_pivot_section(ws, r, "BEBAN", labels)
    r += 1
    exp_rows = []
    for cat in rc.INCOME_CATEGORIES_EXPENSE:
        rc.write_pivot_data_row(
            ws, r, cat, labels,
            lambda label, cat=cat: sumif_across_sheets(_sheets_for_label(month_sheet_map, label), cat),
        )
        exp_rows.append(r)
        r += 1
    rc.write_pivot_data_row(
        ws, r, "Marketing & RnD", labels,
        lambda label: sumif_multi_across_sheets(_sheets_for_label(month_sheet_map, label), rc.MARKETING_RND_CATEGORY_TEXTS),
    )
    exp_rows.append(r)
    r += 1
    all_sheets_quarter = _all_sheets(month_sheet_map)
    nama_bulan_list = [_nama_bulan_short(lb) for lb in labels]
    rc.write_pivot_data_row(
        ws, r, "Gaji Pegawai (rincian per orang: sheet Roster Gaji)", labels,
        lambda label: sumif_gaji_bulan_kuartal(all_sheets_quarter, _nama_bulan_short(label)),
    )
    exp_rows.append(r)
    r += 1
    rc.write_pivot_data_row(
        ws, r, "Gaji Lainnya (accrual luar kuartal - ditotal di kolom bulan pertama)", labels,
        lambda label: "=0" if label != labels[0] else sumif_gaji_lainnya_kuartal(all_sheets_quarter, nama_bulan_list),
    )
    exp_rows.append(r)
    r += 1
    rc.write_pivot_data_row(
        ws, r, "Biaya Admin Bank (termasuk biaya transfer Fliptech)", labels,
        lambda label: sumif_multi_across_sheets(_sheets_for_label(month_sheet_map, label), rc.BANK_FEE_CATEGORY_TEXTS),
    )
    exp_rows.append(r)
    r += 1
    total_exp_row = r
    rc.write_pivot_subtotal_row(ws, r, "Total Beban", labels, exp_rows)
    r += 2

    rc.write_pivot_section(ws, r, "LAIN-LAIN (perlu verifikasi manual)", labels)
    r += 1
    other_rows = []
    for cat in rc.OTHER_CATEGORIES:
        rc.write_pivot_data_row(
            ws, r, cat, labels,
            lambda label, cat=cat: sumif_across_sheets(_sheets_for_label(month_sheet_map, label), cat),
        )
        other_rows.append(r)
        r += 1
    total_other_row = r
    rc.write_pivot_subtotal_row(ws, r, "Total Lain-lain", labels, other_rows)
    r += 2

    net_row = r
    rc.write_pivot_formula_row(
        ws, r, "LABA / RUGI BERSIH", labels,
        lambda cl: f"={cl}{total_rev_row}+{cl}{total_exp_row}+{cl}{total_other_row}",
        bold=True,
    )
    for c in range(1, rc.pivot_total_col(labels) + 1):
        ws.cell(row=r, column=c).font = Font(bold=True, size=12)

    ws.column_dimensions["A"].width = 42
    for i in range(len(labels)):
        ws.column_dimensions[rc.col_letter(2 + i)].width = 18
    ws.column_dimensions[rc.col_letter(rc.pivot_total_col(labels))].width = 18
    ws.freeze_panes = "B5"

    return {"sheet": name, "net": net_row, "total_rev": total_rev_row,
            "total_exp": total_exp_row, "total_other": total_other_row,
            "labels": labels, "total_col": rc.pivot_total_col(labels)}


def _sheets_for_label(month_sheet_map, label):
    for m in month_sheet_map:
        if m["label"] == label:
            return m["sheets"]
    return []


def _opening_for_label(month_sheet_map, label):
    for m in month_sheet_map:
        if m["label"] == label:
            return m["opening"]
    return []


def _all_sheets(month_sheet_map):
    """Semua sheet rekening dari SELURUH kuartal digabung jadi satu list -
    dipakai buat baris Gaji, karena gaji untuk bulan X bisa saja BARU
    tercatat/dibayar di sheet bulan Y (telat bayar). Supaya "Gaji Eva
    Januari 2023" tetap masuk kolom Januari meskipun transaksinya baru
    muncul di sheet Februari, pencarian kategori+bulan harus lintas semua
    sheet kuartal, bukan cuma sheet bulan itu sendiri."""
    sheets = []
    for m in month_sheet_map:
        sheets.extend(m["sheets"])
    return sheets


def _nama_bulan_short(label):
    """'Januari 2025' -> 'Januari' (nama bulan tanpa tahun, dipakai buat
    wildcard pencarian di Keterangan - contoh nyata di data ('Gaji
    Latifatul Husna Januari') kadang tidak menyertakan tahun sama sekali)."""
    return label.rsplit(" ", 1)[0]


def sumif_gaji_bulan_kuartal(all_sheets, nama_bulan):
    """SUMIFS lintas SEMUA sheet kuartal: kategori 'Gaji*' DAN keterangan
    menyebut nama_bulan - menangkap gaji untuk bulan itu di manapun dia
    tercatat (termasuk kalau baru dibayar/tercatat di bulan berikutnya)."""
    parts = [
        (f"SUMIFS('{sheet}'!$J$2:$J${last_row},"
         f"'{sheet}'!$C$2:$C${last_row},\"Gaji*\","
         f"'{sheet}'!$B$2:$B${last_row},\"*{nama_bulan}*\")")
        for sheet, last_row in all_sheets
    ]
    return "=" + "+".join(parts) if parts else "=0"


def sumif_gaji_lainnya_kuartal(all_sheets, nama_bulan_list):
    """Jaring pengaman: total semua baris 'Gaji*' di seluruh kuartal DIKURANGI
    yang sudah tertangkap oleh 3 kolom bulan di atas - supaya gaji yang
    keterangannya menyebut bulan DI LUAR kuartal (mis. accrual dari sebelum
    kuartal ini mulai) tidak diam-diam hilang dari Laba Rugi Kuartal."""
    parts = ["+".join(
        f"SUMIF('{sheet}'!$C$2:$C${last_row},\"Gaji*\",'{sheet}'!$J$2:$J${last_row})"
        for sheet, last_row in all_sheets
    ) or "0"]
    for nama_bulan in nama_bulan_list:
        parts.append(sumif_gaji_bulan_kuartal(all_sheets, nama_bulan)[1:])
    return "=" + parts[0] + "".join(f"-({p})" for p in parts[1:])


def sumifs_employee_bulan_kuartal(all_sheets, employee_name, nama_bulan):
    """Sama seperti sumif_gaji_bulan_kuartal tapi ditambah filter nama
    pegawai (kolom Objek) - dipakai per sel di Roster Gaji."""
    name = employee_name.strip()
    parts = [
        (f"SUMIFS('{sheet}'!$J$2:$J${last_row},"
         f"'{sheet}'!$C$2:$C${last_row},\"Gaji*\","
         f"'{sheet}'!$H$2:$H${last_row},\"*{name}*\","
         f"'{sheet}'!$B$2:$B${last_row},\"*{nama_bulan}*\")")
        for sheet, last_row in all_sheets
    ]
    return "=" + "+".join(parts) if parts else "=0"


# ---------------------------------------------------------------------------
# Helper pivot "snapshot": khusus buat Neraca Kuartal, di mana kolom kanan
# HARUS berarti "posisi di akhir kuartal" (= kolom bulan terakhir), BUKAN
# penjumlahan 3 bulan - beda dengan Laba Rugi/Arus Kas yang memang laporan
# arus (flow) sehingga penjumlahan masuk akal. Neraca itu snapshot (stok),
# menjumlah 3 snapshot saldo akan salah secara akuntansi (mis. Total Aset
# Jan+Feb+Mar tidak berarti apa-apa).
# ---------------------------------------------------------------------------

def write_snapshot_header(ws, row, labels):
    ws.cell(row=row, column=1, value="")
    for i, label in enumerate(labels):
        c = 2 + i
        cell = ws.cell(row=row, column=c, value=label)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    total_col = 2 + len(labels)
    ws.cell(row=row, column=total_col, value=f"AKHIR KUARTAL ({labels[-1]})")
    rc.style_header(ws, row, total_col)
    ws.row_dimensions[row].height = 32


def write_snapshot_data_row(ws, row, label, labels, formula_fn, bold=False):
    ws.cell(row=row, column=1, value=label)
    n = len(labels)
    for i, lb in enumerate(labels):
        c = 2 + i
        ws.cell(row=row, column=c, value=formula_fn(lb))
        ws.cell(row=row, column=c).number_format = rc.NUMBER_FORMAT
    total_col = 2 + n
    last_col_letter = rc.col_letter(total_col - 1)
    ws.cell(row=row, column=total_col, value=f"={last_col_letter}{row}")
    ws.cell(row=row, column=total_col).number_format = rc.NUMBER_FORMAT
    if bold:
        for c in range(1, total_col + 1):
            ws.cell(row=row, column=c).font = Font(bold=True)


def write_snapshot_subtotal_row(ws, row, label, labels, ref_rows, bold=True):
    n = len(labels)
    ws.cell(row=row, column=1, value=label)
    for i in range(n):
        c = 2 + i
        cl = rc.col_letter(c)
        ws.cell(row=row, column=c, value=f"=SUM({cl}{ref_rows[0]}:{cl}{ref_rows[-1]})")
        ws.cell(row=row, column=c).number_format = rc.NUMBER_FORMAT
    total_col = 2 + n
    last_col_letter = rc.col_letter(total_col - 1)
    ws.cell(row=row, column=total_col, value=f"={last_col_letter}{row}")
    ws.cell(row=row, column=total_col).number_format = rc.NUMBER_FORMAT
    if bold:
        for c in range(1, total_col + 1):
            ws.cell(row=row, column=c).font = Font(bold=True)


def write_snapshot_formula_row(ws, row, label, labels, per_col_formula_fn, bold=False):
    n = len(labels)
    ws.cell(row=row, column=1, value=label)
    for i in range(n):
        c = 2 + i
        cl = rc.col_letter(c)
        ws.cell(row=row, column=c, value=per_col_formula_fn(cl))
        ws.cell(row=row, column=c).number_format = rc.NUMBER_FORMAT
    total_col = 2 + n
    last_col_letter = rc.col_letter(total_col - 1)
    ws.cell(row=row, column=total_col, value=f"={last_col_letter}{row}")
    ws.cell(row=row, column=total_col).number_format = rc.NUMBER_FORMAT
    if bold:
        for c in range(1, total_col + 1):
            ws.cell(row=row, column=c).font = Font(bold=True)

# ---------------------------------------------------------------------------
# Neraca Kuartal (kolom = 3 bulan, tiap kolom snapshot akhir bulan itu;
# sisi Ekuitas dihitung KUMULATIF dari bulan pertama supaya nyambung)
# ---------------------------------------------------------------------------

def write_quarterly_balance_sheet(wb, month_sheet_map, income_ref):
    name = "Neraca Kuartal"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    labels = [m["label"] for m in month_sheet_map]
    ws["A1"] = f"NERACA KUARTALAN - PER AKHIR {labels[0].upper()} S.D. {labels[-1].upper()}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Tiap kolom = posisi akhir bulan itu (snapshot). Sisi Ekuitas dihitung kumulatif "
                "dari bulan pertama kuartal supaya nyambung dengan sisi Aset.")
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    r = 4
    write_snapshot_header(ws, r, labels)
    r += 1

    rc.write_pivot_section(ws, r, "ASET (KAS & SETARA KAS)", labels)
    r += 1
    kas_row = r
    write_snapshot_data_row(
        ws, r, "Kas & Setara Kas (Saldo Akhir Bulan)", labels,
        lambda label: sumif_kas_akhir(_sheets_for_label(month_sheet_map, label)),
        bold=True,
    )
    total_asset_row = kas_row
    r += 2

    rc.write_pivot_section(ws, r, "EKUITAS (kumulatif sejak awal kuartal)", labels)
    r += 1
    saldo_awal_kuartal_formula = sumif_saldo_awal(_opening_for_label(month_sheet_map, labels[0]))
    saldo_awal_row = r
    write_snapshot_data_row(ws, r, "Saldo Awal Kuartal (tetap, dari bulan pertama)", labels,
                             lambda label: saldo_awal_kuartal_formula)
    r += 1
    modal_row = r
    write_snapshot_data_row(
        ws, r, "Modal & Setoran Pemilik (kumulatif s.d. bulan ini)", labels,
        lambda label: sumif_modal_kumulatif(month_sheet_map, label),
    )
    r += 1
    laba_row = r
    write_snapshot_formula_row(
        ws, r, "Laba Bersih (kumulatif s.d. bulan ini)", labels,
        lambda cl: laba_kumulatif_formula(income_ref, labels, cl),
    )
    r += 1
    total_equity_row = r
    write_snapshot_subtotal_row(ws, r, "Total Ekuitas", labels, [saldo_awal_row, laba_row])
    r += 2

    balance_check_row = r
    write_snapshot_formula_row(
        ws, r, "CEK KESEIMBANGAN (Aset - Ekuitas)", labels,
        lambda cl: f"={cl}{total_asset_row}-{cl}{total_equity_row}",
        bold=True,
    )
    r += 1
    transfer_row = r
    write_snapshot_data_row(
        ws, r, "Transfer Bersih Kumulatif (info)", labels,
        lambda label: sumif_transfer_kumulatif(month_sheet_map, label),
    )
    r += 1
    residual_row = r
    write_snapshot_formula_row(
        ws, r, "Selisih Belum Terjelaskan", labels,
        lambda cl: f"={cl}{balance_check_row}-{cl}{transfer_row}",
        bold=True,
    )

    ws.column_dimensions["A"].width = 42
    for i in range(len(labels)):
        ws.column_dimensions[rc.col_letter(2 + i)].width = 18
    ws.column_dimensions[rc.col_letter(rc.pivot_total_col(labels))].width = 18
    ws.freeze_panes = "B5"

    return {"sheet": name, "total_asset": total_asset_row, "total_equity": total_equity_row,
            "saldo_awal": saldo_awal_row, "balance_check": balance_check_row,
            "labels": labels, "total_col": rc.pivot_total_col(labels)}


def sumif_kas_akhir(sheets_info):
    parts = [f"'{sheet}'!$K${last_row}" for sheet, last_row in sheets_info]
    return "=" + "+".join(parts) if parts else "=0"


def sumif_saldo_awal(opening_info):
    parts = [f"'{sheet}'!$K${row}" for sheet, row in opening_info]
    return "=" + "+".join(parts) if parts else "=0"


def sumif_modal_kumulatif(month_sheet_map, upto_label):
    parts = []
    for m in month_sheet_map:
        parts.append(sumif_across_sheets(m["sheets"], "Modal & Setoran Pemilik")[1:])
        if m["label"] == upto_label:
            break
    return "=" + "+".join(parts) if parts else "=0"


def sumif_transfer_kumulatif(month_sheet_map, upto_label):
    parts = []
    for m in month_sheet_map:
        parts.append(sumif_multi_across_sheets(m["sheets"], rc.TRANSFER_CATEGORY_TEXTS)[1:])
        if m["label"] == upto_label:
            break
    return "=" + "+".join(parts) if parts else "=0"


def laba_kumulatif_formula(income_ref, labels, current_col_letter):
    """SUM kolom Laba Rugi dari bulan pertama s.d. kolom bulan ini
    (kumulatif). current_col_letter adalah huruf kolom bulan (bukan kolom
    TOTAL - itu ditangani otomatis lewat mirroring di write_snapshot_*)."""
    income_sheet = income_ref["sheet"]
    net_row = income_ref["net"]
    first_letter = rc.col_letter(2)
    return f"=SUM('{income_sheet}'!{first_letter}{net_row}:{current_col_letter}{net_row})"


# ---------------------------------------------------------------------------
# Laporan Arus Kas Kuartal
# ---------------------------------------------------------------------------

def write_quarterly_cash_flow(wb, month_sheet_map, income_ref, balance_ref):
    name = "Arus Kas Kuartal"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    labels = [m["label"] for m in month_sheet_map]
    ws["A1"] = f"LAPORAN ARUS KAS KUARTALAN - {labels[0].upper()} S.D. {labels[-1].upper()}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Metode langsung, kolom = tiap bulan + TOTAL kuartal."
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    r = 4
    rc.write_pivot_header(ws, r, labels)
    r += 1

    rc.write_pivot_section(ws, r, "ARUS KAS DARI AKTIVITAS OPERASI", labels)
    r += 1
    op_row = r
    rc.write_pivot_formula_row(
        ws, r, "Laba Bersih Bulan Ini (basis kas)", labels,
        lambda cl: f"='{income_ref['sheet']}'!{cl}{income_ref['net']}",
    )
    r += 1
    total_op_row = r
    rc.write_pivot_subtotal_row(ws, r, "Kas Bersih dari Operasi", labels, [op_row, op_row])
    r += 2

    rc.write_pivot_section(ws, r, "ARUS KAS DARI AKTIVITAS PENDANAAN", labels)
    r += 1
    fin_row = r
    rc.write_pivot_data_row(
        ws, r, "Modal & Setoran Pemilik", labels,
        lambda label: sumif_across_sheets(_sheets_for_label(month_sheet_map, label), "Modal & Setoran Pemilik"),
    )
    r += 1
    total_fin_row = r
    rc.write_pivot_subtotal_row(ws, r, "Kas Bersih dari Pendanaan", labels, [fin_row, fin_row])
    r += 2

    net_change_row = r
    rc.write_pivot_formula_row(
        ws, r, "KENAIKAN (PENURUNAN) KAS BERSIH", labels,
        lambda cl: f"={cl}{total_op_row}+{cl}{total_fin_row}",
        bold=True,
    )
    r += 1
    saldo_awal_row = r
    rc.write_pivot_formula_row(
        ws, r, "Saldo Kas Awal Bulan", labels,
        lambda cl: (
            f"='{balance_ref['sheet']}'!{rc.col_letter(2)}{balance_ref['saldo_awal']}"
            if cl == rc.col_letter(2)
            else _saldo_awal_chain(cl, labels, net_change_row, saldo_awal_row, balance_ref)
        ),
    )
    r += 1
    saldo_akhir_row = r
    rc.write_pivot_formula_row(
        ws, r, "Saldo Kas Akhir Bulan", labels,
        lambda cl: f"={cl}{net_change_row}+{cl}{saldo_awal_row}",
        bold=True,
    )
    r += 1
    rc.write_pivot_formula_row(
        ws, r, "Cek vs Total Aset di Neraca", labels,
        lambda cl: f"={cl}{saldo_akhir_row}-'{balance_ref['sheet']}'!{cl}{balance_ref['total_asset']}",
    )

    ws.column_dimensions["A"].width = 40
    for i in range(len(labels)):
        ws.column_dimensions[rc.col_letter(2 + i)].width = 18
    ws.column_dimensions[rc.col_letter(rc.pivot_total_col(labels))].width = 18
    ws.freeze_panes = "B5"
    return ws


def _saldo_awal_chain(cl, labels, net_change_row, saldo_awal_row, balance_ref):
    """Saldo awal bulan ke-2 & ke-3 = saldo akhir bulan sebelumnya (kolom
    sebelah kiri di sheet yang sama). Kolom TOTAL kuartal = saldo awal
    bulan pertama (titik mula kuartal)."""
    total_col_letter = rc.col_letter(rc.pivot_total_col(labels))
    if cl == total_col_letter:
        return f"={rc.col_letter(2)}{saldo_awal_row}"
    idx = None
    for i in range(len(labels)):
        if rc.col_letter(2 + i) == cl:
            idx = i
            break
    prev_letter = rc.col_letter(2 + idx - 1)
    saldo_akhir_row = saldo_awal_row + 1
    return f"={prev_letter}{saldo_akhir_row}"


# ---------------------------------------------------------------------------
# Roster Gaji 3 Bulan (matriks: baris = pegawai, kolom = 3 bulan + TOTAL)
# ---------------------------------------------------------------------------

def write_roster_gaji(wb, month_sheet_map):
    name = "Roster Gaji 3 Bulan"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    labels = [m["label"] for m in month_sheet_map]
    ws["A1"] = f"ROSTER GAJI PEGAWAI - {labels[0].upper()} S.D. {labels[-1].upper()}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Rumus SUMIFS: kategori Gaji* + Keterangan menyebut nama bulan kolom itu, dicari "
                "LINTAS SEMUA sheet kuartal (bukan cuma sheet bulan itu) - supaya gaji yang telat "
                f"dibayar/dicatat tetap masuk kolom bulan yang benar. Kalau kolom {labels[-1]} kosong "
                "padahal pegawainya muncul di bulan lain, kemungkinan gajinya belum dibayar - lihat kolom Catatan.")
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    # kumpulkan semua nama pegawai unik dari kolom Objek pada baris berkategori
    # Gaji* - DIKELOMPOKKAN BERDASARKAN BULAN YANG DISEBUT DI KETERANGAN
    # (mis. "Gaji Eva Januari 2023" -> bulan Januari), BUKAN berdasarkan di
    # sheet bulan mana transaksinya fisik tercatat. Gaji sering dibayar
    # telat (gaji Januari baru dibayar/tercatat Februari), jadi kalau
    # dikelompokkan per sheet, itu akan salah masuk kolom Februari padahal
    # seharusnya kolom Januari. Nama pegawai dinormalisasi (spasi dirapikan
    # + case-insensitive) supaya "Nama" dan " nama " tidak dianggap 2 orang.
    def _norm_name(name):
        return " ".join(name.split()).casefold()

    nama_bulan_list = [_nama_bulan_short(lb) for lb in labels]

    def _parse_bulan_dari_keterangan(ket):
        ket_lower = (ket or "").lower()
        for nama in nama_bulan_list:
            if nama.lower() in ket_lower:
                return nama
        return None

    employees_by_month = [set() for _ in labels]  # index selaras dengan labels
    canonical_name = {}
    all_employee_keys = []
    n_gaji_luar_kuartal = 0
    for m in month_sheet_map:
        for sname, txns in m["txns_by_sheet"].items():
            for t in txns:
                if (t.kategori or "").lower().startswith("gaji") and t.objek:
                    bulan_untuk = _parse_bulan_dari_keterangan(t.desc)
                    if bulan_untuk is None:
                        n_gaji_luar_kuartal += 1
                        continue  # gaji untuk bulan di luar kuartal ini, lihat baris "Gaji Lainnya" di Laba Rugi Kuartal
                    idx = nama_bulan_list.index(bulan_untuk)
                    key = _norm_name(t.objek)
                    employees_by_month[idx].add(key)
                    canonical_name.setdefault(key, " ".join(t.objek.split()))
    for idx in range(len(labels)):
        for key in sorted(employees_by_month[idx]):
            if key not in all_employee_keys:
                all_employee_keys.append(key)
    all_employees = [canonical_name[k] for k in all_employee_keys]

    r = 4
    rc.write_pivot_header(ws, r, labels)
    ws.cell(row=r, column=rc.pivot_total_col(labels) + 1, value="Catatan")
    ws.cell(row=r, column=rc.pivot_total_col(labels) + 1).fill = rc.HEADER_FILL
    ws.cell(row=r, column=rc.pivot_total_col(labels) + 1).font = rc.HEADER_FONT
    r += 1

    all_sheets_quarter = _all_sheets(month_sheet_map)
    n_belum_dibayar_bulan_terakhir = 0
    last_label = labels[-1]
    for key, emp in zip(all_employee_keys, all_employees):
        rc.write_pivot_data_row(
            ws, r, emp, labels,
            lambda label, emp=emp: sumifs_employee_bulan_kuartal(all_sheets_quarter, emp, _nama_bulan_short(label)),
        )
        # cek python-side (bukan formula) apakah pegawai ini dapat gaji untuk
        # bulan terakhir kuartal - kalau tidak, tapi muncul di bulan lain, kasih catatan
        paid_last_month = key in employees_by_month[-1]
        if not paid_last_month:
            n_belum_dibayar_bulan_terakhir += 1
            ws.cell(row=r, column=rc.pivot_total_col(labels) + 1,
                    value=(f"Belum ada gaji untuk {last_label} tercatat - kemungkinan "
                           f"dibebankan sebagai accrual di bulan setelah kuartal ini."))
            ws.cell(row=r, column=rc.pivot_total_col(labels) + 1).font = Font(italic=True, color="B45309")
            ws.cell(row=r, column=rc.pivot_total_col(labels) + 1).alignment = Alignment(wrap_text=True)
        r += 1

    total_row = r
    if all_employees:
        first_data_row = 5
        last_data_row_ = r - 1
        rc.write_pivot_subtotal_row(ws, r, "TOTAL", labels, [first_data_row, last_data_row_])
    else:
        ws.cell(row=r, column=1, value="(tidak ada transaksi berkategori Gaji untuk bulan-bulan di kuartal ini)")
        ws.cell(row=r, column=1).font = Font(italic=True, color="6B7280")
    r += 1
    if n_gaji_luar_kuartal:
        ws.cell(row=r, column=1,
                value=(f"Catatan: {n_gaji_luar_kuartal} baris transaksi berkategori Gaji* keterangannya "
                       "tidak menyebut salah satu dari 3 bulan kuartal ini (kemungkinan accrual dari "
                       "bulan sebelum kuartal dimulai) - totalnya tetap dihitung di Laba Rugi Kuartal "
                       "baris 'Gaji Lainnya', tapi tidak muncul di matriks roster ini karena tidak "
                       "jelas masuk kolom bulan yang mana."))
        ws.cell(row=r, column=1).font = Font(italic=True, size=9, color="6B7280")
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=rc.pivot_total_col(labels) + 1)
        ws.row_dimensions[r].height = 40

    ws.column_dimensions["A"].width = 28
    for i in range(len(labels)):
        ws.column_dimensions[rc.col_letter(2 + i)].width = 18
    ws.column_dimensions[rc.col_letter(rc.pivot_total_col(labels))].width = 18
    ws.column_dimensions[rc.col_letter(rc.pivot_total_col(labels) + 1)].width = 46
    ws.freeze_panes = "B5"

    return {"n_pegawai": len(all_employees),
            "n_belum_dibayar_bulan_terakhir": n_belum_dibayar_bulan_terakhir}


# ---------------------------------------------------------------------------
# Orkestrasi utama
# ---------------------------------------------------------------------------

def run_quarterly_report(paths, output_path):
    """paths: list of 3 path file bulanan (urutan upload bebas, akan
    disortir kronologis). Return dict ringkasan."""
    months = [load_month(p) for p in paths]
    months = validate_consecutive(months)

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    month_sheet_map = []  # list per bulan: {'label':.., 'sheets': [(new_name,last_row),...]}
    for m in months:
        label = f"{rc.MONTHS_ID[m['month']]} {m['year']}"
        sheets_info = []
        opening_info = []  # [(new_name, opening_row)]
        for sname in m["account_sheets"]:
            src_ws = m["wb"][sname]
            new_title = sname
            suffix = 2
            while new_title in out_wb.sheetnames:
                new_title = f"{sname[:28]} ({suffix})"
                suffix += 1
            copy_sheet_into(src_ws, out_wb, new_title)
            sheets_info.append((new_title, m["sheets_last_row"][sname]))
            opening_info.append((new_title, m["opening_rows"][sname]))
        month_sheet_map.append({
            "label": label, "sheets": sheets_info, "opening": opening_info,
            "year": m["year"], "month": m["month"], "txns_by_sheet": m["txns_by_sheet"],
        })

    income_ref = write_quarterly_income_statement(out_wb, month_sheet_map)
    balance_ref = write_quarterly_balance_sheet(out_wb, month_sheet_map, income_ref)
    write_quarterly_cash_flow(out_wb, month_sheet_map, income_ref, balance_ref)
    roster_summary = write_roster_gaji(out_wb, month_sheet_map)

    order = []
    for m in month_sheet_map:
        order += [s for s, _ in m["sheets"]]
    order += ["Laba Rugi Kuartal", "Neraca Kuartal", "Arus Kas Kuartal", "Roster Gaji 3 Bulan"]
    out_wb._sheets = [out_wb[s] for s in order]

    out_wb.save(output_path)

    return {
        "periode": [m["label"] for m in month_sheet_map],
        "n_pegawai": roster_summary["n_pegawai"],
        "n_belum_dibayar_bulan_terakhir": roster_summary["n_belum_dibayar_bulan_terakhir"],
    }


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 5:
        print("Usage: python quarterly.py bulan1.xlsx bulan2.xlsx bulan3.xlsx output.xlsx")
        sys.exit(1)
    paths = sys.argv[1:4]
    out = sys.argv[4]
    try:
        s = run_quarterly_report(paths, out)
        print(s)
    except QuarterlyInputError as e:
        print("Error:", e)
        sys.exit(1)
