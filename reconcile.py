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
THIN = Side(style="thin", color="D1D5DB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Kategori yang menandakan perpindahan uang ANTAR rekening sendiri
# (bukan pendapatan/beban riil) -> tidak masuk Laba Rugi, harus saling
# menutup nol di Rekonsiliasi.
TRANSFER_KEYWORDS = [
    "pindah rekening internal",
    "pindang rekening internal",
    "transfer internal",
    "transfer lainnya",
]

# Kategori setoran/penarikan modal pemilik -> masuk Neraca (ekuitas),
# bukan Laba Rugi.
CAPITAL_KEYWORDS = [
    "modal & setoran pemilik",
    "modal dan setoran pemilik",
]

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
        positif jika kredit (uang masuk)."""
        if self.debit:
            return self.debit  # sudah negatif di data sumber
        if self.kredit:
            return self.kredit
        return 0

    @property
    def is_transfer(self):
        k = (self.kategori or "").lower()
        return any(kw in k for kw in TRANSFER_KEYWORDS)

    @property
    def is_capital(self):
        k = (self.kategori or "").lower()
        return any(kw in k for kw in CAPITAL_KEYWORDS)

    @property
    def is_opening(self):
        k = (self.kategori or "").lower()
        return any(kw in k for kw in OPENING_KEYWORDS)

    def cell_ref(self, col):
        return f"'{self.sheet}'!${col}${self.row}"


# ---------------------------------------------------------------------------
# Membaca sheet rekening
# ---------------------------------------------------------------------------

def read_account_sheet(ws):
    """Baca satu sheet rekening jadi list[Txn], berhenti di baris kosong
    pertama setelah header (baris trailing kosong diabaikan)."""
    txns = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        tanggal, ket, kategori, debit, kredit, saldo, subjek, objek, ket_tambahan = (
            (c.value for c in row[:9])
        )
        if tanggal is None and ket is None and debit is None and kredit is None:
            continue
        txns.append(
            Txn(
                sheet=ws.title,
                row=row[0].row,
                date=tanggal,
                desc=ket or "",
                kategori=kategori or "",
                debit=float(debit) if isinstance(debit, (int, float)) else 0,
                kredit=float(kredit) if isinstance(kredit, (int, float)) else 0,
                saldo=float(saldo) if isinstance(saldo, (int, float)) else None,
                subjek=str(subjek) if subjek is not None else "",
                objek=str(objek) if objek is not None else "",
                ket=ket_tambahan or "",
            )
        )
    return txns


def last_data_row(ws):
    last = 1
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        tanggal, ket, kategori, debit, kredit = (c.value for c in row[:5])
        if any(v is not None for v in (tanggal, ket, kategori, debit, kredit)):
            last = row[0].row
    return last


def resolve_account_sheet(hint, sheet_names):
    """Cocokkan teks bebas di kolom Subjek/Objek (mis. 'BRI-567(Biz)',
    'Rekening Jago + Admin Rp200') ke nama sheet rekening sebenarnya."""
    if not hint:
        return None
    h = hint.lower()
    best = None
    for name in sheet_names:
        n = name.lower()
        # ambil token pembeda dari nama sheet, contoh nomor rekening / "jago" / "kasir"
        tokens = [t for t in n.replace("(", " ").replace(")", " ").split() if len(t) >= 3]
        for tok in tokens:
            if tok in h:
                best = name
                break
        if best:
            break
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


def days_between(a, b):
    if isinstance(a, datetime.datetime):
        a = a.date()
    if isinstance(b, datetime.datetime):
        b = b.date()
    if not isinstance(a, datetime.date) or not isinstance(b, datetime.date):
        return 9999
    return abs((a - b).days)


def find_matches(all_txns, sheet_names):
    """Untuk setiap transaksi bertanda transfer internal, cari pasangan di
    rekening tujuan (berdasarkan Subjek/Objek) dalam jendela +-30 hari,
    dengan toleransi selisih nominal untuk biaya admin/pembulatan.
    Tidak pernah menyimpulkan 'tidak ditemukan' tanpa mencoba seluruh
    kandidat di rekening tujuan terlebih dahulu."""
    transfers = [t for t in all_txns if t.is_transfer]
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
            and id(t) not in consumed_ids
            and (counterpart_hint is None or t.sheet == counterpart_hint)
            and (t.nominal > 0) != (src.nominal > 0)  # tanda berlawanan
        ]
        if not candidates and counterpart_hint is None:
            # tidak ada petunjuk rekening tujuan -> perluas ke semua sheet lain
            candidates = [
                t
                for t in all_txns
                if t is not src
                and t.is_transfer
                and id(t) not in consumed_ids
                and t.sheet != src.sheet
                and (t.nominal > 0) != (src.nominal > 0)
            ]

        # skor tiap kandidat: prioritaskan selisih nominal kecil, lalu tanggal dekat
        scored = []
        for c in candidates:
            nominal_diff = abs(abs(src.nominal) - abs(c.nominal))
            date_diff = days_between(src.date, c.date)
            if date_diff > TOLERANCI_HARI:
                continue
            toleransi = max(TOLERANSI_NOMINAL_ABS, abs(src.nominal) * TOLERANSI_NOMINAL_PERSEN)
            if nominal_diff > toleransi:
                continue
            scored.append((date_diff, nominal_diff, c))

        # prioritaskan kedekatan tanggal (perilaku transfer riil biasanya
        # settle dalam 0-3 hari), baru kedekatan nominal sebagai tie-breaker
        scored.sort(key=lambda x: (x[0], x[1]))

        if scored:
            date_diff, nominal_diff, dst = scored[0]
            matched_dst_ids.add(id(dst))
            consumed_ids.add(id(src))
            consumed_ids.add(id(dst))
            if nominal_diff == 0 and date_diff == 0:
                conf = "High"
                reason = "Nominal dan tanggal sama persis di kedua rekening."
            elif nominal_diff == 0 and date_diff <= 3:
                conf = "High"
                reason = f"Nominal sama persis, selisih tanggal {date_diff} hari (wajar untuk settlement bank)."
            elif nominal_diff > 0 and nominal_diff <= TOLERANSI_NOMINAL_ABS:
                conf = "Medium"
                reason = (
                    f"Selisih nominal Rp{nominal_diff:,.0f} kemungkinan biaya admin/transfer "
                    f"(mis. via Fliptech), selisih tanggal {date_diff} hari."
                ).replace(",", ".")
            elif date_diff > 3:
                conf = "Medium"
                reason = f"Nominal cocok (toleransi Rp{nominal_diff:,.0f}) tapi tanggal berbeda {date_diff} hari, kemungkinan delayed posting/settlement.".replace(",", ".")
            else:
                conf = "Low"
                reason = f"Kecocokan hanya berdasarkan toleransi umum, perlu verifikasi manual (selisih Rp{nominal_diff:,.0f}, {date_diff} hari).".replace(",", ".")
            results.append(
                Match(src=src, dst=dst, confidence=conf, reasoning=reason,
                      date_diff=date_diff, nominal_diff=nominal_diff)
            )
        else:
            results.append(
                Match(
                    src=src,
                    dst=None,
                    confidence="Needs manual verification",
                    reasoning=(
                        "Tidak ada kandidat dengan nominal berlawanan dalam jendela "
                        f"±{TOLERANCI_HARI} hari di rekening tujuan yang terindikasi "
                        f"({counterpart_hint or 'tidak teridentifikasi dari Subjek/Objek'}). "
                        "Kemungkinan: dana masih dalam perjalanan (in-transit), tercatat di "
                        "bulan berikutnya, atau salah kategori."
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


def find_minus_flags(all_txns_by_sheet):
    """Kumpulkan indikasi 'minus': saldo kumulatif negatif, kategori
    Tip/Minus/Lebih, atau baris berpenanda flag (⚑) di keterangan.

    Catatan: cross-check baris-per-baris terhadap saldo hasil rekonstruksi
    (opening + akumulasi nominal) sengaja TIDAK dipakai di sini karena pada
    sebagian sheet sumber (mis. BCA) kolom Saldo Kumulatif tidak selalu
    diisi berurutan per baris (beberapa transaksi bertanggal sama dikelompokkan
    dulu), sehingga cross-check per baris menghasilkan banyak false positive.
    Pengecekan keseimbangan yang valid tetap dilakukan di level bulanan lewat
    kolom K (Saldo Kumulatif Rekonstruksi) tiap sheet dan sheet Neraca."""
    flags = []
    for sheet, txns in all_txns_by_sheet.items():
        for t in txns:
            reasons = []
            if t.saldo is not None and t.saldo < 0:
                reasons.append(f"Saldo kumulatif negatif (Rp{t.saldo:,.0f}).".replace(",", "."))
            if "tip/minus/lebih" in (t.kategori or "").lower():
                reasons.append("Kategori Tip/Minus/Lebih (selisih kas fisik vs catatan).")
            if isinstance(t.ket, str) and "⚑" in t.ket:
                reasons.append(f"Ditandai perlu verifikasi manual: {t.ket}")
            if reasons:
                flags.append((t, reasons))
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


def write_rekonsiliasi_sheet(wb, matches, combo_matches, minus_flags):
    if "Rekonsiliasi" in wb.sheetnames:
        del wb["Rekonsiliasi"]
    ws = wb.create_sheet("Rekonsiliasi")

    ws["A1"] = "REKONSILIASI ANTAR REKENING"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Dibuat otomatis. Setiap baris merujuk langsung (link rumus) ke sel asal di sheet rekening."
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
    ]
    hdr_row = r
    for i, h in enumerate(headers, start=1):
        ws.cell(row=hdr_row, column=i, value=h)
    style_header(ws, hdr_row, len(headers))
    r += 1

    for m in matches:
        ws.cell(row=r, column=1, value=m.src.sheet)
        ws.cell(row=r, column=2, value=f"='{m.src.sheet}'!$A${m.src.row}")
        ws.cell(row=r, column=3, value=f"='{m.src.sheet}'!$B${m.src.row}")
        ws.cell(row=r, column=4, value=f"='{m.src.sheet}'!${'D' if m.src.debit else 'E'}${m.src.row}")
        if m.dst:
            ws.cell(row=r, column=5, value=m.dst.sheet)
            ws.cell(row=r, column=6, value=f"='{m.dst.sheet}'!$A${m.dst.row}")
            ws.cell(row=r, column=7, value=f"='{m.dst.sheet}'!$B${m.dst.row}")
            ws.cell(row=r, column=8, value=f"='{m.dst.sheet}'!${'D' if m.dst.debit else 'E'}${m.dst.row}")
            ws.cell(row=r, column=9, value=m.date_diff)
            ws.cell(row=r, column=10, value=round(m.nominal_diff, 2))
        else:
            ws.cell(row=r, column=5, value="(belum ditemukan)")
        ws.cell(row=r, column=11, value=m.confidence)
        ws.cell(row=r, column=12, value=m.reasoning)
        fill = conf_fill(m.confidence)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(c == 12))
        ws.cell(row=r, column=11).fill = fill
        r += 1

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
        ws.cell(row=r, column=3, value=f"='{src.sheet}'!${'D' if src.debit else 'E'}${src.row}")
        ws.cell(row=r, column=4, value=f"{a.sheet} (brs {a.row})")
        ws.cell(row=r, column=5, value=f"='{a.sheet}'!${'D' if a.debit else 'E'}${a.row}")
        ws.cell(row=r, column=6, value=f"{b.sheet} (brs {b.row})")
        ws.cell(row=r, column=7, value=f"='{b.sheet}'!${'D' if b.debit else 'E'}${b.row}")
        ws.cell(row=r, column=8, value=round(cm["diff"], 2))
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
        ws.cell(row=r, column=3, value=f"='{t.sheet}'!$B${t.row}")
        ws.cell(row=r, column=4, value=f"='{t.sheet}'!${'D' if t.debit else 'E'}${t.row}")
        ws.cell(row=r, column=5, value=f"='{t.sheet}'!$F${t.row}")
        ws.cell(row=r, column=6, value=" | ".join(reasons))
        for c in range(1, len(headers2) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=(c == 6))
        ws.cell(row=r, column=6).fill = LOW_FILL
        r += 1

    widths = [22, 12, 26, 14, 22, 12, 26, 14, 10, 14, 14, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A6"
    return ws


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
    L = Selisih vs Saldo Tercatat (cross-check audit: harus 0 setiap kali
        kolom F terisi; kalau tidak, ada transaksi yang lolos/salah catat)
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
                value=f'=IF($F{row}<>"",$K{row}-$F{row},"")')

    for col, width in ((10, 26), (11, 28), (12, 22)):
        ws.column_dimensions[get_column_letter(col)].width = width


# ---------------------------------------------------------------------------
# Laporan Laba Rugi
# ---------------------------------------------------------------------------

INCOME_CATEGORIES_REVENUE = ["Penjualan"]
INCOME_CATEGORIES_EXPENSE = [
    "Belanja Bahan",
    "Belanja Operasional",
    "Belanja Konsumsi",
    "Marketing",
    "Reparasi",
    "Biaya Admin & Pajak Bank",
    "Biaya Admin dan Bunga Bank",
    "Bunga dan Admin Bank",
    "Gaji Desember 2024",
    "Gaji Desember 2025",
    "Gaji Desember 2026",
    "Gaji Desember 2027",
    "Gaji Desember 2028",
]
OTHER_CATEGORIES = ["Tip/Minus/Lebih"]


def sumif_formula(sheets_last_row, category):
    parts = []
    for sheet, last_row in sheets_last_row.items():
        parts.append(
            f"SUMIF('{sheet}'!$C$2:$C${last_row},\"{category}\",'{sheet}'!$J$2:$J${last_row})"
        )
    return "=" + "+".join(parts)


def write_income_statement(wb, sheets_last_row, period_label):
    name = "Laporan Laba Rugi"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws["A1"] = f"LAPORAN LABA RUGI - {period_label.upper()}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Semua angka adalah rumus SUMIF beralamat absolut ke seluruh sheet rekening (termasuk Kasir)."
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    r = 4
    ws.cell(row=r, column=1, value="PENDAPATAN")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    rev_rows = []
    for cat in INCOME_CATEGORIES_REVENUE:
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=2, value=sumif_formula(sheets_last_row, cat))
        rev_rows.append(r)
        r += 1
    total_rev_row = r
    ws.cell(row=r, column=1, value="Total Pendapatan")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"=SUM(B{rev_rows[0]}:B{rev_rows[-1]})")
    ws.cell(row=r, column=2).font = Font(bold=True)
    r += 2

    ws.cell(row=r, column=1, value="BEBAN")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    exp_rows = []
    for cat in INCOME_CATEGORIES_EXPENSE:
        ws.cell(row=r, column=1, value=cat)
        # beban tersimpan sebagai debit negatif -> beri tanda kurung memakai ABS lewat rumus
        ws.cell(row=r, column=2, value=sumif_formula(sheets_last_row, cat))
        exp_rows.append(r)
        r += 1
    total_exp_row = r
    ws.cell(row=r, column=1, value="Total Beban")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"=SUM(B{exp_rows[0]}:B{exp_rows[-1]})")
    ws.cell(row=r, column=2).font = Font(bold=True)
    r += 2

    ws.cell(row=r, column=1, value="LAIN-LAIN (perlu verifikasi manual)")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    other_rows = []
    for cat in OTHER_CATEGORIES:
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=2, value=sumif_formula(sheets_last_row, cat))
        other_rows.append(r)
        r += 1
    total_other_row = r
    ws.cell(row=r, column=1, value="Total Lain-lain")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"=SUM(B{other_rows[0]}:B{other_rows[-1]})" if other_rows else "=0")
    r += 2

    net_row = r
    ws.cell(row=r, column=1, value="LABA / RUGI BERSIH")
    ws.cell(row=r, column=1).font = Font(bold=True, size=12)
    ws.cell(row=r, column=2, value=f"=B{total_rev_row}+B{total_exp_row}+B{total_other_row}")
    ws.cell(row=r, column=2).font = Font(bold=True, size=12)

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 20
    for row in ws.iter_rows(min_row=4, max_row=r, min_col=2, max_col=2):
        for cell in row:
            cell.number_format = "#,##0"
    return ws, {"total_rev": total_rev_row, "total_exp": total_exp_row,
                "total_other": total_other_row, "net": net_row, "sheet": name}


# ---------------------------------------------------------------------------
# Neraca (Balance Sheet)
# ---------------------------------------------------------------------------

def write_balance_sheet(wb, sheets_last_row, opening_rows, income_ref, period_end_label):
    name = "Neraca"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws["A1"] = f"NERACA - PER {period_end_label.upper()}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Saldo kas = rumus merujuk langsung ke sel Saldo Kumulatif terakhir tiap rekening."
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    r = 4
    ws.cell(row=r, column=1, value="ASET (KAS & SETARA KAS)")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    asset_rows = []
    for sheet, last_row in sheets_last_row.items():
        ws.cell(row=r, column=1, value=sheet)
        ws.cell(row=r, column=2, value=f"='{sheet}'!$K${last_row}")
        asset_rows.append(r)
        r += 1
    total_asset_row = r
    ws.cell(row=r, column=1, value="Total Aset")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"=SUM(B{asset_rows[0]}:B{asset_rows[-1]})")
    ws.cell(row=r, column=2).font = Font(bold=True)
    r += 2

    ws.cell(row=r, column=1, value="EKUITAS")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    saldo_awal_row = r
    ws.cell(row=r, column=1, value="Saldo Awal Bulan (seluruh rekening)")
    parts = [f"'{sheet}'!$K${row}" for sheet, row in opening_rows.items()]
    ws.cell(row=r, column=2, value="=" + "+".join(parts))
    r += 1
    modal_row = r
    ws.cell(row=r, column=1, value="Modal & Setoran Pemilik (bulan ini)")
    ws.cell(row=r, column=2, value=sumif_formula(sheets_last_row, "Modal & Setoran Pemilik"))
    r += 1
    laba_row = r
    ws.cell(row=r, column=1, value="Laba Bersih Bulan Ini")
    ws.cell(row=r, column=2, value=f"='{income_ref['sheet']}'!$B${income_ref['net']}")
    r += 1
    total_equity_row = r
    ws.cell(row=r, column=1, value="Total Ekuitas")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"=B{saldo_awal_row}+B{modal_row}+B{laba_row}")
    ws.cell(row=r, column=2).font = Font(bold=True)
    r += 2

    ws.cell(row=r, column=1, value="CEK KESEIMBANGAN (Aset - Ekuitas)")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"=B{total_asset_row}-B{total_equity_row}")
    ws.cell(row=r, column=2).font = Font(bold=True)
    ws.cell(row=r, column=3,
            value='=IF(ABS(B' + str(r) + ')<1,"Balanced","Selisih - perlu telusur transfer/kasir")')
    balance_check_row = r

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 34
    for row in ws.iter_rows(min_row=4, max_row=r, min_col=2, max_col=2):
        for cell in row:
            cell.number_format = "#,##0"
    return ws, {"total_asset": total_asset_row, "total_equity": total_equity_row,
                "saldo_awal": saldo_awal_row, "balance_check": balance_check_row, "sheet": name}


# ---------------------------------------------------------------------------
# Laporan Arus Kas (Cash Flow Statement)
# ---------------------------------------------------------------------------

def write_cash_flow(wb, sheets_last_row, income_ref, balance_ref, period_label):
    name = "Laporan Arus Kas"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws["A1"] = f"LAPORAN ARUS KAS - {period_label.upper()}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Metode langsung (direct). Transfer antar rekening sendiri sengaja tidak dimasukkan karena saling menutup nol (lihat sheet Rekonsiliasi)."
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    r = 4
    ws.cell(row=r, column=1, value="ARUS KAS DARI AKTIVITAS OPERASI")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    op_row = r
    ws.cell(row=r, column=1, value="Laba Bersih Bulan Ini (basis kas)")
    ws.cell(row=r, column=2, value=f"='{income_ref['sheet']}'!$B${income_ref['net']}")
    r += 1
    total_op_row = r
    ws.cell(row=r, column=1, value="Kas Bersih dari Operasi")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"=B{op_row}")
    ws.cell(row=r, column=2).font = Font(bold=True)
    r += 2

    ws.cell(row=r, column=1, value="ARUS KAS DARI AKTIVITAS PENDANAAN")
    ws.cell(row=r, column=1).font = SECTION_FONT
    ws.cell(row=r, column=1).fill = SECTION_FILL
    r += 1
    fin_row = r
    ws.cell(row=r, column=1, value="Modal & Setoran Pemilik")
    ws.cell(row=r, column=2, value=sumif_formula(sheets_last_row, "Modal & Setoran Pemilik"))
    r += 1
    total_fin_row = r
    ws.cell(row=r, column=1, value="Kas Bersih dari Pendanaan")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"=B{fin_row}")
    ws.cell(row=r, column=2).font = Font(bold=True)
    r += 2

    net_change_row = r
    ws.cell(row=r, column=1, value="KENAIKAN (PENURUNAN) KAS BERSIH")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"=B{total_op_row}+B{total_fin_row}")
    ws.cell(row=r, column=2).font = Font(bold=True)
    r += 1
    saldo_awal_row = r
    ws.cell(row=r, column=1, value="Saldo Kas Awal Bulan")
    ws.cell(row=r, column=2, value=f"='{balance_ref['sheet']}'!$B${balance_ref['saldo_awal']}")
    r += 1
    saldo_akhir_row = r
    ws.cell(row=r, column=1, value="Saldo Kas Akhir Bulan")
    ws.cell(row=r, column=1).font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"=B{net_change_row}+B{saldo_awal_row}")
    ws.cell(row=r, column=2).font = Font(bold=True)
    r += 1
    ws.cell(row=r, column=1, value="Cek vs Total Aset di Neraca")
    ws.cell(row=r, column=2, value=f"=B{saldo_akhir_row}-'{balance_ref['sheet']}'!$B${balance_ref['total_asset']}")
    ws.cell(row=r, column=3,
            value='=IF(ABS(B' + str(r) + ')<1,"Cocok dengan Neraca","Selisih - telusur transfer belum matched")')

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 34
    for row in ws.iter_rows(min_row=4, max_row=r, min_col=2, max_col=2):
        for cell in row:
            cell.number_format = "#,##0"
    return ws


MONTHS_ID = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]


def detect_period_label(all_txns):
    """Tebak label bulan/tahun laporan dari tanggal transaksi yang paling
    sering muncul (bukan hardcode), dipakai di judul-judul laporan."""
    from collections import Counter

    counts = Counter()
    for t in all_txns:
        if isinstance(t.date, datetime.datetime):
            counts[(t.date.year, t.date.month)] += 1
    if not counts:
        return "PERIODE TIDAK TERDETEKSI"
    (year, month), _ = counts.most_common(1)[0]
    return f"{MONTHS_ID[month]} {year}"


def detect_period_end_date(all_txns, year_month_label):
    """Tanggal terakhir transaksi pada bulan yang terdeteksi, dipakai untuk
    judul Neraca ('per tanggal X')."""
    dated = [t.date for t in all_txns if isinstance(t.date, datetime.datetime)]
    if not dated:
        return ""
    last = max(dated)
    return f"{last.day} {MONTHS_ID[last.month]} {last.year}"


# ---------------------------------------------------------------------------
# Orkestrasi utama
# ---------------------------------------------------------------------------

def run_reconciliation(input_path, output_path, with_statements=False):
    """Fungsi utama: rekonsiliasi antar rekening (pencocokan transfer,
    deteksi split/merge, indikasi minus/selisih). Ini yang jalan secara
    default setiap ada file masuk.

    with_statements=True akan menambahkan 3 sheet laporan keuangan
    (Laba Rugi, Neraca, Arus Kas) - dibuat opsional supaya proses default
    tetap ringan dan fokus ke rekonsiliasi saja, sesuai kebutuhan.
    """
    wb = openpyxl.load_workbook(input_path)
    account_sheets = [s for s in wb.sheetnames if s != "Rekonsiliasi"]

    all_txns = []
    all_txns_by_sheet = {}
    sheets_last_row = {}
    opening_rows = {}
    for sname in account_sheets:
        ws = wb[sname]
        lr = last_data_row(ws)
        sheets_last_row[sname] = lr
        add_helper_column(ws, lr)
        txns = read_account_sheet(ws)
        all_txns.extend(txns)
        all_txns_by_sheet[sname] = txns
        opening = next((t for t in txns if t.is_opening), None)
        opening_rows[sname] = opening.row if opening else 2

    matches, combo_matches = find_matches(all_txns, account_sheets)
    minus_flags = find_minus_flags(all_txns_by_sheet)

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

    write_rekonsiliasi_sheet(wb, matches_section1, combo_matches, minus_flags)

    order = list(account_sheets) + ["Rekonsiliasi"]

    if with_statements:
        period_label = detect_period_label(all_txns)
        period_end_label = detect_period_end_date(all_txns, period_label)
        income_ws, income_ref = write_income_statement(wb, sheets_last_row, period_label)
        balance_ws, balance_ref = write_balance_sheet(wb, sheets_last_row, opening_rows, income_ref, period_end_label)
        write_cash_flow(wb, sheets_last_row, income_ref, balance_ref, period_label)
        order += [income_ref["sheet"], balance_ref["sheet"], "Laporan Arus Kas"]

    # urutan sheet: rekening dulu, lalu laporan
    wb._sheets = [wb[s] for s in order]

    wb.save(output_path)

    summary = {
        "n_transfer_high": sum(1 for m in matches if m.confidence == "High"),
        "n_transfer_medium": sum(1 for m in matches if m.confidence == "Medium"),
        "n_transfer_low": sum(1 for m in matches if m.confidence == "Low"),
        "n_transfer_split_merge": len(combo_matches),
        "n_transfer_unmatched": sum(1 for m in matches_section1 if m.dst is None),
        "n_minus_flags": len(minus_flags),
        "with_statements": with_statements,
    }
    return summary


if __name__ == "__main__":
    import sys
    inp = sys.argv[1] if len(sys.argv) > 1 else "Recon_Januari_2025.xlsx"
    out = sys.argv[2] if len(sys.argv) > 2 else "Recon_Januari_2025_HASIL.xlsx"
    with_stmt = "--laporan" in sys.argv
    s = run_reconciliation(inp, out, with_statements=with_stmt)
    print(s)
