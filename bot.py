"""
bot.py
Bot Telegram Rekonsiliasi Stoa Space (bot baru, terpisah dari stoabot &
bank-statement-bot).

Fungsi UTAMA (selalu jalan tiap ada file masuk): rekonsiliasi antar rekening
saja - pencocokan transfer, deteksi transfer terpecah/tergabung, dan
indikasi minus/selisih kas yang perlu verifikasi manual.

Fungsi laporan keuangan (Laba Rugi / Neraca / Arus Kas) DIBUAT OPSIONAL,
hanya keluar kalau ditrigger:
1. Kirim file dengan caption mengandung kata "laporan" (mis. "laporan
   lengkap"), ATAU
2. Kirim file dulu (dapat hasil recon-only), lalu kirim /laporan untuk
   generate ulang file yang sama + 3 sheet laporan keuangan.

Fungsi TAMBAHAN: laporan keuangan KUARTALAN dan TAHUNAN. Trigger /kuartal
(3 file) atau /tahunan (12 file) dulu, baru upload file-file hasil
rekonsiliasi bulanan yang SUDAH completed (bulan harus berurutan).
Outputnya laporan ringkasan/kesimpulan (bukan rekonsiliasi baru - itu
sudah beres di masing-masing file bulanan).

Environment variable yang dibutuhkan:
- BOT_TOKEN : token bot dari BotFather
"""

import logging
import os
import shutil
import tempfile
import time

from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters
from telegram.error import Conflict, NetworkError

from reconcile import run_reconciliation
from quarterly import run_quarterly_report, QuarterlyInputError
from annual import run_annual_report, AnnualInputError

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")

# cache file terakhir yang diupload tiap user, supaya /laporan tidak perlu
# upload ulang. Disimpan di disk (bukan cuma di memori) karena worker bisa
# restart; dibersihkan otomatis kalau lebih tua dari CACHE_TTL.
CACHE_DIR = os.path.join(tempfile.gettempdir(), "reconbot_cache")
CACHE_TTL = 6 * 3600  # 6 jam
os.makedirs(CACHE_DIR, exist_ok=True)

# staging file buat mode /kuartal & /tahunan (N file per user, ditumpuk
# sampai lengkap). Satu direktori generik dipakai untuk kedua mode,
# dipisah per-user sekaligus per-mode di path-nya.
MULTI_DIR = os.path.join(tempfile.gettempdir(), "reconbot_multi")
os.makedirs(MULTI_DIR, exist_ok=True)

# konfigurasi tiap mode multi-file: jumlah file yang dibutuhkan, fungsi
# generator laporan, exception khusus mode itu, dan teks-teks tampilan
MULTI_MODE_CONFIG = {
    "kuartal": {
        "n_files": 3,
        "run_fn": run_quarterly_report,
        "error_cls": QuarterlyInputError,
        "label": "kuartalan",
        "output_prefix": "Laporan Kuartalan",
    },
    "tahunan": {
        "n_files": 12,
        "run_fn": run_annual_report,
        "error_cls": AnnualInputError,
        "label": "tahunan",
        "output_prefix": "Laporan Tahunan",
    },
}

LAPORAN_TRIGGER_WORDS = ["laporan", "lengkap", "statement", "financial"]

WELCOME = (
    "Halo! Kirim satu file *.xlsx* gabungan berisi sheet per rekening "
    "(BRI, BCA, Jago, Kasir, dst - format kolom: Tanggal, Keterangan, "
    "Kategori, Debit, Kredit, Saldo Kumulatif, Subjek, Objek, Keterangan "
    "Tambahan).\n\n"
    "Default: bot cuma melakukan *rekonsiliasi antar rekening* - "
    "mencocokkan transfer, deteksi transfer terpecah/tergabung, dan "
    "menandai indikasi minus/selisih kas.\n\n"
    "Kalau butuh Laporan Laba Rugi, Neraca, dan Arus Kas juga (rumus "
    "Excel beralamat absolut), ada 2 cara:\n"
    "1. Tambahkan kata *laporan* di caption waktu upload file, atau\n"
    "2. Kirim /laporan setelah upload (pakai file yang sama, tanpa upload ulang)\n\n"
    "Butuh laporan KUARTALAN (3 bulan sekaligus)? Kirim /kuartal.\n"
    "Butuh laporan TAHUNAN (12 bulan sekaligus)? Kirim /tahunan."
)


def _cache_path(user_id):
    return os.path.join(CACHE_DIR, f"{user_id}_last_upload.xlsx")


def _cleanup_cache():
    now = time.time()
    for fname in os.listdir(CACHE_DIR):
        fpath = os.path.join(CACHE_DIR, fname)
        if os.path.isfile(fpath) and now - os.path.getmtime(fpath) > CACHE_TTL:
            os.remove(fpath)


def _multi_user_dir(user_id, mode):
    return os.path.join(MULTI_DIR, mode, str(user_id))


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(WELCOME, parse_mode=ParseMode.MARKDOWN)


def build_caption(summary):
    lines = [
        "Rekonsiliasi selesai.",
        "",
        f"Transfer cocok: {summary['n_transfer_high']} high, "
        f"{summary['n_transfer_medium']} medium",
        f"Transfer terpecah/tergabung: {summary['n_transfer_split_merge']}",
        f"Transfer belum ketemu pasangan: {summary['n_transfer_unmatched']} "
        "(lihat sheet Rekonsiliasi, bagian 1)",
        f"Indikasi minus/selisih perlu verifikasi: {summary['n_minus_flags']}",
    ]
    if summary["with_statements"]:
        lines += [
            "",
            "Laporan Laba Rugi / Neraca / Arus Kas ikut dibuat. Baris "
            "'CEK KESEIMBANGAN' di Neraca menandai kalau masih ada selisih "
            "yang harus ditelusuri manual.",
        ]
    else:
        lines += [
            "",
            "Belum termasuk Laporan Laba Rugi/Neraca/Arus Kas. Kirim "
            "/laporan kalau butuh itu juga.",
        ]
    return "\n".join(lines)


def build_output_filename(summary, with_statements):
    period = summary.get("period_label") or "Periode Tidak Diketahui"
    if summary.get("no_issues"):
        base = f"Reconciliation Completed - {period}"
    else:
        base = f"Reconciliation - {period} (perlu verifikasi manual)"
    suffix = " + Laporan Keuangan" if with_statements else ""
    return f"{base}{suffix}.xlsx"


async def _process_and_reply(update, input_path, with_statements):
    status_msg = await update.message.reply_text("Memproses, tunggu sebentar...")
    with tempfile.TemporaryDirectory() as tmp:
        # nama file internal sementara - nama final ditentukan setelah tahu
        # hasil analisis (lihat build_output_filename)
        raw_output_path = os.path.join(tmp, "output.xlsx")
        try:
            summary = run_reconciliation(input_path, raw_output_path, with_statements=with_statements)
        except Exception as e:
            logger.exception("Gagal memproses file")
            await status_msg.edit_text(
                f"Gagal memproses file: {e}\n\n"
                "Cek apakah nama kolom dan urutan sheet sesuai format baku."
            )
            return
        final_filename = build_output_filename(summary, with_statements)
        await status_msg.delete()
        with open(raw_output_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=final_filename,
                caption=build_caption(summary),
            )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or not doc.file_name.lower().endswith(".xlsx"):
        await update.message.reply_text("Kirim file .xlsx ya, format lain belum didukung.")
        return

    active_mode = context.user_data.get("multi_mode")
    if active_mode:
        await handle_multi_document(update, context, doc, active_mode)
        return

    _cleanup_cache()
    user_id = update.effective_user.id
    cache_path = _cache_path(user_id)

    tg_file = await doc.get_file()
    await tg_file.download_to_drive(cache_path)  # simpan buat trigger /laporan nanti

    caption = (update.message.caption or "").lower()
    with_statements = any(w in caption for w in LAPORAN_TRIGGER_WORDS)

    await _process_and_reply(update, cache_path, with_statements)


async def laporan_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    cache_path = _cache_path(user_id)
    if not os.path.exists(cache_path):
        await update.message.reply_text(
            "Belum ada file yang diupload. Kirim file .xlsx dulu, baru /laporan."
        )
        return
    await _process_and_reply(update, cache_path, with_statements=True)


# ---------------------------------------------------------------------------
# Mode multi-file (/kuartal, /tahunan): trigger dulu, baru upload N file
# bulanan yang sudah completed (bulan harus berurutan)
# ---------------------------------------------------------------------------

async def _start_multi_mode(update: Update, context: ContextTypes.DEFAULT_TYPE, mode: str):
    cfg = MULTI_MODE_CONFIG[mode]
    user_id = update.effective_user.id
    user_dir = _multi_user_dir(user_id, mode)
    shutil.rmtree(user_dir, ignore_errors=True)
    os.makedirs(user_dir, exist_ok=True)
    context.user_data["multi_mode"] = mode
    context.user_data["multi_files"] = []
    await update.message.reply_text(
        f"Mode laporan {cfg['label']} aktif. Upload {cfg['n_files']} file hasil rekonsiliasi "
        "bulanan yang sudah *completed* (bulan harus berurutan, urutan upload bebas - nanti "
        "diurutkan otomatis).\n\n"
        "Kirim /batal kalau mau keluar dari mode ini.",
        parse_mode=ParseMode.MARKDOWN,
    )


async def kuartal_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await _start_multi_mode(update, context, "kuartal")


async def tahunan_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await _start_multi_mode(update, context, "tahunan")


async def batal_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mode = context.user_data.get("multi_mode")
    if not mode:
        await update.message.reply_text("Tidak ada proses yang bisa dibatalkan.")
        return
    user_id = update.effective_user.id
    shutil.rmtree(_multi_user_dir(user_id, mode), ignore_errors=True)
    context.user_data["multi_mode"] = None
    context.user_data["multi_files"] = []
    cfg = MULTI_MODE_CONFIG[mode]
    await update.message.reply_text(f"Mode laporan {cfg['label']} dibatalkan.")


def _reset_multi_mode(context):
    context.user_data["multi_mode"] = None
    context.user_data["multi_files"] = []


async def handle_multi_document(update: Update, context: ContextTypes.DEFAULT_TYPE, doc, mode):
    cfg = MULTI_MODE_CONFIG[mode]
    n_files = cfg["n_files"]
    user_id = update.effective_user.id
    user_dir = _multi_user_dir(user_id, mode)
    os.makedirs(user_dir, exist_ok=True)
    files = context.user_data.setdefault("multi_files", [])

    idx = len(files) + 1
    dest_path = os.path.join(user_dir, f"bulan_{idx}.xlsx")
    tg_file = await doc.get_file()
    await tg_file.download_to_drive(dest_path)
    files.append(dest_path)

    if len(files) < n_files:
        await update.message.reply_text(f"Diterima ({len(files)}/{n_files}). Upload {n_files - len(files)} file lagi.")
        return

    status_msg = await update.message.reply_text(f"{n_files} file diterima, memproses laporan {cfg['label']}...")
    with tempfile.TemporaryDirectory() as tmp:
        output_path = os.path.join(tmp, "laporan.xlsx")
        try:
            summary = cfg["run_fn"](files, output_path)
        except cfg["error_cls"] as e:
            await status_msg.edit_text(f"Gagal: {e}\n\nKirim /{mode} lagi untuk coba ulang.")
            _reset_multi_mode(context)
            shutil.rmtree(user_dir, ignore_errors=True)
            return
        except Exception as e:
            logger.exception("Gagal memproses laporan %s", cfg["label"])
            await status_msg.edit_text(
                f"Gagal memproses: {e}\n\nCek apakah semua {n_files} file itu benar hasil rekonsiliasi "
                f"bulanan yang sudah completed. Kirim /{mode} lagi untuk coba ulang."
            )
            _reset_multi_mode(context)
            shutil.rmtree(user_dir, ignore_errors=True)
            return

        periode = summary["periode"]
        final_filename = f"{cfg['output_prefix']} - {periode[0]} s.d. {periode[-1]}.xlsx"
        caption_lines = [
            f"Laporan {cfg['label']} {periode[0]} - {periode[-1]} selesai.",
            "",
            f"Jumlah pegawai di roster gaji: {summary['n_pegawai']}",
        ]
        if summary["n_belum_dibayar_bulan_terakhir"] > 0:
            caption_lines.append(
                f"{summary['n_belum_dibayar_bulan_terakhir']} pegawai belum tercatat "
                f"dibayar di bulan terakhir ({periode[-1]}) - lihat kolom Catatan di "
                "sheet Roster Gaji, kemungkinan dibebankan sebagai accrual bulan berikutnya."
            )
        await status_msg.delete()
        with open(output_path, "rb") as f:
            await update.message.reply_document(
                document=f, filename=final_filename, caption="\n".join(caption_lines)
            )

    _reset_multi_mode(context)
    shutil.rmtree(user_dir, ignore_errors=True)


async def error_handler(update, context: ContextTypes.DEFAULT_TYPE):
    """Handler error global - PTB akan spam traceback mentah ke log kalau
    tidak ada ini terdaftar. Conflict (dua instance bot rebutan getUpdates,
    biasanya sisa deploy lama yang belum mati) ditangani beda dari error
    lain: cukup dicatat singkat, PTB sendiri sudah auto-retry sampai
    instance lama berhenti - tidak perlu tindakan lebih lanjut di sini.
    Error lain dicatat lengkap supaya tetap bisa ditelusuri."""
    err = context.error
    if isinstance(err, Conflict):
        logger.warning(
            "409 Conflict: kemungkinan ada instance bot lain yang masih jalan "
            "(sisa deploy lama). PTB akan retry otomatis sampai instance lama mati."
        )
        return
    if isinstance(err, NetworkError):
        logger.warning("Network error sementara, PTB akan retry otomatis: %s", err)
        return
    logger.error("Unhandled exception saat proses update", exc_info=err)


def main():
    if not BOT_TOKEN:
        raise RuntimeError("Env var BOT_TOKEN belum di-set")
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("laporan", laporan_command))
    app.add_handler(CommandHandler("kuartal", kuartal_command))
    app.add_handler(CommandHandler("tahunan", tahunan_command))
    app.add_handler(CommandHandler("batal", batal_command))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_error_handler(error_handler)
    logger.info("Bot rekonsiliasi jalan...")
    # drop_pending_updates: buang antrean update lama saat start - supaya
    # kalau ada command yang "hilang" waktu jendela konflik 2 instance
    # (mis. saat redeploy), instance baru tidak balas telat/dobel ke pesan
    # basi, dan user tinggal kirim ulang command-nya
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
