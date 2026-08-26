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

Environment variable yang dibutuhkan:
- BOT_TOKEN : token bot dari BotFather
"""

import logging
import os
import tempfile
import time

from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters

from reconcile import run_reconciliation

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
    "2. Kirim /laporan setelah upload (pakai file yang sama, tanpa upload ulang)"
)


def _cache_path(user_id):
    return os.path.join(CACHE_DIR, f"{user_id}_last_upload.xlsx")


def _cleanup_cache():
    now = time.time()
    for fname in os.listdir(CACHE_DIR):
        fpath = os.path.join(CACHE_DIR, fname)
        if os.path.isfile(fpath) and now - os.path.getmtime(fpath) > CACHE_TTL:
            os.remove(fpath)


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


def main():
    if not BOT_TOKEN:
        raise RuntimeError("Env var BOT_TOKEN belum di-set")
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("laporan", laporan_command))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    logger.info("Bot rekonsiliasi jalan...")
    app.run_polling()


if __name__ == "__main__":
    main()
