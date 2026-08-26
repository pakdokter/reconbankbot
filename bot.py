"""
bot.py
Bot Telegram Rekonsiliasi Stoa Space (bot baru, terpisah dari stoabot &
bank-statement-bot).

Alur:
1. User kirim /start -> instruksi singkat.
2. User upload 1 file .xlsx gabungan (1 sheet per rekening + Kasir, format
   kolom baku seperti contoh Recon_Januari_2025.xlsx).
3. Bot jalankan reconcile.run_reconciliation() dan kirim balik file hasil
   (berisi sheet Rekonsiliasi, Laporan Laba Rugi, Neraca, Laporan Arus Kas)
   plus ringkasan angka penting di chat.

Environment variable yang dibutuhkan:
- BOT_TOKEN : token bot dari BotFather

Deploy: pola sama seperti stoabot/bank-statement-bot (Railway, git push).
"""

import logging
import os
import tempfile

from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters

from reconcile import run_reconciliation

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")

WELCOME = (
    "Halo! Kirim satu file *.xlsx* gabungan berisi sheet per rekening "
    "(BRI, BCA, Jago, Kasir, dst - format kolom seperti biasa: Tanggal, "
    "Keterangan, Kategori, Debit, Kredit, Saldo Kumulatif, Subjek, Objek, "
    "Keterangan Tambahan).\n\n"
    "Bot akan:\n"
    "1. Mencocokkan transfer antar rekening (termasuk yang terpecah/tergabung)\n"
    "2. Menandai indikasi minus/selisih kas yang perlu verifikasi manual\n"
    "3. Membuat Laporan Laba Rugi, Neraca, dan Laporan Arus Kas otomatis "
    "(rumus Excel beralamat absolut, bukan angka mati)\n\n"
    "Kirim file-nya kapan saja."
)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(WELCOME, parse_mode=ParseMode.MARKDOWN)


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc or not doc.file_name.lower().endswith(".xlsx"):
        await update.message.reply_text("Kirim file .xlsx ya, format lain belum didukung.")
        return

    status_msg = await update.message.reply_text("Memproses rekonsiliasi, tunggu sebentar...")

    with tempfile.TemporaryDirectory() as tmp:
        input_path = os.path.join(tmp, doc.file_name)
        output_path = os.path.join(tmp, doc.file_name.replace(".xlsx", "_HASIL.xlsx"))

        tg_file = await doc.get_file()
        await tg_file.download_to_drive(input_path)

        try:
            summary = run_reconciliation(input_path, output_path)
        except Exception as e:
            logger.exception("Gagal memproses file")
            await status_msg.edit_text(
                f"Gagal memproses file: {e}\n\n"
                "Cek apakah nama kolom dan urutan sheet sesuai format baku."
            )
            return

        caption = (
            "Rekonsiliasi selesai.\n\n"
            f"Transfer cocok: {summary['n_transfer_high']} high, "
            f"{summary['n_transfer_medium']} medium\n"
            f"Transfer terpecah/tergabung: {summary['n_transfer_split_merge']}\n"
            f"Transfer belum ketemu pasangan: {summary['n_transfer_unmatched']} "
            "(lihat sheet Rekonsiliasi, bagian 1)\n"
            f"Indikasi minus/selisih perlu verifikasi: {summary['n_minus_flags']}\n\n"
            "Cek sheet Laporan Laba Rugi / Neraca / Laporan Arus Kas untuk "
            "ringkasan keuangan. Baris 'CEK KESEIMBANGAN' di Neraca menandai "
            "kalau masih ada selisih yang harus ditelusuri manual."
        )

        await status_msg.delete()
        with open(output_path, "rb") as f:
            await update.message.reply_document(document=f, filename=os.path.basename(output_path), caption=caption)


def main():
    if not BOT_TOKEN:
        raise RuntimeError("Env var BOT_TOKEN belum di-set")
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    logger.info("Bot rekonsiliasi jalan...")
    app.run_polling()


if __name__ == "__main__":
    main()
