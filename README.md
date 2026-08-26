# Bot Rekonsiliasi Stoa Space

Bot Telegram baru (terpisah dari stoabot dan bank-statement-bot). Terima satu
file xlsx gabungan (output bank-statement-bot: 1 sheet per rekening + Kasir),
lalu hasilkan:

- **Rekonsiliasi**: pencocokan transfer antar rekening (termasuk yang
  terpecah/digabung), plus daftar indikasi minus/selisih kas yang perlu
  verifikasi manual. Setiap baris confidence: High/Medium/Low/Needs manual
  verification, dengan alasan audit.
- **Laporan Laba Rugi**, **Neraca**, **Laporan Arus Kas**: seluruhnya rumus
  Excel (SUMIF, referensi sel) beralamat absolut ke sheet rekening asli -
  bukan angka hasil hitung Python yang ditulis mati. Kalau kamu edit angka di
  sheet rekening, laporan ikut update begitu file dibuka di Excel.

## Struktur

- `reconcile.py` - mesin inti (bisa dipakai standalone: `python reconcile.py input.xlsx output.xlsx`)
- `bot.py` - wrapper Telegram (python-telegram-bot v21, polling)
- `requirements.txt`

## Cara kerja pencocokan transfer

1. Setiap baris berkategori "Pindah Rekening Internal" / "Transfer Internal" /
   "Transfer Lainnya" dianggap kandidat transfer.
2. Rekening tujuan ditebak dari kolom Subjek/Objek (mis. "BRI-567(Biz)" ->
   sheet "BRI Bis...").
3. Dicari pasangan di rekening tujuan dalam jendela ±30 hari, toleransi
   selisih nominal 2% atau Rp5.000 (mana yang lebih besar) untuk menoleransi
   biaya admin/pembulatan.
4. Kalau tidak ketemu satu-satu, dicoba kombinasi 2 transaksi lain yang
   jumlahnya cocok (transfer terpecah/digabung, mis. setoran yang sebagian
   gagal lalu dikembalikan).
5. Yang masih tidak ketemu ditandai "Needs manual verification" beserta
   alasannya - TIDAK dipaksa cocok.

## Keterbatasan yang perlu kamu tahu

- Kolom Debit/Kredit di beberapa sheet contoh (BCA) ternyata tidak konsisten
  posisinya (nilai positif kadang muncul di kolom Debit). Kolom bantu
  "Nominal Bersih" (J) menangani ini pakai `N($D)+N($E)`, jadi aman selama
  tanda (+/-) di data sumber sudah benar per baris.
- Kolom "Saldo Kumulatif" di sumber data tidak selalu terisi berurutan per
  baris (beberapa transaksi bertanggal sama dikelompokkan). Karena itu saldo
  akhir tiap rekening di Neraca dihitung ulang lewat kolom bantu K (Saldo
  Kumulatif Rekonstruksi = saldo awal + akumulasi berjalan), bukan langsung
  ambil sel terakhir kolom F.
- Baris "CEK KESEIMBANGAN" di Neraca dan "Cek vs Total Aset" di Laporan Arus
  Kas akan menunjukkan selisih kalau masih ada transfer/kasir yang belum
  matched - ini fitur, bukan bug, supaya tidak ada yang "didiamkan".

## Deploy ke Railway

1. Buat bot baru lewat BotFather, catat token-nya.
2. Buat project baru di Railway, push repo ini (Procfile sudah ada:
   `worker: python bot.py`).
3. Set env var `BOT_TOKEN` di Railway ke token dari langkah 1.
4. Railway otomatis install dari `requirements.txt` dan jalankan Procfile.

## Belum dikerjakan (langkah selanjutnya)

- ~~Auto-detect nama bulan/tahun dari data~~ - selesai, judul laporan
  sekarang ambil dari tanggal transaksi paling banyak muncul di data
  (`detect_period_label`), bukan hardcode.
- Simpan histori (Postgres) supaya bisa lintas bulan (deteksi delayed
  settlement yang baru settle bulan berikutnya, sesuai prinsip audit
  "trace beyond month boundaries" - saat ini tracing hanya sebatas data
  yang ada di satu file yang diupload).
- Belum ada testing dengan format sheet lain di luar contoh yang dikasih
  (mis. kalau ada rekening pinjaman/loan beneran, atau modal keluar/prive).
