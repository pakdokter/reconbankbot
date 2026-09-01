"""Data kategori/kata kunci/alias bersama untuk reconbot dan bot-bot Stoa
lainnya (bank-statement-bot, stoabot). Urutan sumber (yang pertama
tersedia dipakai): Postgres (DATABASE_URL, tabel shared_rules) -> file
JSON lokal (shared_rules.json) -> {} (pemanggil fallback ke default
hardcode masing-masing).

Skema tabel Postgres (lihat migrate.sql):
    CREATE TABLE shared_rules (
        key TEXT PRIMARY KEY,
        value JSONB NOT NULL,
        updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
Satu baris per top-level key (category_override_rules, employee_aliases,
dst) - value-nya persis struktur yang sama dengan shared_rules.json.
"""

import json
import os

_JSON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shared_rules.json")
_cache = None


def _load_from_postgres():
    dsn = os.environ.get("DATABASE_URL")
    if not dsn:
        return None
    try:
        import psycopg2
    except ImportError:
        return None
    try:
        conn = psycopg2.connect(dsn, connect_timeout=5)
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT key, value FROM shared_rules")
                rows = cur.fetchall()
        finally:
            conn.close()
    except Exception:
        return None
    return {key: value for key, value in rows}


def _load_from_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def load(path=None):
    global _cache
    if path is None and _cache is not None:
        return _cache

    data = _load_from_postgres()
    if data is None:
        data = _load_from_json(path or _JSON_PATH)
    if data is None:
        data = {}

    if path is None:
        _cache = data
    return data


def get(key, default):
    return load().get(key, default)


def push_json_to_postgres(json_path=None):
    """Utilitas satu-kali: baca shared_rules.json, tulis tiap top-level
    key sebagai baris ke tabel Postgres shared_rules (upsert). Dipakai
    untuk bootstrap tabel dari file JSON yang sudah ada. Butuh
    DATABASE_URL dan tabel shared_rules sudah dibuat (lihat migrate.sql)."""
    import psycopg2

    dsn = os.environ["DATABASE_URL"]
    data = _load_from_json(json_path or _JSON_PATH)
    if data is None:
        raise FileNotFoundError(json_path or _JSON_PATH)
    data.pop("_meta", None)

    conn = psycopg2.connect(dsn)
    try:
        with conn.cursor() as cur:
            for key, value in data.items():
                cur.execute(
                    """
                    INSERT INTO shared_rules (key, value, updated_at)
                    VALUES (%s, %s, now())
                    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = now()
                    """,
                    (key, json.dumps(value)),
                )
        conn.commit()
    finally:
        conn.close()


def _connect():
    import psycopg2
    return psycopg2.connect(os.environ["DATABASE_URL"], connect_timeout=5)


def _read_row(cur, key):
    cur.execute("SELECT value FROM shared_rules WHERE key = %s", (key,))
    row = cur.fetchone()
    return row[0] if row else None


def _write_row(cur, key, value):
    cur.execute(
        """
        INSERT INTO shared_rules (key, value, updated_at)
        VALUES (%s, %s, now())
        ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = now()
        """,
        (key, json.dumps(value)),
    )


def add_category_rule(keywords, category, valid_categories, sheet_contains=None):
    """Tambah satu aturan baru (kata kunci apapun di antara `keywords`
    -> `category`) ke row 'category_override_rules' di Postgres. Rule
    baru ditambahkan di UJUNG list (prioritas paling rendah - kalau ada
    kata kunci yang sama tumpang tindih dengan rule lain yang sudah ada,
    rule lama tetap menang). `category` DIVALIDASI dulu terhadap
    `valid_categories` - kalau tidak dikenal, raise ValueError (mencegah
    kelas bug 'Marketing & RnD' yang pernah terjadi: kategori tujuan yang
    salah tulis bikin uang hilang tanpa error apapun). global _cache
    direset supaya proses berikutnya baca ulang dari Postgres, bukan
    cache lama."""
    global _cache
    if category not in valid_categories:
        raise ValueError(
            f"Kategori {category!r} tidak dikenal. Pilihan yang valid: {sorted(valid_categories)}"
        )
    conn = _connect()
    try:
        with conn.cursor() as cur:
            rules = _read_row(cur, "category_override_rules") or []
            new_rule = {"any": list(keywords), "category": category, "sheet_contains": sheet_contains}
            rules.append(new_rule)
            _write_row(cur, "category_override_rules", rules)
        conn.commit()
    finally:
        conn.close()
    _cache = None
    return new_rule


def add_employee_alias(short_name, full_name):
    """Tambah/timpa satu alias pegawai (short_name -> full_name) di row
    'employee_aliases' di Postgres."""
    global _cache
    conn = _connect()
    try:
        with conn.cursor() as cur:
            aliases = _read_row(cur, "employee_aliases") or {}
            aliases[short_name.strip().lower()] = full_name.strip()
            _write_row(cur, "employee_aliases", aliases)
        conn.commit()
    finally:
        conn.close()
    _cache = None


def list_category_rules():
    conn = _connect()
    try:
        with conn.cursor() as cur:
            return _read_row(cur, "category_override_rules") or []
    finally:
        conn.close()


def list_employee_aliases():
    conn = _connect()
    try:
        with conn.cursor() as cur:
            return _read_row(cur, "employee_aliases") or {}
    finally:
        conn.close()


def remove_category_rule(index):
    """Hapus aturan ke-`index` (0-based, urutan sama seperti /lihataturan)
    dari row 'category_override_rules'."""
    global _cache
    conn = _connect()
    try:
        with conn.cursor() as cur:
            rules = _read_row(cur, "category_override_rules") or []
            if not (0 <= index < len(rules)):
                raise IndexError(f"Index {index} di luar jangkauan (ada {len(rules)} aturan)")
            removed = rules.pop(index)
            _write_row(cur, "category_override_rules", rules)
        conn.commit()
    finally:
        conn.close()
    _cache = None
    return removed


def diagnose_connection_error(exc):
    """Terjemahkan exception koneksi/query Postgres jadi pesan diagnosis
    yang lebih spesifik - dipakai bot.py supaya user langsung tahu apa
    yang perlu dicek, bukan cuma 'gagal terhubung' generik."""
    if isinstance(exc, KeyError) and "DATABASE_URL" in str(exc):
        return "DATABASE_URL belum diset di environment variable reconbot (cek Railway -> project reconbot -> Variables)."
    text = str(exc).lower()
    if "password" in text or "authentication" in text:
        return "Autentikasi Postgres gagal - DATABASE_URL kemungkinan salah atau kredensial sudah berubah (cek connection string terbaru di Railway -> project stoabot -> Postgres -> Variables)."
    if "does not exist" in text and ("relation" in text or "table" in text):
        return "Tabel 'shared_rules' belum dibuat di database ini - jalankan seed_shared_rules.sql dulu di Query tab Postgres."
    if "timeout" in text or "timed out" in text or "could not connect" in text or "connection refused" in text or "could not translate host" in text:
        return "Tidak bisa konek ke Postgres (timeout/refused/host tidak ditemukan) - cek DATABASE_URL benar dan Postgres masih aktif di Railway."
    if "no module named" in text and "psycopg2" in text:
        return "Modul psycopg2 belum terpasang di deployment - cek requirements.txt sudah ter-push ke repo dan reconbot sudah di-redeploy."
    return f"{type(exc).__name__}: {exc}"
