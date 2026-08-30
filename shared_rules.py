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
