"""
Suressă unică pentru structura bazei de date (catalog.db).

Toată crearea de tabele, migrările şi backfill-urile se petrec AICI,
în funcția connect(). Nici un alt module nu mai defineşte schema;
core.catalog.py şi core.prompts.py folosesc connect() din acest module.
"""

from pathlib import Path
import sqlite3


DEFAULT_DATABASE_NAME = "catalog.db"

CATALOG_ITEMS_TABLE_DDL = """
CREATE TABLE IF NOT EXISTS catalog_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    image_hash TEXT NOT NULL UNIQUE,
    analysis_key TEXT,
    original_path TEXT NOT NULL,
    catalog_path TEXT NOT NULL,
    response_text TEXT NOT NULL DEFAULT '',
    thinking_text TEXT NOT NULL DEFAULT '',
    prompt_text TEXT NOT NULL DEFAULT '',
    preset_name TEXT NOT NULL DEFAULT '',
    manual_location TEXT NOT NULL DEFAULT '',
    metadata_json TEXT NOT NULL DEFAULT '{}',
    statistics_json TEXT NOT NULL DEFAULT '{}',
    category TEXT NOT NULL,
    kingdom TEXT,
    genus TEXT,
    species TEXT,
    scientific_name TEXT,
    ro_name TEXT,
    en_name TEXT,
    confidence REAL,
    is_uncertain INTEGER,
    extra_fields_json TEXT NOT NULL,
    manual_override INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
)
"""

SPECIES_PROFILES_TABLE_DDL = """
CREATE TABLE IF NOT EXISTS species_profiles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    scientific_name TEXT NOT NULL UNIQUE,
    category TEXT,
    kingdom TEXT,
    genus TEXT,
    species TEXT,
    ro_name TEXT,
    en_name TEXT,
    profile_response TEXT NOT NULL DEFAULT '',
    profile_fields_json TEXT NOT NULL DEFAULT '{}',
    extra_fields_json TEXT NOT NULL DEFAULT '{}',
    verified_manually INTEGER NOT NULL DEFAULT 0,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
)
"""

PROMPTS_TABLE_DDL = """
CREATE TABLE IF NOT EXISTS prompts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    content TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
)
"""

# Migrările pentru bazele de date vecthi (coloane lipsă din versiuni anterioare).
CATALOG_ITEMS_COLUMN_MIGRATIONS = {
    "response_text": "TEXT NOT NULL DEFAULT ''",
    "analysis_key": "TEXT",
    "manual_override": "INTEGER NOT NULL DEFAULT 0",
    "thinking_text": "TEXT NOT NULL DEFAULT ''",
    "prompt_text": "TEXT NOT NULL DEFAULT ''",
    "preset_name": "TEXT NOT NULL DEFAULT ''",
    "manual_location": "TEXT NOT NULL DEFAULT ''",
    "metadata_json": "TEXT NOT NULL DEFAULT '{}'",
    "statistics_json": "TEXT NOT NULL DEFAULT '{}'",
}

SPECIES_PROFILES_COLUMN_MIGRATIONS = {
    "profile_fields_json": "TEXT NOT NULL DEFAULT '{}'",
}


def database_path_for(output_dir: Path | str) -> Path:
    """
    Returnează calea către fişierul bazei de date în directorul output.
    """
    return Path(output_dir) / DEFAULT_DATABASE_NAME


def connect(database_path: Path | str) -> sqlite3.Connection:
    """
    Deschide (creeând la nevoie) baza de date şi asurează că întrega
    structură există: tabelele catalog_items, species_profiles şi prompts,
    plus toate migrările şi backfill-urile pentru bazele vector.

    Acestă este singurul loc unde schema bazei de date este definită.
    """
    database_path = Path(database_path)
    database_path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(database_path)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")

    connection.execute(CATALOG_ITEMS_TABLE_DDL)
    connection.execute(SPECIES_PROFILES_TABLE_DDL)
    connection.execute(PROMPTS_TABLE_DDL)

    existing_columns = {
        row[1] for row in connection.execute(
            "PRAGMA table_info(catalog_items)"
        )
    }
    for column, definition in CATALOG_ITEMS_COLUMN_MIGRATIONS.items():
        if column not in existing_columns:
            connection.execute(
                f"ALTER TABLE catalog_items ADD COLUMN {column} {definition}"
            )

    existing_sp_columns = {
        row[1] for row in connection.execute(
            "PRAGMA table_info(species_profiles)"
        )
    }
    for column, definition in SPECIES_PROFILES_COLUMN_MIGRATIONS.items():
        if column not in existing_sp_columns:
            connection.execute(
                f"ALTER TABLE species_profiles ADD COLUMN {column} {definition}"
            )

    connection.commit()
    return connection