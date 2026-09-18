"""Importă denumirile populare (engleză + română) din arhiva CoL în col/taxonomy.db.

Rulare:
    python tools/import_vernacular.py

Citește streaming (fără extracție) fișierul VernacularName.tsv din arhiva CoL
din col/ și umple coloanele en_name (limba "eng") și ro_name (limba "ron")
pentru taxonii existenți în taxonomy.db.

Reguli:
  - UPDATE doar pe câmpurile NULL (nu suprascrie datele manuale din Excel).
  - Denumiri multiple: dedublate case/diacritice-insensitiv, unite cu "; ".
    Prioritate: col:preferred, apoi ordinea din arhivă.
  - eng: maxim 5 denumiri distincte (acoperă 99,5% din cazuri);
    ron: toate (max observat: 11).

Durată: ~2-3 minute (streaming ~2M linii). Sursă e arhiva zip, deci
operația e re-rulabilă oricând.
"""
from __future__ import annotations

import sqlite3
import time
import zipfile
from pathlib import Path
from datetime import datetime

APP_ROOT = Path(__file__).resolve().parent.parent
DB_PATH = APP_ROOT / "col" / "taxonomy.db"

MAX_NAMES = {"eng": 5, "ron": 1000}  # ron: fără limită practică
SEP = "; "

_DIA = {"ă": "a", "â": "a", "î": "i", "ș": "s", "ş": "s", "ț": "t", "ţ": "t",
        "Ă": "a", "Â": "a", "Î": "i", "Ș": "s", "Ş": "s", "Ț": "t", "Ţ": "t"}

_FOLD_CACHE: dict[str, str] = {}


def _fold(s: str) -> str:
    """Abollah diacriticele românești și face lowercase pentru deduplificare."""
    cached = _FOLD_CACHE.get(s)
    if cached is not None:
        return cached
    folded = "".join(_DIA.get(c, c) for c in s.lower())
    _FOLD_CACHE[s] = folded
    return folded


def _best_names(rows: list[tuple[bool, str]], limit: int) -> str:
    """Dedubluare diacritice/case-insensitiv; preferred primele; unire cu SEP."""
    seen: set[str] = set()
    preferred: list[str] = []
    others: list[str] = []
    for is_pref, name in rows:
        key = _fold(name)
        if not name or key in seen:
            continue
        seen.add(key)
        (preferred if is_pref else others).append(name)
    ordered = preferred + others
    return SEP.join(ordered[:limit]) if ordered else ""


def _find_col_zip(cache_dir: Path | None = None) -> Path | None:
    col_dir = cache_dir or (APP_ROOT / "col")
    zips = sorted(col_dir.glob("*.zip"))
    return zips[0] if zips else None


def import_names(
    progress=None,
    cache_dir: Path | None = None,
    db_path: Path | None = None,
) -> dict:
    """Importă denumirile din VernacularName.tsv în taxonomy.db.

    Apelabil atât din CLI (main()), cât și din butonul GUI.
    `progress(done, total)` este opțional (done=linii citite, total=actualizați).
    Returnează un dict cu statistici.
    """
    zip_path = _find_col_zip(cache_dir)
    if not zip_path:
        raise FileNotFoundError("Nicio arhivă .zip în col/")
    target_db = db_path or DB_PATH
    if not target_db.is_file():
        raise FileNotFoundError(f"{target_db} nu există.")

    start = time.perf_counter()

    # 1) streaming din zip -> nume per taxonID
    names: dict[str, dict[str, list[tuple[bool, str]]]] = {}
    n_rows = 0
    with zipfile.ZipFile(zip_path) as z:
        inner = next((i.filename for i in z.infolist()
                      if i.filename.endswith("VernacularName.tsv")), None)
        if not inner:
            raise ValueError("VernacularName.tsv nu există în arhiva CoL.")
        with z.open(inner) as fh:
            header = fh.readline().decode("utf-8").rstrip("\n\r").split("\t")
            i_id, i_name = header.index("col:taxonID"), header.index("col:name")
            i_lang, i_pref = header.index("col:language"), header.index("col:preferred")
            for line in fh:
                n_rows += 1
                p = line.decode("utf-8").rstrip("\n\r").split("\t")
                lang = p[i_lang]
                if lang not in MAX_NAMES or len(p) <= i_pref:
                    continue
                langs = names.setdefault(p[i_id], {})
                langs.setdefault(lang, []).append((bool(p[i_pref].strip()), p[i_name]))
                if progress is not None and n_rows % 500_000 == 0:
                    progress(n_rows, 0)
    if progress is not None:
        progress(n_rows, 0)

    # 2) actualizare DB, doar câmpuri NULL
    con = sqlite3.connect(str(target_db))
    try:
        cols = {row[1] for row in con.execute("PRAGMA table_info(taxa)")}
        if "en_name" not in cols:
            con.execute("ALTER TABLE taxa ADD COLUMN en_name TEXT")
        if "ro_name" not in cols:
            con.execute("ALTER TABLE taxa ADD COLUMN ro_name TEXT")

        updated = 0
        batch: list[tuple[str | None, str | None, str]] = []
        for tid, langs in names.items():
            en = _best_names(langs.get("eng", []), MAX_NAMES["eng"])
            ro = _best_names(langs.get("ron", []), MAX_NAMES["ron"])
            if not en and not ro:
                continue
            batch.append((en or None, ro or None, tid))
            if len(batch) >= 20_000:
                updated += con.executemany(
                    """UPDATE taxa
                       SET en_name = COALESCE(en_name, ?),
                           ro_name = COALESCE(ro_name, ?)
                       WHERE id = ?""",
                    batch,
                ).rowcount
                batch.clear()
                if progress is not None:
                    progress(n_rows, updated)
        if batch:
            updated += con.executemany(
                """UPDATE taxa
                   SET en_name = COALESCE(en_name, ?),
                       ro_name = COALESCE(ro_name, ?)
                   WHERE id = ?""",
                batch,
            ).rowcount
            batch.clear()
        con.commit()

        en_count = con.execute(
            "SELECT COUNT(*) FROM taxa WHERE en_name IS NOT NULL AND en_name != ''"
        ).fetchone()[0]
        ro_count = con.execute(
            "SELECT COUNT(*) FROM taxa WHERE ro_name IS NOT NULL AND ro_name != ''"
        ).fetchone()[0]

        con.execute(
            "INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)",
            ("vernacular_imported_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        con.commit()
    finally:
        con.close()

    return {
        "rows": n_rows,
        "updated": updated,
        "en_count": en_count,
        "ro_count": ro_count,
        "seconds": time.perf_counter() - start,
    }


def main(cache_dir: Path | None = None, db_path: Path | None = None) -> int:
    try:
        stats = import_names(
            progress=lambda done, upd: print(
                f"  ... {done:,} linii, {upd:,} actualizați", flush=True),
            cache_dir=cache_dir,
            db_path=db_path,
        )
    except (FileNotFoundError, ValueError) as exc:
        print(f"EROARE: {exc}")
        return 1
    print(f"\nGATA în {stats['seconds']:.0f}s.")
    print(f"  Taxoni actualizați: {stats['updated']:,}")
    print(f"  en_name umplute în DB: {stats['en_count']:,}")
    print(f"  ro_name umplute în DB: {stats['ro_count']:,}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
