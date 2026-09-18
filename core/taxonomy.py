"""Cache taxonomic local din exportul Catalogue of Life (ColDP).

Arhiva CoL (NameUsage.tsv, ~3 GB) se citeste STREAMING direct din zip,
fara extractie. Se pastreaza doar taxonii `accepted` cu graful
ID/parentID, rankul, numele stiintific si linhajul denormalizat.
"""
from __future__ import annotations

import re
import sqlite3
import zipfile
from pathlib import Path

APP_ROOT = Path(__file__).resolve().parent.parent
COL_DIR = APP_ROOT / "col"
TAXONOMY_DB = COL_DIR / "taxonomy.db"

I_ID = 0
I_PARENT = 4
I_STATUS = 6
I_NAME = 7
I_RANK = 9
I_GENUS = 52
I_FAMILY = 56
I_ORDER = 59
I_CLASS = 61
I_PHYLUM = 63
I_KINGDOM = 64

TREE_LEVELS = (
    ("kingdom", "Regn"),
    ("phylum", "Increngatura"),
    ("class", "Clasa"),
    ("order", "Ordinul"),
    ("family", "Familia"),
    ("genus", "Genul"),
    ("species", "Specia"),
)

CHILD_RANK = {
    "kingdom": ("phylum",),
    "phylum": ("class",),
    "class": ("order",),
    "order": ("family",),
    "family": ("genus",),
    "genus": ("species",),
}

RANK_ALIASES = {
    "kingdom": "kingdom", "regnum": "kingdom",
    "phylum": "phylum", "division": "phylum", "divisio": "phylum",
    "class": "class", "classis": "class",
    "order": "order", "ordo": "order",
    "family": "family", "familia": "family",
    "genus": "genus", "species": "species",
}

BRACKET_RE = re.compile(r"\[~(\d[\d\s.,]*)\]")


def find_col_zip():
    if not COL_DIR.is_dir():
        return None
    zips = sorted(COL_DIR.glob("*.zip"))
    return zips[0] if zips else None


def is_cache_ready(db_path=None):
    db = Path(db_path) if db_path else TAXONOMY_DB
    if not db.is_file():
        return False
    try:
        con = sqlite3.connect(str(db))
        try:
            return con.execute("SELECT COUNT(*) FROM taxa").fetchone()[0] > 0
        finally:
            con.close()
    except sqlite3.Error:
        return False


def connect_taxonomy(db_path=None):
    db = Path(db_path) if db_path else TAXONOMY_DB
    con = sqlite3.connect(str(db))
    con.row_factory = sqlite3.Row
    return con


def _norm_rank(rank):
    return RANK_ALIASES.get((rank or "").strip().lower(), "")


def build_cache(zip_path=None, db_path=None, progress=None):
    zip_path = Path(zip_path) if zip_path else find_col_zip()
    if not zip_path or not Path(zip_path).is_file():
        raise FileNotFoundError("Arhiva CoL nu a fost gasita in folderul col/.")
    db_path = Path(db_path) if db_path else TAXONOMY_DB
    db_path.parent.mkdir(parents=True, exist_ok=True)
    if db_path.is_file():
        db_path.unlink()
    con = sqlite3.connect(str(db_path))
    try:
        con.execute(
            "CREATE TABLE taxa (id TEXT PRIMARY KEY, parent_id TEXT,"
            " name TEXT NOT NULL, rank TEXT NOT NULL,"
            " lineage_genus TEXT, lineage_family TEXT, lineage_order TEXT,"
            " lineage_class TEXT, lineage_phylum TEXT, lineage_kingdom TEXT,"
            " ro_name TEXT, en_name TEXT)"
        )
        batch = []
        rows = accepted = 0
        with zipfile.ZipFile(zip_path) as zf:
            inner = next(
                (i.filename for i in zf.infolist()
                 if i.filename.endswith("NameUsage.tsv")), None)
            if not inner:
                raise ValueError("NameUsage.tsv nu exista in arhiva CoL.")
            with zf.open(inner) as fh:
                next(fh)
                for line in fh:
                    rows += 1
                    parts = line.decode("utf-8", "replace").rstrip("\n").split("\t")
                    if len(parts) <= I_KINGDOM:
                        continue
                    if parts[I_STATUS] != "accepted":
                        continue
                    rank = (parts[I_RANK] or "").strip().lower()
                    taxon_name = (parts[I_NAME] or "").strip()
                    if not taxon_name or not rank:
                        continue
                    accepted += 1
                    batch.append((
                        parts[I_ID].strip(), parts[I_PARENT].strip() or None,
                        taxon_name, rank,
                        parts[I_GENUS].strip() or None,
                        parts[I_FAMILY].strip() or None,
                        parts[I_ORDER].strip() or None,
                        parts[I_CLASS].strip() or None,
                        parts[I_PHYLUM].strip() or None,
                        parts[I_KINGDOM].strip() or None,
                    ))
                    if len(batch) >= 50000:
                        con.executemany(
                            "INSERT OR IGNORE INTO taxa (id, parent_id, name, rank,"
                            " lineage_genus, lineage_family, lineage_order,"
                            " lineage_class, lineage_phylum, lineage_kingdom)"
                            " VALUES (?,?,?,?,?,?,?,?,?,?)",
                            batch)
                        batch.clear()
                    if progress is not None and rows % 200000 == 0:
                        progress(rows, accepted)
                if batch:
                    con.executemany(
                        "INSERT OR IGNORE INTO taxa (id, parent_id, name, rank,"
                        " lineage_genus, lineage_family, lineage_order,"
                        " lineage_class, lineage_phylum, lineage_kingdom)"
                        " VALUES (?,?,?,?,?,?,?,?,?,?)",
                        batch)
        con.execute("CREATE INDEX idx_taxa_name ON taxa(name)")
        con.execute("CREATE INDEX idx_taxa_parent ON taxa(parent_id)")
        con.execute("CREATE INDEX idx_taxa_rank ON taxa(rank)")
        con.execute("CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT)")
        con.execute("INSERT INTO meta VALUES ('source','Catalogue of Life XR')")
        con.commit()
    finally:
        con.close()
    return {"rows": rows, "accepted": accepted}

def lookup_catalog_category(scientific_name, db_path=None):
    """Resolve an exact, unique species locally; never scan archives or create a DB."""
    name = (scientific_name or "").strip()
    db = Path(db_path) if db_path else TAXONOMY_DB
    if not name or not db.is_file():
        return "Alta"
    # The fast parser supplies a canonical binomial. Exact equality uses idx_taxa_name.
    name = name[0].upper() + name[1:]
    try:
        con = sqlite3.connect(db.resolve().as_uri() + "?mode=ro", uri=True)
        con.row_factory = sqlite3.Row
        try:
            rows = con.execute(
                "SELECT id, parent_id, name, rank, lineage_kingdom, lineage_class"
                " FROM taxa WHERE name = ? AND rank = 'species' LIMIT 2",
                (name,),
            ).fetchall()
            if len(rows) != 1:
                return "Alta"
            row = rows[0]
            kingdom = (row["lineage_kingdom"] or "").strip().lower()
            tax_class = (row["lineage_class"] or "").strip().lower()
            if not kingdom or (kingdom == "animalia" and not tax_class):
                chain = _walk_up(con, row)
                for level, value in (("kingdom", kingdom), ("class", tax_class)):
                    parent = _pick(chain, level)
                    if not value and parent:
                        if level == "kingdom":
                            kingdom = parent["name"].strip().lower()
                        else:
                            tax_class = parent["name"].strip().lower()
            if kingdom == "plantae":
                return "Planta"
            if kingdom == "fungi":
                return "Ciuperca"
            if kingdom == "animalia" and tax_class:
                return "Pasare" if tax_class == "aves" else "Animal"
            return "Alta"
        finally:
            con.close()
    except (sqlite3.Error, OSError):
        return "Alta"



def find_species(scientific_name, db_path=None):
    name = (scientific_name or "").strip()
    if not name:
        return None
    con = connect_taxonomy(db_path)
    try:
        row = con.execute(
            "SELECT id, parent_id, name, rank FROM taxa"
            " WHERE lower(name) = lower(?) AND rank = 'species' LIMIT 2",
            (name,)).fetchone()
        return dict(row) if row else None
    finally:
        con.close()


def _walk_up(con, row):
    chain = [dict(row)]
    seen = {row["id"]}
    cur = row["parent_id"]
    while cur and cur not in seen:
        seen.add(cur)
        prow = con.execute(
            "SELECT id, parent_id, name, rank FROM taxa WHERE id = ?",
            (cur,)).fetchone()
        if not prow:
            break
        chain.append(dict(prow))
        cur = prow["parent_id"]
        if len(chain) > 40:
            break
    chain.reverse()
    return chain


def _pick(chain, level):
    for row in chain:
        if _norm_rank(row.get("rank")) == level:
            return row
    return None

def _lineage_from_zip(species_name):
    zp = find_col_zip()
    if zp is None:
        return None
    lmap = None
    with zipfile.ZipFile(zp) as zf:
        inner = next(
            (i.filename for i in zf.infolist()
             if i.filename.endswith("NameUsage.tsv")), None)
        if inner is None:
            return None
        with zf.open(inner) as fh:
            next(fh)
            for line in fh:
                parts = line.decode("utf-8", "replace").rstrip("\n").split("\t")
                if len(parts) <= I_KINGDOM:
                    continue
                if parts[I_STATUS] != "accepted":
                    continue
                if parts[I_RANK].strip().lower() != "species":
                    continue
                if parts[I_NAME].strip().lower() != species_name.strip().lower():
                    continue
                lmap = {
                    "kingdom": parts[I_KINGDOM].strip() or None,
                    "phylum": parts[I_PHYLUM].strip() or None,
                    "class": parts[I_CLASS].strip() or None,
                    "order": parts[I_ORDER].strip() or None,
                    "family": parts[I_FAMILY].strip() or None,
                    "genus": parts[I_GENUS].strip() or None,
                }
                break
    return lmap

def lookup_chain(scientific_name, db_path=None):
    con = connect_taxonomy(db_path)
    try:
        name = (scientific_name or "").strip()
        row = con.execute(
            "SELECT id, parent_id, name, rank, ro_name, en_name FROM taxa"
            " WHERE lower(name) = lower(?) AND rank = 'species' LIMIT 2",
            (name,)).fetchone()
        if row is None:
            return None
        sp_name = row["name"]
        ro_name = row["ro_name"] if row["ro_name"] else None
        en_name = row["en_name"] if row["en_name"] else None
        chain = _walk_up(con, row)
        picked = [_pick(chain, level) for level, _lab in TREE_LEVELS[:-1]]
        if any(r is None for r in picked):
            lmap = None
        else:
            lmap = None
        if any(r is None for r in picked):
            lmap = _lineage_from_zip(sp_name)
            if not lmap:
                return None
            for i, (level, _lab) in enumerate(TREE_LEVELS[:-1]):
                if picked[i] is None and lmap[level]:
                    alt = con.execute(
                        "SELECT id, parent_id, name, rank FROM taxa"
                        " WHERE lower(name) = lower(?) AND rank = ? LIMIT 1",
                        (lmap[level], level)).fetchone()
                    if alt:
                        picked[i] = dict(alt)
                    else:
                        picked[i] = {"id": None, "parent_id": None,
                                     "name": lmap[level], "rank": level}
            if any(r is None for r in picked):
                return None
        out = []
        for i, (level, _lab) in enumerate(TREE_LEVELS[:-1]):
            out.append({"rank": level, "name": picked[i]["name"],
                        "child_count": None})
        out.append({"rank": "species", "name": sp_name, "child_count": None})
        for i, (level, _lab) in enumerate(TREE_LEVELS[:-1]):
            wanted = CHILD_RANK[level]
            marks = ",".join("?" for _ in wanted)
            col = "lineage_" + level
            cnt = con.execute(
                "SELECT COUNT(DISTINCT name) FROM taxa WHERE rank IN ("
                + marks + ") AND lower(" + col + ") = lower(?)",
                (*wanted, out[i]["name"])).fetchone()[0]
            out[i]["child_count"] = int(cnt or 0)
        return {
            "chain": out,
            "source": "Catalogue of Life",
            "ro_name": ro_name,
            "en_name": en_name,
        }
    finally:
        con.close()


def lookup_species_simple(scientific_name: str, db_path=None) -> dict | None:
    """Caută o specie și returnează date de bază (fără arbore complet).

    Returnează dict cu kingdom, ro_name, en_name, genus, species sau None.
    """
    name = (scientific_name or "").strip()
    if not name:
        return None
    con = connect_taxonomy(db_path)
    try:
        row = con.execute(
            "SELECT name, ro_name, en_name, lineage_kingdom, lineage_genus"
            " FROM taxa WHERE lower(name) = lower(?) AND rank = 'species' LIMIT 1",
            (name,),
        ).fetchone()
        if not row:
            return None
        return {
            "name": row["name"],
            "kingdom": row["lineage_kingdom"] or "",
            "ro_name": row["ro_name"] or "",
            "en_name": row["en_name"] or "",
            "genus": row["lineage_genus"] or "",
            "species": row["name"].split()[-1] if row["name"] else "",
        }
    finally:
        con.close()


def search_taxonomy_names(query: str, limit: int = 10, db_path=None) -> list[dict]:
    """Caută specii în taxonomy.db după prefixul numelui științific.

    Returnează listă de dict-uri cu cheile: name, ro_name, en_name, kingdom.
    """
    query = (query or "").strip()
    if not query:
        return []
    con = connect_taxonomy(db_path)
    try:
        rows = con.execute(
            "SELECT name, ro_name, en_name, lineage_kingdom FROM taxa"
            " WHERE rank = 'species' AND lower(name) LIKE lower(?)"
            " ORDER BY name LIMIT ?",
            (f"{query}%", limit),
        ).fetchall()
        return [
            {
                "name": r["name"],
                "ro_name": r["ro_name"] or "",
                "en_name": r["en_name"] or "",
                "kingdom": r["lineage_kingdom"] or "",
            }
            for r in rows
        ]
    finally:
        con.close()


def build_verified_context(scientific_name, db_path=None):
    found = lookup_chain(scientific_name, db_path)
    if not found:
        return ""
    bits = []

    # Denumiri verificate (daca exista in baza de date)
    ro_name = found.get("ro_name")
    en_name = found.get("en_name")
    if ro_name or en_name:
        name_line = "Denumiri verificate: "
        parts = []
        if ro_name:
            parts.append(f"româna: {ro_name}")
        if en_name:
            parts.append(f"engleza: {en_name}")
        bits.append(name_line + ", ".join(parts))

    for node in found["chain"]:
        if node["child_count"] is None:
            bits.append(node["rank"] + " " + node["name"] + " (fara numar)")
        else:
            bits.append(node["rank"] + " " + node["name"]
                        + " [~" + str(node["child_count"]) + "]")
    lines = "\n".join("- " + b for b in bits)
    return (
        "Date taxonomice verificate (Catalogue of Life - numar de taxoni "
        "copii directi, status accepted):\n" + lines + "\n"
        "Foloseste EXACT aceste cifre in [~X] si EXACT aceste nume "
        "de taxoni. Nu inventa alte cifre.")


def fix_tree_counts(response_text, scientific_name, db_path=None):
    """Inlocuieste cifrele [~X] din arborele generat cu cele verificate CoL.

    Parcurge textul documentului si inlocuieste fiecare [~X] gasit
    (in ordine: Regn, Increngatura, Clasa, Ordin, Familie, Gen) cu
    valoarea din cache-ul CoL. Daca specia nu este in cache, returneaza
    textul original neschimbat."""
    found = lookup_chain(scientific_name, db_path)
    if not found:
        return response_text

    counts = []
    for node in found["chain"]:
        if node["child_count"] is not None:
            counts.append(str(node["child_count"]))

    idx = 0

    def _repl(m):
        nonlocal idx
        if idx < len(counts):
            result = "[~" + counts[idx] + "]"
            idx += 1
            return result
        idx += 1
        return m.group(0)

    return BRACKET_RE.sub(_repl, response_text)


def verify_tree_counts(tree_text, scientific_name, db_path=None):
    found = lookup_chain(scientific_name, db_path)
    if not found:
        return [{"warning": "specia nu este in cache-ul CoL"}]
    expected = {n["rank"]: n["child_count"] for n in found["chain"]
                if n["child_count"] is not None}
    nums = []
    for m in BRACKET_RE.findall(tree_text or ""):
        try:
            nums.append(int(m.replace(" ", "").replace(".", "").replace(",", "")))
        except ValueError:
            continue
    diffs = []
    levels = [lv for lv, _ in TREE_LEVELS[:-1]]
    for i, level in enumerate(levels):
        exp = expected.get(level)
        got = nums[i] if i < len(nums) else None
        if exp is not None and got is not None and got != exp:
            diffs.append({"rank": level, "expected": exp, "found": got})
    return diffs



