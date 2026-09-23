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

def lookup_catalog_group(scientific_name, db_path=None):
    """Resolve an exact, unique species locally; never scan archives or create a DB."""
    name = (scientific_name or "").strip()
    db = Path(db_path) if db_path else TAXONOMY_DB
    if not name or not db.is_file():
        return "Other"
    # The fast parser supplies a canonical binomial. Exact equality uses idx_taxa_name.
    name = name[0].upper() + name[1:]
    try:
        con = sqlite3.connect(db.resolve().as_uri() + "?mode=ro", uri=True)
        con.row_factory = sqlite3.Row
        try:
            # Only the name belongs in the WHERE clause. Adding `rank = 'species'`
            # makes SQLite prefer idx_taxa_rank and scan millions of rows on every
            # call (measured ~0.75 s); idx_taxa_name answers in under a millisecond.
            rows = con.execute(
                "SELECT id, parent_id, name, rank, lineage_kingdom, lineage_class"
                " FROM taxa WHERE name = ? LIMIT 6",
                (name,),
            ).fetchall()
            species = [r for r in rows
                       if str(r["rank"] or "").strip().lower() == "species"]
            if len(species) != 1:
                return "Other"
            row = species[0]
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
                return "Plant"
            if kingdom == "fungi":
                return "Mushroom"
            if kingdom == "animalia":
                if not tax_class:
                    return "Animal"
                if tax_class == "aves":
                    return "Bird"
                if tax_class == "mammalia":
                    return "Mammal"
                if tax_class == "reptilia":
                    return "Reptile"
                if tax_class == "amphibia":
                    return "Amphibian"
                if tax_class == "insecta":
                    return "Insect"
                if tax_class == "arachnida":
                    return "Arachnid"
                if tax_class in ("gastropoda", "bivalvia", "cephalopoda",
                                 "polyplacophora", "scaphopoda"):
                    return "Mollusc"
                if tax_class in ("teleostei", "chondrichthyes", "actinopterygii",
                                 "elasmobranchii", "myxini", "petromyzontida",
                                 "sarcopterygii", "dipnoi", "coelacanthiformes",
                                 "chondrostei", "holostei"):
                    return "Fish"
                if tax_class in ("malacostraca", "diplopoda",
                                 "chilopoda", "clitellata", "hirudinea",
                                 "polychaeta", "echinoidea", "asteroidea"):
                    return "Animal"
                return "Animal"
            return "Other"
        finally:
            con.close()
    except (sqlite3.Error, OSError):
        return "Other"


def lookup_catalog_category(scientific_name, db_path=None):
    """Backward-compatible alias returning the canonical catalog group."""
    return lookup_catalog_group(scientific_name, db_path=db_path)



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



# ------------------------------------------------------- CoL validation
# Read-only validation of identified names against the CoL snapshot, with
# graduated suggestions. Spec: tools/local_dev/test_validate_taxonomy.py;
# algorithm: tools/local_dev/_probe_validate.py (timings <=1s per query —
# LIKE patterns are interpolated as literals so SQLite keeps its range
# optimization; bound parameters would defeat it).

_VAL_CAMEL_RE = re.compile(r"(?<=[a-z])(?=[A-Z])")

# category -> (kingdom, class) anchors used when the genus row is missing
_VAL_ANCHOR = {
    "plant": ("Plantae", ""),
    "bird": ("Animalia", "Aves"),
    "mammal": ("Animalia", "Mammalia"),
    "insect": ("Animalia", "Insecta"),
    "arachnid": ("Animalia", "Arachnida"),
    "mollusc": ("Animalia", "Mollusca"),
    "fish": ("Animalia", ""),
}


def _val_lev(a, b):
    """Edit distance (Levenshtein), two-row implementation."""
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _val_nkey(candidate, query_lower):
    """Normalized similarity key: lev / max(len) — raw lev favors short
    genera (validated probe finding: 'Potentilla indica' was #124/157)."""
    c = candidate.lower()
    return _val_lev(query_lower, c) / max(len(query_lower), len(c), 1)


def _val_like(pattern):
    """SQL-literal LIKE pattern (escaped), never a bound parameter."""
    return pattern.replace("'", "''")


def _val_split(name):
    """(genus, epithet) of a binomial; camel-split for 'Genusspecies' slips."""
    parts = name.split()
    if len(parts) >= 2:
        return parts[0], parts[-1]
    bits = _VAL_CAMEL_RE.split(name)
    if len(bits) == 2:
        return bits[0], bits[1]
    return name, None


# Genus reassignments confirmed by taxonomic revision — our own local
# knowledge (NOT part of CoL): sunk genus -> currently accepted genus.
# Kept as a plain dict so taxonomy.db remains byte-identical; extend here
# when a new sunk genus is identified by validation.
_VAL_ALIAS_GENUS = {
    "duchesnea": "Potentilla",
    "lavatera": "Malva",
    "pennisetum": "Cenchrus",
    "calystegia": "Convolvulus",
}


def validate_species(species, category=None, db_path=None):
    """Valideaza o specie identificata de model impotriva snapshot-ului CoL.

    Read-only (PRAGMA query_only). Returneaza dict cu: status ('exact' |
    'genus_ok' | 'not_in_col'), genus, epithet, member_count, members,
    closest_in_genus, accepted_genus (ex. Calystegia -> Convolvulus),
    accepted ({name, family, class} | None) — tinta de reincadrare, si
    closest (top 5 alternative). Sugestiile sunt propuneri pentru
    confirmarea umana (Update Species); taxonomy.db nu este modificat.
    Spec: tools/local_dev/test_validate_taxonomy.py.
    """
    con = connect_taxonomy(db_path)
    try:
        con.execute("PRAGMA query_only = ON")
        result = {
            "status": "not_in_col",
            "genus": "",
            "epithet": None,
            "member_count": 0,
            "members": [],
            "closest_in_genus": None,
            "accepted_genus": None,
            "accepted": None,
            "closest": [],
        }
        name = (species or "").strip()
        if not name:
            return result

        # tier 0: exact species (binary compare hits idx_taxa_name; the
        # catalog enforces canonical case at ingestion, LIKE is only the
        # fallback for manually typed names).
        cols = ("lineage_genus, lineage_family, lineage_order, lineage_class,"
                " lineage_phylum, lineage_kingdom")
        row = con.execute(
            f"SELECT {cols} FROM taxa WHERE name = ? AND rank='species' LIMIT 1",
            (name,),
        ).fetchone()
        if row is None:
            row = con.execute(
                f"SELECT {cols} FROM taxa WHERE name LIKE '{_val_like(name)}'"
                " AND rank='species' LIMIT 1",
            ).fetchone()
        if row is not None:
            genus, epithet = _val_split(name)
            result["status"] = "exact"
            result["genus"] = genus
            result["epithet"] = epithet
            return result

        # split the queried name; recover genus prefix for opaque slips
        genus, epithet = _val_split(name)
        if epithet is None and " " not in name:
            pref = con.execute(
                "SELECT name FROM taxa WHERE rank='genus'"
                " AND lower(?) LIKE lower(name || '%')"
                " ORDER BY length(name) DESC LIMIT 1",
                (name,),
            ).fetchone()
            if pref:
                genus = pref[0]
                epithet = name[len(genus):].strip() or None

        # anchor row: the queried genus, if CoL still accepts it
        gen_row = con.execute(
            "SELECT lineage_class, lineage_family FROM taxa"
            " WHERE name = ? AND rank='genus' LIMIT 1",
            (genus,),
        ).fetchone()

        result["genus"] = genus
        result["epithet"] = epithet
        cls = (gen_row[0] if gen_row else None) or ""
        fam = (gen_row[1] if gen_row else None) or ""
        result["status"] = "genus_ok" if gen_row is not None else "not_in_col"

        # members of the (existing) genus — one scan feeds count, list and
        # the closest member
        if gen_row is not None:
            members_all = [
                r[0] for r in con.execute(
                    f"SELECT name FROM taxa WHERE name LIKE"
                    f" '{_val_like(genus)} %' AND rank='species'",
                ).fetchall()
            ]
            members_all.sort(key=str.lower)
            result["member_count"] = len(members_all)
            result["members"] = members_all[:25]
            if members_all and epithet:
                query_lower = name.lower()
                result["closest_in_genus"] = min(
                    members_all, key=lambda m: _val_nkey(m, query_lower)
                )

        # tier F: infraspecific fragments of the full binomial
        # ('Calystegia sepium subsp. ...') -> lineage_genus gives the
        # accepted genus; tier F2: fragments under a genus row that is
        # gone ('Duchesnea ...' surviving under an old combination).
        if epithet:
            frag = con.execute(
                "SELECT lineage_genus FROM taxa"
                f" WHERE name LIKE '{_val_like(name)} %'"
                " AND lineage_genus IS NOT NULL LIMIT 5",
            ).fetchall()
            if gen_row is None:
                frag += con.execute(
                    "SELECT lineage_genus FROM taxa"
                    f" WHERE name LIKE '{_val_like(genus)} %'"
                    " AND lineage_genus IS NOT NULL"
                    f" AND lower(lineage_genus) != lower('{_val_like(genus)}')"
                    " LIMIT 5",
                ).fetchall()
            for (lin,) in frag:
                lin = (lin or "").strip()
                if lin and lin.lower() != genus.lower():
                    result["accepted_genus"] = lin
                    break

        # Genus reassignments confirmed by taxonomic revision — local
        # knowledge, never written into taxonomy.db (CoL stays pristine).
        if result["accepted_genus"] is None and gen_row is None and genus:
            result["accepted_genus"] = _VAL_ALIAS_GENUS.get(genus.lower())

        # Same epithet inside the accepted genus beats any string-distance
        # guess: Duchesnea indica -> Potentilla indica is synonymy, which
        # Levenshtein ranking cannot see. All 4 target rows verified to
        # exist in the snapshot (Convolvulus sepium, Potentilla indica,
        # Cenchrus alopecuroides, Malva trimestris).
        confirmed = None
        if epithet and result["accepted_genus"]:
            agr = result["accepted_genus"]
            confirmed = con.execute(
                "SELECT name, lineage_family, lineage_class FROM taxa"
                " WHERE name = ? AND rank='species' LIMIT 1",
                (f"{agr} {epithet}",),
            ).fetchone()
            if confirmed is None:
                confirmed = con.execute(
                    "SELECT name, lineage_family, lineage_class FROM taxa"
                    f" WHERE name LIKE '{_val_like(agr)} %{_val_like(epithet)}'"
                    " AND rank='species' LIMIT 1",
                ).fetchone()
            if confirmed is not None:
                result["accepted"] = {
                    "name": confirmed[0],
                    "family": confirmed[1],
                    "class": confirmed[2],
                }

        # Epithet match across the whole snapshot, filtered by the anchors
        # we have — genus lineage (class/family) first, then the identified
        # category (class/kingdom) — ranked by normalized Levenshtein.
        # Proposals only: taxonomy.db is never modified.
        if epithet:
            hits = con.execute(
                "SELECT name, lineage_class, lineage_family, lineage_kingdom"
                " FROM taxa"
                f" WHERE name LIKE '% {_val_like(epithet)}'"
                " AND rank='species' LIMIT 3000",
            ).fetchall()
            knd, cls_cat = _VAL_ANCHOR.get(
                (category or "").strip().lower(), ("", "")
            )
            tiers = []
            if fam:
                tiers.append([h for h in hits if (h[2] or "") == fam])
            if cls:
                tiers.append([h for h in hits if (h[1] or "") == cls])
            if cls_cat:
                tiers.append([h for h in hits if (h[1] or "") == cls_cat])
            if knd:
                tiers.append([h for h in hits if (h[3] or "") == knd])
            tiers.append(list(hits))
            seen = set()
            if confirmed is not None:
                seen.add(confirmed[0])
                result["closest"].append(
                    {"name": confirmed[0],
                     "family": confirmed[1], "class": confirmed[2]}
                )
            # Known accepted genus but no confirmable species inside it:
            # propose the genus only — never present a cross-family
            # string guess as "the accepted species".
            pin_genus = (
                result["accepted_genus"] is not None and confirmed is None
            )
            for tier in tiers:
                ranked = sorted(
                    tier, key=lambda h: _val_nkey(h[0], name.lower())
                )
                if ranked and result["accepted"] is None and not pin_genus:
                    best = ranked[0]
                    result["accepted"] = {
                        "name": best[0], "family": best[2], "class": best[1],
                    }
                for hit in ranked:
                    if hit[0] in seen:
                        continue
                    seen.add(hit[0])
                    result["closest"].append(
                        {"name": hit[0], "family": hit[2], "class": hit[1]}
                    )
                    if len(result["closest"]) >= 5:
                        break
                if len(result["closest"]) >= 5 and result["accepted"]:
                    break
        return result
    finally:
        con.close()


