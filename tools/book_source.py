"""Extrage denumirile populare românești dintr-o carte/dicționar PDF.

Rulare din rădăcina proiectului:
    python -m tools.book_source "cale/carte.pdf"

Citește PDF-ul din name_sources/, parsează intrările dicționarului (nume științific +
denumire populară principală, fără variantele din „Alte numiri:"), le potrivește cu
speciile din col/taxonomy.db și generează denumiri_draft.xlsx:

  Sheet „Draft"      — intrări potrivite: A=nume științific (din DB),
                       B=ro_name propus, C=en_name (din DB)
  Sheet „Nepotrivite"— intrări din PDF care nu s-au potrivit cu DB
                       (nomenclatura din 1968 diferă de cea actuală)

Fișierul se revizuiește manual, apoi se importă prin butonul
„Adaugă denumiri din Excel”. În distribuția cu utilitare de dezvoltare se poate
folosi și: python tools/local_dev/add_common_names.py denumiri_draft.xlsx.
"""
import re
import unicodedata
from pathlib import Path

import openpyxl
import sqlite3

APP_ROOT = Path(__file__).resolve().parent.parent
DB_PATH = APP_ROOT / "col" / "taxonomy.db"
OUT_PATH = APP_ROOT / "tools" / "denumiri_draft.xlsx"

# ---------------------------------------------------------------- normalizare

_ORTHO = {"ă": "a", "â": "i", "î": "i", "ș": "s", "ş": "s",
          "ț": "t", "ţ": "t", "Ă": "a", "Â": "i", "Î": "i",
          "Ș": "s", "Ş": "s", "Ț": "t", "Ţ": "t"}


def _skel(s: str) -> str:
    """Schelet ortografic: minuscule, diacritice eliminate (și ortografia 1968:
    â/î -> i), doar litere. „rîuri"/„râuri" -> „riuri"."""
    s = "".join(_ORTHO.get(c, c) for c in s.lower())
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z]", "", s)


# ------------------------------------------------------------------- parsing
# Reguli configurabile (încărcate din tools/parse_rules.json, editabil din GUI)

RULES_PATH = APP_ROOT / "tools" / "parse_rules.json"

_DEFAULT_RULES = {
    "entry": {
        "pattern": r"^([A-ZĂÂÎȘȚŞŢ](?:[A-ZĂÂÎȘȚŞŢ\-]+[a-zăâîșțşţ]?|[a-zăâîșțşţ\-]{2,}))\s+"
                   r"([a-zăâîșțşţ][a-zăâîșțşţ\-]{2,})\b(.*)$",
        "desc": "Detectează linia de antet a unei intrări (gen + specie + rest)"
    },
    "synonym": {
        "pattern": r"\(([A-ZĂÂÎȘȚŞŢ][A-Za-zăâîșțşţ\-]+)\s+([a-zăâîșțşţ][a-zăâîșțşţ\-]+)",
        "desc": "Extrage sinonimele din paranteze (ex: (Prunus armeniaca L.))"
    },
    "citation": {
        "pattern": r"\s*\([^()]*\)\s*",
        "desc": "Elimină citările din denumiri (ex: (180, 603))"
    },
    "stop_line": {
        "pattern": r"^\s*(Magh|Franc|Eng|Germ|Ung|Slov|Ceh|Bul|Serb|Ucr|Rus|Ital|Esp|"
                   r"Neerl|Pol|Dan|Sued|Fam)\b\s*[.:]",
        "flags": ["IGNORECASE"],
        "desc": "Secțiuni străine care termină zona de denumiri românești"
    },
    "alte_numiri": {
        "pattern": r"A\s*l\s*t\s*e\s*n\s*u\s*(?:m|n)\s*i\s*r\s*i\s*:",
        "flags": ["IGNORECASE"],
        "desc": "Oprește extragerea denumirilor înainte de «Alte numiri:» (tolerant la spațieri OCR: «Alte n uniri:»)"
    },
    "fam": {
        "pattern": r"F\s*[ao]\s*(?:m|in)\b\.?",
        "flags": ["IGNORECASE"],
        "desc": "Oprește extragerea denumirilor înainte de «Fam.» (tolerant OCR: «Fain.», «Fom»)"
    },
    "genus_header": {
        "pattern": r"^[A-ZĂÂÎȘȚŞŢ]{3,}\s+[A-ZĂÂÎȘȚŞŢ]",
        "desc": "Antete de GEN (ex: «CONVOLVULUS L.», «EPIPACTIS Zinn.») — nu sunt specii, nu intră în denumiri"
    },
    "dash": {
        "pattern": r"(?:[—–]|\s-\s)\s*(.+)$",
        "flags": ["DOTALL"],
        "desc": "Separatorul — între nume științific și denumirea populară"
    },
    "has_letter": {
        "pattern": r"[A-Za-zăâîșțşţ]",
        "desc": "Verifică dacă textul conține litere (filtru anti-zgomot)"
    },
    "alpha_tab": {
        "pattern": r"[A-ZĂÂÎȘȚŞŢ]{2,4}",
        "desc": "Taburi alfabetice de navigare (BER, DOR...) — se ignoră"
    },
    "page_number": {
        "pattern": r"\d{1,3}",
        "desc": "Numere de pagină — se ignoră"
    },
    "valid_name_noise": {
        "pattern": r"\)|\(|jud\.|Herb\.|exsicc|Trans\.|Mold\.|Banat\)|folosește",
        "desc": "Zgomot OCR care invalidează o denumire"
    }
}


def load_rules() -> dict:
    """Încarcă regulile de parsare din JSON. Creează fișierul implicit dacă lipsește."""
    import json
    if not RULES_PATH.is_file():
        RULES_PATH.write_text(
            json.dumps(_DEFAULT_RULES, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8")
    return json.loads(RULES_PATH.read_text(encoding="utf-8"))


def _compile_rules(rules: dict) -> dict:
    """Compilează pattern-urile din reguli. Returnează dict de regex-uri."""
    compiled = {}
    for key, rule in rules.items():
        flags = 0
        for f in rule.get("flags", []):
            flags |= getattr(re, f)
        compiled[key] = re.compile(rule["pattern"], flags)
    return compiled


# încărcăm o singură dată la import
_rules = load_rules()
_rx = _compile_rules(_rules)

_ENTRY_RE = _rx["entry"]
_SYNONYM_RE = _rx["synonym"]
_CITATION_RE = _rx["citation"]
_STOP_LINE_RE = _rx["stop_line"]
_ALTE_RE = _rx["alte_numiri"]
_FAM_RE = _rx["fam"]
_GENUS_HEADER_RE = _rx.get("genus_header")
_DASH_RE = _rx["dash"]
_HAS_LETTER = _rx["has_letter"]
_ALPHA_TAB_RE = _rx["alpha_tab"]
_PAGE_NUM_RE = _rx["page_number"]
_VALID_NOISE_RE = _rx["valid_name_noise"]


def _parse_page_range(page_range: str | None) -> tuple[int, int]:
    """Normalizează un interval „8-200” (sau „51”) la (start, end), 1-based.

    Acceptă spații și liniuțe tipărite (‐ – —). Gol/None = (1, infinit).
    """
    import math

    if not page_range or not str(page_range).strip():
        return 1, math.inf  # type: ignore[return-value]
    norm = re.sub(r"[‐–—]", "-", str(page_range).strip())
    parts = [p.strip() for p in norm.split("-", 1)]
    try:
        if len(parts) == 1:
            start = end = int(parts[0])
        else:
            start = int(parts[0]) if parts[0] else 1
            end = int(parts[1]) if parts[1] else math.inf  # type: ignore[assignment]
    except ValueError:
        raise ValueError(
            f"Interval de pagini invalid: {page_range!r}. "
            "Folosește forma «8-200» sau o singură pagină, «51».")
    if isinstance(start, int) and isinstance(end, int) and start > end:
        raise ValueError(
            f"Interval de pagini invalid: {page_range!r} (start > end).")
    return start, end


def extract_entries(pdf_path: Path, page_range: str | None = None) -> list[dict]:
    """Parcurge textul PDF și returnează intrările dicționarului.

    `page_range`: interval opțional de pagini PDF, 1-based, de forma
    „8-200” (sau o singură pagină, „51”). Gol = tot documentul.
    """
    import pymupdf

    start, end = _parse_page_range(page_range)

    doc = pymupdf.open(str(pdf_path))
    total = len(doc)
    end_eff = total if end == float("inf") else end
    if start < 1 or end_eff > total or start > end_eff:
        doc.close()
        raise ValueError(
            f"Interval de pagini invalid: {page_range!r} "
            f"(documentul are {total} pagini).")
    lines: list[str] = []
    for page in doc[start - 1:end_eff]:
        lines.extend(page.get_text("text").splitlines())
    doc.close()

    # filtrează liniile de navigație (taburi alfabetice + numere de pagină)
    clean: list[str] = []
    for ln in lines:
        t = ln.strip()
        if not t:
            continue
        if _ALPHA_TAB_RE.fullmatch(t):  # tab alfabetic (BER, DOR...)
            continue
        if _PAGE_NUM_RE.fullmatch(t):  # număr de pagină
            continue
        clean.append(t.replace("\xad", ""))  # soft-hyphen OCR

    entries: list[dict] = []
    current: dict | None = None
    region: list[str] = []

    def flush():
        nonlocal current, region
        if current is not None:
            current["region"] = region
            entries.append(current)
        current, region = None, []

    for ln in clean:
        # Marcatorii nu sunt specii noi; păstrează-i pentru delimitare.
        m = None if (_ALTE_RE.match(ln) or _FAM_RE.match(ln)
                     or (_GENUS_HEADER_RE and _GENUS_HEADER_RE.match(ln))) \
            else _ENTRY_RE.match(ln)
        if m:
            genus, species, rest = m.group(1), m.group(2), m.group(3)
            if len(species) < 3:
                continue
            # falsuri: rest doar cu citări „(154,", referințe „vezi", liste
            # continuare „...: Căpușnic", fragmente fără litere
            r0 = rest.lstrip()
            if not _HAS_LETTER.search(r0):
                continue
            if r0.startswith((":", ",")):
                continue
            if re.search(r"\bvezi\b", rest, re.IGNORECASE):
                # referință încrucișată („ATRAGERE vezi CLEMA TIS")
                flush()
                continue
            flush()
            current = {"genus": genus, "species": species,
                       "sc_printed": f"{genus} {species}",
                       "rest": rest, "synonyms": []}
            for sm in _SYNONYM_RE.finditer(rest):
                current["synonyms"].append(f"{sm.group(1)} {sm.group(2)}")
        elif current is not None:
            region.append(ln)
    flush()
    return entries

def _clean_name(raw: str) -> str:
    n = _CITATION_RE.sub("", raw)
    n = re.sub(r"\s+", " ", n)
    # citare trunchiată la finalul liniei: „Tășculiță (319"
    n = re.sub(r"\([^()]*$", "", n)
    n = n.strip(" ;,:.-–—\t")
    return n


def _valid_name(n: str) -> bool:
    """Filtru de zgomot OCR pentru denumirile populare extrase."""
    if not (2 <= len(n) <= 60):
        return False
    if re.fullmatch(r"[\d\s,.]+", n):
        return False
    if _VALID_NOISE_RE.search(n):
        return False
    return True


def popular_names(rest: str, region: list[str]) -> list[str]:
    """Scoate denumirile populare românești dintr-o intrare.

    Format: numele principal după „—"/„ -" pe linia de antet (posibil
    continuat pe liniile următoare), până la „Alte numiri:”.
    Variantele de după marcator și zonele străine sunt excluse.
    """
    text = (rest + "\n" + "\n".join(region)).strip()
    # Folosește poziția marcatorilor, indiferent de grupurile din regex.
    stops = [m.start() for m in (_ALTE_RE.search(text), _FAM_RE.search(text)) if m]
    primary_part = text[:min(stops)] if stops else text
    # taie partea principală la prima secțiune străină / antet de gen / Fam.
    kept: list[str] = []
    for ln in primary_part.splitlines():
        if _STOP_LINE_RE.match(ln):
            break
        if _GENUS_HEADER_RE and _GENUS_HEADER_RE.match(ln):
            break
        kept.append(ln)
    primary_text = "\n".join(kept)

    names: list[str] = []
    m = _DASH_RE.search(primary_text)
    if m:
        tail = m.group(1)
        for part in re.split(r"[;,]|\n", tail):
            n = _clean_name(part)
            if _valid_name(n):
                names.append(n)

    # dedupe case/diacritice-insensitiv, păstrând ordinea
    seen: set[str] = set()
    out: list[str] = []
    for n in names:
        k = _skel(n)
        if k and k not in seen:
            seen.add(k)
            out.append(n)
    return out


# -------------------------------------------------------------------- API GUI

def extract_draft(pdf_path: Path | None = None, progress=None,
                  page_range: str | None = None) -> dict:
    """Extrage denumirile din PDF, potrivește cu taxonomy.db și scrie draftul.

    Apelabil din CLI (main()) și din butonul GUI. `progress(done, total, hits)`
    este opțional. `page_range` limitează paginile PDF parcurse („8-200”
    sau o singură pagină, „51”); gol = tot documentul. Returnează un dict
    cu statistici + căile scrise.
    """
    if not pdf_path:
        raise ValueError("Selectează un PDF sursă pentru extragerea denumirilor.")
    pdf_path = Path(pdf_path)
    if not pdf_path.is_file():
        raise FileNotFoundError(f"PDF-ul nu există: {pdf_path}")
    if not DB_PATH.is_file():
        raise FileNotFoundError(f"{DB_PATH} nu există.")

    entries = extract_entries(pdf_path, page_range=page_range)

    # index DB: (genus_skel -> {species_skel: nume_db}) + info pe nume
    con = sqlite3.connect(str(DB_PATH))
    genus_map: dict[str, dict[str, str]] = {}
    info: dict[str, tuple[str | None, str | None]] = {}
    for tid, name, ro, en in con.execute(
            "SELECT id, name, ro_name, en_name FROM taxa WHERE rank='species'"):
        parts = name.split()
        if len(parts) < 2:
            continue
        g, s = _skel(parts[0]), _skel(parts[1])
        if not g or not s:
            continue
        # Coliziune pe același binom (ex: „Abies alba" vs „Abies alba × concolor"):
        # binomul exact de 2 cuvinte câștigă; între două forme extinse,
        # prima întâlnită rămâne (comportamentul vechi).
        prev = genus_map.setdefault(g, {}).get(s)
        if prev is None or (len(prev.split()) > 2 and len(parts) == 2):
            genus_map[g][s] = name
        info[name] = (ro, en)
    genus_list = list(genus_map.keys())
    # blocare pe prima literă pentru căutarea fuzzy de genuri (OCR nu corupe
    # de obicei prima literă) -> scade de la ~35k comparații la ~1-2k
    genus_by_first: dict[str, list[str]] = {}
    for g in genus_list:
        genus_by_first.setdefault(g[0], []).append(g)

    from difflib import get_close_matches

    def find_match(genus: str, species: str, synonyms: list[str]) -> tuple[str | None, str]:
        """Potrivire exactă, apoi fuzzy (OCR). Returnează (nume_db, mod)."""
        cands = [(genus, species)] + [(s.split()[0], s.split()[1]) for s in synonyms
                                      if len(s.split()) >= 2]
        for g, s in cands:
            g_skel = _skel(g)
            s_skel = _skel(s)
            grp = genus_map.get(g_skel)
            if grp:
                if s_skel in grp:
                    return grp[s_skel], "exact"
                close = get_close_matches(s_skel, grp.keys(), n=1, cutoff=0.75)
                if close:
                    return grp[close[0]], "OCR"
        # genul însuși e corupt de OCR -> genuri apropiate (blocate pe litera 1)
        for g, s in cands:
            g_skel = _skel(g)
            if g_skel in genus_map:
                continue  # deja tratat mai sus
            s_skel = _skel(s)
            for gs in get_close_matches(g_skel, genus_by_first.get(g_skel[0], []),
                                        n=2, cutoff=0.85):
                grp = genus_map[gs]
                if s_skel in grp:
                    return grp[s_skel], "OCR"
                close = get_close_matches(s_skel, grp.keys(), n=1, cutoff=0.8)
                if close:
                    return grp[close[0]], "OCR"
        return None, ""

    draft: dict[str, tuple[str, str, str, str]] = {}  # db_name -> (sc, ro, en, mod)
    unmatched: list[tuple[str, str]] = []
    seen_unmatched: set[str] = set()

    total = len(entries)
    for n, e in enumerate(entries, 1):
        if progress is not None and n % 250 == 0:
            progress(n, total, len(draft))
        names = popular_names(e.get("rest", ""), e["region"])
        if not names:
            continue
        ro = "; ".join(names)
        db_name, mode = find_match(e["genus"], e["species"], e["synonyms"])
        if db_name:
            prev = draft.get(db_name)
            if prev is None or len(ro) > len(prev[1]):
                draft[db_name] = (db_name, ro, info[db_name][1] or "", mode)
        else:
            if e["sc_printed"] not in seen_unmatched:
                seen_unmatched.add(e["sc_printed"])
                unmatched.append((e["sc_printed"], ro))

    con.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Draft"
    ws.append(["Denumire științifică", "ro_name", "en_name", "Potrivire"])
    for db_name, (sc, ro, en, mode) in sorted(draft.items()):
        ws.append([db_name, ro, en or None, mode])

    ws2 = wb.create_sheet("Nepotrivite")
    ws2.append(["Denumire din PDF", "ro_name propus", "Motiv"])
    for sc, ro in sorted(unmatched):
        ws2.append([sc, ro, "specia nu a fost găsită în taxonomy.db "
                            "(nomenclatură 1968; caută numele actual)"])
    wb.save(OUT_PATH)

    n_exact = sum(1 for v in draft.values() if v[3] == "exact")
    return {
        "entries": total,
        "matched": len(draft),
        "exact": n_exact,
        "ocr": len(draft) - n_exact,
        "unmatched": len(unmatched),
        "out_path": str(OUT_PATH),
        "pdf_path": str(pdf_path),
        "page_range": page_range or "",
    }


# -------------------------------------------------------------------- CLI

def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Extrage denumiri dintr-o carte PDF.")
    parser.add_argument("pdf_path", type=Path, help="PDF sursă cu text extractabil")
    parser.add_argument("--pages", default=None,
                        help="Interval de pagini PDF, 1-based (ex: 8-200 sau 51). Gol = tot documentul.")
    args = parser.parse_args()
    try:
        stats = extract_draft(pdf_path=args.pdf_path, page_range=args.pages,
                              progress=lambda done, total, hits: print(
            f"  ... {done}/{total} potrivite până acum: {hits}", flush=True))
    except (FileNotFoundError, ValueError) as exc:
        print(f"EROARE: {exc}")
        return 1
    print(f"Intrări dicționar detectate: {stats['entries']}")
    print(f"Potrivite:   {stats['matched']}  "
          f"(exact: {stats['exact']}, OCR-fuzzy: {stats['ocr']})")
    print(f"Nepotrivite: {stats['unmatched']}")
    print(f"Draft scris: {stats['out_path']}")
    print('Revizuiește-l (mai ales sheet-ul Nepotrivite), apoi importă-l prin '
          'butonul „Adaugă denumiri din Excel".')
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

