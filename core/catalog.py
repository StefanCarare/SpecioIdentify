from __future__ import annotations

import hashlib
import json
import re
import shutil
import unicodedata
from pathlib import Path

from core.database import connect
from core.taxonomy import lookup_catalog_category
from core.catalog_paths import (
    store_catalog_path, store_original_path, resolve_catalog_path, resolved_catalog_item,
)


STRUCTURED_DATA_MARKER = "=== STRUCTURED DATA ==="

SPECIES_PROFILE_PROMPT_NAME = "species_profile_deep"

STANDARD_FIELDS = (
    "category",
    "kingdom",
    "genus",
    "species",
    "scientific_name",
    "ro_name",
    "en_name",
    "confidence",
    "is_uncertain",
)


def normalize_preset_name(value: str | None) -> str:
    if not value:
        return ""
    # Strip the optional ordinal prefix used for taxonomic ordering
    # (e.g. "02_1.plant_fast.txt" -> "plant_fast"); legacy names pass through.
    name = re.sub(r"^\d+(?:[._]\d+)*\.", "", str(value).strip())
    return Path(name).stem.casefold()


def append_structured_data_instruction(
    prompt: str,
    compact: bool = False,
) -> str:
    """Asks the model for stable catalog fields without changing the prompt's content."""
    if compact:
        instruction = f"""

{STRUCTURED_DATA_MARKER}
Răspunde doar cu un singur obiect JSON valid, fără Markdown și fără explicații,
folosind câmpurile: category, genus, species, scientific_name, ro_name,
en_name, confidence, is_uncertain. Valorile necunoscute sunt null.
""".strip()
        return f"{prompt.rstrip()}\n\n{instruction}"

    instruction = f"""

{STRUCTURED_DATA_MARKER}
La finalul răspunsului adaugă acest marker și apoi un singur obiect JSON valid.
Nu pune Markdown în obiectul JSON. Folosește exact aceste câmpuri standard:
category, kingdom, genus, species, scientific_name, ro_name, en_name,
confidence, is_uncertain.
`category` trebuie să fie una dintre: Planta, Animal, Pasare, Ciuperca, Alta.
`confidence` este un număr între 0 și 1. Dacă o valoare nu poate fi stabilită,
folosește null. Câmpurile suplimentare specifice acestui prompt pot fi adăugate
în obiect, dar nu elimina câmpurile standard.
""".strip()
    return f"{prompt.rstrip()}\n\n{instruction}"


# ============================================================
# SPECIES PROFILE (date structurate la nivel de specie)
# ============================================================

SPECIES_PROFILE_FIELDS_COMMON = (
    "raspandire",
    "dimensiuni",
    "habitat",
    "identificare",
    "descriere",
)

SPECIES_PROFILE_FIELDS_BIRD = (
    "cuib",
    "oua",
    "hrana",
)

SPECIES_PROFILE_FIELDS_PLANT = (
    "port",
    "tulpina",
    "frunza_dispozitie",
    "frunza_forma",
    "frunza_margine",
    "inflorescenta",
    "involucru",
    "petale",
    "simetrie",
    "stil",
    "ovar",
    "fruct",
    "inflorire",
    "culoare_floare",
    "polenizare",
    "fapt_divers",
)

SPECIES_PROFILE_FIELDS_ANIMAL = (
    "hrana",
    "acoperire",
    "comportament",
)


def append_species_profile_instruction(prompt: str) -> str:
    """
    Adaugă cererea pentru date structurate la nivel de specie
    (răspândire, dimensiuni, habitat, cuib, oua, hrana, etc.).
    Aceste date vor fi preluate ulterior de aplicație și trimise
    în tabela species_profiles.
    """
    instruction = f"""
{STRUCTURED_DATA_MARKER}
La finalul răspunsului adaugă exact acest marker (pe o linie separată, fără
asteriscuri sau alte formatări) și apoi un singur obiect JSON valid, complet.
Reguli obligatorii pentru JSON:
- fiecare valoare este un singur string JSON între ghilimele (de exemplu
  \"3-5 cm\"), nu obiect imbricat, nu număr gol, nu text în afara ghilimelelor;
- nu folosi Markdown în obiectul JSON și nu adăuga comentarii;
- închide toate acoladele deschise și nu adăuga nimic după obiect.
Folosește câmpurile comune:
{", ".join(SPECIES_PROFILE_FIELDS_COMMON)},
plus `category`.
Pentru Păsări adaugă și: {", ".join(SPECIES_PROFILE_FIELDS_BIRD)}.
Pentru Plante adaugă și: {", ".join(SPECIES_PROFILE_FIELDS_PLANT)}.
Pentru Animale adaugă și: {", ".join(SPECIES_PROFILE_FIELDS_ANIMAL)}.
`category` trebuie să fie una dintre: Planta, Animal, Pasare, Ciuperca, Alta.
Dacă o valoare nu poate fi stabilită pentru această specie sau nu este relevantă
pentru categorie, folosește null. Nu elimina câmpurile comune; pe lângă ele,
poți adăuga și altele utile pentru cineva care întâlnește specia pe teren.
""".strip()
    return f"{prompt.rstrip()}\n\n{instruction}"


def extract_structured_data(response: str) -> dict:
    """Extracts the JSON block while keeping the human-readable response intact."""
    candidates = []
    if STRUCTURED_DATA_MARKER in response:
        candidates.append(response.split(STRUCTURED_DATA_MARKER, 1)[1].strip())

    # Toleranță: unele modele scriu markerul cu Bold (**STRUCTURED DATA**)
    bolded_marker = re.search(
        r"\*\*\s*STRUCTURED DATA\s*\*\*",
        response,
        flags=re.IGNORECASE,
    )
    if bolded_marker:
        candidates.append(response[bolded_marker.end():].strip())

    # Toleranță: orice linie „heading" de tip Structured Data (ex. "## Structured Data")
    heading_marker = re.search(
        r"(?im)^[\s>#*_`~=\-]*structured[\s_\-]*data[\s:#*_`~=\-]*$",
        response,
    )
    if heading_marker:
        candidates.append(response[heading_marker.end():].strip())

    # Toleranță: JSON gol la final, fără niciun marker
    tail_dict = _tail_json_dict(response)
    if tail_dict is not None:
        candidates.append(json.dumps(tail_dict))

    candidates.extend(
        match.group(1).strip()
        for match in re.finditer(
            r"```(?:json)?\s*(\{.*?\})\s*```",
            response,
            flags=re.IGNORECASE | re.DOTALL,
        )
    )

    for candidate in candidates:
        candidate = re.sub(r"^```(?:json)?\s*", "", candidate, flags=re.IGNORECASE)
        candidate = re.sub(r"\s*```\s*$", "", candidate)
        try:
            data = json.loads(candidate)
        except json.JSONDecodeError:
            continue
        if isinstance(data, dict):
            return data

    scientific_match = re.search(
        r"(?im)^.*(?:scientific|stiin|știin).*?:\s*[*`]?([A-Za-z][\w-]+)\s+([A-Za-z][\w-]+)",
        response,
        flags=re.IGNORECASE,
    )
    if scientific_match:
        scientific_name = " ".join(
            scientific_match.group(index).strip()
            for index in (1, 2)
        )
        parts = scientific_name.split()
        common_match = re.search(
            r"(?:Nume[^:\n]*comun|ro_name)\s*:\s*([^\n]+)",
            response,
            flags=re.IGNORECASE,
        )
        return {
            "category": "Planta",
            "genus": parts[0],
            "species": parts[1],
            "scientific_name": scientific_name,
            "ro_name": common_match.group(1).strip() if common_match else None,
            "en_name": None,
            "confidence": None,
            "is_uncertain": False,
        }

    binomial_match = re.search(
        r"(?im)^\s*[*`]?([A-Z][A-Za-z-]+)\s+([a-z][A-Za-z-]+)[*`]?\s*$",
        response.strip(),
    )
    if binomial_match:
        genus, species = binomial_match.groups()
        return {
            "category": "Planta",
            "genus": genus,
            "species": species,
            "scientific_name": f"{genus} {species}",
            "ro_name": None,
            "en_name": None,
            "confidence": None,
            "is_uncertain": False,
        }

    return {}


def _tail_json_dict(text: str) -> dict | None:
    """Detectează un obiect JSON „gol" (fără marker, fără code fences) la finalul textului."""
    stripped = text.rstrip()
    if not stripped.endswith("}"):
        return None
    # Încearcă pozițiile { din coadă (max 12), de la ultima spre prima.
    brace_positions = [m.start() for m in re.finditer(r"\{", stripped)]
    for start in reversed(brace_positions[-12:]):
        candidate = stripped[start:]
        try:
            data = json.loads(candidate)
        except (json.JSONDecodeError, ValueError):
            continue
        if isinstance(data, dict) and (
            {"category", "scientific_name", "genus", "kingdom", "ro_name",
             "en_name", "confidence", "is_uncertain"} & set(data)
        ):
            return data
    return None


def strip_structured_data(response: str) -> str:
    """Removes the technical JSON block from text shown to the user."""
    response = re.sub(
        r"(?im)^\s*(?:to|from|recipient)\s*(?:=|:)\s*[^\n]+\s*$",
        "",
        response,
    )

    response_markers = [
        marker for marker in ("=== RESPONSE ===", "=== RESPONSE ===\n")
        if marker in response
    ]
    if response_markers:
        response = response.rsplit(response_markers[0], 1)[1]

    if STRUCTURED_DATA_MARKER in response:
        response = response.split(STRUCTURED_DATA_MARKER, 1)[0]
    else:
        # Toleranță: unele modele scriu markerul cu Bold (**STRUCTURED DATA**)
        bolded_marker = re.search(
            r"\*\*\s*STRUCTURED DATA\s*\*\*",
            response,
            flags=re.IGNORECASE,
        )
        if bolded_marker:
            response = response[: bolded_marker.start()]
        else:
            # Toleranță: orice linie „heading" de tip Structured Data,
            # ex.: "## Structured Data", "**Structured data:**", "STRUCTURED DATA:",
            # dar doar dacă urmată de un bloc JSON.
            heading = re.search(
                r"(?im)^[\s>#*_`~=\-]*structured[\s_\-]*data[\s:#*_`~=\-]*$",
                response,
            )
            if heading and re.match(r"\s*\{", response[heading.end():]):
                response = response[: heading.start()]
            elif _tail_json_dict(response) is not None:
                # JSON gol la final (fără marker deloc)
                stripped = response.rstrip()
                brace_positions = [m.start() for m in re.finditer(r"\{", stripped)]
                for start in reversed(brace_positions[-12:]):
                    try:
                        data = json.loads(stripped[start:])
                    except (json.JSONDecodeError, ValueError):
                        continue
                    if isinstance(data, dict) and (
                        {"category", "scientific_name", "genus", "kingdom",
                         "ro_name", "en_name", "confidence",
                         "is_uncertain"} & set(data)
                    ):
                        response = stripped[:start]
                        break

    if "=== STATISTICS ===" in response:
        response = response.split("=== STATISTICS ===", 1)[0]

    if "=== THINKING ===" in response:
        response = response.split("=== THINKING ===", 1)[0]

    response = re.sub(
        r"\n?```(?:json)?\s*\{.*?\}\s*```\s*",
        "\n",
        response,
        flags=re.IGNORECASE | re.DOTALL,
    )
    return response.strip()


def _image_hash(image_path: Path) -> str:
    digest = hashlib.sha256()
    with Path(image_path).open("rb") as image_file:
        for chunk in iter(lambda: image_file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def analysis_key(
    prompt: str,
    model: str,
    num_ctx: int,
    num_predict: int,
    thinking: bool,
) -> str:
    value = json.dumps(
        {
            "prompt": prompt,
            "model": model,
            "num_ctx": num_ctx,
            "num_predict": num_predict,
            "thinking": thinking,
        },
        ensure_ascii=False,
        sort_keys=True,
    )
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _safe_part(value: object, fallback: str) -> str:
    value = str(value or "").strip()
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode()
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._-")
    return value or fallback


def _category_folder(category: object) -> str:
    value = str(category or "Alta").strip().casefold()
    if "plant" in value:
        return "plante"
    if "pas" in value or "bird" in value:
        return "pasari"
    if "animal" in value:
        return "animale"
    if "ciuper" in value or "fung" in value:
        return "ciuperci"
    return "altele"


def is_fast_preset(value: str | None) -> bool:
    name = normalize_preset_name(value)
    return name.endswith("_fast") or name == "botanic_quick"


def fast_analysis_fields(response: str, preset_name: str) -> dict:
    """Parse a short binomial locally, without assuming every organism is a plant."""
    raw = str(preset_name or "").strip().casefold()
    if raw.endswith(".txt"):
        raw = raw[:-4]
    name = raw  # keep dots and digits: tokens are matched one by one below
    if name == "botanic_quick":
        return extract_structured_data(response)
    match = re.fullmatch(r"\s*[*`]*([A-Z][a-z-]+)\s+([a-z][a-z-]+)[*`]*\s*", response)
    if not match or match.group(1).lower() in {"unknown", "identificare", "identification"}:
        return {}
    genus, species = match.groups()
    categories = {
        "botanic": "Planta", "plant": "Planta", "flower": "Planta",
        "bird": "Pasare", "animal": "Animal", "insect": "Animal",
        "reptile": "Animal", "reptil": "Animal",
        "amphibian": "Animal", "amfibian": "Animal", "mammal": "Animal",
        "fungi": "Ciuperca", "mushroom": "Ciuperca",
    }
    category = "Alta"
    for token in re.findall(r"[a-z]+", name):
        if token == "default":
            category = lookup_catalog_category(f"{genus} {species}")
            break
        if token in categories:
            category = categories[token]
            break
    return {
        "category": category,
        "scientific_name": f"{genus} {species}", "genus": genus, "species": species,
        "ro_name": None, "en_name": None, "confidence": None,
        "is_uncertain": False,
    }


def fast_analysis_prompt(prompt: str, preset_name: str) -> str:
    if normalize_preset_name(preset_name) == "botanic_quick":
        return prompt
    return (
        f"{prompt.rstrip()}\n\nMod rapid ({normalize_preset_name(preset_name)}): "
        "indiferent de formatul cerut mai sus, răspunde doar cu un singur binom "
        "Gen specie, fără JSON, denumiri populare, explicații sau profil. "
        "Dacă specia nu poate fi stabilită, răspunde doar: Identificare nesigură."
    )


def catalogize_analysis(
    image_path: Path,
    response: str,
    output_dir: Path,
    analysis_key_value: str | None = None,
    thinking: str = "",
    prompt: str = "",
    preset_name: str = "",
    manual_location: str = "",
    image_metadata: dict | None = None,
    statistics: dict | None = None,
) -> Path | None:
    """Copies a newly identified image into the catalog and records its fields."""
    fields = (fast_analysis_fields(response, preset_name)
              if is_fast_preset(preset_name) else extract_structured_data(response))
    if not fields:
        return None

    image_path = Path(image_path)
    output_dir = Path(output_dir)
    database_path = output_dir / "catalog.db"
    image_hash = _image_hash(image_path)
    category = _category_folder(fields.get("category"))
    scientific_name = fields.get("scientific_name") or "Necunoscut"
    species_key = _safe_part(scientific_name, "Necunoscut")
    catalog_folder = output_dir / category / species_key
    catalog_folder.mkdir(parents=True, exist_ok=True)

    connection = connect(database_path)
    try:
        existing = connection.execute(
            "SELECT catalog_path FROM catalog_items WHERE image_hash = ?",
            (image_hash,),
        ).fetchone()
        if existing:
            clean_response = strip_structured_data(response)
            connection.execute(
                """
                UPDATE catalog_items
                SET analysis_key = ?, response_text = ?, category = ?,
                    kingdom = ?, genus = ?, species = ?, scientific_name = ?,
                    ro_name = ?, en_name = ?, confidence = ?, is_uncertain = ?,
                    extra_fields_json = ?
                    , manual_override = 0, thinking_text = ?, prompt_text = ?,
                    preset_name = ?, manual_location = ?, metadata_json = ?,
                    statistics_json = ?
                WHERE image_hash = ?
                """,
                (
                    analysis_key_value,
                    clean_response,
                    _category_folder(fields.get("category")),
                    fields.get("kingdom"),
                    fields.get("genus"),
                    fields.get("species"),
                    fields.get("scientific_name"),
                    fields.get("ro_name"),
                    fields.get("en_name"),
                    fields.get("confidence"),
                    int(bool(fields.get("is_uncertain"))),
                    json.dumps(
                        {
                            key: value for key, value in fields.items()
                            if key not in STANDARD_FIELDS
                        },
                        ensure_ascii=False,
                    ),
                    thinking,
                    prompt,
                    preset_name,
                    manual_location,
                    json.dumps(image_metadata or {}, ensure_ascii=False),
                    json.dumps(statistics or {}, ensure_ascii=False),
                    image_hash,
                ),
            )
            connection.commit()
            return resolve_catalog_path(existing["catalog_path"], output_dir)

        next_number = connection.execute(
            "SELECT COUNT(*) FROM catalog_items WHERE category = ? AND scientific_name = ?",
            (category, scientific_name),
        ).fetchone()[0] + 1
        destination = catalog_folder / (
            f"{_safe_part(scientific_name, 'Necunoscut')}_{next_number:04d}"
            f"{image_path.suffix.lower()}"
        )
        shutil.copy2(image_path, destination)

        extra_fields = {
            key: value for key, value in fields.items()
            if key not in STANDARD_FIELDS
        }
        clean_response = strip_structured_data(response)

        # Protejăm profile-urile bogate generate explicit (prin butonul
        # "Generează Profil Specie"): nu le suprascriem cu răspunsurile
        # foto-specifice venite din analiză automată.
        rich_profile_row = connection.execute(
            "SELECT profile_fields_json FROM species_profiles "
            "WHERE lower(scientific_name) = lower(?)",
            (scientific_name,),
        ).fetchone()
        has_rich_profile = bool(
            rich_profile_row
            and rich_profile_row["profile_fields_json"].strip() not in ("", "{}")
        )

        # NOTA: Generarea automată de profile din analiza de imagine a fost
        # eliminată — profilurile se generează DOAR prin butoanele dedicate
        # GEN PROFILE / GEN PROFILE ALL. Astfel ANALYZE ALL nu mai declanșează
        # în spate generarea de profiluri (dublarea timpului de rulare).
        connection.execute(
            """
            INSERT INTO catalog_items (
                image_hash, analysis_key, original_path, catalog_path, response_text,
                thinking_text, prompt_text, preset_name, manual_location,
                metadata_json, statistics_json,
                category, kingdom, genus, species, scientific_name, ro_name,
                en_name, confidence, is_uncertain, extra_fields_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                image_hash,
                analysis_key_value,
                store_original_path(image_path, output_dir),
                store_catalog_path(destination, output_dir),
                clean_response,
                thinking,
                prompt,
                preset_name,
                manual_location,
                json.dumps(image_metadata or {}, ensure_ascii=False),
                json.dumps(statistics or {}, ensure_ascii=False),
                category,
                fields.get("kingdom"),
                fields.get("genus"),
                fields.get("species"),
                scientific_name,
                fields.get("ro_name"),
                fields.get("en_name"),
                fields.get("confidence"),
                int(bool(fields.get("is_uncertain"))),
                json.dumps(extra_fields, ensure_ascii=False),
            ),
        )
        connection.commit()
        return destination
    finally:
        connection.close()


def get_cached_analysis(
    image_path: Path,
    output_dir: Path,
    analysis_key_value: str,
) -> dict | None:
    """Returns a previously cataloged response for the exact same image."""
    image_path = Path(image_path)
    connection = connect(Path(output_dir) / "catalog.db")
    try:
        row = connection.execute(
            "SELECT * FROM catalog_items WHERE image_hash = ? AND analysis_key = ?",
            (_image_hash(image_path), analysis_key_value),
        ).fetchone()
        if not row or not row["response_text"]:
            return None
        result = resolved_catalog_item(row, output_dir)
        result["response_text"] = strip_structured_data(
            result["response_text"]
        )
        return result
    finally:
        connection.close()


def get_species_profile(output_dir: Path, scientific_name: str) -> dict | None:
    connection = connect(Path(output_dir) / "catalog.db")
    try:
        row = connection.execute(
            "SELECT * FROM species_profiles WHERE lower(scientific_name) = lower(?)",
            (scientific_name.strip(),),
        ).fetchone()
        return dict(row) if row else None
    finally:
        connection.close()


def update_species_profile_response(
    output_dir: Path,
    scientific_name: str,
    profile_response: str,
) -> bool:
    """
    Actualizează textul de afișare al profilului unei specii
    (profile_response), păstrând restul câmpurilor. Se folosește
    când utilizatorul editează profilul în zona Output și apasă
    SAVE PROFILE. Returnează True dacă a existat un profil de
    actualizat, False în caz contrar.
    """
    output_dir = Path(output_dir)
    scientific_name = (scientific_name or "").strip()

    if not scientific_name:
        raise ValueError("Numele științific nu poate fi gol.")

    connection = connect(output_dir / "catalog.db")
    try:
        cursor = connection.execute(
            """
            UPDATE species_profiles
            SET profile_response = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE lower(scientific_name) = lower(?)
            """,
            (profile_response, scientific_name),
        )
        connection.commit()
        return cursor.rowcount > 0
    finally:
        connection.close()


def get_profile_for_response(output_dir: Path, response: str) -> dict | None:
    fields = extract_structured_data(response)
    scientific_name = fields.get("scientific_name")
    if not scientific_name:
        return None
    return get_species_profile(output_dir, scientific_name)


def save_species_profile(
    output_dir: Path,
    scientific_name: str,
    profile_response: str,
    profile_fields: dict | None = None,
    category: str = "",
    kingdom: str | None = None,
    genus: str | None = None,
    species: str | None = None,
    ro_name: str | None = None,
    en_name: str | None = None,
) -> dict:
    """
    Salvează sau actualizeză profilul unei specii în species_profiles.

    profile_response este textul bogat al profilului pentru afișare;
    profile_fields este dict-ul flexibil JSON (rapsândire, dimensiuni,
    habitat, cuib, oua, hrana etc.) generat la nivel de specie.
    """
    output_dir = Path(output_dir)
    scientific_name = (scientific_name or "").strip()

    if not scientific_name:
        raise ValueError("Numele științific nu poate fi gol.")

    profile_fields = profile_fields or {}

    connection = connect(output_dir / "catalog.db")
    try:
        connection.execute(
            """
            INSERT INTO species_profiles (
                scientific_name, category, kingdom, genus, species,
                ro_name, en_name, profile_response, profile_fields_json,
                extra_fields_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(scientific_name) DO UPDATE SET
                category = excluded.category,
                kingdom = excluded.kingdom,
                genus = excluded.genus,
                species = excluded.species,
                ro_name = excluded.ro_name,
                en_name = excluded.en_name,
                profile_response = excluded.profile_response,
                profile_fields_json = excluded.profile_fields_json,
                updated_at = CURRENT_TIMESTAMP
            """,
            (
                scientific_name,
                category or "",
                kingdom,
                genus,
                species,
                ro_name,
                en_name,
                profile_response,
                json.dumps(profile_fields, ensure_ascii=False),
                "{}",
            ),
        )
        connection.commit()
    finally:
        connection.close()

    return {
        "scientific_name": scientific_name,
        "category": category,
        "profile_response": profile_response,
        "profile_fields": profile_fields,
    }


def get_catalog_item(image_path: Path, output_dir: Path) -> dict | None:
    connection = connect(Path(output_dir) / "catalog.db")
    try:
        row = connection.execute(
            "SELECT * FROM catalog_items WHERE image_hash = ?",
            (_image_hash(Path(image_path)),),
        ).fetchone()
        return resolved_catalog_item(row, output_dir) if row else None
    finally:
        connection.close()


def update_catalog_response(
    image_path: Path,
    output_dir: Path,
    response_text: str,
) -> bool:
    """Actualizează response_text în catalog_items (SQLite).

    Folosit de butonul SAVE (panoul Output Item) când utilizatorul editează
    analiza afișată. Returnează True dacă s-a actualizat o înregistrare.
    """
    item = get_catalog_item(image_path, output_dir)
    if not item:
        return False

    response_text = (response_text or "").strip()
    if not response_text:
        raise ValueError("Rezultatul nu poate fi gol.")

    connection = connect(Path(output_dir) / "catalog.db")
    try:
        connection.execute(
            "UPDATE catalog_items SET response_text = ? WHERE image_hash = ?",
            (response_text, item["image_hash"]),
        )
        connection.commit()
    finally:
        connection.close()

    return True


def _update_response_identity(
    response: str,
    scientific_name: str,
    ro_name: str,
    en_name: str,
) -> str:
    identification = "\n".join([
        "1. **Identificare**",
        f"Nume comun în limba română: {ro_name or 'Nespecificat'}",
        f"Nume comun în limba engleză: {en_name or 'Nespecificat'}",
        f"Nume științific: *{scientific_name}*",
        "Identificare corectată manual.",
        "",
    ])
    pattern = r"1\.\s*\*\*Identificare\*\*.*?(?=\n\s*2\.\s*\*\*|\Z)"
    updated, count = re.subn(
        pattern,
        identification,
        response,
        count=1,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if count:
        return updated.strip()
    return f"{identification}\n{response.strip()}".strip()


def update_catalog_classification(
    image_path: Path,
    output_dir: Path,
    category: str,
    scientific_name: str,
    ro_name: str = "",
    en_name: str = "",
) -> Path:
    """Moves a catalog item to a corrected category/species folder."""
    image_path = Path(image_path)
    output_dir = Path(output_dir)
    category_folder = _category_folder(category)
    scientific_name = scientific_name.strip()
    if not scientific_name:
        raise ValueError("Numele științific nu poate fi gol.")

    item = get_catalog_item(image_path, output_dir)
    if not item:
        raise ValueError("Imaginea nu are încă o înregistrare în catalog.")

    species_key = _safe_part(scientific_name, "Necunoscut")
    destination_folder = output_dir / category_folder / species_key
    destination_folder.mkdir(parents=True, exist_ok=True)
    old_image = Path(item["catalog_path"])
    if old_image.parent.resolve() == destination_folder.resolve() and old_image.exists():
        destination = old_image
    else:
        number = 1
        while True:
            destination = destination_folder / (
                f"{species_key}_{number:04d}{old_image.suffix.lower()}"
            )
            if not destination.exists():
                break
            number += 1

    if old_image.resolve() != destination.resolve() and old_image.exists():
        shutil.move(str(old_image), str(destination))

    parts = scientific_name.split()
    genus = parts[0] if parts else None
    species = parts[1] if len(parts) > 1 else None
    connection = connect(output_dir / "catalog.db")
    try:
        updated_response = _update_response_identity(
            item.get("response_text", ""),
            scientific_name,
            ro_name.strip(),
            en_name.strip(),
        )
        connection.execute(
            """
            UPDATE catalog_items
            SET category = ?, scientific_name = ?, genus = ?, species = ?,
                ro_name = ?, en_name = ?, response_text = ?, analysis_key = NULL,
                catalog_path = ?, manual_override = 1
            WHERE image_hash = ?
            """,
            (
                category_folder,
                scientific_name,
                genus,
                species,
                ro_name.strip(),
                en_name.strip(),
                updated_response,
                store_catalog_path(destination, output_dir),
                item["image_hash"],
            ),
        )
        connection.commit()
    finally:
        connection.close()

    return destination

def update_catalog_common_names(
    output_dir: Path,
    image_path: Path,
    ro_name: str | None,
    en_name: str | None,
) -> bool:
    """Actualizează ro_name și en_name în catalog_items pentru o imagine.

    Folosit după ANALYZE pentru a sincroniza denumirile cu taxonomy.db.
    Returnează True dacă s-a actualizat o înregistrare.
    """
    item = get_catalog_item(image_path, output_dir)
    if not item:
        return False

    ro_name = (ro_name or "").strip() or None
    en_name = (en_name or "").strip() or None

    if not ro_name and not en_name:
        return False

    connection = connect(output_dir / "catalog.db")
    try:
        updates = []
        params = []
        if ro_name:
            updates.append("ro_name = ?")
            params.append(ro_name)
        if en_name:
            updates.append("en_name = ?")
            params.append(en_name)
        params.append(item["image_hash"])

        connection.execute(
            f"UPDATE catalog_items SET {', '.join(updates)} WHERE image_hash = ?",
            params,
        )
        connection.commit()
        return True
    finally:
        connection.close()


def apply_common_name_markers(
    response_text: str,
    ro_name: str | None = None,
    en_name: str | None = None,
) -> str:
    """Înlocuiește marcajele #...# dintre prompturile de analiză.

    Regulă (convenția curentă în prompturi):
      - #nume comun#, #română#, #romana#, #România#  → denumirea populară RO
      - #engleză#, #engleza#                          → denumirea populară EN

    Dacă limba corespunzătoare nu are valoare în DB, marcajul rămâne
    neschimbat (utilizatorul/promptul spune să se scrie „Necunoscut”).
    """
    ro_name = (ro_name or "").strip()
    en_name = (en_name or "").strip()

    if not ro_name and not en_name:
        return response_text

    ro_keys = {"nume comun", "română", "romana", "românia"}
    en_keys = {"engleză", "engleza"}

    def _replace(match: re.Match) -> str:
        marker = match.group(1)
        key = marker.casefold().strip()
        if ro_name and key in ro_keys:
            return ro_name
        if en_name and key in en_keys:
            return en_name
        return f"#{marker}#"

    return re.sub(
        r"#([^#\n]+)#",
        _replace,
        response_text,
    )


def _replace_common_name_bullet(
    text: str,
    ro_name: str | None,
    en_name: str | None,
) -> str:
    """Înlocuiește blocul „Numele comun”/„Numele popular” din response_text."""
    parts = []
    if ro_name:
        parts.append(f"română: {ro_name}")
    if en_name:
        parts.append(f"engleză: {en_name}")
    if not parts:
        return text
    replacement = " / ".join(parts)

    lines = text.split("\n")
    new_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if re.search(r"Nume(?:le)?\s*(comun|popular)", line, re.IGNORECASE):
            # Păstram doar eticheta (până la „:”), taem orice valoare după ea
            label = line
            m2 = re.match(
                r"^(.*?Nume(?:le)?\s*(?:comun|popular)[^:\n]*:?\*{0,2})",
                line,
                re.IGNORECASE,
            )
            if m2:
                label = m2.group(1).rstrip()
            # A language-specific label must never receive the other language.
            # With no verified name for that language, preserve the model's text.
            folded_label = unicodedata.normalize("NFKD", label.casefold())
            folded_label = "".join(c for c in folded_label if not unicodedata.combining(c))
            if re.search(r"\b(?:romana|romanesc|ro)\b", folded_label):
                value = ro_name
            elif re.search(r"\b(?:engleza|englezesc|en)\b", folded_label):
                value = en_name
            else:
                value = replacement
            if not value:
                new_lines.append(line)
                i += 1
                continue
            new_lines.append(label)
            i += 1
            while (
                i < len(lines)
                and re.match(r"^\s+\S", lines[i])
                and not lines[i].strip().startswith("*")
            ):
                i += 1
            new_lines.append(f"    {value}")
        else:
            new_lines.append(line)
            i += 1
    return "\n".join(new_lines)


def update_catalog_response_names(
    output_dir: Path,
    image_path: Path,
    ro_name: str | None = None,
    en_name: str | None = None,
) -> bool:
    """Actualizează ro_name/en_name și patch-uiește response_text.

    După ANALYZE, modelul poate returna denumiri populare greșite.
    Această funcție:
    1. Actualizează câmpurile ro_name și en_name
    2. Înlocuiește denumirile din secțiunea Identificare din response_text
       (marcajele #...#; compatibilitate cu blocuri [NUME_COMUN] vechi;
       fallback pe bullet „Numele comun”)
    """
    item = get_catalog_item(image_path, output_dir)
    if not item:
        return False

    ro_name = (ro_name or "").strip() or None
    en_name = (en_name or "").strip() or None

    if not ro_name and not en_name:
        return False

    response = item.get("response_text") or ""
    response_text = response

    # 1. Compatibilitate: bloculi vechi [NUME_COMUN]...[/NUME_COMUN]
    if re.search(r"\[NUME_COMUN\].*?\[/NUME_COMUN\]", response_text, re.DOTALL):
        parts = []
        if ro_name:
            parts.append(f"român: {ro_name}")
        if en_name:
            parts.append(f"engleză: {en_name}")
        content = " / ".join(parts)
        response_text = re.sub(
            r"\[NUME_COMUN\].*?\[/NUME_COMUN\]",
            content,
            response_text,
            flags=re.DOTALL,
        )

    # 2. Marcajele #...# (convenția curentă)
    response_text = apply_common_name_markers(response_text, ro_name, en_name)

    # 3. Fallback: nici marcaje, nici bloc vechi → înlocuim bullet direct
    if response_text == response:
        response_text = _replace_common_name_bullet(response, ro_name, en_name)

    connection = connect(output_dir / "catalog.db")
    try:
        updates = ["response_text = ?"]
        params = [response_text]
        if ro_name:
            updates.append("ro_name = ?")
            params.append(ro_name)
        if en_name:
            updates.append("en_name = ?")
            params.append(en_name)
        params.append(item["image_hash"])

        connection.execute(
            f"UPDATE catalog_items SET {', '.join(updates)} WHERE image_hash = ?",
            params,
        )
        connection.commit()
        return True
    finally:
        connection.close()


def quick_catalog_images(
    image_paths: list[Path],
    scientific_name: str,
    category: str,
    output_dir: Path,
    manual_location: str = "",
) -> dict:
    """Catalogheaza direct imagini fara analiza AI."""
    from core.taxonomy import lookup_species_simple

    output_dir = Path(output_dir)
    scientific_name = (scientific_name or "").strip()
    category = (category or "altele").strip().casefold()

    if not scientific_name:
        raise ValueError("Numele stiintific nu poate fi gol.")

    info = lookup_species_simple(scientific_name)
    if not info:
        raise ValueError(
            f"\u201e{scientific_name}\u201d nu exista in taxonomy.db."
        )

    db_ro = (info.get("ro_name") or "").strip()
    db_en = (info.get("en_name") or "").strip()
    kingdom = (info.get("kingdom") or "").strip()
    genus = info.get("genus", "")
    species_epithet = info.get("species", "")

    species_key = _safe_part(scientific_name, "Necunoscut")
    catalog_folder = output_dir / category / species_key
    catalog_folder.mkdir(parents=True, exist_ok=True)

    # Numără fișierele reale din folder (nu intrările din catalog),
    # pentru a evitarea suprascrierii celor deja existente.
    existing_files = list(catalog_folder.glob(f"{species_key}_*"))
    next_number = len(existing_files) + 1

    connection = connect(output_dir / "catalog.db")
    try:

        copied = 0
        errors = []
        for src in image_paths:
            src = Path(src)
            if not src.is_file():
                errors.append(f"{src.name}: fisierul nu exista")
                continue
            try:
                image_hash = _image_hash(src)
                dest = catalog_folder / (
                    f"{species_key}_{next_number:04d}{src.suffix.lower()}"
                )
                shutil.copy2(src, dest)

                connection.execute(
                    """INSERT OR IGNORE INTO catalog_items (
                        image_hash, original_path, catalog_path,
                        scientific_name, category, kingdom, genus, species,
                        ro_name, en_name, prompt_text, preset_name,
                        analysis_key, response_text, extra_fields_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (image_hash, store_original_path(src, output_dir),
                     store_catalog_path(dest, output_dir), scientific_name, category,
                     kingdom, genus, species_epithet, db_ro, db_en,
                     "[quick-catalog]", "quick-catalog",
                     f"quick-{image_hash[:12]}",
                     f"Catalogat manual ca {scientific_name}.", "{}"),
                )
                connection.commit()
                copied += 1
                next_number += 1
            except Exception as exc:
                errors.append(f"{src.name}: {exc}")

        return {"copied": copied, "errors": errors}
    finally:
        connection.close()


def get_manual_classification_context(
    image_path: Path,
    output_dir: Path,
) -> str:
    item = get_catalog_item(image_path, output_dir)
    if not item or not item.get("manual_override"):
        return ""
    values = [item.get("category"), item.get("scientific_name"),
              item.get("ro_name"), item.get("en_name")]
    return (
        "=== IDENTIFICARE MANUALA PROPUSA ===\n"
        f"Categoria: {values[0] or 'necunoscuta'}\n"
        f"Numele stiintific propus: {values[1] or 'necunoscut'}\n"
        f"Numele romanesc propus: {values[2] or 'necunoscut'}\n"
        f"Numele englezesc propus: {values[3] or 'necunoscut'}\n\n"
        "Utilizeaza aceastere identificare manuala ca ipoteza prioritară pentru "
        "reanalizare. Verific-o în imagine și actualizează identificarea doar "
        "dacă există dovezi vizuale clare împotriva ei."
    )



    stripped = response.strip()
    if stripped.startswith("{") and stripped.endswith("}"):
        candidates.append(stripped)