from pathlib import Path

from core.database import DEFAULT_DATABASE_NAME, connect, database_path_for


DEFAULT_PROMPT = (
    "Analizeaza imaginea cu atentie. "
    "Descrie ceea ce vezi si raspunde la intrebarea utilizatorului. "
    "Raspunde in romana. "
    "Nu inventa caracteristici care nu sunt vizibile."
)

# Numele presetului implicit (fișierul prompts/00_2.default_standard.txt).
DEFAULT_PROMPT_NAME = "00_2.default_standard"
# Intrare legacy din bazele de date create înainte de redenumire;
# folosită doar ca fallback la încărcarea promptului implicit.
LEGACY_DEFAULT_PROMPT_NAME = "default"


def load_prompt(prompt_path: Path) -> str:
    """
    Citește un prompt dintr-un fișier text.
    """
    if not prompt_path.is_file():
        raise FileNotFoundError(
            f"Fișierul prompt nu există: {prompt_path}"
        )

    prompt = prompt_path.read_text(encoding="utf-8").strip()

    if not prompt:
        raise ValueError(
            f"Fișierul prompt este gol: {prompt_path}"
        )

    return prompt


def save_prompt(prompt_path: Path, prompt: str) -> None:
    """
    Salvează un prompt într-un fișier text.
    """
    prompt = prompt.strip()

    if not prompt:
        raise ValueError("Promptul nu poate fi gol.")

    prompt_path.parent.mkdir(parents=True, exist_ok=True)
    prompt_path.write_text(prompt, encoding="utf-8")


def get_default_prompt() -> str:
    """
    Returnează promptul implicit al aplicației.
    """
    return DEFAULT_PROMPT


# ============================================================
# STOCARE ÎN BAZA DE DATE (catalog.db)
#
# NOTĂ: Toată structura bazei de date (tabele + migrări) este
# definită într-un singur loc: core/database.py. Aici folosim
# funcția connect() din acest module; nu mai existe schema locală.
# ============================================================


def normalize_prompt_name(value: str) -> str:
    """
    Normalizează numele unui prompt: transformă în miniscule și
    elimină extensia .txt pentru a unifica fișierele și baza de date.
    """
    name = str(value or "").strip().casefold()
    if name.endswith(".txt"):
        name = name[:-len(".txt")]
    return name.strip()


def list_prompts(database_path: Path) -> list[dict]:
    """
    Returnează toate prompt-urile din baza de date, sortate alfabetic.
    Fiecare item este un dicționar cu cheie: id, name, content,
    created_at, updated_at.
    """
    connection = connect(database_path)
    try:
        rows = connection.execute(
            "SELECT id, name, content, created_at, updated_at "
            "FROM prompts "
            "ORDER BY name COLLATE NOCASE"
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        connection.close()


def load_prompt_from_db(database_path: Path, name: str) -> dict | None:
    """
    Citește un prompt dintr-un baza de date după numele său.
    Returnează None dacă nu există.
    """
    connection = connect(database_path)
    try:
        row = connection.execute(
            "SELECT id, name, content, created_at, updated_at "
            "FROM prompts WHERE name = ?",
            (normalize_prompt_name(name),),
        ).fetchone()
        return dict(row) if row else None
    finally:
        connection.close()


def save_prompt_to_db(database_path: Path, name: str, content: str) -> None:
    """
    Salvează (creează sau atualizează) un prompt în baza de date.
    """
    content = (content or "").strip()

    if not content:
        raise ValueError("Promptul nu poate fi gol.")

    connection = connect(database_path)
    try:
        connection.execute(
            """
            INSERT INTO prompts (name, content) VALUES (?, ?)
            ON CONFLICT(name) DO UPDATE SET
                content = excluded.content,
                updated_at = CURRENT_TIMESTAMP
            """,
            (normalize_prompt_name(name), content),
        )
        connection.commit()
    finally:
        connection.close()


def create_prompt_in_db(database_path: Path, name: str, content: str = "") -> None:
    """
    Creează un prompt nou în baza de date.
    Aruncă ValueError dacă există deja un prompt cu acest nume.
    """
    name = normalize_prompt_name(name)

    if not name:
        raise ValueError("Numele promptului nu poate fi gol.")

    connection = connect(database_path)
    try:
        existing = connection.execute(
            "SELECT 1 FROM prompts WHERE name = ?",
            (name,),
        ).fetchone()

        if existing:
            raise ValueError(f"Promptul '{name}' există deja în baza de date.")

        connection.execute(
            "INSERT INTO prompts (name, content) VALUES (?, ?)",
            (name, content or ""),
        )
        connection.commit()
    finally:
        connection.close()


def sync_prompts_from_files(database_path: Path, prompts_dir: Path) -> int:
    """
    Importează în baza de date prompt-urile .txt din directorul prompts/
    care nu există încă în DB (migrare automat, fără suprascriere).

    Returnează numărul de prompt-uri importăte.
    """
    prompts_dir = Path(prompts_dir)

    if not prompts_dir.is_dir():
        return 0

    connection = connect(database_path)
    imported = 0
    try:
        for prompt_file in sorted(prompts_dir.glob("*.txt")):
            name = normalize_prompt_name(prompt_file.stem)

            existing = connection.execute(
                "SELECT 1 FROM prompts WHERE name = ?",
                (name,),
            ).fetchone()

            if existing:
                continue

            content = prompt_file.read_text(encoding="utf-8").strip()

            if not content:
                continue

            connection.execute(
                "INSERT INTO prompts (name, content) VALUES (?, ?)",
                (name, content),
            )
            imported += 1

        connection.commit()
        return imported
    finally:
        connection.close()