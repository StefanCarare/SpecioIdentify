from pathlib import Path
import json
import os
import re


_DIA = {"ă":"a","â":"a","î":"i","ș":"s","ş":"s","ț":"t","ţ":"t","Ă":"a","Â":"a","Î":"i","Ș":"s","Ş":"s","Ț":"t","Ţ":"t"}
def _fold(s):
    """Coboară la minuscule și elimină diacriticele românești."""
    return "".join(_DIA.get(c, c) for c in s.lower())

import threading
import time
import traceback
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, simpledialog, ttk


from PIL import Image, ImageTk

# Adaugă rădăcina proiectului în sys.path pentru rularea directă din gui/
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from core.images import (
    image_to_base64,
    get_image_metadata,
)

from core.ollama import (
    check_ollama,
    generate,
    list_models,
    get_response,
    get_thinking,
    get_statistics,
    OllamaError,
)


from core.catalog import (
    fast_analysis_prompt,
    is_fast_preset,
    append_structured_data_instruction,
    append_species_profile_instruction,
    analysis_key,
    catalogize_analysis,
    get_cached_analysis,
    strip_structured_data,
    extract_structured_data,
    get_catalog_item,
    get_manual_classification_context,
    get_profile_for_response,
    get_species_profile,
    save_species_profile,
    update_catalog_classification,
    update_species_profile_response,
    update_catalog_response,
    apply_common_name_markers,
    STANDARD_FIELDS,
    SPECIES_PROFILE_PROMPT_NAME,
)

from core.prompts import (
    load_prompt,
    save_prompt,
    list_prompts,
    load_prompt_from_db,
    save_prompt_to_db,
    create_prompt_in_db,
    sync_prompts_from_files,
    DEFAULT_PROMPT_NAME,
    LEGACY_DEFAULT_PROMPT_NAME,
)

from core.database import connect

from core.taxonomy import (
    COL_DIR,
    is_cache_ready as taxonomy_cache_ready,
    lookup_chain as taxonomy_lookup_chain,
    build_verified_context as taxonomy_verified_context,
    fix_tree_counts as taxonomy_fix_tree_counts,
)

from core.lang import t, get_language, set_language, SUPPORTED_LANGUAGES


def read_excel_entries(path: Path) -> list[tuple[str, str | None, str | None]]:
    """Citește foaia activă a unui Excel: A=nume științific, B=RO, C=EN.

    Sare peste o singură linie de antet, dacă există. Returnează perechile
    (stiintific, ro, en) cu cel puțin o denumire populară completă.
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        entries = []
        header_seen = False
        header_words = ("denumire", "nume", "name", "scientific", "stiintific", "common")
        for row in ws.iter_rows(min_row=1, values_only=True):
            if not row or not row[0]:
                continue
            sc = str(row[0]).strip()
            # Sare peste linia de header (o dată), dacă există
            if not header_seen:
                header_seen = True
                others = _fold(" ".join(str(c) for c in row[1:] if c))
                if (any(re.search(rf"\b{re.escape(w)}", _fold(sc)) for w in header_words)
                        or any(re.search(rf"\b{re.escape(w)}", others) for w in header_words)):
                    continue
            ro = str(row[1]).strip() if row[1] and str(row[1]).strip() else None
            en = str(row[2]).strip() if row[2] and str(row[2]).strip() else None
            if sc and (ro or en):
                entries.append((sc, ro, en))
    finally:
        wb.close()
    return entries


def write_names_template(path: Path) -> None:
    """Creează șablonul de import: foaia „Import" (activă, cu antet) + „Exemplu"."""
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Import"
    ws.append(("Denumire stiintifica", "Denumire populara RO", "Denumire populara EN"))
    example = wb.create_sheet("Exemplu")
    example.append(("FOAIA EXEMPLU NU SE IMPORTĂ — se citește doar foaia „Import", "", ""))
    example.append(("Rosa canina", "Măceș", "dog rose"))
    example.append(("Passer domesticus", "Vrabie de casă", "house sparrow"))
    example.append(("Amanita muscaria", "Burete roșu", "fly agaric"))
    wb.active = 0
    wb.save(path)


# ============================================================
# PATHS
# ============================================================

APP_ROOT = Path(__file__).resolve().parent.parent

CONFIG_DIR = APP_ROOT / "config"
OUTPUT_DIR = APP_ROOT / "output"
PROMPTS_DIR = APP_ROOT / "prompts"

ROMANIAN_COUNTIES = [
    "Alba",
    "Arad",
    "Argeș",
    "Bacău",
    "Bihor",
    "Bistrița-Năsăud",
    "Botoșani",
    "Brașov",
    "Brăila",
    "București",
    "Buzău",
    "Caraș-Severin",
    "Călărași",
    "Cluj",
    "Constanța",
    "Covasna",
    "Dâmbovița",
    "Dolj",
    "Galați",
    "Giurgiu",
    "Gorj",
    "Harghita",
    "Hunedoara",
    "Ialomița",
    "Iași",
    "Ilfov",
    "Maramureș",
    "Mehedinți",
    "Mureș",
    "Neamț",
    "Olt",
    "Prahova",
    "Sălaj",
    "Satu Mare",
    "Sibiu",
    "Suceava",
    "Teleorman",
    "Timiș",
    "Tulcea",
    "Vaslui",
    "Vâlcea",
    "Vrancea",
]

# ============================================================
# SETTINGS
# ============================================================

def load_settings(settings_path: Path) -> dict:
    """
    Citește setările dintr-un fișier KEY=value.
    """

    if not settings_path.is_file():
        return {}

    settings = {}

    for line in settings_path.read_text(
        encoding="utf-8"
    ).splitlines():

        line = line.strip()

        if not line or line.startswith("#"):
            continue

        if "=" not in line:
            continue

        key, value = line.split("=", 1)

        settings[key.strip()] = value.strip()

    return settings


def save_setting(settings_path: Path, key: str, value: str) -> None:
    """
    Actualizează (sau adaugă) o singură cheie KEY=value în fișierul
    de setări, păstrând restul liniilor și comentariile existente.
    """

    settings_path = Path(settings_path)
    value = str(value)

    if settings_path.is_file():
        lines = settings_path.read_text(
            encoding="utf-8"
        ).splitlines()
    else:
        lines = []

    new_lines = []
    updated = False

    for line in lines:

        stripped = line.strip()

        if (
            stripped
            and not stripped.startswith("#")
            and "=" in line
            and line.split("=", 1)[0].strip() == key
        ):
            new_lines.append(f"{key}={value}")
            updated = True
            continue

        new_lines.append(line)

    if not updated:
        new_lines.append(f"{key}={value}")

    settings_path.write_text(
        "\n".join(new_lines) + "\n",
        encoding="utf-8",
    )


def load_prompt_file(prompt_path: Path) -> str:
    """
    Citește un prompt dintr-un fișier text UTF-8.
    """

    if not prompt_path.is_file():
        return ""

    return prompt_path.read_text(
        encoding="utf-8"
    ).strip()


def normalize_preset_name(value: str | None) -> str:
    if not value:
        return ""
    return Path(str(value).strip()).stem.casefold()


def fast_predict_limit(value: str | None, configured: int) -> int:
    ceiling = 1024 if normalize_preset_name(value) == "botanic_quick" else 512
    return min(configured, ceiling) if configured > 0 else ceiling


# ============================================================
# MAIN WINDOW
# ============================================================

class MainWindow:

    def __init__(self, root: tk.Tk):

        self.root = root

        self.root.title("Specio Identify")
        self.root.geometry("1100x800")
        self.root.minsize(900, 650)

        # ----------------------------------------------------
        # Settings
        # ----------------------------------------------------

        settings = load_settings(
            CONFIG_DIR / "settings.txt"
        )

        # ----------------------------------------------------
        # Limbă UI (persistată în settings.txt: LANGUAGE=ro|en)
        # ----------------------------------------------------

        set_language(
            settings.get("LANGUAGE", "ro")
        )

        self.ollama_url = settings.get(
            "OLLAMA_URL",
            "http://localhost:11434"
        )

        self.model = settings.get(
            "MODEL",
            "muse-glimmer"
        )

        try:
            self.num_ctx = int(
                settings.get("NUM_CTX", "8192")
            )
        except ValueError:
            self.num_ctx = 8192

        try:
            self.num_predict = int(
                settings.get("NUM_PREDICT", "4000")
            )      
        except ValueError:
            self.num_predict = 4000

        try:
            self.thinking = (
                settings.get("THINKING", "true").lower()
                in ("1", "true", "yes", "on")
            )            
        except ValueError:
            self.thinking = 4000

        # ----------------------------------------------------
        # Setări separate pentru generărire profilului speciei
        # (text-only, la nivel de specie). Cantitate mai mică
        # de tokenuri => răspuns mai rapid, la exactitudine.
        # ----------------------------------------------------

        try:
            self.profile_num_predict = int(
                settings.get("PROFILE_NUM_PREDICT", "4000")
            )
        except ValueError:
            self.profile_num_predict = 4000

        try:
            self.profile_thinking = (
                settings.get("PROFILE_THINKING", "false").lower()
                in ("1", "true", "yes", "on")
            )
        except ValueError:
            self.profile_thinking = False

        # Model text-only pentru profilul speciei (ales din drop-down;
        # gol în settings => folosește MODEL).
        self.profile_model = (
            (settings.get("PROFILE_MODEL", "") or "").strip()
            or self.model
        )

        # ----------------------------------------------------
        # State
        # ----------------------------------------------------

        self.image_paths = []
        self.image_status = {}
        self.image_locations = {}
        self.image_path: Path | None = None
        self.image_metadata = {}

        # Panoul Output are două zone: Item (stânga) și Species Profile
        # (dreapta), fiecare cu conținutul și butonul lui de salvare.
        self.profile_response_text = ""
        self.profile_is_rendered = False

        # Editorul Species Prompt (sus, lângă Prompt) — șablonul
        # editabil al profilului speciei, salvat în baza de date.
        self.species_prompt_raw_text = ""
        self.species_prompt_is_rendered = False
        self.image_base64 = None

        self.response_text = ""
        self.thinking_text = ""
        self.statistics = {}
        self.prompt_raw_text = ""
        self.prompt_is_rendered = False
        self.output_is_rendered = False
        self.text_zoom_size = 10

        self.preview_image = None
        self.analysis_start_time = None
        self.profile_start_time = None
        self.status_update_timer = None
        self.batch_current = 0
        self.batch_total = 0
        self.batch_filename = ""

        # ----------------------------------------------------
        # Baza de date — inițializare centrală într-un singur loc
        # (creează întrega structură: catalog_items, species_profiles,
        #  prompts, plus migrările — vezi core/database.py)
        # ----------------------------------------------------

        connect(
            OUTPUT_DIR / "catalog.db"
        )

        # ----------------------------------------------------
        # GUI
        # ----------------------------------------------------
        
        self.build_ui()

        self.load_default_prompt()

        # Folderul input/ se (re)creează la pornire dacă lipsește, astfel încât
        # utilizatorul are mereu unde pune fotografiile, iar dialogul Add
        # Images îl deschide direct. Conținutul folderului este ignorat prin
        # propriul .gitignore imbricat; nu există placeholder de șters.
        (APP_ROOT / "input").mkdir(parents=True, exist_ok=True)
        # name_sources/ găzduiește PDF-urile pentru „Extract denumiri";
        # îl (re)creăm la pornire ca dialogul să aibă mereu destinația.
        (APP_ROOT / "name_sources").mkdir(parents=True, exist_ok=True)

        self.set_status("Ready")

        # Verificare la pornire: ce lipsește față de un setup complet?
        # Arată un singur popup informativ (nu la fiecare pornire dacă
        # utilizatorul îl ignoră — doar când lipsește ceva esențial).
        self.root.after(600, self.check_first_run_setup)

    # ========================================================
    # UI
    # ========================================================

    def build_ui(self):

        # ----------------------------------------------------
        # Main container
        # ----------------------------------------------------

        main = ttk.Frame(
            self.root,
            padding=12
        )

        main.pack(
            fill="both",
            expand=True
        )

        # ----------------------------------------------------
        # Title
        # ----------------------------------------------------

        title = ttk.Label(
            main,
            text="Specio Identify",
            font=("Segoe UI", 20, "bold")
        )

        title.pack(
            anchor="w",
            pady=(0, 10)
        )


        # ----------------------------------------------------
        # Top section: Image + Prompt
        # ----------------------------------------------------

        top_paned = ttk.Panedwindow(
            main,
            orient="horizontal"
        )

        top_paned.pack(
            fill="x",
            expand=False,
            pady=(0, 10)
        )


        # ----------------------------------------------------
        # Image section
        # ----------------------------------------------------

        image_frame = ttk.LabelFrame(
            top_paned,
            text=t("panel_images"),
            padding=10
        )

        top_paned.add(image_frame)

        # ----------------------------------------------------
        # Image list
        # ----------------------------------------------------

        list_frame = ttk.Frame(
            image_frame
        )

        list_frame.pack(
            fill="x"
        )

        self.image_listbox = tk.Listbox(
            list_frame,
            height=5,
            selectmode=tk.SINGLE,
            font=("Segoe UI", 10)
        )

        self.image_listbox.pack(
            side="left",
            fill="both",
            expand=True
        )

        image_scroll = ttk.Scrollbar(
            list_frame,
            orient="vertical",
            command=self.image_listbox.yview
        )

        image_scroll.pack(
            side="right",
            fill="y"
        )

        self.image_listbox.configure(
            yscrollcommand=image_scroll.set
        )

        self.image_listbox.bind(
            "<<ListboxSelect>>",
            self.on_image_selected
        )

        # ----------------------------------------------------
        # Image buttons
        # ----------------------------------------------------

        image_buttons = ttk.Frame(
            image_frame
        )

        image_buttons.pack(
            fill="x",
            pady=(8, 0)
        )

        self.add_images_button = ttk.Button(
            image_buttons,
            text="Add Images...",
            command=self.add_images
        )

        self.add_images_button.pack(
            side="left"
        )

        self.add_folder_button = ttk.Button(
            image_buttons,
            text="Add Folder...",
            command=self.add_folder
        )

        self.add_folder_button.pack(
            side="left",
            padx=(8, 0)
        )

        self.remove_image_button = ttk.Button(
            image_buttons,
            text="Remove",
            command=self.remove_selected_image
        )

        self.remove_image_button.pack(
            side="left",
            padx=(8, 0)
        )

        self.clear_images_button = ttk.Button(
            image_buttons,
            text="Clear",
            command=self.clear_images
        )

        self.clear_images_button.pack(
            side="left",
            padx=(8, 0)
        )

        # ----------------------------------------------------
        # Preview + Photo Info
        # ----------------------------------------------------

        preview_info_frame = ttk.Frame(
            image_frame
        )

        preview_info_frame.pack(
            fill="x",
            pady=(10, 0)
        )

        # ----------------------------------------------------
        # Preview
        # ----------------------------------------------------

        preview_frame = ttk.Frame(
            preview_info_frame
        )

        preview_frame.pack(
            side="left",
            fill="both",
            expand=True
        )

        self.preview_label = ttk.Label(
            preview_frame,
            text="No image selected",
            anchor="center"
        )

        self.preview_label.pack(
            fill="both",
            expand=True
        )

        # ----------------------------------------------------
        # Photo Info
        # ----------------------------------------------------

        photo_info_frame = ttk.LabelFrame(
            preview_info_frame,
            text="Photo Info",
            padding=10
        )

        photo_info_frame.pack(
            side="left",
            fill="both",
            padx=(15, 0)
        )

        self.photo_file_label = ttk.Label(
            photo_info_frame,
            text="File: —",
            anchor="w"
        )

        self.photo_file_label.pack(
            anchor="w",
            pady=(0, 6)
        )

        self.photo_date_label = ttk.Label(
            photo_info_frame,
            text="Date: —",
            anchor="w"
        )

        self.photo_date_label.pack(
            anchor="w",
            pady=(0, 6)
        )

        self.photo_camera_label = ttk.Label(
            photo_info_frame,
            text="Camera: —",
            anchor="w"
        )

        self.photo_camera_label.pack(
            anchor="w",
            pady=(0, 6)
        )

        self.photo_model_label = ttk.Label(
            photo_info_frame,
            text="Model: —",
            anchor="w"
        )

        self.photo_model_label.pack(
            anchor="w",
            pady=(0, 6)
        )

        self.photo_location_label = ttk.Label(
            photo_info_frame,
            text="GPS: —",
            anchor="w"
        )

        self.photo_location_label.pack(
            anchor="w",
            pady=(0, 8)
        )

        ttk.Label(
            photo_info_frame,
            text="Location:"
        ).pack(
            anchor="w"
        )

        self.location_var = tk.StringVar(
            value="București"
        )

        self.location_combo = ttk.Combobox(
            photo_info_frame,
            textvariable=self.location_var,
            values=ROMANIAN_COUNTIES + ["__ALTCEVA__"],
            state="readonly",
            width=24
        )

        self.location_combo.pack(
            anchor="w"
        )

        self.location_combo.bind(
            "<<ComboboxSelected>>",
            self.on_location_selected
        )

        self.custom_location_var = tk.StringVar()

        self.custom_location_frame = ttk.Frame(
            photo_info_frame
        )

        ttk.Label(
            self.custom_location_frame,
            text="Localitate / locație:"
        ).pack(
            anchor="w",
            pady=(6, 2)
        )

        self.custom_location_entry = ttk.Entry(
            self.custom_location_frame,
            textvariable=self.custom_location_var,
            width=27
        )

        self.custom_location_entry.pack(
            anchor="w"
        )

        self.custom_location_entry.bind(
            "<FocusOut>",
            self.on_custom_location_changed
        )

        self.custom_location_entry.bind(
            "<Return>",
            self.on_custom_location_changed
        )

        # ----------------------------------------------------
        # Prompt section
        # ----------------------------------------------------

        prompt_frame = ttk.LabelFrame(
            top_paned,
            text=t("panel_prompt"),
            padding=10
        )

        top_paned.add(prompt_frame)

        # Prompt selector
        prompt_controls = ttk.Frame(
            prompt_frame
        )

        prompt_controls.pack(
            fill="x",
            pady=(0, 8)
        )

        ttk.Label(
            prompt_controls,
            text="Preset:"
        ).pack(
            side="left"
        )

        self.prompt_var = tk.StringVar()

        self.prompt_combo = ttk.Combobox(
            prompt_controls,
            textvariable=self.prompt_var,
            state="readonly",
            width=30
        )

        self.prompt_combo.pack(
            side="left",
            padx=(8, 8)
        )

        self.new_prompt_button = ttk.Button(
            prompt_controls,
            text="New Prompt",
            command=self.new_prompt
        )

        self.new_prompt_button.pack(
            side="left",
            padx=(0, 6)
        )

        self.update_prompt_button = ttk.Button(
            prompt_controls,
            text="Update",
            command=self.update_prompt
        )

        self.update_prompt_button.pack(
            side="left"
        )

        self.prompt_combo.bind(
            "<<ComboboxSelected>>",
            self.on_prompt_selected
        )

        # Prompt editor        
        self.prompt_text = tk.Text(
            prompt_frame,
            height=7,
            wrap="word",
            font=("Segoe UI", 10)
        )
        self.configure_rich_text(
            self.prompt_text
        )
        self.prompt_text.bind(
            "<FocusIn>",
            self.edit_prompt_text
        )
        self.prompt_text.bind(
            "<FocusOut>",
            self.render_prompt_text
        )

        self.prompt_text.pack(
            fill="both",
            expand=True
        )

        self.populate_prompt_list()

        # ----------------------------------------------------
        # Species Prompt (în oglinda zonei Prompt; editabil,
        # salvat în baza de date ca "species_profile_deep")
        # ----------------------------------------------------

        species_prompt_frame = ttk.LabelFrame(
            top_paned,
            text=t("panel_species_prompt"),
            padding=10
        )

        top_paned.add(species_prompt_frame)

        species_prompt_controls = ttk.Frame(
            species_prompt_frame
        )

        species_prompt_controls.pack(
            fill="x",
            pady=(0, 8)
        )

        self.update_species_prompt_button = ttk.Button(
            species_prompt_controls,
            text="Update",
            command=self.update_species_prompt
        )

        self.update_species_prompt_button.pack(
            side="left"
        )

        ttk.Label(
            species_prompt_controls,
            text="species_profile_deep (din baza de date)"
        ).pack(
            side="left",
            padx=(8, 0)
        )

        self.species_prompt_text = tk.Text(
            species_prompt_frame,
            height=7,
            wrap="word",
            font=("Segoe UI", 10)
        )
        self.configure_rich_text(
            self.species_prompt_text
        )
        self.species_prompt_text.bind(
            "<FocusIn>",
            self.edit_species_prompt_text
        )
        self.species_prompt_text.bind(
            "<FocusOut>",
            self.render_species_prompt_text
        )

        self.species_prompt_text.pack(
            fill="both",
            expand=True
        )

        self.set_species_prompt_text(
            self.load_species_prompt_template()
        )


        # ----------------------------------------------------
        # Buttons
        # ----------------------------------------------------

        button_frame = ttk.Frame(main)

        button_frame.pack(
            fill="x",
            pady=(0, 10)
        )

        # ----------------------------------------------------
        # Grup stânga: fluxul ITEM (analiza imaginii)
        # ----------------------------------------------------

        analysis_buttons = ttk.Frame(button_frame)
        analysis_buttons.pack(side="left")

        self.analyze_button = ttk.Button(
            analysis_buttons,
            text=t("btn_analyze"),
            command=self.start_analysis
        )

        self.analyze_button.pack(
            side="left"
        )

        self.analyze_all_button = ttk.Button(
            analysis_buttons,
            text=t("btn_analyze_all"),
            command=self.start_batch_analysis
        )

        self.analyze_all_button.pack(
            side="left",
            padx=(8, 0)
        )

        self.save_button = ttk.Button(
            analysis_buttons,
            text=t("btn_save"),
            command=self.save_current_result,
            state="disabled"
        )

        self.save_button.pack(
            side="left",
            padx=(8, 0)
        )

        self.open_species_button = ttk.Button(
            analysis_buttons,
            text=t("btn_open_species_folder"),
            command=self.open_species_folder,
            state="disabled"
        )
        self.open_species_button.pack(side="left", padx=(8, 0))

        self.quick_catalog_button = ttk.Button(
            analysis_buttons,
            text=t("btn_quick_catalog"),
            command=self.open_quick_catalog_dialog,
            state="disabled",
        )
        self.quick_catalog_button.pack(side="left", padx=(8, 0))

        self.update_species_button = ttk.Button(
            analysis_buttons,
            text=t("btn_update_species"),
            command=self.update_species_classification,
            state="disabled"
        )
        self.update_species_button.pack(side="left", padx=(8, 0))

        # ----------------------------------------------------
        # Grup dreapta: fluxul SPECIES PROFILE
        # ----------------------------------------------------

        profile_buttons = ttk.Frame(button_frame)
        profile_buttons.pack(side="right")

        self.generate_profile_button = ttk.Button(
            profile_buttons,
            text=t("btn_gen_profile"),
            command=self.generate_species_profile,
            state="disabled"
        )
        self.generate_profile_button.pack(side="left")

        self.load_profile_button = ttk.Button(
            profile_buttons,
            text=t("btn_load_profile"),
            command=self.load_saved_profile,
            state="disabled"
        )
        self.load_profile_button.pack(side="left", padx=(8, 0))

        self.save_profile_button = ttk.Button(
            profile_buttons,
            text=t("btn_save_profile"),
            command=self.save_current_profile,
            state="disabled"
        )
        self.save_profile_button.pack(side="left", padx=(8, 0))

        self.generate_all_profiles_button = ttk.Button(
            profile_buttons,
            text=t("btn_gen_profile_all"),
            command=self.start_batch_profile_generation
        )
        self.generate_all_profiles_button.pack(side="left", padx=(8, 0))

        # ----------------------------------------------------
        # Model pentru profilul speciei (text-only, mai rapid).
        # Lista se încarcă de pe serverul Ollama în fundal.
        # ----------------------------------------------------

        self.profile_model_label = ttk.Label(
            profile_buttons,
            text=t("profile_model_label")
        )
        self.profile_model_label.pack(side="left", padx=(12, 2))

        self.profile_model_var = tk.StringVar(
            value=self.profile_model
        )

        self.profile_model_combo = ttk.Combobox(
            profile_buttons,
            textvariable=self.profile_model_var,
            state="readonly",
            width=18
        )
        self.profile_model_combo.pack(side="left", padx=(0, 4))

        self.profile_model_combo.bind(
            "<<ComboboxSelected>>",
            self.on_profile_model_selected
        )

        self.refresh_profile_models()

        # ----------------------------------------------------
        # Progress
        # ----------------------------------------------------

        self.progress = ttk.Progressbar(
            button_frame,
            mode="indeterminate",
            length=150
        )

        self.progress.pack(
            side="right"
        )

        # ----------------------------------------------------
        # Batch progress group
        # ----------------------------------------------------

        self.batch_progress_frame = ttk.Frame(
            button_frame
        )

        self.batch_progress_label = ttk.Label(
            self.batch_progress_frame,
            text=t("batch_progress").format(cur=0, total=0)
        )

        self.batch_progress_label.pack(
            side="left",
            padx=(0, 8)
        )

        self.batch_progress = ttk.Progressbar(
            self.batch_progress_frame,
            mode="determinate",
            length=150
        )

        self.batch_progress.pack(
            side="left"
        )

        # Hidden until ANALYZE ALL
        self.batch_progress_frame.pack_forget()

        # ----------------------------------------------------
        # Status bar
        # ----------------------------------------------------

        status_frame = ttk.Frame(main)

        status_frame.pack(
            side="bottom",
            fill="x",
            pady=(8, 0)
        )

        self.status_label = ttk.Label(
            status_frame,
            text=t("status_ready")
        )

        self.status_label.pack(
            side="left"
        )

        # Comutator de limbă (live, persistat în settings.txt)
        self.lang_var = tk.StringVar(value=get_language())

        ttk.Label(
            status_frame,
            text=t("lang_label")
        ).pack(side="right", padx=(8, 0))

        self.lang_combo = ttk.Combobox(
            status_frame,
            textvariable=self.lang_var,
            values=list(SUPPORTED_LANGUAGES),
            width=4,
            state="readonly"
        )
        self.lang_combo.pack(side="right")
        self.lang_combo.bind(
            "<<ComboboxSelected>>",
            self.on_language_changed
        )

        self.help_button = ttk.Button(
            status_frame, text=t("help_button"), command=self.show_help
        )
        self.help_button.pack(side="right", padx=(8, 8))

        self.model_label = ttk.Label(
            status_frame,
            text=f"Model: {self.model}"
        )

        self.model_label.pack(
            side="right",
            padx=(0, 10)
        )

        # ----------------------------------------------------
        # Administrare (ascuns implicit, show/hide)
        # ----------------------------------------------------

        self.admin_frame = ttk.LabelFrame(
            main,
            text=t("admin_title"),
            padding=10
        )
        # Informativ: initial ascuns; se poate toggle cu metoda de mai jos
        self.admin_visible = False
        self.admin_frame.pack_forget()

        admin_inner = ttk.Frame(self.admin_frame)
        admin_inner.pack(fill="x")

        self.admin_toggle_button = ttk.Button(
            status_frame,
            text=t("admin_show"),
            command=self.toggle_admin,
            width=14
        )
        self.admin_toggle_button.pack(side="left", padx=(12, 0))

        # Butoane administrare (ordinea naturală a fluxului)
        ttk.Button(
            admin_inner,
            text=t("btn_rebuild_taxonomy"),
            command=self.admin_rebuild_taxonomy,
            width=28
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            admin_inner,
            text=t("btn_import_vernacular"),
            command=self.admin_import_vernacular,
            width=26
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            admin_inner,
            text=t("btn_extract_book"),
            command=self.admin_extract_book,
            width=26
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            admin_inner,
            text=t("btn_add_common_names"),
            command=self.admin_add_common_names,
            width=26
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            admin_inner,
            text=t("btn_parse_rules"),
            command=self.admin_edit_parse_rules,
            width=16
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            admin_inner,
            text=t("btn_open_col"),
            command=self.admin_open_col_folder,
            width=20
        ).pack(side="left")

        self.admin_status = ttk.Label(
            admin_inner,
            text=""
        )
        self.admin_status.pack(side="left", padx=(12, 0))


        # ----------------------------------------------------
        # Output section — două panouri: Item | Species Profile
        # ----------------------------------------------------

        output_paned = ttk.Panedwindow(
            main,
            orient="horizontal"
        )

        output_paned.pack(
            fill="both",
            expand=True,
            side="top"
        )

        # ----------------------------------------------------
        # Panoul stâng: Output Item (analiza imaginii)
        # ----------------------------------------------------

        output_frame = ttk.LabelFrame(
            output_paned,
            text=t("panel_output_item"),
            padding=10
        )

        self.output_text = tk.Text(
            output_frame,
            wrap="word",
            font=("Segoe UI", 10)
        )
        self.configure_rich_text(
            self.output_text
        )
        self.output_text.bind(
            "<FocusIn>",
            self.edit_output_text
        )
        self.output_text.bind(
            "<FocusOut>",
            self.render_output_text
        )

        output_scroll = ttk.Scrollbar(
            output_frame,
            orient="vertical",
            command=self.output_text.yview
        )

        output_scroll.pack(
            side="right",
            fill="y"
        )

        self.output_text.pack(
            side="left",
            fill="both",
            expand=True
        )

        self.output_text.configure(
            yscrollcommand=output_scroll.set
        )

        # ----------------------------------------------------
        # Panoul drept: Output Species Profile
        # ----------------------------------------------------

        profile_frame = ttk.LabelFrame(
            output_paned,
            text=t("panel_output_profile"),
            padding=10
        )

        self.profile_output_text = tk.Text(
            profile_frame,
            wrap="word",
            font=("Segoe UI", 10)
        )
        self.configure_rich_text(
            self.profile_output_text
        )
        self.profile_output_text.bind(
            "<FocusIn>",
            self.edit_profile_output_text
        )
        self.profile_output_text.bind(
            "<FocusOut>",
            self.render_profile_output_text
        )

        profile_scroll = ttk.Scrollbar(
            profile_frame,
            orient="vertical",
            command=self.profile_output_text.yview
        )

        profile_scroll.pack(
            side="right",
            fill="y"
        )

        self.profile_output_text.pack(
            side="left",
            fill="both",
            expand=True
        )

        self.profile_output_text.configure(
            yscrollcommand=profile_scroll.set
        )

        output_paned.add(output_frame, weight=1)
        output_paned.add(profile_frame, weight=1)

        self.set_profile_output_text(
            "Selectează o imagine; profilul speciei va apărea aici.",
            saveable=False,
        )

    def configure_rich_text(
        self,
        widget: tk.Text
    ):
        """
        Configurează taguri vizuale pentru Markdown simplu.
        """

        self.apply_text_zoom(widget)

        widget.bind(
            "<Control-MouseWheel>",
            self.zoom_text_areas
        )

        widget.bind(
            "<Control-Button-4>",
            self.zoom_text_areas
        )

        widget.bind(
            "<Control-Button-5>",
            self.zoom_text_areas
        )

    def apply_text_zoom(
        self,
        widget: tk.Text
    ):
        """
        Aplică dimensiunea curentă a fontului pe zona text și taguri.
        """

        size = self.text_zoom_size

        body_font = tkfont.Font(
            family="Segoe UI",
            size=size
        )
        bold_font = tkfont.Font(
            family="Segoe UI",
            size=size,
            weight="bold"
        )
        italic_font = tkfont.Font(
            family="Segoe UI",
            size=size,
            slant="italic"
        )
        h1_font = tkfont.Font(
            family="Segoe UI",
            size=size + 5,
            weight="bold"
        )
        h2_font = tkfont.Font(
            family="Segoe UI",
            size=size + 3,
            weight="bold"
        )
        h3_font = tkfont.Font(
            family="Segoe UI",
            size=size + 1,
            weight="bold"
        )

        widget._rich_text_fonts = (
            body_font,
            bold_font,
            italic_font,
            h1_font,
            h2_font,
            h3_font
        )

        widget.configure(
            font=body_font
        )

        widget.tag_configure(
            "h1",
            font=h1_font,
            spacing1=10,
            spacing3=6
        )

        widget.tag_configure(
            "h2",
            font=h2_font,
            spacing1=8,
            spacing3=5
        )

        widget.tag_configure(
            "h3",
            font=h3_font,
            spacing1=6,
            spacing3=4
        )

        widget.tag_configure(
            "bold",
            font=bold_font
        )

        widget.tag_configure(
            "italic",
            font=italic_font
        )

        widget.tag_configure(
            "body",
            spacing1=2,
            spacing3=2
        )

        widget.tag_configure(
            "list",
            lmargin1=0,
            lmargin2=0,
            spacing1=2,
            spacing3=2
        )

        widget.tag_configure(
            "rule",
            foreground="#777777",
            spacing1=4,
            spacing3=4
        )

    def zoom_text_areas(self, event):
        """
        Ctrl + rotița mouse-ului mărește sau micșorează Prompt și Output.
        """

        if getattr(event, "num", None) == 4 or getattr(event, "delta", 0) > 0:
            delta = 1
        else:
            delta = -1

        new_size = max(
            8,
            min(
                24,
                self.text_zoom_size + delta
            )
        )

        if new_size == self.text_zoom_size:
            return "break"

        self.text_zoom_size = new_size

        self.apply_text_zoom(
            self.prompt_text
        )
        self.apply_text_zoom(
            self.species_prompt_text
        )
        self.apply_text_zoom(
            self.output_text
        )
        self.apply_text_zoom(
            self.profile_output_text
        )

        return "break"

    def set_rich_text(
        self,
        widget: tk.Text,
        text: str
    ):
        """
        Înlocuiește conținutul și aplică formatarea vizuală.
        """

        widget.delete(
            "1.0",
            tk.END
        )

        self.insert_rich_text(
            widget,
            text
        )

    def set_plain_text(
        self,
        widget: tk.Text,
        text: str
    ):
        """
        Înlocuiește conținutul fără să modifice textul original.
        """

        widget.delete(
            "1.0",
            tk.END
        )

        widget.insert(
            "1.0",
            text
        )

    def set_prompt_text(
        self,
        text: str,
        render: bool = True
    ):
        """
        Păstrează sursa promptului și afișează opțional varianta formatată.
        """

        self.prompt_raw_text = text

        if render and text:
            self.set_rich_text(
                self.prompt_text,
                text
            )
            self.prompt_is_rendered = True
            return

        self.set_plain_text(
            self.prompt_text,
            text
        )
        self.prompt_is_rendered = False

    def get_prompt_text(self) -> str:
        """
        Returnează promptul sursă, nu textul curățat pentru afișare.
        """

        if self.prompt_is_rendered:
            return self.prompt_raw_text.strip()

        return self.prompt_text.get(
            "1.0",
            tk.END
        ).strip()

    def edit_prompt_text(self, _event=None):
        """
        La editare, arată Markdown-ul original.
        """

        if not self.prompt_is_rendered:
            return

        self.set_plain_text(
            self.prompt_text,
            self.prompt_raw_text
        )
        self.prompt_is_rendered = False

    def render_prompt_text(self, _event=None):
        """
        La ieșirea din câmp, păstrează sursa și afișează formatat.
        """

        if self.prompt_is_rendered:
            return

        text = self.prompt_text.get(
            "1.0",
            tk.END
        ).strip()

        self.set_prompt_text(
            text,
            render=True
        )

    def set_output_text(
        self,
        text: str,
        render: bool = True,
        saveable: bool = True
    ):
        """
        Păstrează sursa outputului și afișează opțional varianta formatată.
        """

        self.response_text = text if saveable else ""

        if render and text:
            self.set_rich_text(
                self.output_text,
                text
            )
            self.output_is_rendered = saveable
        else:
            self.set_plain_text(
                self.output_text,
                text
            )
            self.output_is_rendered = False

        self.update_save_button_state()

    def get_output_text(self) -> str:
        """
        Returnează outputul sursă, inclusiv modificările utilizatorului.
        """

        if self.output_is_rendered:
            return self.response_text.strip()

        return self.output_text.get(
            "1.0",
            tk.END
        ).strip()

    def edit_output_text(self, _event=None):
        """
        La editare, arată textul original al rezultatului.
        """

        if not self.output_is_rendered:
            return

        self.set_plain_text(
            self.output_text,
            self.response_text
        )
        self.output_is_rendered = False

    def render_output_text(self, _event=None):
        """
        La ieșirea din câmp, păstrează modificările și afișează formatat.
        """

        if self.output_is_rendered:
            return

        text = self.output_text.get(
            "1.0",
            tk.END
        ).strip()

        if not text:
            self.set_output_text(
                "",
                render=False,
                saveable=False
            )
            return

        self.set_output_text(
            text,
            render=True,
            saveable=True
        )

    def set_profile_output_text(
        self,
        text: str,
        render: bool = True,
        saveable: bool = True
    ):
        """
        Păstrează sursa profilului și afișează opțional varianta formatată.
        """

        self.profile_response_text = text if saveable else ""

        if render and text:
            self.set_rich_text(
                self.profile_output_text,
                text
            )
            self.profile_is_rendered = saveable
        else:
            self.set_plain_text(
                self.profile_output_text,
                text
            )
            self.profile_is_rendered = False

        self.update_save_profile_button_state()

    def get_profile_output_text(self) -> str:
        """
        Returnează textul sursă al profilului, inclusiv modificările
        utilizatorului.
        """

        if self.profile_is_rendered:
            return self.profile_response_text.strip()

        return self.profile_output_text.get(
            "1.0",
            tk.END
        ).strip()

    def edit_profile_output_text(self, _event=None):
        """
        La editare, arată textul original al profilului.
        """

        if not self.profile_is_rendered:
            return

        self.set_plain_text(
            self.profile_output_text,
            self.profile_response_text
        )
        self.profile_is_rendered = False

    def render_profile_output_text(self, _event=None):
        """
        La ieșirea din câmp, păstrează modificările și afișează formatat.
        """

        if self.profile_is_rendered:
            return

        text = self.profile_output_text.get(
            "1.0",
            tk.END
        ).strip()

        if not text:
            self.set_profile_output_text(
                "",
                render=False,
                saveable=False
            )
            return

        self.set_profile_output_text(
            text,
            render=True,
            saveable=True
        )

    def update_save_profile_button_state(self):
        """
        SAVE PROFILE este activ doar când există profil afișat
        pentru imaginea selectată.
        """

        if not hasattr(self, "save_profile_button"):
            return

        has_result = (
            self.image_path is not None
            and bool(self.profile_response_text.strip())
        )

        self.save_profile_button.configure(
            state="normal" if has_result else "disabled"
        )

    def load_species_prompt_template(self) -> str:
        """
        Încarcă șablonul profilului speciei: întâi din baza de date
        (intrarea "species_profile_deep"), cu fallback la fișierul
        prompts/species_profile_deep.txt și apoi la un text minimal.
        """

        try:
            record = load_prompt_from_db(
                OUTPUT_DIR / "catalog.db",
                SPECIES_PROFILE_PROMPT_NAME,
            )
            if record and (record.get("content") or "").strip():
                return record["content"].strip()
        except Exception:
            pass

        try:
            return load_prompt(
                PROMPTS_DIR / f"{SPECIES_PROFILE_PROMPT_NAME}.txt"
            )
        except (FileNotFoundError, ValueError):
            return (
                "Generează un profil enciclopedic detaliat al speciei "
                "indicate, valabil pentru specia în general, nu pentru "
                "o fotografie anume."
            )

    def set_species_prompt_text(self, text: str, render: bool = True):
        """
        Păstrează sursa șablonului și afișează opțional varianta formatată.
        """

        self.species_prompt_raw_text = text

        if render and text:
            self.set_rich_text(
                self.species_prompt_text,
                text
            )
            self.species_prompt_is_rendered = True
        else:
            self.set_plain_text(
                self.species_prompt_text,
                text
            )
            self.species_prompt_is_rendered = False

    def get_species_prompt_text(self) -> str:
        """
        Returnează șablonul sursă, inclusiv modificările utilizatorului.
        """

        if self.species_prompt_is_rendered:
            return self.species_prompt_raw_text.strip()

        return self.species_prompt_text.get(
            "1.0",
            tk.END
        ).strip()

    def edit_species_prompt_text(self, _event=None):
        """
        La editare, arată Markdown-ul original al șablonului.
        """

        if not self.species_prompt_is_rendered:
            return

        self.set_plain_text(
            self.species_prompt_text,
            self.species_prompt_raw_text
        )
        self.species_prompt_is_rendered = False

    def render_species_prompt_text(self, _event=None):
        """
        La ieșirea din câmp, păstrează modificările și afișează formatat.
        """

        if self.species_prompt_is_rendered:
            return

        text = self.species_prompt_text.get(
            "1.0",
            tk.END
        ).strip()

        self.set_species_prompt_text(
            text,
            render=True
        )

    def update_species_prompt(self):
        """
        Salvează șablonul profilului speciei (editat în panoul
        Species Prompt) în baza de date, sub numele
        "species_profile_deep". De aici este folosit de GEN PROFILE.
        """

        text = self.get_species_prompt_text()

        if not text:
            messagebox.showwarning(
                "Empty prompt",
                "The prompt cannot be empty.",
                parent=self.root,
            )
            return

        try:
            save_prompt_to_db(
                OUTPUT_DIR / "catalog.db",
                SPECIES_PROFILE_PROMPT_NAME,
                text,
            )
        except Exception as exc:
            messagebox.showerror(
                "Update prompt error",
                str(exc),
                parent=self.root,
            )
            return

        self.set_species_prompt_text(text)
        self.set_status("Species prompt updated")

    def build_species_profile_prompt(
        self,
        scientific_name: str,
        ro_name: str = "",
        en_name: str = "",
    ) -> str:
        """
        Construiește promptul complet trimis la GEN PROFILE:
        contextul speciei + șablonul "species_profile_deep" (din baza
        de date, editabil în panoul Species Prompt) + cererea de date
        structurate. Exact acest text este trimis la model la GEN PROFILE.
        """

        template = self.load_species_prompt_template()

        # Preluare denumiri din taxonomy.db (prioritate maximaa)
        try:
            chain_info = taxonomy_lookup_chain(scientific_name)
            if chain_info:
                db_ro = chain_info.get("ro_name") or ""
                db_en = chain_info.get("en_name") or ""
                ro_name = db_ro if db_ro else ro_name
                en_name = db_en if db_en else en_name
        except Exception:
            pass

        context = (
            f"Specia: {scientific_name}\n"
            f"Nume românesc: {ro_name or '—'}\n"
            f"Nume englezesc: {en_name or '—'}\n\n"
        )

        # Date verificate din Catalogue of Life (daca exista cache-ul).
        # Rezolva halucinarea cifrelor [~X] din arbore: modelul primeste
        # exact numele taxonilor si numarul copiilor directi.
        try:
            verified = taxonomy_verified_context(scientific_name)
        except Exception:
            verified = ""

        if verified:
            context = f"{context}{verified}\n\n"

        # Marcajele #română# / #engleză# din șablon se înlocuiesc cu
        # denumirile din taxonomy.db — altfel modelul le copiază literal
        # în profilul generat.
        try:
            template = apply_common_name_markers(template, ro_name, en_name)
        except Exception:
            pass

        return append_species_profile_instruction(
            f"{context}{template}"
        )

    def _load_profile_into_pane(
        self,
        scientific_name: str,
        profile: dict | None,
    ):
        """
        Încarcă profilul în panoul din dreapta, silențios (fără mesaje).
        Folosit la selecția imaginii și de butonul LOAD PROFILE.
        """

        if not profile or not (profile.get("profile_response") or "").strip():
            self.set_profile_output_text(
                "No species profile saved for this species.\n\n"
                "Apasă GEN PROFILE pentru a genera unul.",
                saveable=False,
            )
            return

        stored_response = profile.get("profile_response") or ""

        # Profilurile salvate de versiunile noi conțin textul combinat
        # (antet + card + narativ) exact cum a fost afișat/salvat de
        # utilizator; cele vechi conțineau doar narativul, caz în care
        # îl reformăm cu antet + card de fapte.
        is_combined = any(
            line.strip() == f"*{scientific_name}*"
            for line in stored_response.splitlines()[:3]
        )

        if is_combined:
            combined = stored_response
        else:
            try:
                profile_fields = json.loads(
                    profile.get("profile_fields_json") or "{}"
                )
            except (TypeError, ValueError):
                profile_fields = {}

            combined = self._format_species_profile_output(
                scientific_name=scientific_name,
                display_response=stored_response,
                profile_fields=profile_fields,
                ro_name=profile.get("ro_name") or "",
                en_name=profile.get("en_name") or "",
                category=profile.get("category") or "",
            )

        self.set_profile_output_text(combined)

    def insert_rich_text(
        self,
        widget: tk.Text,
        text: str
    ):
        """
        Redă un subset mic de Markdown în tk.Text.
        """

        for raw_line in text.splitlines():
            line = raw_line.rstrip()
            stripped = line.strip()

            if not stripped:
                widget.insert(
                    tk.END,
                    "\n"
                )
                continue

            # Preserve leading spaces for indentation
            leading_spaces = len(line) - len(line.lstrip())

            if re.fullmatch(r"[-*_]{3,}", stripped):
                if leading_spaces > 0:
                    widget.insert(
                        tk.END,
                        " " * leading_spaces
                    )
                widget.insert(
                    tk.END,
                    "------------------------------\n",
                    ("rule",)
                )
                continue

            heading = re.match(
                r"^(#{1,3})\s+(.*)$",
                stripped
            )

            if heading:
                if leading_spaces > 0:
                    widget.insert(
                        tk.END,
                        " " * leading_spaces
                    )
                level = len(heading.group(1))
                tag = f"h{level}"
                self.insert_inline_rich_text(
                    widget,
                    heading.group(2).strip(),
                    (tag,)
                )
                widget.insert(
                    tk.END,
                    "\n"
                )
                continue

            list_item = re.match(
                r"^([-*+]|\d+[.)])\s+(.*)$",
                stripped
            )

            if list_item:
                if leading_spaces > 0:
                    widget.insert(
                        tk.END,
                        " " * leading_spaces
                    )
                marker = list_item.group(1)
                if marker in {"*", "+"}:
                    marker = "-"

                widget.insert(
                    tk.END,
                    f"{marker} ",
                    ("list",)
                )
                self.insert_inline_rich_text(
                    widget,
                    list_item.group(2).strip(),
                    ("list",)
                )
                widget.insert(
                    tk.END,
                    "\n"
                )
                continue

            if leading_spaces > 0:
                widget.insert(
                    tk.END,
                    " " * leading_spaces
                )
            self.insert_inline_rich_text(
                widget,
                stripped,
                ("body",)
            )
            widget.insert(
                tk.END,
                "\n"
            )

    def insert_inline_rich_text(
        self,
        widget: tk.Text,
        text: str,
        base_tags: tuple[str, ...]
    ):
        """
        Aplică bold și italic pentru Markdown inline simplu.
        """

        pattern = re.compile(
            r"(\*\*|__)(.+?)\1|(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|(?<!_)_(?!_)(.+?)(?<!_)_(?!_)"
        )
        cursor = 0

        for match in pattern.finditer(text):
            if match.start() > cursor:
                widget.insert(
                    tk.END,
                    text[cursor:match.start()],
                    base_tags
                )

            if match.group(1):
                formatted_text = match.group(2)
                formatted_tags = base_tags + ("bold",)
            elif match.group(3) is not None:
                formatted_text = match.group(3)
                formatted_tags = base_tags + ("italic",)
            else:
                formatted_text = match.group(4)
                formatted_tags = base_tags + ("italic",)

            widget.insert(
                tk.END,
                formatted_text,
                formatted_tags
            )

            cursor = match.end()

        if cursor < len(text):
            widget.insert(
                tk.END,
                text[cursor:],
                base_tags
            )

    # ========================================================
    # FIRST-RUN SETUP CHECK
    # ========================================================

    def check_first_run_setup(self):
        """Verifică la pornire ce componente lipsesc și ghidează utilizatorul.

        Arată UN SINGUR popup (nu câte unul pe problemă) cu lista exactă
        a pașilor lipsă + buton de deschidere a secțiunii Administrare.
        Nu blochează aplicația — utilizatorul poate lucra și fără taxonomy.db
        (funcțiile taxonomice folosesc fallback la modelul LLM).
        """
        from core.taxonomy import find_col_zip

        missing: list[str] = []

        # Arhiva este necesară pentru construire, nu pentru pornirea cu o bază validă.
        col_zip = find_col_zip()
        cache_ready = taxonomy_cache_ready()
        if not cache_ready:
            missing.append(t("setup_missing_zip" if col_zip is None else "setup_missing_db"))

        # Denumirile populare se verifică separat de existența arhivei.
        if cache_ready:
            try:
                from core.taxonomy import connect_taxonomy
                con = connect_taxonomy()
                try:
                    row = con.execute(
                        "SELECT COUNT(*) AS c FROM taxa "
                        "WHERE (ro_name IS NOT NULL AND ro_name <> '') "
                        "OR (en_name IS NOT NULL AND en_name <> '')"
                    ).fetchone()
                    if not row or not row["c"]:
                        missing.append(t("setup_missing_names"))
                finally:
                    con.close()
            except Exception:
                pass  # nu blocăm pornirea pentru o verificare auxiliară

        if not missing:
            return

        msg = (
            t("setup_welcome")
            + "\n\n".join(missing)
            + t("setup_readme_hint")
        )
        if messagebox.askyesno(t("setup_title"), msg, parent=self.root):
            if not self.admin_visible:
                self.toggle_admin()
            self._open_col_readme()

    def _open_col_readme(self):
        """Deschide col/README.md în editorul asociat sistemului (.md).

        Fallback: deschide folderul col/ în Explorer dacă fișierul
        nu poate fi deschis direct.
        """
        readme = APP_ROOT / "col" / "README.md"
        try:
            if readme.is_file():
                os.startfile(str(readme))
                return
        except Exception:
            pass
        try:
            COL_DIR.mkdir(parents=True, exist_ok=True)
            os.startfile(str(COL_DIR))
        except Exception as exc:
            messagebox.showerror(
                "Configurare inițială",
                f"Nu pot deschide ghidul:\n{readme}\n\n{exc}",
                parent=self.root,
            )

    # ========================================================
    # PROMPTS
    # ========================================================

    def load_default_prompt(self):
        """
        Încarcă automat promptul implicit la pornirea aplicației.
        Prioritate: baza de date ("00_2.default_standard", apoi legacy
        "default"), apoi fișierul prompts/00_2.default_standard.txt.
        """

        item = load_prompt_from_db(
            OUTPUT_DIR / "catalog.db",
            DEFAULT_PROMPT_NAME,
        )

        if item and item.get("content"):
            self.set_prompt_text(
                item["content"]
            )
            return

        item = load_prompt_from_db(
            OUTPUT_DIR / "catalog.db",
            LEGACY_DEFAULT_PROMPT_NAME,
        )

        if item and item.get("content"):
            self.set_prompt_text(
                item["content"]
            )
            return

        prompt = load_prompt_file(
            PROMPTS_DIR / f"{DEFAULT_PROMPT_NAME}.txt"
        )

        if prompt:
            self.set_prompt_text(
                prompt
            )


    def extract_location_from_description(self, image_path: Path) -> str | None:
        """
        Extrage locația manuală din fișierul _description dacă există.
        Returnează None dacă nu există sau dacă nu este locație manuală.
        """
        item = get_catalog_item(image_path, OUTPUT_DIR)
        return item.get("manual_location") if item else None


    def extract_prompt_from_description(self, image_path: Path) -> str | None:
        """
        Extrage promptul din fișierul _description dacă există.
        Returnează None dacă nu există.
        """
        item = get_catalog_item(image_path, OUTPUT_DIR)
        return item.get("prompt_text") if item else None


    def extract_preset_name_from_description(self, image_path: Path) -> str | None:
        """
        Extrage numele preset-ului din fișierul _description dacă există.
        Returnează None dacă nu există.
        """
        item = get_catalog_item(image_path, OUTPUT_DIR)
        return item.get("preset_name") if item else None


    # ========================================================
    # IMAGES
    # ========================================================

    def add_images(self):

        paths = filedialog.askopenfilenames(
            title="Select images",
            initialdir=str(APP_ROOT / "input"),
            filetypes=[
                (
                    "Image files",
                    "*.jpg *.jpeg *.png *.webp *.bmp *.tif *.tiff"
                ),
                (
                    "All files",
                    "*.*"
                ),
            ]
        )

        if not paths:
            return

        for path in paths:

            path = Path(path)

            if path not in self.image_paths:
                self.image_paths.append(path)

        self.refresh_image_list()

        if self.image_paths:
            self.select_image(0)


    def add_folder(self):

        folder = filedialog.askdirectory(
            title="Select image folder"
        )

        if not folder:
            return

        folder = Path(folder)

        extensions = {
            ".jpg",
            ".jpeg",
            ".png",
            ".webp",
            ".bmp",
            ".tif",
            ".tiff",
        }

        paths = sorted(
            path
            for path in folder.iterdir()
            if path.is_file()
            and path.suffix.lower() in extensions
        )

        added = 0

        for path in paths:

            if path not in self.image_paths:

                self.image_paths.append(path)
                added += 1

        self.refresh_image_list()

        if self.image_paths:

            self.select_image(0)

        self.set_status(
            f"Added {added} image(s)"
        )


    def refresh_image_list(self):

        self.image_listbox.delete(
            0,
            tk.END
        )

        for path in self.image_paths:

            status = self.image_status.get(
                path
            )

            if status is None:
                catalog_item = get_catalog_item(path, OUTPUT_DIR)
                status = "done" if catalog_item else ""

            if status == "done":
                prefix = "✓ "

            elif status == "processing":
                prefix = "... "

            elif status == "error":
                prefix = "✗ "

            else:
                prefix = "   "

            self.image_listbox.insert(
                tk.END,
                prefix + path.name
            )


    def select_image(self, index: int):

        if not self.image_paths:
            return

        if index < 0 or index >= len(self.image_paths):
            return

        self.image_listbox.selection_clear(
            0,
            tk.END
        )

        self.image_listbox.selection_set(
            index
        )

        self.image_listbox.see(
            index
        )

        self.load_selected_image(
            self.image_paths[index]
        )


    def on_image_selected(self, event=None):

        selection = self.image_listbox.curselection()

        if not selection:
            return

        index = selection[0]

        self.load_selected_image(
            self.image_paths[index]
        )


    def load_selected_image(self, path: Path):

        try:

            self.image_path = path
            self.update_catalog_buttons()
            
            # ------------------------------------------------
            # Manual location for this image
            # ------------------------------------------------

            # Locația manuală salvată în catalog (SQLite)
            saved_location = self.extract_location_from_description(path)
            
            if saved_location:
                # Salvează în dicționar pentru utilizare ulterioară
                self.image_locations[path] = saved_location
            else:
                # Fallback la dicționar sau la București
                saved_location = self.image_locations.get(
                    path,
                    "București"
                )

            if saved_location in ROMANIAN_COUNTIES:
                self.location_var.set(
                    saved_location
                )
                self.custom_location_var.set("")
                self.custom_location_frame.pack_forget()
            else:
                self.location_var.set("__ALTCEVA__")
                self.custom_location_var.set(
                    saved_location
                )
                self.custom_location_frame.pack(
                    anchor="w"
                )
            
            # ------------------------------------------------
            # Metadata
            # ------------------------------------------------

            self.image_metadata = get_image_metadata(
                path
            )

            # ------------------------------------------------
            # Photo Info
            # ------------------------------------------------

            photo_date = self.image_metadata.get(
                "photo_date"
            )

            camera_make = self.image_metadata.get(
                "camera_make"
            )

            camera_model = self.image_metadata.get(
                "camera_model"
            )

            latitude = self.image_metadata.get(
                "latitude"
            )

            longitude = self.image_metadata.get(
                "longitude"
            )

            self.photo_file_label.configure(
                text=f"File: {path.name}"
            )

            self.photo_date_label.configure(
                text=f"Date: {photo_date or '—'}"
            )

            self.photo_camera_label.configure(
                text=f"Camera: {camera_make or '—'}"
            )

            self.photo_model_label.configure(
                text=f"Model: {camera_model or '—'}"
            )

            if (
                latitude is not None
                and longitude is not None
            ):
                location_text = (
                    f"{latitude:.6f}, {longitude:.6f}"
                )
            else:
                location_text = "—"

            self.photo_location_label.configure(
                text=f"Location: {location_text}"
            )

            # ------------------------------------------------
            # Base64
            # ------------------------------------------------

            self.image_base64 = image_to_base64(
                path
            )

            # ------------------------------------------------
            # Preview
            # ------------------------------------------------

            with Image.open(path) as image:

                image.thumbnail(
                    (500, 250)
                )

                preview = image.copy()

            self.preview_image = ImageTk.PhotoImage(
                preview
            )

            self.preview_label.configure(
                image=self.preview_image,
                text=""
            )

            self.set_status(
                f"Selected: {path.name}"
            )
            
            self.load_saved_result(path)
            
            # ------------------------------------------------
            # Load prompt from description file if exists
            # ------------------------------------------------
            
            saved_prompt = self.extract_prompt_from_description(path)
            if saved_prompt:
                self.set_prompt_text(saved_prompt)
            else:
                self.load_default_prompt()
            
            # ------------------------------------------------
            # Load preset name from description file if exists
            # ------------------------------------------------
            
            saved_preset_name = self.extract_preset_name_from_description(path)
            if saved_preset_name and saved_preset_name in self.prompt_combo['values']:
                self.prompt_var.set(saved_preset_name)
            elif LEGACY_DEFAULT_PROMPT_NAME in self.prompt_combo['values']:
                self.prompt_var.set(LEGACY_DEFAULT_PROMPT_NAME)

        except Exception as exc:

            self.image_path = None
            self.image_base64 = None

            messagebox.showerror(
                "Image error",
                str(exc)
            )

    def update_catalog_buttons(self):
        state = "disabled"
        profile_state = "disabled"
        load_profile_state = "disabled"

        if self.image_path is not None:
            try:
                item = get_catalog_item(self.image_path, OUTPUT_DIR)
                if item:
                    state = "normal"
                    scientific_name = (item.get("scientific_name") or "").strip()
                    if scientific_name and scientific_name not in ("Necunoscut", "Unknown"):
                        profile_state = "normal"

                        # LOAD PROFILE activ doar dacă există deja
                        # un profil salvat în baza de date.
                        if get_species_profile(OUTPUT_DIR, scientific_name):
                            load_profile_state = "normal"
            except Exception:
                pass

        self.open_species_button.configure(state=state)
        self.update_species_button.configure(state=state)
        self.generate_profile_button.configure(state=profile_state)
        self.generate_all_profiles_button.configure(
            state="normal" if self.image_paths else "disabled"
        )
        self.load_profile_button.configure(state=load_profile_state)
        self.quick_catalog_button.configure(
            state="normal" if self.image_paths else "disabled"
        )

    def update_catalog_buttons_for_path(self, path: Path):
        if self.image_path == path:
            self.update_catalog_buttons()

    def on_profile_model_selected(self, _event=None):
        """
        La alegerea unui model din drop-down, îl folosește pentru
        GEN PROFILE și îl salvează în settings.txt.
        """

        model = self.profile_model_var.get().strip()

        if not model:
            return

        self.profile_model = model

        try:
            save_setting(
                CONFIG_DIR / "settings.txt",
                "PROFILE_MODEL",
                model,
            )
        except OSError:
            pass

        self.set_status(f"Profile model: {model}")

    def refresh_profile_models(self):
        """
        Încarcă lista modelelor disponibile de pe serverul Ollama
        într-un thread separat, fără să blocheze interfața.
        Dacă serverul nu răspunde, rămâne doar modelul curent.

        Notă: thread-ul scrie doar într-o variabilă partajată;
        actualizarea widget-ului se face exclusiv din main thread
        (root.after), ca să evităm apelurile tkinter cross-thread.
        """

        ollama_url = self.ollama_url

        self._models_done = False
        self._models_result = []

        def worker():
            self._models_result = list_models(ollama_url)
            self._models_done = True

        threading.Thread(
            target=worker,
            daemon=True,
        ).start()

        def apply(attempt: int = 0):
            if not self._models_done and attempt < 50:
                # Thread-ul încă lucrează (max ~10s); reîncearcă.
                self.root.after(
                    200,
                    lambda: apply(attempt + 1),
                )
                return

            models = list(self._models_result)
            current = self.profile_model_var.get().strip()

            if models:
                if current and current not in models:
                    models.append(current)
                self.profile_model_combo["values"] = models
            else:
                self.profile_model_combo["values"] = (
                    [current] if current else []
                )

        self.root.after(
            100,
            apply,
        )

    # ========================================================
    # Species profile generation
    # ========================================================

    SPECIES_PROFILE_FIELD_LABELS = {
        "raspandire": "Răspândire",
        "dimensiuni": "Dimensiuni",
        "habitat": "Habitat",
        "cuib": "Cuib",
        "oua": "Ouă",
        "hrana": "Hrană",
        "port": "Port",
        "tulpina": "Tulpina",
        "frunza_dispozitie": "Frunza — dispoziție",
        "frunza_forma": "Frunza — formă",
        "frunza_margine": "Frunza — margine",
        "inflorescenta": "Inflorescență",
        "involucru": "Involucru/Involucel",
        "petale": "Petale",
        "simetrie": "Simetrie",
        "stil": "Stil",
        "ovar": "Ovar",
        "fruct": "Fruct",
        "inflorire": "Înflorire",
        "culoare_floare": "Culoarea florii",
        "polenizare": "Polenizare",
        "fapt_divers": "Fapt divers",
        "acoperire": "Acoperire",
        "comportament": "Comportament",
    }

    SPECIES_PROFILE_FIELD_ORDER = (
        "raspandire",
        "dimensiuni",
        "habitat",
        "cuib",
        "oua",
        "hrana",
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
        "acoperire",
        "comportament",
    )

    # Aceste câmpuri sunt în general prosă (venite în narativul modelului),
    # deci nu le duplicăm în cardul de fapte.
    SPECIES_PROFILE_PROSE_FIELDS = {"descriere", "identificare"}

    def _format_species_profile_output(
        self,
        scientific_name: str,
        display_response: str,
        profile_fields: dict | None = None,
        ro_name: str = "",
        en_name: str = "",
        category: str = "",
    ) -> str:
        """
        Compune textul profilului pentru zona Output:
        antetul + narativul (prosa de model cu cele 8 secțiuni).
        """
        lines = []

        # ----------------------------------------------------
        # Antet
        # ----------------------------------------------------
        names = []
        if ro_name:
            names.append(f"**{ro_name}**")
        if en_name:
            names.append(f"*{en_name}*")
        if names:
            lines.append(" | ".join(names))
        lines.append(f"*{scientific_name}*")
        if category:
            lines.append(f"*{category}*")

        # ----------------------------------------------------
        # Narativul (prosa de model: Identificare, Descriere...)
        # ----------------------------------------------------
        narrative = (display_response or "").strip()

        if narrative:
            lines.append("")
            lines.append("---")
            lines.append("")
            lines.append(narrative)

        return "\n".join(lines).strip()

    def generate_species_profile(self):

        if self.image_path is None:
            return

        item = get_catalog_item(
            self.image_path,
            OUTPUT_DIR,
        )

        if not item:
            messagebox.showinfo(
                "Species profile",
                "Analițează imaginea înainte de a genera profilul speciei.",
                parent=self.root,
            )
            return

        scientific_name = (item.get("scientific_name") or "").strip()

        if not scientific_name or scientific_name in ("Necunoscut", "Unknown"):
            messagebox.showinfo(
                "Species profile",
                "Specia nu este încă identificată.",
                parent=self.root,
            )
            return

        confirmed = messagebox.askyesno(
            "Generate species profile",
            f"Vei genera profilul speciei:\n\n"
            f"{scientific_name}\n\n"
            "Profilul se bazează pe cunoștințele modelului despre specia în general, "
            "nu pe fotografia curentă.\n\n"
            f"Model (text-only): {self.profile_model or self.model}\n\n"
            "Continui?",
            parent=self.root,
        )

        if not confirmed:
            return

        self.analyze_button.configure(state="disabled")
        self.analyze_all_button.configure(state="disabled")
        self.generate_profile_button.configure(state="disabled")

        # ----------------------------------------------------
        # Indicator de progres: progress bar + timer live
        # (asemănat mecanismului ANALYZE)
        # ----------------------------------------------------

        self.progress.pack(
            side="right"
        )
        self.progress.start(
            200
        )

        self.profile_start_time = time.perf_counter()
        self.set_status(
            f"Generating species profile: {scientific_name}..."
        )
        self.update_status_with_time()

        # Placeholder în panoul Species Profile — panoul Item (analiza)
        # rămâne neatins. Imaginea NU este trimisă la model.
        self.set_profile_output_text(
            "Generating species profile...\n\n"
            f"Specia: {scientific_name}\n\n"
            "*Text-only — imaginea nu este trimisă la model.*\n"
            "Aceasta poate dura câteva minute pentru un model local.",
            saveable=False,
        )

        def worker():

            try:

                if not check_ollama(self.ollama_url):
                    raise OllamaError(
                        "Serverul Ollama nu este accesibil."
                    )

                ro_name = item.get("ro_name") or ""
                en_name = item.get("en_name") or ""

                full_prompt = self.build_species_profile_prompt(
                    scientific_name,
                    ro_name,
                    en_name,
                )

                data = generate(
                    ollama_url=self.ollama_url,
                    model=self.profile_model or self.model,
                    prompt=full_prompt,
                    image_base64=None,
                    num_ctx=self.num_ctx,
                    num_predict=self.profile_num_predict,
                    thinking=self.profile_thinking,
                )

                # Răspuns trunchiat de limita de tokenuri => JSON-ul
                # structurat ar ieși incomplet. Raportăm clar cauza.
                if data.get("done_reason") == "length":
                    raise OllamaError(
                        "Răspunsul a fost trunchiat de limita de tokenuri "
                        f"(PROFILE_NUM_PREDICT={self.profile_num_predict}). "
                        "Mărește PROFILE_NUM_PREDICT în config/settings.txt "
                        "și încearcă din nou."
                    )

                response = get_response(data)

                if not response.strip():
                    raise OllamaError(
                        "Modelul nu a returnat nimic."
                    )

                fields = extract_structured_data(response)

                profile_fields = {
                    key: value
                    for key, value in fields.items()
                    if key not in STANDARD_FIELDS
                }

                display_response = strip_structured_data(response)

                profile_category = (
                    fields.get("category") or item.get("category") or ""
                )
                profile_ro = fields.get("ro_name") or ro_name
                profile_en = fields.get("en_name") or en_name

                save_species_profile(
                    OUTPUT_DIR,
                    scientific_name=scientific_name,
                    category=profile_category,
                    kingdom=fields.get("kingdom") or item.get("kingdom"),
                    genus=fields.get("genus") or item.get("genus"),
                    species=fields.get("species") or item.get("species"),
                    ro_name=profile_ro,
                    en_name=profile_en,
                    profile_response=display_response,
                    profile_fields=profile_fields,
                )

                self.root.after(
                    0,
                    lambda: self.profile_generated(
                        display_response,
                        scientific_name,
                        profile_fields,
                        profile_ro,
                        profile_en,
                        profile_category,
                        time.perf_counter() - self.profile_start_time
                        if self.profile_start_time is not None
                        else None,
                    ),
                )

            except Exception as exc:

                self.root.after(
                    0,
                    # exc=exc: binding la creare — Python șterge variabila
                    # `exc` la ieșirea din except, iar callback-ul rulează
                    # abia după (altfel: NameError în popup-ul de eroare).
                    lambda exc=exc: self.profile_generation_failed(exc),
                )

        threading.Thread(
            target=worker,
            daemon=True,
        ).start()

    def profile_generated(
        self,
        display_response: str,
        scientific_name: str,
        profile_fields: dict | None = None,
        ro_name: str = "",
        en_name: str = "",
        category: str = "",
        elapsed: float | None = None,
    ):

        # Corectam cifrele [~X] din arborele taxonomic: inlocuim cu
        # valorile verificate din Catalogue of Life (daca exista cache).
        display_response = taxonomy_fix_tree_counts(
            display_response, scientific_name
        )

        combined = self._format_species_profile_output(
            scientific_name=scientific_name,
            display_response=display_response,
            profile_fields=profile_fields or {},
            ro_name=ro_name,
            en_name=en_name,
            category=category,
        )

        # Profilul merge in panoul din dreapta; analiza din stanga
        # ramane neatinsa.
        self.set_profile_output_text(combined)

        self.stop_status_timer()
        self.profile_start_time = None

        self.progress.stop()
        self.progress["value"] = 0
        self.progress.pack_forget()

        self.analyze_button.configure(state="normal")
        self.analyze_all_button.configure(state="normal")
        self.update_catalog_buttons()

        if elapsed is not None:
            elapsed_str = self.format_seconds(elapsed)

            self.set_status(
                f"Species profile generated: {scientific_name}"
                f" — {elapsed_str}"
            )

            # Pop-up de finalizare, ca la ANALYZE
            messagebox.showinfo(
                "Profil finalizat",
                f"Profilul a fost generat pentru:\n\n"
                f"{scientific_name}\n\n"
                f"Durată: {elapsed_str}",
                parent=self.root,
            )
        else:
            self.set_status(
                f"Species profile generated: {scientific_name}"
            )

    def profile_generation_failed(self, exc: Exception):

        self.stop_status_timer()
        self.profile_start_time = None

        self.progress.stop()
        self.progress["value"] = 0
        self.progress.pack_forget()

        self.analyze_button.configure(state="normal")
        self.analyze_all_button.configure(state="normal")
        self.update_catalog_buttons()

        messagebox.showerror(
            "Species profile error",
            str(exc),
            parent=self.root,
        )

        self.set_status(
            "Species profile generation failed"
        )

    def load_saved_profile(self):
        """
        Reafișează în zona Output profilul speciei salvat în baza de
        date (generat anterior cu GEN PROFILE), pentru imaginea
        selectată curent.

        Într-un singur mesaj în toate cazurile în care nu se poate
        încărca profilul — nu se afișează popup-uri intermediare.
        """

        if self.image_path is None:
            return

        item = get_catalog_item(
            self.image_path,
            OUTPUT_DIR,
        )

        # Determinăm motivul pentru care nu putem încărca profilul
        if not item:
            messagebox.showinfo(
                "Species profile",
                "Imaginea nu este în catalog.\n\n"
                "Rulează ANALYZE mai întâi pentru a o identifica.",
                parent=self.root,
            )
            return

        scientific_name = (item.get("scientific_name") or "").strip()

        if not scientific_name or scientific_name in ("Necunoscut", "Unknown"):
            messagebox.showinfo(
                "Species profile",
                "Specia nu este încă identificată.\n\n"
                "Rulează ANALYZE pentru a o identifica.",
                parent=self.root,
            )
            return

        profile = get_species_profile(
            OUTPUT_DIR,
            scientific_name,
        )

        if not profile or not (profile.get("profile_response") or "").strip():
            messagebox.showinfo(
                "Species profile",
                f"Nu există profil salvat pentru:\n\n{scientific_name}\n\n"
                "Apasă GEN PROFILE pentru a genera unul.",
                parent=self.root,
            )
            return

        self._load_profile_into_pane(scientific_name, profile)

        self.set_status(
            f"Species profile loaded: {scientific_name}"
        )

    def start_batch_profile_generation(self):
        """
        Colectează speciile unice identificate din imaginile încărcate
        și generează profilul pentru fiecare (sărite pe cele deja
        existente). Simetric cu ANALYZE ALL.
        """

        if not self.image_paths:
            messagebox.showwarning(
                "No images",
                "Add at least one image first.",
                parent=self.root,
            )
            return

        species_map = self._collect_unique_species()
        existing_with_profile = [
            info for info in species_map.values()
            if info["has_profile"]
        ]
        pending = [
            info for info in species_map.values()
            if not info["has_profile"]
        ]

        # Un singur dialog pentru confirmare — mesaj adaptat automat
        if existing_with_profile:
            action_desc = (
                f"{len(existing_with_profile)} specie(i) au deja un profil salvat "
                "în baza de date. "
                "Dorești să regenerezi TOATE (inclusiv cele salvate)?\n"
                f"Se vor procesa toate cele {len(species_map)} specie(i) încărcate."
            )
        else:
            action_desc = (
                f"Vei genera profilul pentru {len(pending)} specii "
                f"(din {len(species_map)} identificate)."
            )

        confirmed = messagebox.askyesno(
            "Generate all species profiles",
            f"{action_desc}\n\n"
            "Poți alege modelul și din dropdown-ul MODEL.\n"
            "Procesul poate dura câteva minute pentru un model local.\n\n"
            "Continui?",
            parent=self.root,
        )

        if not confirmed:
            return

        # Dacă am profiluri salvate și utilizatorul a acceptat,
        # vom regenra toate — nu doar pe cele fără profil
        if existing_with_profile:
            pending = list(species_map.values())

        self.analyze_button.configure(state="disabled")
        self.analyze_all_button.configure(state="disabled")
        self.generate_profile_button.configure(state="disabled")
        self.generate_all_profiles_button.configure(state="disabled")
        self.load_profile_button.configure(state="disabled")
        self.save_profile_button.configure(state="disabled")

        self.progress.pack(side="right")
        self.progress.start(200)

        self.profile_start_time = time.perf_counter()
        self.batch_total = len(pending)
        self.batch_current = 0
        self.batch_filename = None
        self.set_status(
            f"Generating species profiles: 0 / {len(pending)}..."
        )
        self.update_status_with_time()

        self.set_profile_output_text(
            f"Generating {len(pending)} species profiles...\n\n"
            "*Text-only — imaginile nu sunt trimise la model.*",
            saveable=False,
        )

        thread = threading.Thread(
            target=self.run_batch_profile_generation,
            args=(pending,),
            daemon=True,
        )
        thread.start()

    def run_batch_profile_generation(self, pending: list[dict]):
        """
        Worker: generează profilul pentru fiecare specie din lista
        `pending`. La final apelează batch_profile_finished.
        """

        total = len(pending)
        completed = 0
        errors = []

        for index, info in enumerate(pending):

            scientific_name = info["scientific_name"]

            self.root.after(
                0,
                self._batch_profile_status,
                index,
                total,
                scientific_name,
            )

            try:

                if not check_ollama(self.ollama_url):
                    raise OllamaError(
                        "Serverul Ollama nu este accesibil."
                    )

                full_prompt = self.build_species_profile_prompt(
                    scientific_name,
                    info["ro_name"],
                    info["en_name"],
                )

                data = generate(
                    ollama_url=self.ollama_url,
                    model=self.profile_model or self.model,
                    prompt=full_prompt,
                    image_base64=None,
                    num_ctx=self.num_ctx,
                    num_predict=self.profile_num_predict,
                    thinking=self.profile_thinking,
                )

                if data.get("done_reason") == "length":
                    raise OllamaError(
                        "Răspunsul a fost trunchiat de limita de tokenuri "
                        f"(PROFILE_NUM_PREDICT={self.profile_num_predict})."
                    )

                response = get_response(data)

                if not response.strip():
                    raise OllamaError(
                        "Modelul nu a returnat nimic."
                    )

                fields = extract_structured_data(response)

                profile_fields = {
                    key: value
                    for key, value in fields.items()
                    if key not in STANDARD_FIELDS
                }

                display_response = strip_structured_data(response)

                # Corectam cifrele [~X] din arborele taxonomic: inlocuim
                # cu valorile verificate din Catalogue of Life.
                display_response = taxonomy_fix_tree_counts(
                    display_response, scientific_name
                )

                save_species_profile(
                    OUTPUT_DIR,
                    scientific_name=scientific_name,
                    category=info["category"],
                    kingdom=info["kingdom"],
                    genus=info["genus"],
                    species=info["species"],
                    ro_name=info["ro_name"],
                    en_name=info["en_name"],
                    profile_response=display_response,
                    profile_fields=profile_fields,
                )

                completed += 1

            except Exception as exc:

                errors.append(f"{scientific_name}: {exc}")

        batch_elapsed = time.perf_counter() - self.profile_start_time

        self.root.after(
            0,
            self.batch_profile_finished,
            total,
            completed,
            errors,
            batch_elapsed,
        )

    def _batch_profile_status(
        self,
        current: int,
        total: int,
        scientific_name: str,
    ):
        """
        Actualizează progresul batch-ului de profile (main-thread).
        """

        self.batch_current = current
        self.batch_filename = scientific_name

        self.progress.stop()
        self.progress.configure(mode="indeterminate")
        self.progress.start(200)

        self.set_status(
            f"Generating species profile {current + 1} / {total}: "
            f"{scientific_name}..."
        )

    def batch_profile_finished(
        self,
        total: int,
        completed: int,
        errors: list[str],
        elapsed: float,
    ):
        """
        Callback final pentru GEN PROFILE ALL: restaurează controalele,
        afișează sumarul și un singur popup.
        """

        self.stop_status_timer()
        self.profile_start_time = None

        self.progress.stop()
        self.progress.pack_forget()

        self.analyze_button.configure(state="normal")
        self.analyze_all_button.configure(state="normal")
        self.update_catalog_buttons()

        duration = self.format_seconds(elapsed)

        if errors:

            self.set_status(
                f"Species profiles: {completed}/{total} completed, "
                f"errors: {len(errors)} — {duration}"
            )

            error_text = "\n".join(errors)

            messagebox.showwarning(
                "Species profiles finished with errors",
                f"Completed: {completed}/{total}\n\n"
                f"Errors: {len(errors)}\n\n"
                f"Total duration: {duration}\n\n"
                f"{error_text}",
                parent=self.root,
            )

        else:

            self.set_status(
                f"Species profiles completed: {completed}/{total} — "
                f"{duration}"
            )

            messagebox.showinfo(
                "Species profiles complete",
                f"All {total} species profiles were generated "
                f"successfully.\n\n"
                f"Total duration: {duration}",
                parent=self.root,
            )

    def _collect_unique_species(self) -> dict[str, dict]:
        """
        Parcurge imaginile încărcate, citește catalog_items și returnează
        un dicționar indexat pe scientific_name (doar speciile
        identificate). Pentru fiecare indică dacă are deja profil.
        """

        species_map: dict[str, dict] = {}

        for path in self.image_paths:

            item = get_catalog_item(path, OUTPUT_DIR)

            if not item:
                continue

            scientific_name = (
                item.get("scientific_name") or ""
            ).strip()

            if not scientific_name or scientific_name in (
                "Necunoscut",
                "Unknown",
            ):
                continue

            if scientific_name.lower() in {
                name.lower() for name in species_map
            }:
                continue

            existing = get_species_profile(OUTPUT_DIR, scientific_name)

            species_map[scientific_name] = {
                "scientific_name": scientific_name,
                "ro_name": item.get("ro_name") or "",
                "en_name": item.get("en_name") or "",
                "category": item.get("category") or "",
                "kingdom": item.get("kingdom"),
                "genus": item.get("genus"),
                "species": item.get("species"),
                "has_profile": bool(
                    existing
                    and (existing.get("profile_response") or "").strip()
                ),
            }

        return species_map

    def remove_selected_image(self):

        selection = self.image_listbox.curselection()

        if not selection:
            return

        index = selection[0]
        
        path = self.image_paths[index]

        self.image_locations.pop(
            path,
            None
        )

        del self.image_paths[index]        

        self.refresh_image_list()

        if self.image_paths:

            new_index = min(
                index,
                len(self.image_paths) - 1
            )

            self.select_image(
                new_index
            )

        else:

            self.image_path = None
            self.image_base64 = None
            self.thinking_text = ""

            self.preview_label.configure(
                image="",
                text="No image selected"
            )

            self.set_output_text(
                "",
                render=False,
                saveable=False
            )

            self.set_status(
                "No images"
            )


    def clear_images(self):

        self.image_paths.clear()
        self.image_locations.clear()

        self.image_path = None
        self.image_base64 = None
        self.thinking_text = ""

        self.image_listbox.delete(
            0,
            tk.END
        )

        self.preview_label.configure(
            image="",
            text="No image selected"
        )

        self.set_status(
            "Images cleared"
        )

        self.set_output_text(
            "",
            render=False,
            saveable=False
        )


    def batch_status(
        self,
        current,
        total,
        filename
    ):

        self.progress.stop()

        self.progress.configure(
            mode="indeterminate"
        )

        self.progress.start(
            200
        )

        self.batch_progress.configure(
            mode="determinate",
            maximum=total,
            value=current + 1
        )

        self.batch_progress_label.configure(
            text=t("batch_progress").format(cur=current, total=total)
        )

        self.batch_current = current
        self.batch_filename = filename
        self.update_status_with_time()


    def batch_item_error(
        self,
        current,
        total,
        filename,
        error
    ):

        self.progress.stop()

        self.progress.configure(
            mode="indeterminate"
        )

        self.progress.start(
            200
        )

        self.batch_progress.configure(
            mode="determinate",
            maximum=total,
            value=current + 1
        )

        self.batch_progress_label.configure(
            text=t("batch_progress").format(cur=current, total=total)
        )

        self.set_output_text(
            f"ERROR\n\n{filename}\n\n{error}",
            saveable=False
        )

        self.set_status(
            f"Error {current} / {total}: {filename}"
        )


    def batch_item_finished(
        self,
        current,
        total,
        filename,
        response
    ):

        self.progress.stop()

        self.progress.configure(
            mode="indeterminate"
        )

        self.progress.start(
            200
        )

        self.batch_progress.configure(
            mode="determinate",
            maximum=total,
            value=current + 1
        )

        self.batch_progress_label.configure(
            text=t("batch_progress").format(cur=current, total=total)
        )

        self.set_output_text(
            response
        )

        self.set_status(
            f"Completed {current} / {total}: {filename}"
        )


    def batch_finished(
        self,
        total,
        completed,
        errors,
        elapsed
    ):

        self.stop_status_timer()
        self.analysis_start_time = None
        self.batch_current = 0
        self.batch_total = 0

        self.progress.stop()

        self.batch_progress.configure(
            mode="determinate",
            maximum=total,
            value=total
        )

        self.batch_progress_label.configure(
            text=t("batch_progress").format(cur=total, total=total)
        )

        self.batch_progress_frame.pack_forget()

        self.progress.pack(
            side="right"
        )

        # ----------------------------------------------------
        # Re-enable controls
        # ----------------------------------------------------

        self.analyze_button.configure(
            state="normal"
        )

        self.analyze_all_button.configure(
            state="normal"
        )

        self.add_images_button.configure(
            state="normal"
        )

        self.add_folder_button.configure(
            state="normal"
        )

        self.remove_image_button.configure(
            state="normal"
        )

        self.clear_images_button.configure(
            state="normal"
        )

        # ----------------------------------------------------
        # Status
        # ----------------------------------------------------

        duration = self.format_seconds(
            elapsed
        )

        if errors:

            self.set_status(
                f"Batch finished: {completed}/{total}, "
                f"errors: {len(errors)} — "
                f"{duration}"
            )

            error_text = "\n".join(
                errors
            )

            messagebox.showwarning(
                "Batch finished with errors",
                f"Completed: {completed}/{total}\n\n"
                f"Errors: {len(errors)}\n\n"
                f"Total duration: {duration}\n\n"
                f"{error_text}"
            )

        else:

            self.set_status(
                f"Batch completed: {completed}/{total} — "
                f"{duration}"
            )

            messagebox.showinfo(
                "Batch complete",
                f"All {total} images were analyzed successfully.\n\n"
                f"Total duration: {duration}"
            )


    def load_image(self, path: Path):

        try:

            self.image_path = path

            self.image_entry.delete(
                0,
                tk.END
            )

            self.image_entry.insert(
                0,
                str(path)
            )

            # -----------------------------------------------
            # Metadata
            # -----------------------------------------------

            self.image_metadata = get_image_metadata(
                path
            )

            # -----------------------------------------------
            # Base64
            # -----------------------------------------------

            self.image_base64 = image_to_base64(
                path
            )

            # -----------------------------------------------
            # Preview
            # -----------------------------------------------

            with Image.open(path) as image:

                image.thumbnail(
                    (500, 250)
                )

                preview = image.copy()

            self.preview_image = ImageTk.PhotoImage(
                preview
            )

            self.preview_label.configure(
                image=self.preview_image,
                text=""
            )

            self.set_status(
                f"Loaded: {path.name}"
            )

        except Exception as exc:

            self.image_path = None
            self.image_base64 = None

            messagebox.showerror(
                "Image error",
                str(exc)
            )


    def load_saved_result(self, image_path: Path):

        item = get_catalog_item(image_path, OUTPUT_DIR)

        # ----------------------------------------------------
        # Panoul drept: prompt + profilul speciei (silențios)
        # ----------------------------------------------------

        scientific_name = (
            (item.get("scientific_name") or "").strip()
            if item
            else ""
        )

        if scientific_name and scientific_name not in ("Necunoscut", "Unknown"):
            self._load_profile_into_pane(
                scientific_name,
                get_species_profile(OUTPUT_DIR, scientific_name),
            )
        else:
            self._load_profile_into_pane(scientific_name, None)

        if not item or not item.get("response_text"):
            self.thinking_text = ""

            self.set_output_text(
                "No saved result for this image.",
                saveable=False
            )
            return

        try:
            response = item["response_text"]

            # ------------------------------------------------
            # Display only response
            # ------------------------------------------------

            self.thinking_text = ""

            self.set_output_text(
                response
            )

        except Exception as exc:

            self.thinking_text = ""

            self.set_output_text(
                f"Could not load saved result:\n\n{exc}",
                saveable=False
            )


    def build_analysis_prompt(
        self,
        prompt: str,
        metadata: dict,
        manual_location: str = "",
        image_path: Path | None = None,
    ) -> str:
        """
        Contruiește promptul final trimis modelului,
        adăugând contextul disponibil al fotografiei.

        Contextul poate conține:
        - data fotografierii
        - producătorul camerei
        - modelul camerei
        - coordonate GPS
        - locație introdusă manual

        Metadata este folosită ca informație contextuală pentru
        evaluarea plauzibilității identificării, nu ca dovadă vizuală.

        Imaginea rămâne sursa principală pentru identificare.

        Parametrul opțional `image_path` suprascrie `self.image_path` la
        căutarea denumirilor verificate în taxonomy.db — folosit în batch
        (`run_batch_analysis`) unde `self.image_path` nu coincide cu imaginea
        curentă din iteratie.
        """

        metadata = metadata or {}

        photo_date = metadata.get(
            "photo_date"
        )

        camera_make = metadata.get(
            "camera_make"
        )

        camera_model = metadata.get(
            "camera_model"
        )

        latitude = metadata.get(
            "latitude"
        )

        longitude = metadata.get(
            "longitude"
        )

        manual_location = (
            manual_location or ""
        ).strip()

        # Denumiri populare verificate din taxonomy.db (dacă imaginea e în
        # catalog și specia a fost deja identificată înt-r-o analiză anterioară).
        # Modelul le vede în context și le folosește în locul marcajelor #...#.
        db_ro = db_en = ""
        try:
            _img = image_path if image_path is not None else self.image_path
            if _img is not None:
                _item = get_catalog_item(_img, OUTPUT_DIR)
                _sci = (
                    (_item.get("scientific_name") or "").strip()
                    if _item
                    else ""
                )
                if _sci and _sci not in ("Necunoscut", "Unknown"):
                    _chain = taxonomy_lookup_chain(_sci)
                    if _chain:
                        db_ro = (_chain.get("ro_name") or "").strip()
                        db_en = (_chain.get("en_name") or "").strip()
        except Exception:
            pass

        context_lines = [
            "=== CONTEXT FOTOGRAFIE ==="
        ]

        # ------------------------------------------------
        # Data fotografierii
        # ------------------------------------------------

        if photo_date:
            context_lines.append(
                f"Data fotografierii: {photo_date}"
            )

        # ------------------------------------------------
        # Camera
        # ------------------------------------------------

        if camera_make:
            context_lines.append(
                f"Producător cameră: {camera_make}"
            )

        if camera_model:
            context_lines.append(
                f"Model cameră: {camera_model}"
            )

        # ------------------------------------------------
        # Locație
        # ------------------------------------------------

        if (
            latitude is not None
            and longitude is not None
        ):
            context_lines.append(
                f"Locație GPS: {latitude:.6f}, {longitude:.6f}"
            )

        elif manual_location:
            context_lines.append(
                f"Locație: {manual_location} (introdusă manual)"
            )

        # ------------------------------------------------
        # Denumiri populare verificate (din taxonomy.db)
        # ------------------------------------------------

        if db_ro or db_en:
            context_lines.extend([
                "",
                "=== DENUMIRI POPULARE VERIFICATE ===",
                f"română: {db_ro or 'Necunoscut'}",
                f"engleză: {db_en or 'Necunoscut'}",
            ])

        # ------------------------------------------------
        # Dacă nu există context
        # ------------------------------------------------

        if len(context_lines) == 1:
            return prompt

        # ------------------------------------------------
        # Reguli pentru folosirea contextului
        # ------------------------------------------------

        context_lines.extend([
            "",
            "Folosește aceste informații ca date contextuale "
            "pentru evaluarea plauzibilității identificării.",
            "",
            "Data fotografierii poate fi folosită pentru a ține "
            "cont de sezon și de perioada în care anumite specii "
            "pot fi întâlnite.",
            "",
            "Locația poate fi folosită pentru a ține cont de "
            "distribuția geografică și de speciile plauzibile "
            "în zona respectivă.",
            "",
            "NU considera aceste informații drept dovezi vizuale.",
            "NU afirma că data, locația sau datele camerei sunt "
            "vizibile în fotografie.",
            "",
            "Identificarea trebuie să se bazeze în primul rând "
            "pe caracteristicile observabile în imagine.",
            "",
            "Dacă informațiile contextuale sugerează o identificare "
            "diferită de ceea ce este vizibil în imagine, acordă "
            "prioritate caracteristicilor vizibile și menționează "
            "incertitudinea atunci când este relevant.",
            "",
            "DENUMIRI POPULARE: În răspunsul tău, înlocuiește fiecare marcaj "
            "#...# dintre prompt (ex. #română#, #engleză#, #nume comun#) cu "
            "denumirea populară exactă în limba corespunzătoare. "
            "Dacă nu cunoști o denumire, scrie „Necunoscut”.",
            "Folosește denumirile verificate din blocul „=== DENUMIRI POPULARE VERIFICATE ===” ",
            "dacă este prezente și nu le inventa; „Necunoscut” doar dacă limba lipsește complet.",
            "",
            "=== INSTRUCȚIUNEA UTILIZATORULUI ==="
        ])

        return (
            "\n".join(context_lines)
            + "\n"
            + prompt
        )


    def on_location_selected(self, event=None):

        if self.image_path is None:
            return

        location = self.location_var.get().strip()

        if location == "__ALTCEVA__":
            self.custom_location_frame.pack(
                anchor="w"
            )
            self.custom_location_entry.focus_set()

            custom_location = self.custom_location_var.get().strip()

            if custom_location:
                self.image_locations[
                    self.image_path
                ] = custom_location

            return

        self.custom_location_frame.pack_forget()
        self.custom_location_var.set("")

        self.image_locations[
            self.image_path
        ] = location


    def on_custom_location_changed(self, event=None):

        if self.image_path is None:
            return

        custom_location = self.custom_location_var.get().strip()

        if self.location_var.get().strip() != "__ALTCEVA__":
            return

        if custom_location:
            self.image_locations[
                self.image_path
            ] = custom_location


    # ========================================================
    # ANALYSIS
    # ========================================================

    def start_analysis(self):

        self.image_status[self.image_path] = "processing"

        self.refresh_image_list()

        if self.image_path is None:
            messagebox.showwarning(
                "No image",
                "Select an image first."
            )
            return

        prompt = self.get_prompt_text()

        if not prompt:
            messagebox.showwarning(
                "No prompt",
                "Enter a prompt first."
            )
            return

        self.batch_progress_frame.pack_forget()
        self.batch_progress_label.pack_forget()

        # ----------------------------------------------------
        # Disable controls
        # ----------------------------------------------------

        self.analyze_button.configure(
            state="disabled"
        )

        self.add_images_button.configure(
            state="disabled"
        )

        self.add_folder_button.configure(
            state="disabled"
        )

        self.remove_image_button.configure(
            state="disabled"
        )

        self.clear_images_button.configure(
            state="disabled"
        )

        self.save_button.configure(
            state="disabled"
        )
        
        self.progress.pack(
            side="right"
        )        

        self.progress.start(
            200
        )

        self.analysis_start_time = time.perf_counter()
        self.profile_start_time = None
        self.set_status(
            "Muse Glimmer is working..."
        )
        self.update_status_with_time()

        # ----------------------------------------------------
        # Capture preset name
        # ----------------------------------------------------

        preset_name = self.prompt_var.get()

        # ----------------------------------------------------
        # Thread
        # ----------------------------------------------------

        thread = threading.Thread(
            target=self.run_analysis,
            args=(prompt, preset_name),
            daemon=True
        )

        thread.start()

    # --------------------------------------------------------

    def start_batch_analysis(self):

        if not self.image_paths:
            messagebox.showwarning(
                "No images",
                "Add at least one image first."
            )
            return

        prompt = self.get_prompt_text()

        if not prompt:
            messagebox.showwarning(
                "No prompt",
                "Enter a prompt first."
            )
            return

        # ----------------------------------------------------
        # Disable controls
        # ----------------------------------------------------

        self.analyze_button.configure(
            state="disabled"
        )

        self.analyze_all_button.configure(
            state="disabled"
        )

        self.add_images_button.configure(
            state="disabled"
        )

        self.add_folder_button.configure(
            state="disabled"
        )

        self.remove_image_button.configure(
            state="disabled"
        )

        self.clear_images_button.configure(
            state="disabled"
        )

        self.save_button.configure(
            state="disabled"
        )

        # ----------------------------------------------------
        # Progress
        # ----------------------------------------------------

        total = len(self.image_paths)

        # Current image progress
        self.progress.pack(
            side="right"
        )

        self.progress.configure(
            mode="indeterminate"
        )

        self.progress.start(
            200
        )

        # Batch progress
        self.batch_progress.configure(
            mode="determinate",
            maximum=total,
            value=0
        )

        self.batch_progress_label.configure(
            text=t("batch_progress").format(cur=0, total=total)
        )

        self.batch_progress_frame.pack(
            side="right",
            padx=(0, 8)
        )

        self.batch_total = total
        self.batch_current = 0
        self.analysis_start_time = time.perf_counter()
        self.set_status(
            f"Processing batch: 0 / {total}"
        )
        self.update_status_with_time()

        # ----------------------------------------------------
        # Snapshot of the list and preset
        # ----------------------------------------------------

        paths = list(
            self.image_paths
        )

        preset_name = self.prompt_var.get()

        # ----------------------------------------------------
        # Start worker
        # ----------------------------------------------------

        thread = threading.Thread(
            target=self.run_batch_analysis,
            args=(paths, prompt, preset_name),
            daemon=True
        )

        thread.start()


    def run_batch_analysis(
        self,
        paths,
        prompt,
        preset_name
    ):

        batch_start_time = time.perf_counter()
        
        total = len(paths)        

        completed = 0
        errors = []

        for index, path in enumerate(paths, start=0):

            self.image_status[path] = "processing"

            self.root.after(
                0,
                self.refresh_image_list
            )

            # Update status immediately before processing
            self.batch_current = index
            self.batch_filename = path.name
            self.root.after(
                0,
                self.batch_status,
                index,
                total,
                path.name
            )

            try:
                # ------------------------------------------------
                # Read image
                # ------------------------------------------------

                image_base64 = image_to_base64(
                    path
                )

                metadata = get_image_metadata(
                    path
                )
                
                manual_location = self.image_locations.get(
                    path,
                    "București"
                )                

                analysis_prompt = self.build_analysis_prompt(
                    prompt,
                    metadata,
                    manual_location,
                    image_path=path,
                )

                manual_context = get_manual_classification_context(
                    path,
                    OUTPUT_DIR,
                )
                if manual_context:
                    analysis_prompt = f"{manual_context}\n\n{analysis_prompt}"

                is_fast_prompt = is_fast_preset(preset_name)
                catalog_prompt = (
                    fast_analysis_prompt(analysis_prompt, preset_name)
                    if is_fast_prompt
                    else append_structured_data_instruction(analysis_prompt)
                )

                current_analysis_key = analysis_key(
                    catalog_prompt,
                    self.model,
                    self.num_ctx,
                    fast_predict_limit(preset_name, self.num_predict) if is_fast_prompt else self.num_predict,
                    False if is_fast_prompt else self.thinking,
                )
                cached = get_cached_analysis(
                    path,
                    OUTPUT_DIR,
                    current_analysis_key,
                )

                if cached:
                    response = cached["response_text"]
                    profile = (
                        get_profile_for_response(OUTPUT_DIR, response)
                        if normalize_preset_name(preset_name) == "botanic_quick"
                        else None
                    )
                    if profile and profile.get("profile_response"):
                        response = profile["profile_response"]
                    completed += 1
                    self.image_status[path] = "done"
                    self.root.after(0, self.refresh_image_list)
                    self.root.after(0, self.batch_item_finished, index, total, path.name, response)
                    continue

                # ------------------------------------------------
                # Ollama
                # ------------------------------------------------

                if not check_ollama(
                    self.ollama_url
                ):
                    raise OllamaError(
                        "Ollama is not accessible."
                    )

                data = generate(
                    ollama_url=self.ollama_url,
                    model=self.model,
                    prompt=catalog_prompt,
                    image_base64=image_base64,
                    num_ctx=self.num_ctx,
                    num_predict=fast_predict_limit(preset_name, self.num_predict) if is_fast_prompt else self.num_predict,
                    thinking=False if is_fast_prompt else self.thinking,
                )

                response = get_response(
                    data
                )

                if not response.strip():
                    retry_prompt = (
                        "Identifică planta din imagine. "
                        "Răspunde doar cu genul și specia în latină. "
                        "Nu oferi explicații."
                        if normalize_preset_name(preset_name) == "botanic_quick"
                        else catalog_prompt
                    )
                    retry_data = generate(
                        ollama_url=self.ollama_url,
                        model=self.model,
                        prompt=retry_prompt,
                        image_base64=image_base64,
                        num_ctx=self.num_ctx,
                        num_predict=fast_predict_limit(preset_name, self.num_predict) if is_fast_prompt else min(self.num_predict, 1024),
                        thinking=False,
                    )
                    response = get_response(retry_data)

                if not response.strip():
                    keys = ", ".join(sorted(data.keys()))
                    raise OllamaError(
                        "Ollama a finalizat cererea fără text. "
                        f"Model: {self.model}; preset: {preset_name}; "
                        f"câmpuri primite: {keys or 'niciunul'}. "
                        "Răspunsul primit a fost doar protocol sau gol."
                    )

                profile = (
                    get_profile_for_response(OUTPUT_DIR, response)
                    if normalize_preset_name(preset_name) == "botanic_quick"
                    else None
                )
                display_response = (
                    profile["profile_response"]
                    if profile and profile.get("profile_response")
                    else strip_structured_data(response)
                )

                thinking = get_thinking(
                    data
                )

                statistics = get_statistics(
                    data
                )

                # ------------------------------------------------
                # Output path
                # ------------------------------------------------

                catalogize_analysis(
                    image_path=path,
                    response=response,
                    output_dir=OUTPUT_DIR,
                    analysis_key_value=current_analysis_key,
                    thinking=thinking if self.thinking else "",
                    prompt=prompt,
                    preset_name=preset_name,
                    manual_location=manual_location,
                    image_metadata=metadata,
                    statistics=statistics,
                )

                # Patch: la fel ca la ANALYZE simplu, căutăm denumirile
                # populare verificate în taxonomy.db și le injectăm în
                # response_text (înlocuind „Necunoscut” / marcajele #...#).
                try:
                    item = get_catalog_item(path, OUTPUT_DIR)
                    if item:
                        scientific = (item.get("scientific_name") or "").strip()
                        if scientific and scientific not in ("Necunoscut", "Unknown"):
                            chain_info = taxonomy_lookup_chain(scientific)
                            if chain_info and (chain_info.get("ro_name") or chain_info.get("en_name")):
                                from core.catalog import update_catalog_response_names
                                update_catalog_response_names(
                                    OUTPUT_DIR,
                                    path,
                                    chain_info.get("ro_name"),
                                    chain_info.get("en_name"),
                                )
                except Exception:
                    pass

                self.root.after(
                    0,
                    self.update_catalog_buttons_for_path,
                    path,
                )

                completed += 1
                
                self.image_status[path] = "done"

                self.root.after(
                    0,
                    self.refresh_image_list
                )                

                self.root.after(
                    0,
                    self.batch_item_finished,
                    index,
                    total,
                    path.name,
                    display_response
                )

            except Exception as exc:

                errors.append(
                    f"{path.name}: {exc}"
                )

                self.root.after(
                    0,
                    self.batch_item_error,
                    index,
                    total,
                    path.name,
                    str(exc)
                )
                
                self.image_status[path] = "error"

                self.root.after(
                    0,
                    self.refresh_image_list
                )                

        # --------------------------------------------------------
        # Batch finished
        # --------------------------------------------------------

        batch_elapsed = time.perf_counter() - batch_start_time

        self.root.after(
            0,
            self.batch_finished,
            total,
            completed,
            errors,
            batch_elapsed
        )

    def run_analysis(self, prompt: str, preset_name: str):

        start_time = time.perf_counter()

        try:
            # ------------------------------------------------
            # Ollama
            # ------------------------------------------------

            if not check_ollama(
                self.ollama_url
            ):
                raise OllamaError(
                    "Ollama is not accessible."
                )

            # ------------------------------------------------
            # Generate
            # ------------------------------------------------
           
            manual_location = self.image_locations.get(
                self.image_path,
                self.location_var.get().strip()
            )            

            analysis_prompt = self.build_analysis_prompt(
                prompt,
                self.image_metadata,
                manual_location
            )            

            manual_context = get_manual_classification_context(
                self.image_path,
                OUTPUT_DIR,
            )
            if manual_context:
                analysis_prompt = f"{manual_context}\n\n{analysis_prompt}"

            is_fast_prompt = is_fast_preset(preset_name)
            catalog_prompt = (
                fast_analysis_prompt(analysis_prompt, preset_name)
                if is_fast_prompt
                else append_structured_data_instruction(analysis_prompt)
            )

            current_analysis_key = analysis_key(
                catalog_prompt,
                self.model,
                self.num_ctx,
                fast_predict_limit(preset_name, self.num_predict) if is_fast_prompt else self.num_predict,
                False if is_fast_prompt else self.thinking,
            )
            cached = get_cached_analysis(
                self.image_path,
                OUTPUT_DIR,
                current_analysis_key,
            )
            if cached:
                response = cached["response_text"]
                elapsed = time.perf_counter() - start_time
                self.response_text = response
                self.thinking_text = ""
                self.statistics = {"cache": "hit"}
                self.root.after(
                    0,
                    self.analysis_finished,
                    response,
                    elapsed,
                    prompt,
                    preset_name
                )
                return

            data = generate(
                ollama_url=self.ollama_url,
                model=self.model,
                prompt=catalog_prompt,
                image_base64=self.image_base64,
                num_ctx=self.num_ctx,
                num_predict=fast_predict_limit(preset_name, self.num_predict) if is_fast_prompt else self.num_predict,
                thinking=False if is_fast_prompt else self.thinking,
            )

            # ------------------------------------------------
            # Results
            # ------------------------------------------------

            response = get_response(
                data
            )

            if not response.strip():
                retry_prompt = (
                    "Identifică planta din imagine. "
                    "Răspunde doar cu genul și specia în latină. "
                    "Nu oferi explicații."
                    if normalize_preset_name(preset_name) == "botanic_quick"
                    else catalog_prompt
                )
                retry_data = generate(
                    ollama_url=self.ollama_url,
                    model=self.model,
                    prompt=retry_prompt,
                    image_base64=self.image_base64,
                    num_ctx=self.num_ctx,
                    num_predict=fast_predict_limit(preset_name, self.num_predict) if is_fast_prompt else min(self.num_predict, 1024),
                    thinking=False,
                )
                response = get_response(retry_data)

                if not response.strip():
                    keys = ", ".join(sorted(data.keys()))
                    raise OllamaError(
                        "Ollama a finalizat cererea fără text. "
                        f"Câmpuri primite: {keys or 'niciunul'}."
                    )

            profile = (
                get_profile_for_response(OUTPUT_DIR, response)
                if normalize_preset_name(preset_name) == "botanic_quick"
                else None
            )
            display_response = (
                profile["profile_response"]
                if profile and profile.get("profile_response")
                else strip_structured_data(response)
            )

            thinking = get_thinking(
                data
            )

            statistics = get_statistics(
                data
            )

            catalogize_analysis(
                image_path=self.image_path,
                response=response,
                output_dir=OUTPUT_DIR,
                analysis_key_value=current_analysis_key,
                thinking=thinking if self.thinking else "",
                prompt=prompt,
                preset_name=preset_name,
                manual_location=manual_location,
                image_metadata=self.image_metadata,
                statistics=statistics,
            )
            # Actualizeaza denumirile din taxonomy.db in catalog + response_text
            patched_response = None
            try:
                item = get_catalog_item(self.image_path, OUTPUT_DIR)
                if item:
                    scientific = (item.get("scientific_name") or "").strip()
                    if scientific:
                        chain_info = taxonomy_lookup_chain(scientific)
                        if chain_info and (chain_info.get("ro_name") or chain_info.get("en_name")):
                            from core.catalog import update_catalog_response_names
                            update_catalog_response_names(
                                OUTPUT_DIR,
                                self.image_path,
                                chain_info.get("ro_name"),
                                chain_info.get("en_name"),
                            )
                            # Re-citim response_text patch-uit din catalog
                            item2 = get_catalog_item(self.image_path, OUTPUT_DIR)
                            if item2:
                                patched_response = item2.get("response_text")
            except Exception:
                pass

            self.root.after(
                0,
                self.update_catalog_buttons,
            )

            elapsed = time.perf_counter() - start_time

            # Folosim response_text patch-uit dacă exista, altfel originalul
            final_response = patched_response if patched_response else response
            self.response_text = final_response
            self.thinking_text = thinking
            self.statistics = statistics

            self.root.after(
                0,
                self.analysis_finished,
                final_response,
                elapsed,
                prompt,
                preset_name
            )

        except Exception as exc:

            log_dir = APP_ROOT / "logs"
            log_dir.mkdir(parents=True, exist_ok=True)
            (log_dir / "analysis_error.log").write_text(
                traceback.format_exc(),
                encoding="utf-8",
            )

            self.root.after(
                0,
                self.analysis_failed,
                str(exc)
            )

    # ========================================================
    # ANALYSIS FINISHED
    # ========================================================

    def analysis_finished(
        self,
        response: str,
        elapsed: float,
        prompt: str,
        preset_name: str
    ):

        if self.image_path is not None:
            self.image_status[self.image_path] = "done"

        self.update_catalog_buttons()

        self.refresh_image_list()

        self.stop_status_timer()
        self.analysis_start_time = None
        self.profile_start_time = None

        self.progress.stop()
        self.progress["value"] = 0
        self.progress.pack_forget()

        self.analyze_button.configure(
            state="normal"
        )

        self.add_images_button.configure(
            state="normal"
        )

        self.add_folder_button.configure(
            state="normal"
        )

        self.remove_image_button.configure(
            state="normal"
        )

        self.clear_images_button.configure(
            state="normal"
        )

        self.save_button.configure(
            state="normal"
        )

        # ----------------------------------------------------
        # Output
        # ----------------------------------------------------

        self.set_output_text(
            response
        )

        try:
            self.save_analysis_result(response, prompt, preset_name)

        except Exception as exc:
            self.set_status(
                f"Result displayed, but could not be saved: {exc}"
            )

        # ----------------------------------------------------
        # Status
        # ----------------------------------------------------

        self.set_status(
            f"Done — {self.format_seconds(elapsed)}"
        )

        # ----------------------------------------------------
        # Completion popup
        # ----------------------------------------------------

        messagebox.showinfo(
            "Analiză finalizată",
            f"Analiza a fost finalizată pentru:\n\n"
            f"{self.image_path.name}\n\n"
            f"Durată: {self.format_seconds(elapsed)}"
        )

    def analysis_failed(
        self,
        error
    ):

        if self.image_path is not None:
            self.image_status[self.image_path] = "error"

        self.refresh_image_list()

        self.stop_status_timer()
        self.analysis_start_time = None
        self.profile_start_time = None

        self.progress.stop()
        self.progress["value"] = 0
        self.progress.pack_forget()

        self.analyze_button.configure(
            state="normal"
        )

        self.add_images_button.configure(
            state="normal"
        )

        self.add_folder_button.configure(
            state="normal"
        )

        self.remove_image_button.configure(
            state="normal"
        )

        self.clear_images_button.configure(
            state="normal"
        )

        self.set_status(
            "Error"
        )

        messagebox.showerror(
            "Analysis failed",
            str(error)
        )

    # ========================================================
    # SAVE
    # ========================================================

    def open_species_folder(self):
        if self.image_path is None:
            return
        item = get_catalog_item(self.image_path, OUTPUT_DIR)
        if not item:
            messagebox.showinfo("Catalog", "Imaginea nu este încă în catalog.")
            return
        folder = Path(item["catalog_path"]).parent
        if not folder.is_dir():
            messagebox.showwarning("Catalog", "Folderul catalogului nu mai există.")
            return
        os.startfile(str(folder))

    def open_quick_catalog_dialog(self):
        """Dialog pentru catalogare rapidă: alege imagini + specie din taxonomy.db."""
        from core.catalog import quick_catalog_images, get_catalog_item
        from core.taxonomy import search_taxonomy_names

        if not self.image_paths:
            messagebox.showwarning(t("qc_title"), t("qc_no_images"), parent=self.root)
            return

        win = tk.Toplevel(self.root)
        win.title("Catalog rapid")
        win.geometry("640x560")
        win.transient(self.root)
        win.grab_set()

        # Canvas scrollabil principal (conține TOTUL)
        main_canvas = tk.Canvas(win, highlightthickness=0)
        vscroll = ttk.Scrollbar(win, orient="vertical", command=main_canvas.yview)
        scrollable = tk.Frame(main_canvas)

        scrollable.bind("<Configure>",
                        lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
        main_canvas.create_window((0, 0), window=scrollable, anchor="nw", width=600)
        main_canvas.configure(yscrollcommand=vscroll.set)

        main_canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        # --- Imaginea țintă (cea selectată în lista principală) ---
        # Default: catalogăm DOAR imaginea selectată, fără grid de poze.
        target_path = self.image_path
        if target_path is None and self.image_paths:
            target_path = self.image_paths[0]
        target_box = {"path": target_path}

        info_top = ttk.Label(
            scrollable,
            text=(t("qc_selected_image").format(name=Path(target_path).name) if target_path
                  else t("qc_no_selection")),
            font=("", 10, "bold"), wraplength=560, justify="left")
        info_top.pack(fill="x", padx=10, pady=(10, 4))

        # --- Container poze (ascuns implicit, se încarcă LA CERERE) ---
        pics_header = ttk.Frame(scrollable)
        pics_label = ttk.Label(pics_header, text=t("qc_pics_header"),
                               font=("", 10, "bold"))
        pics_label.pack(side="left")
        toggle_pics_btn = ttk.Button(pics_header, text=t("qc_show_pics"))
        toggle_pics_btn.pack(side="right")
        pics_header.pack(fill="x", padx=10, pady=(6, 2))

        pics_box = tk.Frame(scrollable, bg="#fafafa")
        # NU pack la început — apare doar la cerere.

        THUMB = 80
        CELL_W = THUMB + 12  # lățime estimată per celulă
        img_vars = []
        photo_refs = []
        all_cells = []  # refereță la toate celulele pentru reflow
        pics_state = {"built": False, "visible": False}
        grid_frame = None

        def _build_grid():
            """Construiește grid-ul de miniaturi (o singură dată, la cerere)."""
            nonlocal grid_frame
            grid_frame = tk.Frame(pics_box, bg="#fafafa")
            grid_frame.pack(fill="x", padx=10, pady=4)

            for idx, p in enumerate(self.image_paths):
                already = get_catalog_item(p, OUTPUT_DIR) is not None
                var = tk.BooleanVar(value=False)
                img_vars.append((p, var))

                cell = tk.Frame(grid_frame, bd=1, relief="solid",
                                bg="#ddd" if already else "#fff")
                all_cells.append((cell, var, already))

                try:
                    im = Image.open(p)
                    im.thumbnail((THUMB, THUMB))
                    photo = ImageTk.PhotoImage(im)
                    img_label = tk.Label(cell, bg="#eee",
                                         cursor="hand2" if not already else "arrow")
                    img_label.image = photo
                    img_label.configure(image=photo)
                    img_label.pack(padx=2, pady=(2, 0))
                except Exception:
                    img_label = tk.Label(cell, text="?", bg="#eee", width=10, height=5)
                    img_label.pack(padx=2, pady=(2, 0))

                def _toggle(event, v=var, c=cell, catalogued=already, path=p):
                    if catalogued:
                        # Afișează unde e catalogată poza
                        item = get_catalog_item(path, OUTPUT_DIR)
                        if item:
                            sci = item.get("scientific_name", "?")
                            cat = item.get("category", "?")
                            cpath = item.get("catalog_path", "?")
                            folder = Path(cpath).parent if cpath else "?"
                            msg = "\n".join([
                                f"{t('lbl_species')} {sci or '?'}",
                                f"{t('lbl_category')} {cat or '?'}",
                                f"{t('lbl_folder')} {folder}",
                            ])
                            messagebox.showinfo(t("qc_catalogued_photo"), msg, parent=win)
                        return
                    v.set(not v.get())
                    c.configure(bd=3 if v.get() else 1,
                                bg="#90caf9" if v.get() else "#ffffff")

                img_label.bind("<Button-1>", _toggle)

                name_short = p.stem[:14] + ("…" if len(p.stem) > 14 else "")
                caption = name_short + (" ✓" if already else "")
                cap = tk.Label(cell, text=caption, font=("", 8),
                               bg="#ddd" if already else "#fff",
                               fg="#888" if already else "#333")
                cap.pack(fill="x", padx=2, pady=(0, 2))
                cap.bind("<Button-1>", _toggle)

        # Reflow la redimensionare
        def _reflow(event=None):
            if grid_frame is None:
                return
            width = grid_frame.winfo_width()
            cols = max(1, width // CELL_W)
            for i, (c, v, cat) in enumerate(all_cells):
                c.grid(row=i // cols, column=i % cols, padx=4, pady=4, sticky="nsew")

        def _toggle_pics():
            if not pics_state["built"]:
                _build_grid()
                grid_frame.bind("<Configure>", _reflow)
                pics_state["built"] = True
            if pics_state["visible"]:
                pics_box.pack_forget()
                toggle_pics_btn.configure(text=t("qc_show_pics"))
                pics_state["visible"] = False
            else:
                pics_box.pack(fill="x", padx=0, pady=0)
                toggle_pics_btn.configure(text=t("qc_hide_pics"))
                pics_state["visible"] = True

        toggle_pics_btn.configure(command=_toggle_pics)

        # --- Câmp specie științifică ---
        ttk.Label(scrollable, text=t("qc_sci_name_label")).pack(
            fill="x", padx=10, pady=(8, 2))

        species_var = tk.StringVar()
        species_entry = ttk.Entry(scrollable, textvariable=species_var, font=("", 11))
        species_entry.pack(fill="x", padx=10)

        # Listbox pentru sugestii
        suggest_frame = ttk.Frame(scrollable)
        suggest_frame.pack(fill="x", padx=10, pady=2)
        suggest_listbox = tk.Listbox(suggest_frame, height=4, font=("", 10))
        suggest_listbox.pack(fill="x")

        # Specia validată (afișată proeminent)
        validated_var = tk.StringVar(value="")
        validated_label = ttk.Label(scrollable, textvariable=validated_var, font=("", 10, "bold"),
                                    foreground="#2e7d32", wraplength=560, justify="left")
        validated_label.pack(fill="x", padx=10, pady=(2, 0))

        # Info specie
        info_frame = ttk.LabelFrame(scrollable, text=t("qc_info_frame"))
        info_frame.pack(fill="x", padx=10, pady=6)

        cat_var = tk.StringVar(value="—")
        ro_var = tk.StringVar(value="—")
        en_var = tk.StringVar(value="—")

        ttk.Label(info_frame, text=t("lbl_category")).grid(row=0, column=0, sticky="w", padx=8, pady=2)
        cat_combo = ttk.Combobox(info_frame, textvariable=cat_var,
                                 values=["Planta", "Animal", "Pasare", "Ciuperca", "Alta"],
                                 state="readonly", width=12)
        cat_combo.grid(row=0, column=1, sticky="w", padx=4, pady=2)

        ttk.Label(info_frame, text="Ro:").grid(row=0, column=2, sticky="w", padx=8, pady=2)
        ttk.Label(info_frame, textvariable=ro_var, foreground="#1565c0").grid(row=0, column=3, sticky="w", padx=4, pady=2)

        ttk.Label(info_frame, text="En:").grid(row=0, column=4, sticky="w", padx=8, pady=2)
        ttk.Label(info_frame, textvariable=en_var, foreground="#1565c0").grid(row=0, column=5, sticky="w", padx=4, pady=2)

        status_var = tk.StringVar(value=t("qc_type_to_search"))
        status_label = ttk.Label(scrollable, textvariable=status_var, foreground="#888")
        status_label.pack(fill="x", padx=10, pady=(0, 4))

        # --- Callbacks ---
        _search_timer = None

        def on_type(*args):
            nonlocal _search_timer
            if _search_timer:
                win.after_cancel(_search_timer)
            _search_timer = win.after(250, do_search)

        def do_search():
            query = species_var.get().strip()
            suggest_listbox.delete(0, "end")
            if len(query) < 2:
                suggest_listbox.insert("end", t("qc_min_chars"))
                return
            results = search_taxonomy_names(query, limit=8)
            if not results:
                suggest_listbox.insert("end", t("qc_no_results"))
                return
            for r in results:
                name = r.get("name", "")
                ro = r.get("ro_name", "") or "—"
                en = r.get("en_name", "") or "—"
                suggest_listbox.insert("end", f"{name}  |  ro: {ro}  |  en: {en}")

        # Mapping kingdom (taxonomy.db) → categorie (dropdown)
        _KINGDOM_MAP = {
            "plantae": "Planta",
            "animalia": "Animal",
            "fungi": "Ciuperca",
        }

        def _kingdom_to_category(kingdom: str) -> str:
            k = (kingdom or "").strip().casefold()
            for key, val in _KINGDOM_MAP.items():
                if key in k:
                    return val
            # fallback: detectează din pasare/bird
            if "pasar" in k or "bird" in k:
                return "Pasare"
            return "Alta"

        def validate_current():
            """Validează numele curent din entry (la Enter sau focus out)."""
            name = species_var.get().strip()
            if not name:
                return
            from core.taxonomy import lookup_species_simple
            info = lookup_species_simple(name)
            if info:
                cat_var.set(_kingdom_to_category(info.get("kingdom", "")))
                ro_var.set(info.get("ro_name", "—") or "—")
                en_var.set(info.get("en_name", "—") or "—")
                validated_var.set(t("qc_validated").format(name=info.get('name', name)))
                status_var.set(t("qc_valid"))
                status_label.configure(foreground="#2e7d32")
                catalog_btn.configure(state="normal")
            else:
                cat_var.set("—")
                ro_var.set("—")
                en_var.set("—")
                validated_var.set("")
                status_var.set(t("qc_invalid"))
                status_label.configure(foreground="#c62828")
                catalog_btn.configure(state="disabled")

        def on_select_suggestion(event):
            sel = suggest_listbox.curselection()
            if not sel:
                return
            text = suggest_listbox.get(sel[0])
            if text.startswith("(") or text.startswith("Scrie"):
                return
            name = text.split("  |  ")[0].strip()
            species_var.set(name)
            suggest_listbox.delete(0, "end")
            validate_current()

        def do_catalog():
            # Default: catalogăm DOAR imaginea selectată. Dacă grid-ul de
            # referință e vizibil și userul a bifat poze acolo, le luăm pe acelea.
            if pics_state["visible"] and any(v.get() for _, v in img_vars):
                selected = [p for p, v in img_vars if v.get()]
            elif target_box["path"] is not None:
                selected = [target_box["path"]]
            else:
                messagebox.showwarning(t("qc_title"), t("qc_select_image"), parent=win)
                return
            scientific = species_var.get().strip()
            if not scientific:
                messagebox.showwarning(t("qc_title"), t("qc_enter_name"), parent=win)
                return
            category = cat_var.get()
            if category == "—":
                messagebox.showwarning(t("qc_title"), t("qc_choose_category"), parent=win)
                return

            # Mapare nume categorie → folder (consistent cu catalogize_analysis)
            _CAT_TO_FOLDER = {
                "Planta": "plante",
                "Animal": "animale",
                "Pasare": "pasari",
                "Ciuperca": "ciuperci",
                "Alta": "altele",
            }
            folder_category = _CAT_TO_FOLDER.get(category, "altele")

            try:
                result = quick_catalog_images(selected, scientific, folder_category, OUTPUT_DIR)
                copied = result["copied"]
                errors = result["errors"]
                dest_folder = OUTPUT_DIR / folder_category / scientific.replace(" ", "_").replace("×", "x").lower()
                msg = t("qc_result_msg").format(n=copied, folder=dest_folder)
                if errors:
                    msg += f"\n\nErori ({len(errors)}):\n" + "\n".join(errors[:5])
                # Afișează rezultat + buton Deschide folder
                dlg = tk.Toplevel(win)
                dlg.title(t("qc_result_title"))
                dlg.geometry("560x280")
                dlg.transient(win)
                dlg.grab_set()
                ttk.Label(dlg, text=msg, justify="left").pack(fill="both", expand=True, padx=12, pady=10)
                btn_bar = ttk.Frame(dlg)
                btn_bar.pack(fill="x", padx=12, pady=(0, 10))
                ttk.Button(btn_bar, text=t("btn_open_folder"),
                           command=lambda: os.startfile(str(dest_folder))).pack(side="left")
                ttk.Button(btn_bar, text=t("btn_ok"), command=dlg.destroy).pack(side="right")
                if copied > 0:
                    for p in selected:
                        try:
                            if p in self.image_paths:
                                self.image_paths.remove(p)
                        except ValueError:
                            pass
                        self.image_status.pop(p, None)
                        self.image_locations.pop(p, None)
                    if target_box["path"] in selected:
                        target_box["path"] = (
                            self.image_paths[0] if self.image_paths else None
                        )
                        if target_box["path"] is not None:
                            info_top.configure(
                                text=t("qc_selected_image").format(name=Path(target_box['path']).name))
                        else:
                            info_top.configure(text=t("qc_no_selection"))
                    self.refresh_image_list()
            except Exception as exc:
                messagebox.showerror(t("qc_title"), str(exc), parent=win)

        species_var.trace_add("write", on_type)
        suggest_listbox.bind("<<ListboxSelect>>", on_select_suggestion)
        species_entry.bind("<FocusOut>", lambda e: validate_current())
        species_entry.bind("<Return>", lambda e: validate_current())

        # --- Butoane ---
        btn_frame = ttk.Frame(scrollable)
        btn_frame.pack(fill="x", padx=10, pady=(4, 10))
        catalog_btn = ttk.Button(btn_frame, text=t("qc_catalog_btn"), command=do_catalog, state="disabled")
        catalog_btn.pack(side="left")
        ttk.Button(btn_frame, text=t("btn_cancel"), command=win.destroy).pack(side="right")

    def show_help(self):
        """Open a reusable, non-modal bilingual guide without touching user data."""
        win = getattr(self, "help_window", None)
        if win is not None and win.winfo_exists():
            win.deiconify()
            win.lift()
            return
        win = tk.Toplevel(self.root)
        self.help_window = win
        win.geometry("800x600")
        win.minsize(520, 360)
        win.transient(self.root)
        win.bind("<Escape>", lambda event: win.destroy())
        self.help_intro_label = ttk.Label(win, padding=10, wraplength=490)
        self.help_intro_label.pack(fill="x")
        self.help_font_size = getattr(self, "help_font_size", 12)
        self.help_text_font = tkfont.Font(
            root=self.root, **tkfont.nametofont("TkDefaultFont", root=self.root).actual()
        )
        self.help_text_font.configure(size=self.help_font_size)
        font_bar = ttk.Frame(win)
        font_bar.pack(fill="x", padx=10, pady=(0, 8))
        self.help_font_label = ttk.Label(font_bar)
        self.help_font_label.pack(side="left", padx=(0, 8))
        self.help_font_smaller_button = ttk.Button(
            font_bar, width=4, command=lambda: self.resize_help_font(-1)
        )
        self.help_font_smaller_button.pack(side="left")
        self.help_font_value = ttk.Label(font_bar, width=6, anchor="center")
        self.help_font_value.pack(side="left")
        self.help_font_larger_button = ttk.Button(
            font_bar, width=4, command=lambda: self.resize_help_font(1)
        )
        self.help_font_larger_button.pack(side="left")
        self.help_font_reset_button = ttk.Button(
            font_bar, command=lambda: self.resize_help_font(reset=True)
        )
        self.help_font_reset_button.pack(side="left", padx=8)
        self.resize_help_font()
        self.help_notebook = ttk.Notebook(win)
        self.help_notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.help_pages = []
        for section in ("start", "catalog", "books", "data"):
            page = ttk.Frame(self.help_notebook)
            self.help_notebook.add(page, text="")
            scrollbar = ttk.Scrollbar(page, orient="vertical")
            scrollbar.pack(side="right", fill="y")
            text = tk.Text(
                page, wrap="word", font=self.help_text_font, padx=12, pady=12,
                yscrollcommand=scrollbar.set, state="disabled"
            )
            text.pack(fill="both", expand=True)
            scrollbar.configure(command=text.yview)
            self.help_pages.append((section, page, text))
        self.help_close_button = ttk.Button(win, command=win.destroy)
        self.help_close_button.pack(anchor="e", padx=10, pady=(0, 10))
        self.refresh_help_language()

    def resize_help_font(self, delta=0, reset=False):
        """Resize only guide body text; keep the choice for this session."""
        self.help_font_size = 12 if reset else max(9, min(24, self.help_font_size + delta))
        self.help_text_font.configure(size=self.help_font_size)
        self.help_font_value.configure(text=f"{self.help_font_size} pt")
        self.help_font_smaller_button.configure(
            state="disabled" if self.help_font_size == 9 else "normal"
        )
        self.help_font_larger_button.configure(
            state="disabled" if self.help_font_size == 24 else "normal"
        )

    def refresh_help_language(self):
        win = getattr(self, "help_window", None)
        if win is None or not win.winfo_exists():
            return
        win.title(t("help_title"))
        self.help_intro_label.configure(text=t("help_intro"))
        self.help_close_button.configure(text=t("help_close"))
        self.help_font_label.configure(text=t("help_font_size"))
        self.help_font_smaller_button.configure(text=t("help_font_smaller"))
        self.help_font_larger_button.configure(text=t("help_font_larger"))
        self.help_font_reset_button.configure(text=t("help_font_reset"))
        for section, page, text in self.help_pages:
            self.help_notebook.tab(page, text=t(f"help_{section}_title"))
            position = text.yview()[0]
            text.configure(state="normal")
            text.delete("1.0", "end")
            text.insert("1.0", t(f"help_{section}_body"))
            text.configure(state="disabled")
            text.yview_moveto(position)

    def on_language_changed(self, event=None):
        """Comută limba UI (ro/en), persistă în settings.txt și actualizează
        elementele deja construite. Restul șirurilor se aplică la restart
        până când toate textele vor fi legate de t()."""
        new_lang = (self.lang_var.get() or "ro").strip().lower()
        if new_lang == get_language():
            return
        set_language(new_lang)
        try:
            save_setting(CONFIG_DIR / "settings.txt", "LANGUAGE", new_lang)
        except Exception:
            pass
        # Actualizează elementele vizibile care sunt deja internaționalizate
        self.status_label.configure(text=t("status_ready"))
        self.admin_toggle_button.configure(
            text=t("admin_show") if not self.admin_visible else t("admin_hide")
        )
        self.admin_frame.configure(text=t("admin_title"))
        self.help_button.configure(text=t("help_button"))
        self.refresh_help_language()
        self.set_status(
            "Limba interfeței: RO. (Se aplică complet la restart.)"
            if new_lang == "ro"
            else "UI language: EN. (Fully applied after restart.)"
        )

    def toggle_admin(self):
        self.admin_visible = not self.admin_visible
        if self.admin_visible:
            self.admin_frame.pack(fill="x", pady=(6, 0), before=self.status_label.master)
            self.admin_toggle_button.configure(text=t("admin_hide"))
        else:
            self.admin_frame.pack_forget()
            self.admin_toggle_button.configure(text=t("admin_show"))

    def admin_set_status(self, text):
        self.admin_status.configure(text=text)

    def admin_rebuild_taxonomy(self):
        from core.taxonomy import build_cache, find_col_zip
        zip_path = find_col_zip()
        if not zip_path:
            messagebox.showerror("Administrare", "Nicio arhiva .zip in col/.", parent=self.root)
            return
        if not messagebox.askyesno("Administrare",
            f"Se reconstruiește taxonomy.db din:\n\n{zip_path.name}\n\nOperația poate dura câteva minute. Continui?",
            parent=self.root):
            return
        self.admin_set_status("Reconstruiesc taxonomy.db...")
        def worker():
            try:
                stats = build_cache(zip_path)
                def _done():
                    text = f"Gata: {stats['accepted']:,} taxoni încărcați în taxonomy.db."
                    self.admin_set_status(text)
                    messagebox.showinfo("Administrare", text, parent=self.root)
                self.root.after(0, _done)
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror("Administrare", str(exc), parent=self.root))
        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------------------
    # Helper: încarcă un modul din tools/ (fără să fie pachet importabil)
    # ------------------------------------------------------------------
    def _load_tool_module(self, filename: str):
        import importlib.util
        path = APP_ROOT / "tools" / filename
        spec = importlib.util.spec_from_file_location(
            f"_appai_tool_{path.stem}", path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod

    def admin_import_vernacular(self):
        """Buton: 🌐 Import denumiri din CoL (eng + ron din VernacularName.tsv)."""
        from core.taxonomy import find_col_zip
        if not find_col_zip():
            messagebox.showerror("Administrare", "Nicio arhivă .zip în col/.", parent=self.root)
            return
        if not messagebox.askyesno("Administrare",
            "Se importă denumirile populare din arhiva CoL (eng + rom).\n"
            "Actualizează DOAR câmpurile goale, nu strică datele manuale.\n\nContinui?",
            parent=self.root):
            return
        self.admin_set_status("Import denumiri din CoL...")
        def worker():
            try:
                mod = self._load_tool_module("import_vernacular.py")
                stats = mod.import_names(
                    progress=lambda done, upd: self.root.after(
                        0, lambda d=done, u=upd: self.admin_set_status(
                            f"Import CoL: {d:,} linii, {u:,} actualizați...")))
                self.root.after(0, lambda: (
                    self.admin_set_status(
                        f"Import CoL gata: {stats['updated']:,} actualizați "
                        f"(en: {stats['en_count']:,}, ro: {stats['ro_count']:,})."),
                    messagebox.showinfo(
                        "Administrare",
                        f"Import CoL gata: {stats['updated']:,} taxoni actualizați "
                        f"(en: {stats['en_count']:,}, ro: {stats['ro_count']:,}).",
                        parent=self.root)))
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror(
                    "Administrare", f"Import CoL eșuat:\n\n{exc}", parent=self.root))
        threading.Thread(target=worker, daemon=True).start()

    def admin_extract_book(self):
        """Buton: 📄 Extract denumiri din PDF (carte/dicționar din name_sources/)."""
        pdf = filedialog.askopenfilename(
            title=t("extract_pdf_choose_title"),
            initialdir=str(APP_ROOT / "name_sources"),
            filetypes=[("PDF", "*.pdf")], parent=self.root)
        if not pdf:
            return
        page_range = simpledialog.askstring(
            t("extract_pdf_pages_title"),
            t("extract_pdf_pages_prompt"),
            parent=self.root)
        if page_range is None:  # Cancel la dialogul de pagini
            return
        self.admin_set_status(t("extract_pdf_preparing"))
        def worker():
            try:
                mod = self._load_tool_module("book_source.py")
                stats = mod.extract_draft(
                    pdf_path=pdf,
                    page_range=(page_range or "").strip() or None,
                    progress=lambda done, total, hits: self.root.after(
                        0, lambda d=done, t=total, h=hits: self.admin_set_status(
                            t("extract_pdf_progress").format(done=d, total=t, hits=h))))
                shown_range = stats.get("page_range") or t("extract_pdf_all_pages")
                self.root.after(0, lambda: (
                    self.admin_set_status(t("extract_pdf_done")),
                    messagebox.showinfo(
                        t("admin_title"),
                        t("extract_pdf_result_msg").format(
                            out_path=stats["out_path"],
                            pages=shown_range,
                            entries=stats["entries"],
                            matched=stats["matched"],
                            exact=stats["exact"],
                            ocr=stats["ocr"],
                            unmatched=stats["unmatched"]),
                        parent=self.root)))
            except Exception as exc:
                error = str(exc)
                self.root.after(0, lambda error=error: (
                    self.admin_set_status(t("extract_pdf_failed_status")),
                    messagebox.showerror(
                        t("admin_title"),
                        t("extract_pdf_failed_msg").format(error=error), parent=self.root)))
        threading.Thread(target=worker, daemon=True).start()

    def admin_edit_parse_rules(self):
        """Buton: ⚙ Reguli parsare — editează regex-urile pentru extract PDF."""
        try:
            mod = self._load_tool_module("book_source.py")
            mod.load_rules()  # creează fișierul implicit dacă lipsește
        except Exception as exc:
            messagebox.showerror("Administrare", f"Nu pot pregăti regulile:\n{exc}",
                                 parent=self.root)
            return
        rules_path = mod.RULES_PATH

        win = tk.Toplevel(self.root)
        win.title("Reguli parsare PDF")
        win.geometry("820x580")
        win.transient(self.root)

        ttk.Label(win, text=(
            "Aceste regex-uri definesc cum se recunosc intrările în PDF "
            "(antet specie, denumire populară, secțiuni străine etc.).\n"
            "Modifică doar dacă știi ce faci — fișierul e JSON, iar după salvare "
            "extractul următor va folosi noile reguli."
        ), wraplength=780, justify="left").pack(fill="x", padx=10, pady=(8, 4))

        txt = tk.Text(win, wrap="none", font=("Consolas", 10),
                      undo=True, width=110)
        txt.pack(fill="both", expand=True, padx=10, pady=4)
        txt.insert("1.0", rules_path.read_text(encoding="utf-8"))

        def save():
            content = txt.get("1.0", "end").strip()
            try:
                json.loads(content)
            except Exception as exc:
                messagebox.showerror("Reguli parsare", f"JSON invalid:\n\n{exc}",
                                     parent=win)
                return
            rules_path.write_text(content + "\n", encoding="utf-8")
            self.admin_set_status("Reguli parsare salvate.")
            win.destroy()

        bar = ttk.Frame(win)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(bar, text="💾 Salvează", command=save).pack(side="left", padx=(10, 6))
        ttk.Button(bar, text=t("btn_cancel"), command=win.destroy).pack(side="left")

    def admin_add_common_names(self):
        excel_dir = APP_ROOT / "tools"

        def create_template():
            out = filedialog.asksaveasfilename(
                title=t("excel_import_template_title"),
                initialdir=str(excel_dir),
                defaultextension=".xlsx",
                initialfile="denumiri_sablon.xlsx",
                filetypes=[("Excel", "*.xlsx")], parent=self.root)
            if not out:
                return
            try:
                write_names_template(Path(out))
            except Exception as exc:
                messagebox.showerror(t("admin_title"),
                    t("excel_import_template_failed").format(exc=exc), parent=self.root)
                return
            messagebox.showinfo(t("admin_title"),
                t("excel_import_template_saved").format(out=out), parent=self.root)

        try:
            import openpyxl  # noqa: F401 — verificăm dependența devreme
        except ImportError:
            messagebox.showerror(t("admin_title"), t("excel_import_need_openpyxl"), parent=self.root)
            return
        try:
            excel_dir.mkdir(parents=True, exist_ok=True)
            has_excel = any(p.is_file() and p.suffix.lower() == ".xlsx"
                            and not p.name.startswith("~$") for p in excel_dir.iterdir())
        except OSError as exc:
            messagebox.showerror(t("admin_title"),
                t("excel_import_read_error").format(exc=exc), parent=self.root)
            return
        if not has_excel:
            choice = messagebox.askyesnocancel(t("admin_title"),
                t("excel_import_missing_offer").format(folder=excel_dir), parent=self.root)
            if choice is None:
                return
            if choice:
                create_template()
                return
        path = filedialog.askopenfilename(title=t("excel_import_file_title"),
            initialdir=str(excel_dir),
            filetypes=[("Excel", "*.xlsx")], parent=self.root)
        if not path:
            return
        try:
            entries = read_excel_entries(Path(path))
        except Exception as exc:
            messagebox.showerror(t("admin_title"),
                t("excel_import_read_error").format(exc=exc), parent=self.root)
            return
        if not entries:
            if messagebox.askyesno(t("admin_title"), t("excel_import_none_offer"), parent=self.root):
                create_template()
            return
        if not messagebox.askyesno(t("admin_title"),
            t("excel_import_confirm").format(n=len(entries)), parent=self.root):
            return
        self.admin_set_status(t("excel_import_status").format(n=len(entries)))
        def worker():
            try:
                from core.taxonomy import connect_taxonomy
                con = connect_taxonomy()
                cols = {r[1] for r in con.execute("PRAGMA table_info(taxa)")}
                if "ro_name" not in cols: con.execute("ALTER TABLE taxa ADD COLUMN ro_name TEXT")
                if "en_name" not in cols: con.execute("ALTER TABLE taxa ADD COLUMN en_name TEXT")
                con.commit()
                updated = inserted = 0
                for sc, ro, en in entries:
                    row = con.execute("SELECT id, rank FROM taxa WHERE lower(name)=lower(?)", (sc,)).fetchone()
                    if row and row[1] == "species":
                        con.execute("UPDATE taxa SET ro_name=?, en_name=? WHERE id=?", (ro, en, row[0]))
                        updated += 1
                    else:
                        genus = sc.split()[0] if sc.split() else None
                        con.execute(
                            "INSERT OR IGNORE INTO taxa"
                            " (id, parent_id, name, rank, lineage_genus, ro_name, en_name)"
                            " VALUES (?, 'no-parent', ?, 'species', ?, ?, ?)",
                            (f"__common__{sc}", sc, genus, ro, en))
                        inserted += 1
                con.commit(); con.close()
                def _done(u=updated, i=inserted):
                    text = t("excel_import_done").format(u=u, i=i)
                    self.admin_set_status(text)
                    messagebox.showinfo(t("admin_title"), text, parent=self.root)
                self.root.after(0, _done)
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror(
                    t("admin_title"), t("excel_import_failed").format(exc=exc), parent=self.root))
        threading.Thread(target=worker, daemon=True).start()

    def admin_open_col_folder(self):
        """Deschide folderul col/ (arhiva CoL + taxonomy.db) în Explorer."""
        try:
            COL_DIR.mkdir(parents=True, exist_ok=True)
            os.startfile(str(COL_DIR))
        except Exception as exc:
            messagebox.showerror(
                "Administrare",
                f"Nu pot deschide folderul:\n{COL_DIR}\n\n{exc}",
                parent=self.root,
            )


    def update_species_classification(self):
        """
        Prezintă toate câmpurile de clasificare într-o singură
        fereastră cu formular — în loc de 4 ferestre simpledialog
        consecutive.
        """
        if self.image_path is None:
            return

        item = get_catalog_item(self.image_path, OUTPUT_DIR)
        if not item:
            messagebox.showinfo(
                "Catalog",
                "Analizează imaginea înainte de reclasificare.",
                parent=self.root,
            )
            return

        # Formular într-o singură fereastră
        form = tk.Toplevel(self.root)
        form.title("Update species")
        form.transient(self.root)
        form.grab_set()
        form.geometry("480x300")
        form.resizable(True, True)

        values = {}

        prompt_labels = {
            "category": t("form_category_hint"),
            "scientific_name": t("form_scientific_example"),
            "ro_name": t("form_ro_name"),
            "en_name": t("form_en_name"),
        }

        for idx, key in enumerate(["category", "scientific_name", "ro_name", "en_name"]):
            label = prompt_labels[key]

        for idx, (label, key) in enumerate([
            (t("form_category_hint"), "category"),
            (t("form_scientific_example"), "scientific_name"),
            (t("form_ro_name"), "ro_name"),
            (t("form_en_name"), "en_name"),
        ]):
            ttk.Label(form, text=label).pack(anchor="w", padx=12, pady=(8 if idx == 0 else 4, 0))
            initial = item.get(key, "") or ""
            if key == "category":
                initial = initial or "plante"
            entry = ttk.Entry(form, width=50)
            entry.insert(0, initial)
            entry.pack(padx=12, pady=2)
            values[key] = entry

        def on_ok():
            form.result = {k: entry.get() for k, entry in values.items()}
            form.destroy()

        def on_cancel():
            form.result = None
            form.destroy()

        buttons = ttk.Frame(form)
        buttons.pack(side="bottom", pady=12)
        ttk.Button(buttons, text=t("btn_ok"), command=on_ok).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text=t("btn_cancel"), command=on_cancel).pack(side="left")

        form.result = None
        form.wait_window()

        if form.result is None:
            return

        category = form.result["category"]
        scientific_name = form.result["scientific_name"].strip()
        ro_name = form.result["ro_name"].strip()
        en_name = form.result["en_name"].strip()

        if not scientific_name:
            messagebox.showwarning(
                "Update species",
                "Numele științific este obligatoriu.",
                parent=self.root,
            )
            return

        try:
            destination = update_catalog_classification(
                self.image_path,
                OUTPUT_DIR,
                category,
                scientific_name,
                ro_name,
                en_name,
            )
            self.update_catalog_buttons()
            self.set_status(f"Catalog updated: {destination.parent.name}")

            # Singurul pop-up de confirmare pentru reîncărcare
            if messagebox.askyesno(
                "Catalog updated",
                "Încadrarea a fost actualizată.\n\n"
                "Vrei să rulezi din nou promptul pentru a regenera descrierea?",
                parent=self.root,
            ):
                self.start_analysis()
            else:
                self.set_status(
                    f"Catalog updated; analysis pending: {destination.parent.name}"
                )
        except Exception as exc:
            messagebox.showerror("Catalog update failed", str(exc), parent=self.root)

    # În interiorul clasei MainWindow – adiacent altor metode ca get_catalog_item etc.
    # Sugerat: imediat după metoda update_species_classification()

    def on_thumbnail_click(self, image_path: str):
        """Afișează informațiile despre o poză catalogată când se dă click pe miniatură."""
        item = get_catalog_item(image_path, OUTPUT_DIR) or {}

        scientific = item.get("scientific_name", "Necunoscut")
        ro_name = item.get("ro_name", "")
        en_name = item.get("en_name", "")
        category_raw = item.get("category", "")
        category = map_category(category_raw)
        folder = get_catalog_folder(image_path, OUTPUT_DIR)

        info_lines = [
            f"Specie: {scientific}",
        ]
        if ro_name or en_name:
            info_lines.append(f"Nume: {ro_name} / {en_name}")
        info_lines.extend([
            f"Categorie: {category}",
            f"Folder: {folder}"
        ])

        info_text = "\n".join(info_lines)

        if hasattr(self, "catalog_info_label"):
            self.catalog_info_label.config(text=info_text)
        else:
            messagebox.showinfo("Info poză", info_text)


    def save_analysis_result(self, response: str, prompt: str, preset_name: str):
        # Analizele sunt salvate în SQLite în run_analysis.
        return

    def save_current_result(self):
        """Salvează analiza afișată în panoul Output Item.

        Actualizează response_text în catalog_items (sursa afișată la
        selecția imaginii). Fără catalog (imagine neanalizată) nu e nimic
        de salvat.
        """

        if self.image_path is None:
            return

        response = self.get_output_text()

        if not response:
            messagebox.showwarning(
                "Nothing to save",
                "There is no result to save."
            )
            return

        try:
            updated = update_catalog_response(
                self.image_path,
                OUTPUT_DIR,
                response.strip(),
            )
        except Exception as exc:
            messagebox.showerror(
                "Save error",
                str(exc)
            )
            return

        if not updated:
            messagebox.showinfo(
                "Save",
                "Imaginea nu este în catalog.\n\n"
                "Rulează ANALYZE mai întâi pentru a o identifica.",
                parent=self.root,
            )
            return

        self.set_output_text(
            response
        )

        self.set_status(
            f"Saved: {self.image_path.name}"
        )

    def save_current_profile(self):
        """
        Salvează profilul afișat în panoul Species Profile (inclusiv
        modificările făcute de utilizator) în tabela species_profiles,
        pentru specia imaginii selectate.
        """

        response = self.get_profile_output_text()

        if not response.strip():
            messagebox.showwarning(
                "Nothing to save",
                "There is no profile to save."
            )
            return

        item = get_catalog_item(
            self.image_path,
            OUTPUT_DIR,
        )

        if not item:
            messagebox.showinfo(
                "Species profile",
                "Analizează imaginea înainte de a salva profilul.",
                parent=self.root,
            )
            return

        scientific_name = (item.get("scientific_name") or "").strip()

        if not scientific_name or scientific_name in ("Necunoscut", "Unknown"):
            messagebox.showinfo(
                "Species profile",
                "Specia nu este încă identificată.",
                parent=self.root,
            )
            return

        try:
            updated = update_species_profile_response(
                OUTPUT_DIR,
                scientific_name,
                response.strip(),
            )
        except Exception as exc:
            messagebox.showerror(
                "Save error",
                str(exc),
                parent=self.root,
            )
            return

        if not updated:
            messagebox.showinfo(
                "Species profile",
                f"Nu există profil salvat pentru:\n\n{scientific_name}\n\n"
                "Apasă GEN PROFILE pentru a genera unul.",
                parent=self.root,
            )
            return

        self.set_status(
            f"Species profile saved: {scientific_name}"
        )

    # ========================================================
    # STATUS
    # ========================================================

    def update_save_button_state(self):

        if not hasattr(self, "save_button"):
            return

        has_result = (
            self.image_path is not None
            and bool(self.response_text.strip())
        )

        self.save_button.configure(
            state="normal" if has_result else "disabled"
        )

    def set_status(self, text: str):

        self.status_label.configure(
            text=text
        )

    def update_status_with_time(self):
        """
        Update status with elapsed time every 1 second.
        Constructs message dynamically based on current state.
        """
        if self.analysis_start_time is None and self.profile_start_time is None:
            return

        # Stop any existing timer before starting a new one
        if self.status_update_timer:
            self.root.after_cancel(self.status_update_timer)
            self.status_update_timer = None

        # Preferim timer-ul analizei; dacă nu este activ,
        # folosim timer-ul generări profilului speciei.
        if self.analysis_start_time is not None:
            start_time = self.analysis_start_time
        else:
            start_time = self.profile_start_time

        elapsed = time.perf_counter() - start_time
        elapsed_str = self.format_seconds(elapsed)

        # Construct message dynamically for batch mode
        if self.batch_total > 0:
            if self.batch_filename:
                base_message = f"Processing {self.batch_current + 1} / {self.batch_total}: {self.batch_filename}"
            else:
                base_message = f"Processing batch: {self.batch_current + 1} / {self.batch_total}"
        elif self.profile_start_time is not None:
            base_message = "Generating species profile"
        else:
            base_message = "Muse Glimmer is working"

        self.set_status(f"{base_message} — {elapsed_str}")

        # Schedule next update in 1 second
        self.status_update_timer = self.root.after(
            1000,
            self.update_status_with_time
        )

    def stop_status_timer(self):
        """
        Stop the status update timer.
        """
        if self.status_update_timer:
            self.root.after_cancel(self.status_update_timer)
            self.status_update_timer = None

    # ========================================================
    # HELPERS
    # ========================================================

    @staticmethod
    def format_seconds(seconds: float) -> str:

        if seconds < 60:
            return f"{seconds:.1f}s"

        minutes = int(seconds // 60)
        remaining = seconds % 60

        if minutes < 60:
            return f"{minutes}m {remaining:.1f}s"

        hours = int(minutes // 60)
        minutes = minutes % 60

        return (
            f"{hours}h "
            f"{minutes}m "
            f"{remaining:.1f}s"
        )

    def populate_prompt_list(self):
        """
        Listează prompt-urile în selector.

        La prima rulare, prompt-urile .txt din prompts/ sunt importate
        automat în baza de date (catalog.db). Sursa de verdade este
        baza de date; fișierele .txt rămâne pentru compatibilitate.
        """

        PROMPTS_DIR.mkdir(
            parents=True,
            exist_ok=True
        )

        database_path = OUTPUT_DIR / "catalog.db"

        sync_prompts_from_files(
            database_path,
            PROMPTS_DIR,
        )

        prompts = list_prompts(
            database_path
        )

        names = [
            item["name"]
            for item in prompts
            if item["name"] != SPECIES_PROFILE_PROMPT_NAME
        ]

        self.prompt_files = {
            item["name"]: PROMPTS_DIR / f"{item['name']}.txt"
            for item in prompts
        }

        self.prompt_combo["values"] = names

        if DEFAULT_PROMPT_NAME in names:
            self.prompt_var.set(DEFAULT_PROMPT_NAME)

        elif names:
            self.prompt_var.set(names[0])

    def load_selected_prompt(self):
        """
        Încarcă promptul selectat în editor din baza de date.
        """

        name = self.prompt_var.get()

        if not name:
            return

        item = load_prompt_from_db(
            OUTPUT_DIR / "catalog.db",
            name,
        )

        prompt = item.get("content", "") if item else ""

        if not prompt:
            self.set_status(
                f"Prompt not found: {name}"
            )
            return

        self.set_prompt_text(
            prompt
        )

        self.set_status(
            f"Prompt loaded: {name}"
        )

    def on_prompt_selected(self, event=None):
        """
        Când utilizatorul selectează un prompt,
        îl încărcăm automat.
        """

        self.load_selected_prompt()

    def update_prompt(self):

        name = self.prompt_var.get().strip()

        if not name:
            messagebox.showwarning(
                "No prompt selected",
                "Select a prompt first."
            )
            return

        prompt = self.get_prompt_text()

        if not prompt:
            messagebox.showwarning(
                "Empty prompt",
                "The prompt cannot be empty."
            )
            return

        try:

            save_prompt_to_db(
                OUTPUT_DIR / "catalog.db",
                name,
                prompt,
            )

            # Export automat în fișierul .txt corespunzător,
            # pentru compatibilitate cu CLI și versiunile anterioare.
            try:
                save_prompt(
                    PROMPTS_DIR / f"{name}.txt",
                    prompt,
                )
            except Exception:
                pass

            self.set_status(
                f"Prompt updated: {name}"
            )

        except Exception as exc:

            messagebox.showerror(
                "Update prompt error",
                str(exc)
            )

    def new_prompt(self):

        name = simpledialog.askstring(
            "New Prompt",
            "Enter prompt name:"
        )

        if name is None:
            return

        name = name.strip()

        if not name:
            messagebox.showwarning(
                "Invalid name",
                "The prompt name cannot be empty."
            )
            return

        if name.endswith(".txt"):
            name = name[:-len(".txt")]

        try:

            create_prompt_in_db(
                OUTPUT_DIR / "catalog.db",
                name,
                "",
            )

        except ValueError as exc:

            messagebox.showwarning(
                "Prompt already exists",
                str(exc)
            )
            return

        except Exception as exc:

            messagebox.showerror(
                "Create prompt error",
                str(exc)
            )

            return

        self.populate_prompt_list()

        self.prompt_var.set(
            name
        )

        self.set_prompt_text(
            ""
        )

        self.set_status(
            f"New prompt created: {name}"
        )
        

# ============================================================
# APPLICATION ENTRY POINT
# ============================================================

def main():

    root = tk.Tk()

    MainWindow(root)

    root.mainloop()


if __name__ == "__main__":
    main()
