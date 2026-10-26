from pathlib import Path
import json
import os
import re
import webbrowser
from urllib.parse import quote


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
    set_catalog_archived,
    count_archived_older_than,
    purge_archived,
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
    validate_species,
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

# ----------------------------------------------------
# Web lookups — verificarea externă a speciei identificate
# ----------------------------------------------------

COL_SEARCH_URL = "https://www.catalogueoflife.org/data/search?q="
WIKIPEDIA_SPECIES_URL = "https://en.wikipedia.org/wiki/"
INATURALIST_SEARCH_URL = "https://www.inaturalist.org/taxa/search?q="

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
# ANALYZE ALL — species that are missing from the taxonomy database
# ============================================================

def collect_batch_unverified_species(
    image_names,
    output_dir: Path = None,
) -> dict:
    """Group the batch species that are not verified in Catalogue of Life.

    ANALYZE ALL runs unattended, so the per-image CoL dialog never appears.
    This walks the freshly analyzed images instead and returns

        {scientific_name: {"images": [...], "status": ..., "accepted": ...}}

    so the caller can show a single list the user can walk through, image by
    image. A species the user already corrected through UPDATE SPECIES is
    never reported: the manual decision is the authoritative one.

    Returns an empty dict when every species is verified (or undecidable).
    """
    output_dir = Path(output_dir) if output_dir is not None else OUTPUT_DIR
    report: dict = {}
    seen: set = set()

    for image_name in image_names:
        if image_name in seen:
            continue
        seen.add(image_name)
        try:
            item = get_catalog_item(Path(image_name), output_dir)
        except Exception:
            continue
        if not item or item.get("manual_override"):
            continue
        name = (item.get("scientific_name") or "").strip()
        if not name or name in ("Necunoscut", "Unknown"):
            continue
        entry = report.get(name)
        if entry is None:
            entry = report[name] = {"images": [], "status": "", "accepted": None}
        entry["images"].append(Path(image_name).name)
        if entry["status"]:
            continue  # already validated this species
        try:
            result = validate_species(name, (item.get("category") or "").strip())
        except Exception:
            continue
        entry["status"] = result.get("status") or ""
        accepted = result.get("accepted") or {}
        entry["accepted"] = accepted.get("name") or None

    return {
        name: entry for name, entry in report.items()
        if entry["status"] and entry["status"] != "exact"
    }


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


class _RelayoutOnVisibility:
    """Mixin: tells a container when this widget is packed or unpacked.

    The analysis code shows and hides the progress bars with plain
    ``pack()``/``pack_forget()`` calls in a dozen places. Listening to
    <Map>/<Unmap> does not work here: a hidden group is unmapped, so its
    children never fire those events. Overriding the two pack methods keeps
    every existing call site untouched and still lets the bar re-flow.
    """

    on_visibility_change = None

    def pack(self, *args, **kwargs):
        result = super().pack(*args, **kwargs)
        self._notify_visibility()
        return result

    def pack_forget(self):
        result = super().pack_forget()
        self._notify_visibility()
        return result

    def _notify_visibility(self) -> None:
        callback = self.on_visibility_change
        if callback is not None:
            self.after_idle(callback)


class _FlowProgressBar(_RelayoutOnVisibility, ttk.Progressbar):
    """Progressbar that re-flows the button bar when shown or hidden."""


class _FlowProgressFrame(_RelayoutOnVisibility, ttk.Frame):
    """Frame that re-flows the button bar when shown or hidden."""


# ============================================================
# BUTTON BAR WITH REFLOW
# ============================================================

class _FlowBar(ttk.Frame):
    """Places button groups on rows that adapt to the window width.

    Tk's pack never wraps: with thirteen wide buttons the groups packed with
    side="right" (GEN PROFILE, LOAD PROFILE, SAVE PROFILE, GEN PROFILE ALL)
    were allocated whatever width was left, which was nothing, so they simply
    disappeared instead of being clipped. This container measures the groups
    and moves the ones that no longer fit onto a second row, so every button
    stays reachable at any window width.

    The groups are positioned with grid on purpose: ``pack(in_=...)`` does NOT
    re-parent a widget in Tk, it only borrows another container's cavity, so a
    nested pack layout leaves the groups outside the bar where the surrounding
    panels paint over them. Grid only changes row/column, so the groups keep
    their parent, their state (disabled/enabled) and their widgets.

    While everything fits on a single row the last group keeps the old
    flush-right look, so the bar is unchanged at the default window size.
    """

    GROUP_BORDER = "#b9c0cc"
    GROUP_BORDER_ACTIVE = "#8b95a3"
    LINE_GAP = 4


    def __init__(self, master, **kwargs) -> None:
        super().__init__(master, **kwargs)
        # Each entry: [group, padx, dynamic, shown]
        self._groups: list[list] = []
        self._plan: tuple | None = None
        self.bind("<Configure>", self._on_configure, add="+")

    def add(
        self,
        group: tk.Widget,
        padx: tuple = (0, 0),
        dynamic: bool = False,
    ) -> None:
        """Registers a group. The group must be a child of this container.

        A *dynamic* group is shown or hidden by its own contents (the progress
        bars are packed/packed_forget all over the analysis code), so it only
        takes part in the layout while one of its children is managed.
        """
        self._groups.append([group, padx, dynamic, True])
        self.relayout(force=True)

    def row_count(self) -> int:
        return len(self._plan) if self._plan else 0

    def groups(self) -> list[tk.Widget]:
        return [entry[0] for entry in self._groups]

    def _on_configure(self, event) -> None:
        if event.widget is self:
            self.relayout()

    def _is_shown(self, entry: list) -> bool:
        group, _, dynamic, shown = entry
        if not shown:
            return False
        if not dynamic:
            return True
        children = group.winfo_children()
        if not children:
            return False
        return any(child.winfo_manager() for child in children)

    def _compute_plan(self, available: int) -> tuple:
        """Greedy wrap: a group moves down only when it no longer fits."""
        rows: list[tuple] = []
        current: list[int] = []
        used = 0
        for index, entry in enumerate(self._groups):
            if not self._is_shown(entry):
                continue
            group, padx, _, _ = entry
            width = group.winfo_reqwidth() + padx[0] + padx[1]
            if current and used + width > available:
                rows.append(tuple(current))
                current, used = [], 0
            current.append(index)
            used += width
        if current:
            rows.append(tuple(current))
        return tuple(rows)

    def relayout(self, force: bool = False) -> None:
        available = self.winfo_width()
        if available <= 1:
            return  # not laid out yet; <Configure> will call again
        plan = self._compute_plan(available)
        if plan == self._plan and not force:
            return
        self._plan = plan

        # Groups are positioned with place(), not grid: grid shares one set of
        # column widths across every row, so a group placed on a second line
        # would start at the x of the widest group above it instead of right
        # after its neighbour. place() keeps every group a child of this bar
        # (no re-parenting, widgets and their state survive) while allowing a
        # true left-to-right flow, and the bar's height is set by hand.
        for entry in self._groups:
            entry[0].place_forget()

        single_row = len(plan) == 1
        y = 0
        for line in plan:
            x = 0
            line_height = 0
            for position, index in enumerate(line):
                group, padx, _, _ = self._groups[index]
                width = group.winfo_reqwidth()
                height = group.winfo_reqheight()
                offset = x + padx[0]
                if single_row and position == len(line) - 1:
                    # Preserve the original flush-right profile group.
                    offset = max(0, available - width - padx[1])
                group.place(x=offset, y=y, width=width, height=height)
                x += width + padx[0] + padx[1]
                line_height = max(line_height, height)
            y += line_height + self.LINE_GAP

        # place() does not grow the container, so the height is imposed here.
        self.configure(height=max(0, y - self.LINE_GAP))
        self.pack_propagate(False)


# ============================================================
# MAIN WINDOW
# ============================================================

def _listbox_text(box, index: int) -> str:
    """The text of one listbox row.

    tk.Listbox has no getter for a single item (item() is a Treeview API), so
    the list is read and sliced. Only used on short image lists, and only
    when a prefix actually changed.
    """
    return box.get(index)


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

        # PROFILE_NUM_PREDICT este o PLAFON, nu o rezervă: Ollama se oprește
        # la finalul natural al răspunsului. O valoare mai mică nu face
        # generarea mai rapidă — taie doar profilele lungi, iar contractul
        # JSON rămâne incomplet, deci profilul e pierdut.
        try:
            self.profile_num_predict = int(
                settings.get("PROFILE_NUM_PREDICT", "12000")
            )
        except ValueError:
            self.profile_num_predict = 12000

        # Bugetul de rezervă: când răspunsul se trunchiază, profilul este
        # regenerat o singură dată cu acest plafon, în loc de a eșua.
        try:
            self.profile_num_predict_retry = int(
                settings.get("PROFILE_NUM_PREDICT_RETRY", "20000")
            )
        except ValueError:
            self.profile_num_predict_retry = 20000
        # Speciile care au avut nevoie de rezervă: raportate la final, ca să
        # fie vizibil când plafonul de bază e prea mic.
        self.profile_retried = 0

        # Profilul de specie trimite doar text, dar șablonul taxonomic
        # (~21k caractere) + contractul JSON (~3.7k) ocupă aproape tot
        # contextul, iar răspunsul (descriere + JSON) nu mai încapea.
        # De aceea are un context propriu, independent de analiza cu
        # imagine, care rămâne la NUM_CTX pentru viteză și VRAM.
        try:
            self.profile_num_ctx = int(
                settings.get("PROFILE_NUM_CTX", "32768")
            )
        except ValueError:
            self.profile_num_ctx = 32768

        # Orice plafon mai mare decât spațiul rămas în context ar produce
        # aceeași trunchiere, doar mai lent.
        self.profile_num_predict_retry = min(
            self.profile_num_predict_retry,
            max(self.profile_num_predict, self.profile_num_ctx - 8000),
        )

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
        # Cooperative stop: set by the STOP button or Ctrl+. and read at the
        # top of every batch loop. False means "no batch running".
        self._cancel_batch = False
        self._batch_running = False
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
        # Per-image timing for ANALYZE ALL: when the current image started,
        # so the status can show time-on-this-image plus a running average.
        self.batch_image_start_time = None
        # Running timing stats over REAL (non-cached) analyses only, so
        # cached items (served in ~0s) don't drag the average and ETA down.
        self.batch_analyzed_count = 0
        self.batch_analyzed_time = 0.0
        self.batch_cached_current = False

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

        self.set_status(t("status_ready_idle"))

        # Verificare la pornire: ce lipsește față de un setup complet?
        # Arată un singur popup informativ (nu la fiecare pornire dacă
        # utilizatorul îl ignoră — doar când lipsește ceva esențial).
        self.root.after(600, self.check_first_run_setup)

    # ========================================================
    # UI
    # ========================================================

    def _flow_group(
        self,
        padx: tuple = (0, 0),
        dynamic: bool = False,
    ) -> ttk.Frame:
        """Creates an outlined group for the button bar.

        The blocks (analysis, web lookups, progress, species profile) get a
        thin border so they read as separate units instead of one long row of
        buttons. The border colours live on _FlowBar, so the look is tuned in
        one place.

        A *dynamic* group only takes part in the layout while it has a visible
        child — that is how the progress group stays out of the bar until an
        analysis actually shows a progress bar.
        """
        frame = tk.Frame(
            self.button_bar,
            highlightthickness=1,
            highlightbackground=_FlowBar.GROUP_BORDER,
            highlightcolor=_FlowBar.GROUP_BORDER_ACTIVE,
            bd=0,
        )
        inner = ttk.Frame(frame, padding=(6, 3))
        inner.pack(fill="both", expand=True)
        self.button_bar.add(frame, padx=padx, dynamic=dynamic)
        return inner

    def _progress_visibility_changed(self) -> None:
        """A progress bar appeared or disappeared: re-flow the button bar.

        The group is only in the layout while one of its children is packed,
        so the space it needs is reserved before the bar becomes visible
        instead of pushing the following group over.
        """
        bar = getattr(self, "button_bar", None)
        if bar is not None:
            bar.relayout(force=True)

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

        # weight: prompt_frame e deja îngust (width=42 pe Text) și nu absoarbe
        # spațiu extra; restul merge la species_prompt_frame, ca textele lui
        # să fie vizibile la deschidere.
        top_paned.add(image_frame, weight=1)

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
            # EXTENDED gives the Explorer semantics for free through Tk's own
            # class bindings: plain click keeps one row, Ctrl toggles, Shift
            # extends. The selection is the subset for batch work; the photo
            # still follows the row that was actually clicked.
            selectmode=tk.EXTENDED,
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

        # Ctrl+. is the browser's "stop loading" convention, and unlike Esc it
        # is not already taken by the dialogs. It cannot fire by accident.
        self.root.bind_all(
            "<Control-period>", self.cancel_batch, add="+"
        )

        self.image_listbox.bind(
            "<<ListboxSelect>>",
            self.on_image_selected
        )
        # Records the row the pointer hit: curselection() reports the set, not
        # the row acted on. Bound with add="+" so the Tk class binding still
        # performs the Ctrl / Shift selection.
        self.image_listbox.bind(
            "<Button-1>",
            self.on_image_list_clicked,
            add="+"
        )

        # Keyboard navigation: <<ListboxSelect>> is not delivered when the
        # highlight moves with the arrow keys, so the keys are handled here
        # and the photo is loaded once the listbox has moved its selection.
        for key in ("<Up>", "<Down>", "<Prior>", "<Next>", "<Home>", "<End>"):
            self.image_listbox.bind(
                key,
                self._on_image_navigate,
                add="+"
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

        # Click pe miniatura din Images: deschide poza MARE într-o fereastră
        # separată (același popup ca în Quick Catalog) — util când thumb-ul
        # de 500x250 e prea mic pentru determinare.
        self.preview_label.bind("<Button-1>", self._open_main_preview)
        self.preview_label.configure(cursor="hand2")

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

        top_paned.add(prompt_frame, weight=0)

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
        # width=42: fără o lățime explicită, tk.Text cere 80 caractere
        # (~640px) și lățimea cadrelui e dominată de caseta de text, nu de
        # rândul cu butoane — de aceea Promptul acoperea Species Prompt.
        self.prompt_text = tk.Text(
            prompt_frame,
            height=7,
            width=42,
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

        top_paned.add(species_prompt_frame, weight=3)

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

        # Plasează divizorul (sash) dintre Prompt și Species Prompt exact
        # după rândul de controale (adică după butonul Update). Weight=0 pe
        # prompt_frame îl ține la cererea lui; snap-ul acesta îl fixează
        # precis, indiferent de font/DPI.
        # ATENȚIE: after_idle rulează ÎNAINTE ca fereastra să fie mapată pe
        # ecran — winfo_width()/winfo_rootx() întorc valori goale (≈0) și
        # sashpos primea o poziție lângă zero, strângând Images + Prompt în
        # o bandă invizibilă. De aceea reîncercăm la 50ms până când fereastra
        # e mapată și măsura e plauzibilă (max 40 × 50ms = 2s).
        def _snap_prompt_sash(tries=0):
            try:
                top_paned.update_idletasks()
                if not self.root.winfo_ismapped():
                    raise RuntimeError("window not mapped yet")
                total = top_paned.winfo_width()
                # marginea dreaptă a rândului de controale, în coordonate
                # relative ferestrei Panedwindow
                right_edge = (prompt_controls.winfo_rootx()
                              - top_paned.winfo_rootx()
                              + prompt_controls.winfo_width())
                # măsurătoare plauzibilă: controalele chiar au o poziție
                # în interiorul panoului (nu 0/0 de dinainte de mapare)
                if total <= 200 or right_edge < 100 or right_edge >= total:
                    raise RuntimeError(
                        f"bad measure right_edge={right_edge} total={total}")
                top_paned.sashpos(1, right_edge + 14)
            except Exception:
                if tries < 40:
                    self.root.after(
                        50, lambda: _snap_prompt_sash(tries + 1))

        self.root.after_idle(_snap_prompt_sash)


        # ----------------------------------------------------
        # Buttons
        # ----------------------------------------------------

        button_frame = ttk.Frame(main)

        button_frame.pack(
            fill="x",
            pady=(0, 10)
        )

        # Butoanele se reflow-ează pe linii noi când fereastra îngustează
        # (pack nu are reflow, iar grupul din dreapta dispărea complet).
        self.button_bar = _FlowBar(button_frame)
        self.button_bar.pack(side="top", fill="x")

        # ----------------------------------------------------
        # Grup stânga: fluxul ITEM (analiza imaginii)
        # ----------------------------------------------------

        analysis_buttons = self._flow_group()

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

        # Soft-delete: ascunde poza din SpecioCatalog fără a șterge datele.
        # Textul se schimbă între „Arhivează" și „Restaurează" după stare.
        self.archive_button = ttk.Button(
            analysis_buttons,
            text=t("btn_archive"),
            command=self.toggle_archive_current,
            state="disabled"
        )
        self.archive_button.pack(side="left", padx=(8, 0))

        # ----------------------------------------------------
        # Grup mijloc: verificare web (CoL + Wikipedia)
        # ----------------------------------------------------

        web_buttons = self._flow_group(padx=(6, 0))

        # Progress group: created here, still empty, so it keeps its place in
        # the flow between the web lookups and the species profile. The bars
        # themselves are built further down, with the rest of the progress
        # code. Being dynamic, the group takes no space while both bars are
        # hidden, so the bar never reserves an empty row.
        progress_row = self._flow_group(padx=(6, 0), dynamic=True)

        # Câmp editabil: se preumplește cu specia curentă, dar primește
        # și alternative scrise de utilizator (ex. un sinonim de verificat).
        self.web_entry = ttk.Entry(web_buttons, width=20)
        self.web_entry.pack(side="left")

        self.col_button = ttk.Button(
            web_buttons,
            text=t("btn_open_col"),
            command=self.open_col_page
        )
        self.col_button.pack(side="left", padx=(6, 0))

        # Mereu activ: căutarea nu depinde de starea catalogului.
        self.wikipedia_button = ttk.Button(
            web_buttons,
            text=t("btn_open_wikipedia"),
            command=self.open_wikipedia_page
        )
        self.wikipedia_button.pack(side="left", padx=(6, 0))

        # iNaturalist: pagina de taxoni (identificare + ilustrații). Spre
        # deosebire de CoL, acceptă query direct în URL, dar păstrăm și
        # numele în clipboard ca la CoL, pentru lipire manuală dacă e cazul.
        self.inaturalist_button = ttk.Button(
            web_buttons,
            text=t("btn_open_inaturalist"),
            command=self.open_inaturalist_page
        )
        self.inaturalist_button.pack(side="left", padx=(6, 0))

        # ----------------------------------------------------
        # Grup dreapta: fluxul SPECIES PROFILE
        # ----------------------------------------------------

        profile_buttons = self._flow_group(padx=(6, 0))

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
        # The group itself was created with the other blocks, so it sits
        # between the web lookups and the species profile.
        # Stop lives in its own holder, packed *before* the progress bar so it
        # ends up to its right — the last thing in the group, immediately
        # before the row wraps to GEN PROFILE. Inside the batch group it
        # would have landed between the two progress bars.
        #
        # The holder stays managed but is empty when idle, so it costs no
        # width between batches.
        self.cancel_holder = ttk.Frame(progress_row)
        self.cancel_holder.pack(side="right")

        self.cancel_batch_button = ttk.Button(
            self.cancel_holder,
            text=t("btn_stop"),
            command=self.cancel_batch,
        )
        # Shown only while a batch runs.
        self.cancel_batch_button.pack_forget()

        self.progress = _FlowProgressBar(
            progress_row,
            mode="indeterminate",
            length=150
        )
        self.progress.on_visibility_change = self._progress_visibility_changed

        self.progress.pack(
            side="right"
        )

        # ----------------------------------------------------
        # Batch progress group
        # ----------------------------------------------------

        self.batch_progress_frame = _FlowProgressFrame(
            progress_row
        )
        self.batch_progress_frame.on_visibility_change = (
            self._progress_visibility_changed
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

        ttk.Button(
            admin_inner,
            text=t("btn_purge_archived"),
            command=self.admin_purge_archived,
            width=24
        ).pack(side="left", padx=(8, 0))

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

        # Primul calcul al rândurilor după ce toate etichetele au fost puse
        # (lățimile se schimbă odată cu fontul și cu limba).
        self.button_bar.relayout(force=True)

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
                t("warn_empty_prompt_title"),
                t("warn_prompt_empty_body"),
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
                t("err_update_prompt"),
                str(exc),
                parent=self.root,
            )
            return

        self.set_species_prompt_text(text)
        self.set_status(t("status_species_prompt_updated"))

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
                t("err_initial_config"),
                t("err_cannot_open_readme").format(readme, exc),
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
            t("status_added_images").format(added)
        )


    def refresh_image_list(self):
        """Paints the status prefix of each row.

        The rows are updated in place. Deleting and re-inserting them — the
        obvious implementation — threw away the selection, the highlight and
        the keyboard focus on every call, and ANALYZE ALL calls this after
        each image. The effect was that the arrow keys stopped working and
        any selection vanished several times per minute.
        """

        labels = []
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

            labels.append(prefix + path.name)

        box = self.image_listbox

        if box.size() == len(labels):
            # Same rows, maybe a different prefix. A Listbox has no item()
            # setter, so a row is replaced by delete + insert at the same
            # index — and deleting the *active* row drops the selection and
            # moves the anchor. During a batch that row is usually the one
            # being processed, i.e. the one under the user's highlight, so the
            # state is saved and put back.
            selection = [int(i) for i in box.curselection()]
            try:
                active = int(box.index("active"))
            except (tk.TclError, ValueError):
                active = -1

            for index, text in enumerate(labels):
                if _listbox_text(box, index) != text:
                    box.delete(index)
                    box.insert(index, text)

            for index in selection:
                if 0 <= index < len(labels):
                    box.selection_set(index)
            if 0 <= active < len(labels):
                box.activate(active)
            return

        # The list itself changed (added, removed, cleared). Keep the user
        # where they were instead of snapping back to the first row.
        selection = [
            int(i) for i in box.curselection()
        ]
        try:
            active = int(box.index("active"))
        except (tk.TclError, ValueError):
            active = -1
        had_focus = self.root.focus_get() is box

        box.delete(0, tk.END)
        for text in labels:
            box.insert(tk.END, text)

        for index in selection:
            if 0 <= index < len(labels):
                box.selection_set(index)
        if 0 <= active < len(labels):
            box.activate(active)
        if had_focus:
            box.focus_set()


    def select_image(self, index: int):

        if not self.image_paths:
            return

        if index < 0 or index >= len(self.image_paths):
            return

        # Same path as a click: it also re-anchors the listbox, so a later
        # click cannot be dragged to a stale row.
        self._show_index(index, select=True)


    def _on_image_navigate(self, _event=None) -> None:
        """Shows the photo when moving through the image list with the keys.

        The selection is read on an idle callback: at key-press time the
        listbox has not moved its selection yet, so loading here would always
        show the previous photo.
        """
        if getattr(self, "image_listbox", None) is None:
            return
        # A key press means the mouse is no longer driving the list, so any
        # remembered click must not outlive this event.
        self._clicked_index = -1
        self.root.after_idle(self.load_image_from_selection)

    def _selected_index(self) -> int:
        """The row the user acted on, or -1 when nothing is selected.

        curselection()[0] is the *first* selected row, which is the wrong
        answer as soon as several rows are selected: after Ctrl+clicking row
        9 with row 3 already selected, the photo must follow row 9, not row 3.
        A click records its own row, and the keyboard keeps using the
        highlight.
        """
        clicked = getattr(self, "_clicked_index", -1)
        if 0 <= clicked < len(self.image_paths):
            return clicked
        selection = self.image_listbox.curselection()
        return int(selection[0]) if selection else -1

    def selected_indices(self) -> list[int]:
        """Every selected row, in list order."""
        return sorted(int(i) for i in self.image_listbox.curselection())

    def _batch_indices(self) -> list[int]:
        """Rows a batch command works on: the selection, or everything.

        Three cases, and the middle one is why a single selected row does not
        count as a filter:

        * no selection, or a full one — every image, as before the feature;
        * a lone selected row — that is just the *current* image.
          add_folder() and add_images() end with select_image(0), so a freshly
          loaded list always has exactly one row selected. Treating that as a
          filter made GEN PROFILE ALL process a single photo.
        * several rows — a deliberate Ctrl/Shift selection, so honour it.
        """
        chosen = self.selected_indices()
        if 1 < len(chosen) < len(self.image_paths):
            return chosen
        return list(range(len(self.image_paths)))

    def _note_selection_scope(self) -> int:
        """Tells the user how many images a batch command will touch.

        Without this, a leftover selection silently changes what ANALYZE ALL
        does — the exact surprise a multi-select list is prone to.
        """
        total = len(self.image_paths)
        chosen = len(self._batch_indices())
        if chosen < total:
            self.set_status(
                t("status_selection_scope").format(chosen=chosen, total=total)
            )
        return chosen

    def _highlighted_index(self) -> int:
        """The item under the keyboard highlight, or -1 when there is none.

        index("active") is what the arrow keys move, so the keyboard path
        follows it. It is deliberately NOT used for mouse clicks: the anchor
        keeps its old value when the list is only refilled, and honouring it
        then scrolled the list to another row and loaded the wrong photo.
        """
        try:
            return int(self.image_listbox.index("active"))
        except (tk.TclError, ValueError):
            return self._selected_index()

    def _show_index(self, index: int, select: bool) -> None:
        """Selects (when asked) and loads the image at *index*.

        Both entry points go through here, so the mouse and the keyboard end
        up in the same state: a properly highlighted item and the photo on
        screen. Loading is skipped when the photo is already current, which
        also stops the keyboard path from loading twice (selecting the item
        re-enters through <<ListboxSelect>>).
        """
        if index < 0 or index >= len(self.image_paths):
            return
        path = self.image_paths[index]
        if select:
            self.image_listbox.selection_clear(0, tk.END)
            self.image_listbox.selection_set(index)
        # Keep the anchor on the item we just showed. Without this the anchor
        # could stay on an old row and hijack the next mouse click.
        try:
            self.image_listbox.activate(index)
        except tk.TclError:
            pass
        try:
            self.image_listbox.see(index)
        except tk.TclError:
            pass
        if path == self.image_path:
            return
        self.load_selected_image(path)

    def load_image_from_selection(self) -> None:
        """Keyboard path: highlight the item and show its photo."""
        self._show_index(self._highlighted_index(), select=True)

    def on_image_list_clicked(self, event=None):
        """Remembers the row under the pointer.

        With a multi-selection there is no longer a single "selected" row:
        Ctrl+click adds a row, Shift+click extends a range, and in both cases
        the photo must follow the row that was clicked, not the first selected
        one.
        """
        if event is None:
            self._clicked_index = -1
            return
        try:
            self._clicked_index = int(self.image_listbox.nearest(event.y))
        except tk.TclError:
            self._clicked_index = -1

    def on_image_selected(self, event=None):

        # Mouse path: the click already chose the row, so the recorded index
        # is the truth — but only for this event. Widget bindings run before
        # Tk's class bindings, so a click that is left in place would pin
        # the list to that row and swallow every later arrow key.
        index = self._selected_index()
        self._clicked_index = -1
        self._show_index(index, select=False)


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
                t("status_selected").format(path.name)
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
                t("err_image"),
                str(exc)
            )

    def update_catalog_buttons(self):
        state = "disabled"
        profile_state = "disabled"
        load_profile_state = "disabled"
        web_name = ""
        archived = False

        if self.image_path is not None:
            try:
                item = get_catalog_item(self.image_path, OUTPUT_DIR)
                if item:
                    state = "normal"
                    archived = bool(item.get("archived"))
                    scientific_name = (item.get("scientific_name") or "").strip()
                    if scientific_name and scientific_name not in ("Necunoscut", "Unknown"):
                        profile_state = "normal"
                        # Câmpul de căutare web urmează specia curentă.
                        web_name = scientific_name

                        # LOAD PROFILE activ doar dacă există deja
                        # un profil salvat în baza de date.
                        if get_species_profile(OUTPUT_DIR, scientific_name):
                            load_profile_state = "normal"
            except Exception:
                pass

        self.open_species_button.configure(state=state)
        self.update_species_button.configure(state=state)
        # Arhivarea e disponibilă pentru orice poză catalogată; textul
        # comută între „Arhivează" și „Restaurează" după starea curentă.
        # Butonul rămâne activ cât există o imagine selectată, ca la apăsare
        # pe o poză necatalogată să arate mesajul de ghidare (cum se încarcă
        # poza din output/), nu să fie inactiv fără explicație.
        self.archive_button.configure(
            state="normal" if self.image_path is not None else "disabled",
            text=t("btn_restore") if archived else t("btn_archive"),
        )
        # Sincronizează câmpul de căutare (specie nouă selectată /
        # reanalyzată); textul scris manual se înlocuiește la schimbarea
        # contextului — utilizatorul îl poate reintroduce.
        try:
            self.web_entry.delete(0, "end")
            self.web_entry.insert(0, web_name)
        except Exception:
            pass
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

    def toggle_archive_current(self):
        """Comută starea de arhivare a pozei curente (soft-delete reversibil).

        Poza rămâne pe disc și în baza de date; doar flag-ul ``archived``
        schimbă, astfel încât SpecioCatalog o ascunde (Arhivează) sau o
        reafișează (Restaurează). Nu șterge nimic — decizia e reversibilă.
        """
        if self.image_path is None:
            return
        try:
            item = get_catalog_item(self.image_path, OUTPUT_DIR)
        except Exception:
            item = None
        if not item:
            messagebox.showinfo(
                t("title_catalog"), t("info_archive_howto"),
                parent=self.root,
            )
            return

        was_archived = bool(item.get("archived"))
        try:
            changed = set_catalog_archived(
                self.image_path, OUTPUT_DIR, archived=not was_archived
            )
        except Exception as exc:
            messagebox.showerror(
                t("err_catalog_update"), str(exc), parent=self.root
            )
            return

        if changed:
            self.update_catalog_buttons()
            if was_archived:
                self.set_status(t("status_restored").format(self.image_path.name))
            else:
                self.set_status(t("status_archived").format(self.image_path.name))

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

        self.set_status(t("status_profile_model").format(model))

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

    def generate_profile_response(self, full_prompt: str) -> str:
        """Runs the profile model, retrying once when the answer is cut off.

        A truncated answer is not a partial success: the JSON contract closes
        at the very end of the response, so nothing can be saved. Rather than
        failing the species and forcing a manual re-run of the whole batch,
        the profile is generated once more with the reserve budget.

        num_predict is a ceiling, not a reservation — the retry costs time
        only for the species that actually need it.
        """
        budget = self.profile_num_predict
        last_error = None

        for attempt in (1, 2):
            data = generate(
                ollama_url=self.ollama_url,
                model=self.profile_model or self.model,
                prompt=full_prompt,
                image_base64=None,
                num_ctx=self.profile_num_ctx,
                num_predict=budget,
                thinking=self.profile_thinking,
            )

            if data.get("done_reason") != "length":
                response = get_response(data)
                if not response.strip():
                    raise OllamaError("Modelul nu a returnat nimic.")
                return response

            last_error = (
                f"Răspunsul a fost trunchiat de limita de tokenuri "
                f"(PROFILE_NUM_PREDICT={budget})."
            )
            reserve = self.profile_num_predict_retry
            if attempt == 1 and reserve > budget:
                budget = reserve
                self.profile_retried += 1
                continue
            break

        raise OllamaError(
            f"{last_error} Specia are nevoie de mai mult de "
            f"{budget} tokeni; mărește PROFILE_NUM_PREDICT_RETRY în "
            f"config/settings.txt sau PROFILE_NUM_CTX."
        )

    def cancel_batch(self, event=None):
        """Asks the running batch to stop after the current item.

        The flag is only read at the top of the loop, so an image already
        handed to the model still finishes instead of being abandoned
        half-written. Closing the window (X) still stops everything at once,
        but that throws the whole session away.
        """
        if getattr(self, "_batch_running", False):
            self._cancel_batch = True
            self.cancel_batch_button.configure(
                text=t("btn_stopping"), state="disabled"
            )
            self.set_status(t("status_stopping"))

    def batch_was_cancelled(self) -> bool:
        return bool(getattr(self, "_cancel_batch", False))

    def _begin_batch(self) -> None:
        self._cancel_batch = False
        self._batch_running = True
        self.cancel_batch_button.configure(
            text=t("btn_stop"), state="normal"
        )
        # winfo_manager, not winfo_ismapped: a packed widget inside an
        # unmapped window still reports "not mapped", and the button would
        # then never be hidden again.
        if self.cancel_batch_button.winfo_manager() != "pack":
            self.cancel_batch_button.pack(side="left", padx=(10, 0))

    def _end_batch(self) -> bool:
        """Clears the running state and reports whether it was stopped."""
        self._batch_running = False
        stopped = self.batch_was_cancelled()
        self._cancel_batch = False
        if self.cancel_batch_button.winfo_manager() == "pack":
            self.cancel_batch_button.pack_forget()
        return stopped

    def generate_species_profile(self):

        if self.image_path is None:
            return

        item = get_catalog_item(
            self.image_path,
            OUTPUT_DIR,
        )

        if not item:
            messagebox.showinfo(
                t("title_species_profile"),
                t("info_analyze_before_gen_profile"),
                parent=self.root,
            )
            return

        scientific_name = (item.get("scientific_name") or "").strip()

        if not scientific_name or scientific_name in ("Necunoscut", "Unknown"):
            messagebox.showinfo(
                t("title_species_profile"),
                t("info_species_unknown"),
                parent=self.root,
            )
            return

        confirmed = messagebox.askyesno(
            t("title_gen_profile"),
            t("confirm_gen_profile").format(scientific_name, self.profile_model or self.model),
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
            t("status_generating_profile").format(scientific_name)
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

                response = self.generate_profile_response(full_prompt)

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
                t("status_profile_generated_time").format(scientific_name, elapsed_str)
            )

            # Pop-up de finalizare, ca la ANALYZE
            messagebox.showinfo(
                t("info_profile_done_title"),
                t("info_profile_generated").format(scientific_name, elapsed_str),
                parent=self.root,
            )
        else:
            self.set_status(
                t("status_profile_generated_short").format(scientific_name)
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
            t("err_species_profile"),
            str(exc),
            parent=self.root,
        )

        self.set_status(
            t("status_profile_failed")
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
                t("title_species_profile"),
                t("info_not_in_catalog"),
                parent=self.root,
            )
            return

        scientific_name = (item.get("scientific_name") or "").strip()

        if not scientific_name or scientific_name in ("Necunoscut", "Unknown"):
            messagebox.showinfo(
                t("title_species_profile"),
                t("info_species_unknown_run"),
                parent=self.root,
            )
            return

        profile = get_species_profile(
            OUTPUT_DIR,
            scientific_name,
        )

        if not profile or not (profile.get("profile_response") or "").strip():
            messagebox.showinfo(
                t("title_species_profile"),
                t("info_no_profile_saved").format(scientific_name),
                parent=self.root,
            )
            return

        self._load_profile_into_pane(scientific_name, profile)

        self.set_status(
            t("status_profile_loaded").format(scientific_name)
        )

    def ask_profile_scope(self, saved: int, missing: int) -> str | None:
        """Asks whether to regenerate everything or only the gaps.

        Returns "all", "missing", or None (cancelled). A Yes/No pair cannot
        express this: "No" used to abort the whole batch, so a user with 9
        good profiles and 11 missing ones had no way to fill just the gaps
        without risking a pointless overwrite of the 9.
        """
        dialog = tk.Toplevel(self.root)
        dialog.title(t("profile_scope_title"))
        dialog.transient(self.root)
        dialog.resizable(False, False)

        ttk.Label(
            dialog,
            text=t("profile_scope_intro").format(
                saved=saved, total=saved + missing
            ),
            justify="left",
            wraplength=420,
        ).pack(padx=16, pady=(14, 8))

        buttons = ttk.Frame(dialog)
        buttons.pack(padx=16, pady=(0, 10), fill="x")

        # Cancel first, so it keeps focus and Enter stays the safe answer.
        ttk.Button(
            buttons, text=t("btn_cancel"),
            command=dialog.destroy,
        ).pack(side="right", padx=(6, 0))

        def choose(value):
            dialog.result = value
            dialog.destroy()

        ttk.Button(
            buttons,
            text=t("btn_profiles_all").format(n=saved + missing),
            command=lambda: choose("all"),
        ).pack(side="right", padx=(6, 0))
        ttk.Button(
            buttons,
            text=t("btn_profiles_missing").format(n=missing),
            command=lambda: choose("missing"),
        ).pack(side="right")

        ttk.Label(
            dialog,
            text=t("profile_scope_note"),
            justify="left",
            wraplength=420,
        ).pack(padx=16, pady=(0, 12))

        dialog.bind("<Escape>", lambda event: dialog.destroy())
        dialog.bind(
            "<Return>", lambda event: dialog.destroy()
        )   # default action: cancel, nothing is overwritten

        dialog.result = None
        dialog.update_idletasks()
        dialog.grab_set()
        dialog.wait_window()
        return getattr(dialog, "result", None)

    def start_batch_profile_generation(self):
        """
        Colectează speciile unice identificate din imaginile încărcate
        și generează profilul pentru fiecare (sărite pe cele deja
        existente). Simetric cu ANALYZE ALL.
        """

        if not self.image_paths:
            messagebox.showwarning(
                t("warn_no_images_title"),
                t("warn_add_image_first"),
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

        if existing_with_profile:
            # Trei răspunsuri, nu două: refacă tot, umple doar golurile,
            # sau renunță. „Nu" înseamnă altfel anulează tot lotul, deci
            # utilizatorul nu putea genera doar speciile lipsă.
            choice = self.ask_profile_scope(
                len(existing_with_profile), len(pending)
            )
            if choice == "all":
                pending = list(species_map.values())
            elif choice != "missing":
                return
            if not pending:
                return
        else:
            # A partial selection silently shrinks the lot, so say so.
            chosen = self._note_selection_scope()
            body = t("confirm_gen_profiles_all").format(
                len(pending), len(species_map)
            )
            if chosen < len(self.image_paths):
                body += "\n\n" + t("confirm_selection_used").format(
                    chosen=chosen, total=len(self.image_paths)
                )
            confirmed = messagebox.askyesno(
                t("title_gen_profiles_all"),
                body,
                parent=self.root,
            )
            if not confirmed:
                return

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
        self._begin_batch()
        self.set_status(
            t("status_generating_profiles").format(len(pending))
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
        stopped = False
        self.profile_retried = 0

        for index, info in enumerate(pending):

            # Same contract as ANALYZE ALL: finish what is in flight, never
            # start the next one after a stop.
            if self.batch_was_cancelled():
                stopped = True
                break

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

                response = self.generate_profile_response(full_prompt)

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

        self._end_batch()

        self.root.after(
            0,
            self.batch_profile_finished,
            total,
            completed,
            errors,
            batch_elapsed,
            stopped,
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
            t("status_generating_profile_n").format(current + 1, total, scientific_name)
        )

    def batch_profile_finished(
        self,
        total: int,
        completed: int,
        errors: list[str],
        elapsed: float,
        stopped: bool = False,
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
        self.batch_progress_frame.pack_forget()

        duration = self.format_seconds(elapsed)

        if stopped:
            # A stopped batch is not a failed one: say so plainly and keep
            # whatever finished, instead of reporting "0 errors" as if the
            # work had completed.
            self.set_status(
                t("status_stopped").format(
                    done=completed, total=total, duration=duration
                )
            )
            messagebox.showinfo(
                t("profiles_done_title"),
                t("batch_stopped").format(done=completed, total=total),
                parent=self.root,
            )
            return

        # Speciile salvate doar datorită rezervei arată că plafonul de bază
        # e prea mic pentru tipul de specie procesat.
        reserve_note = ""
        if getattr(self, "profile_retried", 0):
            reserve_note = "\n\n" + t("profiles_reserve").format(
                n=self.profile_retried,
                tokens=self.profile_num_predict_retry,
            )
        self.profile_retried = 0

        summary_head = (
            t("summary_completed").format(done=completed, total=total)
            + "\n\n"
            + t("summary_errors").format(n=len(errors))
            + "\n\n"
            + t("summary_duration").format(duration=duration)
            + "\n\n"
        )

        if errors:

            self.set_status(
                t("status_profiles_errors").format(completed, total, len(errors), duration)
            )

            error_text = "\n".join(errors)

            messagebox.showwarning(
                t("profiles_errors_title"),
                summary_head + error_text + reserve_note,
                parent=self.root,
            )

        else:

            self.set_status(
                t("status_profiles_completed").format(completed, total, duration)
            )

            messagebox.showinfo(
                t("profiles_done_title"),
                t("profiles_done").format(
                    total=total, duration=duration
                ) + reserve_note,
                parent=self.root,
            )

    def _collect_unique_species(self) -> dict[str, dict]:
        """
        Parcurge imaginile încărcate, citește catalog_items și returnează
        un dicționar indexat pe scientific_name (doar speciile
        identificate). Pentru fiecare indică dacă are deja profil.
        """

        species_map: dict[str, dict] = {}

        for path in (self.image_paths[i] for i in self._batch_indices()):

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

        # Every selected row goes, not just the first: removing one photo at
        # a time from a multi-selection would be a trap.
        indices = self.selected_indices()

        if not indices:
            return

        current = self.image_path
        for index in reversed(indices):
            path = self.image_paths[index]
            if path == current:
                current = None
            self.image_locations.pop(
                path,
                None
            )
            del self.image_paths[index]

        self.refresh_image_list()

        if self.image_paths:

            # Land on whatever took the place of the first removed row.
            new_index = min(
                indices[0],
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
                t("warn_no_images_title")
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
            t("status_images_cleared")
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
            t("status_step_error").format(current, total, filename)
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
            t("status_step_completed").format(current, total, filename)
        )


    def batch_finished(
        self,
        total,
        completed,
        errors,
        elapsed,
        image_names=None,
        stopped: bool = False,
    ):

        self.stop_status_timer()
        self.analysis_start_time = None
        self.batch_current = 0
        self.batch_total = 0
        self.batch_image_start_time = None
        self.batch_analyzed_count = 0
        self.batch_analyzed_time = 0.0
        self.batch_cached_current = False

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

        summary_head = (
            t("summary_completed").format(done=completed, total=total)
            + "\n\n"
            + t("summary_errors").format(n=len(errors))
            + "\n\n"
            + t("summary_duration").format(duration=duration)
            + "\n\n"
        )

        if stopped:
            # A stopped batch is neither a success nor a failure. Reporting
            # "Errors: 0" would claim the work finished, and the bar reading
            # "50 / 50" would claim all of it was done.
            self.batch_progress_label.configure(
                text=t("batch_progress").format(cur=completed, total=total)
            )
            self.set_status(
                t("status_stopped").format(
                    done=completed, total=total, duration=duration
                )
            )
            messagebox.showinfo(
                t("batch_done_title"),
                t("batch_stopped").format(done=completed, total=total),
                parent=self.root,
            )
            return

        if errors:

            self.set_status(
                t("status_batch_finished").format(completed, total, len(errors), duration)
            )

            error_text = "\n".join(
                errors
            )

            messagebox.showwarning(
                t("batch_errors_title"),
                summary_head + error_text,
                parent=self.root,
            )

        else:

            self.set_status(
                t("status_batch_completed").format(completed, total, duration)
            )

            messagebox.showinfo(
                t("batch_done_title"),
                t("batch_done").format(total=total, duration=duration),
            )

        # ----------------------------------------------------
        # Species missing from the taxonomy database
        # ----------------------------------------------------

        # ANALYZE ALL runs unattended, so the per-image CoL dialog never
        # appears. Instead, collect the whole lot and show one list the user
        # can walk through image by image. Validation opens the 900 MB
        # taxonomy cache, so it runs in the background and only the dialog
        # is scheduled back on the main thread.
        if image_names:
            self._report_batch_unverified_species(image_names)


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
                t("status_loaded").format(path.name)
            )

        except Exception as exc:

            self.image_path = None
            self.image_base64 = None

            messagebox.showerror(
                t("err_image"),
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
                t("warn_no_image_title"),
                t("info_select_image")
            )
            return

        prompt = self.get_prompt_text()

        if not prompt:
            messagebox.showwarning(
                t("warn_no_prompt_title"),
                t("info_enter_prompt")
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
            t("status_working")
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
                t("warn_no_images_title"),
                t("warn_add_image_first")
            )
            return

        prompt = self.get_prompt_text()

        if not prompt:
            messagebox.showwarning(
                t("warn_no_prompt_title"),
                t("info_enter_prompt")
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
        # Only the chosen subset takes part when the user made a selection.
        indices = self._batch_indices()
        paths = [self.image_paths[i] for i in indices]

        # Progress
        # ----------------------------------------------------

        total = len(paths)

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
        self._begin_batch()
        self.set_status(
            t("status_processing_batch").format(total)
        )
        self.update_status_with_time()

        # ----------------------------------------------------
        # Snapshot of the list and preset
        # ----------------------------------------------------

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
        stopped = False

        for index, path in enumerate(paths, start=0):

            # Read at the top of the loop: the image in flight always
            # finishes, the next one is never started.
            if self.batch_was_cancelled():
                stopped = True
                break

            self.image_status[path] = "processing"

            self.root.after(
                0,
                self.refresh_image_list
            )

            # Update status immediately before processing
            self.batch_current = index
            self.batch_filename = path.name
            # Timestamp the start of this image for per-image timing in the
            # status bar (time on current image + running average).
            self.batch_image_start_time = time.perf_counter()
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
                self.batch_cached_current = bool(cached)

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

                # Count per-image time only for real (non-cached) analyses;
                # cached items finish in ~0s and would skew the average/ETA.
                _item_elapsed = time.perf_counter() - self.batch_image_start_time
                if not self.batch_cached_current:
                    self.batch_analyzed_time += _item_elapsed
                    self.batch_analyzed_count += 1

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

        self._end_batch()

        self.root.after(
            0,
            self.batch_finished,
            total,
            completed,
            errors,
            batch_elapsed,
            list(paths),
            stopped,
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
                t("status_result_not_saved").format(exc)
            )

        # ----------------------------------------------------
        # Status
        # ----------------------------------------------------

        self.set_status(
            t("status_done").format(self.format_seconds(elapsed))
        )

        # ----------------------------------------------------
        # Completion popup — chained after the CoL check
        # (validate_species runs in a worker thread; the popup
        # appears after it so dialogs never stack out of order).
        # ----------------------------------------------------

        self._finish_analysis_chain(elapsed)

    def _finish_analysis_chain(self, elapsed: float) -> None:
        """CoL validation (background) → warning dialog → completion popup.

        ``validate_species`` reads the 900 MB taxonomy.db through per-call
        connections, so it is safe in a daemon thread; worst-case epitax
        scans (~5 s) never freeze the UI. The completion popup is shown
        only after validation returns, so the more informative CoL dialog
        (when the name is not verified) always appears first.
        """
        box: dict = {}

        def worker() -> None:
            try:
                box["v"] = self._validate_current_species()
            except Exception:
                box["v"] = None
            try:
                self.root.after(0, show)
            except Exception:
                pass  # window closed while validating

        def show() -> None:
            payload = box.get("v")
            if payload:
                try:
                    self._show_col_dialog(payload["result"], payload["name"])
                except Exception:
                    pass  # never block the completion popup on the dialog
            # Show the species actually stored in the catalog, not the file
            # name: after UPDATE SPECIES the two differ, and the file name
            # still carries the old binomial — which read as "the rename
            # did not stick" and triggered a second, pointless rename.
            species = ""
            if self.image_path is not None:
                try:
                    item = get_catalog_item(self.image_path, OUTPUT_DIR) or {}
                    species = (item.get("scientific_name") or "").strip()
                except Exception:
                    species = ""
            label = species or (self.image_path.name if self.image_path else "—")
            messagebox.showinfo(
                t("info_analysis_done_title"),
                t("info_analysis_done_body").format(label, self.format_seconds(elapsed))
            )

        threading.Thread(target=worker, daemon=True).start()

    def _validate_current_species(self):
        """Read-only CoL check for the current catalog item, or None.

        Returns ``{"result": ..., "name": str}`` when the identified name
        is NOT verified in Catalogue of Life, and None otherwise (no item,
        unknown name, or exact match — nothing worth interrupting for).
        """
        if self.image_path is None:
            return None
        item = get_catalog_item(self.image_path, OUTPUT_DIR) or {}
        name = (item.get("scientific_name") or "").strip()
        category = (item.get("category") or "").strip()
        if not name or name in ("Necunoscut", "Unknown"):
            return None
        result = validate_species(name, category)
        if result.get("status") == "exact":
            return None
        return {"result": result, "name": name}

    def _format_col_message(self, name: str, result: dict):
        """Graduated CoL message → ``(text, accepted_name_or_None)``.

        Tier structure:
        - ``genus_ok``:   genus members + closest-in-genus + accepted hit;
        - ``not_in_col``: accepted species/genus, else closest fallback;
        - anything else:  honest "no plausible match" note.
        """
        lines = [t("col_intro") + name, ""]
        status = result.get("status")
        accepted = result.get("accepted")

        if status == "genus_ok":
            members = result.get("members") or []
            if members:
                lines.append(t("col_genus_list").format(
                    genus=result.get("genus") or "?",
                    n=result.get("member_count") or len(members),
                    shown=len(members),
                    names=", ".join(members),
                ))
            if result.get("closest_in_genus"):
                lines.append(t("col_closest_genus").format(
                    name=result["closest_in_genus"]
                ))

        if accepted:
            acc_genus = accepted["name"].split()[0]
            lines.append(t("col_accepted").format(
                name=accepted["name"],
                genus=acc_genus,
                family=accepted.get("family") or "—",
            ))
            if result.get("genus") and acc_genus != result["genus"]:
                lines.append(t("col_genus_changed").format(
                    old=result["genus"],
                    new=acc_genus,
                    epithet=result.get("epithet") or "",
                ))
        else:
            if (result.get("accepted_genus") and result.get("genus")
                    and result["accepted_genus"] != result["genus"]):
                lines.append(t("col_genus_changed").format(
                    old=result["genus"],
                    new=result["accepted_genus"],
                    epithet=result.get("epithet") or "",
                ))
            closest = result.get("closest") or []
            if closest:
                lines.append(t("col_k_closest").format(
                    name=closest[0]["name"]
                ))
            else:
                lines.append(t("col_no_hint"))

        return "\n".join(lines), (accepted["name"] if accepted else None)

    def _show_unverified_species_dialog(self, report: dict) -> None:
        """Single list of the batch species missing from the database.

        ANALYZE ALL is unattended, so this replaces the per-image CoL dialog
        that would otherwise never appear. Each row keeps every image the
        model named that species on.

        The window is deliberately NON-modal: the user has to reach
        "Update species" in the main window for every photo, so the list only
        follows along, selecting the current image and stepping to the next
        one with the "Next" button.
        """
        if not report:
            return

        # One list at a time; a second ANALYZE ALL replaces the old one.
        previous = getattr(self, "_unverified_dialog", None)
        if previous is not None and previous.winfo_exists():
            previous.destroy()
            self._unverified_dialog = None

        dlg = tk.Toplevel(self.root)
        dlg.title(t("col_batch_title"))
        dlg.transient(self.root)
        dlg.geometry("760x460")
        self._unverified_dialog = dlg

        pending: dict = {
            name: list((entry.get("images") or []))
            for name, entry in report.items()
        }
        # Flat, stable order of every image still to be checked, plus the
        # reverse lookup used to highlight the species row in the list.
        queue: list = []
        owner: dict = {}
        for name in sorted(pending, key=str.casefold):
            for image_name in pending[name]:
                queue.append(image_name)
                owner[image_name] = name
        queue_index = 0
        current_image = ""

        ttk.Label(
            dlg,
            text=t("col_batch_intro").format(n=len(report)),
            justify="left",
            wraplength=720,
        ).pack(fill="x", padx=12, pady=(12, 8))

        tree = ttk.Treeview(
            dlg,
            columns=("species", "count", "files"),
            show="headings",
            height=min(len(report), 12),
        )
        tree.heading("species", text=t("col_batch_species"))
        tree.heading("count", text="#")
        tree.heading("files", text=t("col_batch_images"))
        tree.column("species", width=220, anchor="w")
        tree.column("count", width=48, anchor="center", stretch=False)
        tree.column("files", width=440, anchor="w")
        tree.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        def refresh_tree() -> None:
            tree.delete(*tree.get_children())
            for name in sorted(pending, key=str.casefold):
                files = pending[name]
                tree.insert(
                    "", tk.END, iid=name,
                    values=(name, len(files), ", ".join(files)),
                )

        refresh_tree()

        def select_image(image_name: str) -> bool:
            for index, candidate in enumerate(self.image_paths):
                if candidate.name == image_name:
                    self.select_image(index)
                    return True
            return False

        def step(delta: int) -> None:
            """Select the next/previous image still to be checked."""
            nonlocal queue_index
            if not queue:
                return
            queue_index = max(0, min(len(queue) - 1, queue_index + delta))
            show(queue[queue_index])

        def show(image_name: str) -> None:
            """Select *image_name* in the main list and highlight its row."""
            nonlocal current_image
            if not select_image(image_name):
                return
            current_image = image_name
            row = owner.get(image_name)
            if row and tree.exists(row):
                tree.selection_set(row)
                tree.see(row)

        def mark_checked() -> None:
            """Drop the current species: the user handled it by hand."""
            nonlocal queue_index
            selection = tree.selection()
            if not selection:
                return
            name = selection[0]
            if name not in pending:
                return
            pending.pop(name, None)
            # Rebuild the walk order so "Next" never stops on a done species.
            for key in [key for key, value in owner.items() if value == name]:
                owner.pop(key, None)
                if key in queue:
                    queue.remove(key)
            # Re-anchor on the image still on screen: removing items from
            # ``queue`` would otherwise shift it and "Next" would skip one.
            if current_image in queue:
                queue_index = queue.index(current_image)
            else:
                queue_index = max(0, min(len(queue) - 1, queue_index))
            refresh_tree()
            if not pending:
                dlg.destroy()

        buttons = ttk.Frame(dlg)
        buttons.pack(fill="x", padx=12, pady=(0, 12))

        ttk.Button(
            buttons, text=t("col_batch_next"),
            command=lambda: step(1),
        ).pack(side="left")
        ttk.Button(
            buttons, text=t("col_batch_done"), command=mark_checked
        ).pack(side="left", padx=(6, 0))
        ttk.Button(
            buttons, text=t("col_batch_close"),
            command=lambda: dlg.destroy(),
        ).pack(side="right")

        def on_row(_event=None) -> None:
            """Double-click / Enter jumps to the first image of that species."""
            selection = tree.selection()
            if not selection:
                return
            for image_name in pending.get(selection[0]) or []:
                if select_image(image_name):
                    break

        tree.bind("<Double-1>", on_row)
        tree.bind("<Return>", on_row)

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        step(0)

    def _report_batch_unverified_species(self, image_names) -> None:
        """Validate the whole batch in the background, then show the list."""
        if not image_names:
            return

        def worker() -> None:
            try:
                report = collect_batch_unverified_species(image_names)
            except Exception:
                report = {}
            try:
                self.root.after(0, self._show_unverified_species_dialog, report)
            except Exception:
                pass  # window closed while validating

        threading.Thread(target=worker, daemon=True).start()

    def _show_col_dialog(self, result: dict, name: str) -> None:
        """Notice for a non-verified name + optional pre-filled Update form."""
        text, accepted = self._format_col_message(name, result)

        dlg = tk.Toplevel(self.root)
        dlg.title(t("col_title"))
        dlg.transient(self.root)
        dlg.geometry("660x440")

        body = tk.Text(dlg, wrap="word", relief="flat", height=18)
        body.insert("1.0", text)
        body.configure(state="disabled")
        body.pack(fill="both", expand=True, padx=12, pady=(12, 8))

        buttons = ttk.Frame(dlg)
        buttons.pack(fill="x", padx=12, pady=(0, 12))

        # Verificare web a numelui respins — disponibilă IMEDIAT, fără
        # să închizi dialogul (fereastra principală e blocată de grab_set).
        ttk.Button(
            buttons, text=t("btn_open_col"),
            command=lambda: self._browse_col_for(name)
        ).pack(side="left")
        ttk.Button(
            buttons, text=t("btn_open_wikipedia"),
            command=lambda: self._browse_wikipedia_for(name)
        ).pack(side="left", padx=(6, 0))
        ttk.Button(
            buttons, text=t("btn_open_inaturalist"),
            command=lambda: self._browse_inaturalist_for(name)
        ).pack(side="left", padx=(6, 0))

        if accepted:
            def on_update() -> None:
                dlg.destroy()
                if self.image_path is not None:
                    self.update_species_classification(
                        prefill={"scientific_name": accepted}
                    )

            ttk.Button(
                buttons, text=t("col_btn_update"), command=on_update
            ).pack(side="right")

        ttk.Button(
            buttons, text=t("col_btn_later"), command=dlg.destroy
        ).pack(side="right", padx=(0, 8))

        dlg.grab_set()
        dlg.wait_window()

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
            t("title_error")
        )

        messagebox.showerror(
            t("err_analysis_failed"),
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
            messagebox.showinfo(t("title_catalog"), t("info_image_not_cataloged"))
            return
        folder = Path(item["catalog_path"]).parent
        if not folder.is_dir():
            messagebox.showwarning(t("title_catalog"), t("warn_catalog_folder_gone"))
            return
        os.startfile(str(folder))

    def _web_search_name(self) -> str:
        """Numele din căutarea web (câmp editabil), cu fallback pe catalog."""
        name = ""
        try:
            name = self.web_entry.get().strip()
        except Exception:
            pass
        if not name and self.image_path is not None:
            try:
                item = get_catalog_item(self.image_path, OUTPUT_DIR) or {}
                name = (item.get("scientific_name") or "").strip()
            except Exception:
                name = ""
        return name

    def _browse_col_for(self, name: str) -> None:
        """Deschide pagina de căutare CoL cu ``name`` ca query direct în URL.

        Site-ul CoL acceptă ``?q=`` în URL, deci căutarea pornește direct —
        nu mai e nevoie de lipire manuală din clipboard. Numele se copiază
        totuși și în clipboard, pentru ajustare manuală dacă e nevoie.
        ``name`` e numele de căutat; ignorat dacă e gol.
        """
        name = (name or "").strip()
        if name:
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(name)
                self.root.update()  # menține clipboard-ul după închidere
            except Exception:
                pass
            self.set_status(t("web_copied_clipboard").format(name=name))
        webbrowser.open(COL_SEARCH_URL + quote(name))

    def _browse_wikipedia_for(self, name: str) -> None:
        """Deschide articolul Wikipedia; redirecționează sinonimele."""
        name = (name or "").strip()
        if not name or name in ("Necunoscut", "Unknown"):
            messagebox.showinfo(t("web_title"), t("web_no_species"),
                                parent=self.root)
            return
        # Speciile incerte („Lycogala sp.”) nu au articol propriu —
        # scoatem „sp.” din link și deschidem direct pagina genului.
        url_name = name
        parts = url_name.split()
        if (len(parts) > 1
                and parts[-1].lower().rstrip(".") in ("sp", "spp")):
            url_name = " ".join(parts[:-1])
        webbrowser.open(WIKIPEDIA_SPECIES_URL + quote(url_name.replace(" ", "_")))

    def open_col_page(self):
        """Căutare CoL pentru textul din câmpul de căutare web."""
        self._browse_col_for(self._web_search_name())

    def open_wikipedia_page(self):
        """Wikipedia pentru textul din câmpul de căutare web."""
        self._browse_wikipedia_for(self._web_search_name())

    def _browse_inaturalist_for(self, name: str) -> None:
        """Deschide pagina de taxoni iNaturalist pentru ``name``.

        Spre deosebire de CoL, iNaturalist acceptă query direct în URL, deci
        numele e pus și în link. Îl copiem totuși și în clipboard — ca la CoL —
        pentru lipire manuală, dacă utilizatorul vrea să ajusteze căutarea.
        ``name`` e numele de căutat; ignorat dacă e gol.
        """
        name = (name or "").strip()
        if not name or name in ("Necunoscut", "Unknown"):
            messagebox.showinfo(t("web_title"), t("web_no_species"),
                                parent=self.root)
            return
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(name)
            self.root.update()  # menține clipboard-ul după închidere
        except Exception:
            pass
        webbrowser.open(INATURALIST_SEARCH_URL + quote(name) + "&view=list")

    def open_inaturalist_page(self):
        """iNaturalist pentru textul din câmpul de căutare web."""
        self._browse_inaturalist_for(self._web_search_name())

    def _open_main_preview(self, event=None):
        """Click pe miniatura curentă (panoul Images): poza mare, popup.

        Reutilizează _open_photo_preview — aceeași fereastră ca în Quick
        Catalog (poză la ~80% ecran, text copiabil, butoane web).
        """
        if self.image_path is None:
            return
        # details=False: doar poza mare, fără info și fără butoane web —
        # în fereastra principală oricum există Photo Info (dreapta) și
        # butoanele de căutare.
        self._open_photo_preview(self.image_path, parent=self.root,
                                 details=False)

    def _open_photo_preview(self, image_path, name=None, parent=None,
                            details=True):
        """Previzualizare mare a unei poze (+ detalii opționale).

        - poza e scalată să încapă fără să depășească ~80% din ecran;
        - details=True (Quick Catalog): text selectabil/copiabil cu specia,
          folderul, calea + butoanele @CoL / @WIKIPEDIA / @iNAT;
        - details=False (click pe poza din panoul Images): DOAR poza mare,
          fără informații și fără butoane web (alea există în fereastra
          principală); Escape sau OK închide.
        """
        from core.catalog import get_catalog_item, _category_folder

        path = Path(image_path)
        item = get_catalog_item(str(path), OUTPUT_DIR) or {}

        sci = (item.get("scientific_name") or "").strip()
        query = (name or "").strip() or sci or path.stem
        ro_name = item.get("ro_name", "") or ""
        en_name = item.get("en_name", "") or ""
        category = _category_folder(item.get("category", "")) if item else ""
        folder = (Path(item["catalog_path"]).parent
                  if item.get("catalog_path") else path.parent)

        prev = tk.Toplevel(parent or self.root)
        prev.title(t("qc_preview_title"))
        prev.transient(parent or self.root)

        # Fereastra e plafonată la ecran, iar conținutul (poză + text +
        # butoane) stă într-un canvas scrollabil, deci butoanele de jos
        # rămân accesibile cu rotița chiar și la poze foarte mari.
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        # Poza se încarcă întâi, ca fereastra să se încadreze pe lățimea
        # ei (landscape), nu pe tot ecranul — rămâne loc dedesubt pt. text.
        max_w = int(screen_w * 0.55)
        max_h = int(screen_h * 0.62)
        photo = None
        try:
            with Image.open(path) as im:
                im.thumbnail((max_w, max_h))
                photo = ImageTk.PhotoImage(im)
        except Exception:
            photo = None
        img_w = photo.width() if photo else 480
        img_h = photo.height() if photo else 360
        win_w = min(img_w + 28, int(screen_w * 0.6))
        # details=False: doar poza, deci nu mai e nevoie de cele ~270px
        # rezervate textului/butoanelor de sub ea.
        win_h = min(img_h + (270 if details else 64), int(screen_h * 0.86))
        prev.geometry(f"{win_w}x{win_h}")
        prev.minsize(420, 360)

        pcanvas = tk.Canvas(prev, highlightthickness=0)
        pvscroll = ttk.Scrollbar(prev, orient="vertical", command=pcanvas.yview)
        scrollable = tk.Frame(pcanvas)
        scrollable.bind("<Configure>",
                        lambda e: pcanvas.configure(scrollregion=pcanvas.bbox("all")))
        canvas_window = pcanvas.create_window((0, 0), window=scrollable, anchor="nw")
        pcanvas.configure(yscrollcommand=pvscroll.set)
        pcanvas.pack(side="left", fill="both", expand=True)
        pvscroll.pack(side="right", fill="y")

        # Lățimea conținutului urmează lățimea ferestrei.
        pcanvas.bind("<Configure>",
                     lambda e: pcanvas.itemconfigure(canvas_window, width=e.width))

        # Scroll cu rotița mouse-ului (Windows + Linux).
        def _on_prev_wheel(event):
            if getattr(event, "num", None) == 4:
                pcanvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                pcanvas.yview_scroll(1, "units")
            else:
                pcanvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

        prev.bind("<MouseWheel>", _on_prev_wheel, add="+")
        prev.bind("<Button-4>", _on_prev_wheel, add="+")
        prev.bind("<Button-5>", _on_prev_wheel, add="+")

        # Pozarea pozei ( deja scalată mai sus ).
        if photo is not None:
            img_label = tk.Label(scrollable, image=photo, bg="#222")
            img_label.image = photo  # referință vie (altfel GC o șterge)
            img_label.pack(padx=8, pady=(8, 4))
        else:
            ttk.Label(scrollable, text=t("qc_preview_missing"),
                      padding=8).pack(padx=8, pady=8)

        # Informații — text selectabil/copiabil, nu editabil.
        # Doar în modul cu detalii (Quick Catalog); din panoul Images
        # utilizatorul a cerut strict poza mare, fără text sub ea.
        if details:
            info_lines = [f"{t('lbl_species')} {query}"]
            if sci and sci != query:
                info_lines.append(f"Științific: {sci}")
            if ro_name or en_name:
                info_lines.append(f"Nume: {ro_name} / {en_name}".rstrip(" /"))
            if category:
                info_lines.append(f"{t('lbl_category')} {category}")
            info_lines.append(f"{t('lbl_folder')} {folder}")
            info_lines.append(f"Fișier: {path}")

            info_text = tk.Text(scrollable, height=len(info_lines) + 1,
                                wrap="word",
                                font=("Segoe UI", 10), relief="flat",
                                background=prev.cget("background"))
            info_text.pack(fill="x", padx=8, pady=4)
            info_text.insert("1.0", "\n".join(info_lines))
            info_text.configure(state="disabled")  # rămâne selectabil + copiabil
            ttk.Label(scrollable, text=t("qc_copy_hint"),
                      foreground="#888").pack(fill="x", padx=8)

            # Căutare pe internet după specie.
            web_bar = ttk.Frame(scrollable)
            web_bar.pack(fill="x", padx=8, pady=(6, 4))
            ttk.Label(web_bar, text=t("qc_web_search")).pack(side="left")
            ttk.Button(web_bar, text=t("btn_open_col"),
                       command=lambda: self._browse_col_for(query)).pack(
                side="left", padx=(6, 0))
            ttk.Button(web_bar, text=t("btn_open_wikipedia"),
                       command=lambda: self._browse_wikipedia_for(query)).pack(
                side="left", padx=(6, 0))
            ttk.Button(web_bar, text=t("btn_open_inaturalist"),
                       command=lambda: self._browse_inaturalist_for(query)).pack(
                side="left", padx=(6, 0))

        # Bara de jos: (opțional) copiază calea + deschide folder + închide.
        btn_bar = ttk.Frame(scrollable)
        btn_bar.pack(fill="x", padx=8, pady=(0, 8))

        if details:
            def _copy_path():
                prev.clipboard_clear()
                prev.clipboard_append(str(path))

            ttk.Button(btn_bar, text=t("qc_copy_path"),
                       command=_copy_path).pack(side="left")
            ttk.Button(btn_bar, text=t("btn_open_folder"), command=lambda: (
                os.startfile(str(folder)) if folder.exists() else None
            )).pack(side="left", padx=(6, 0))
        ttk.Button(btn_bar, text=t("btn_ok"),
                   command=prev.destroy).pack(side="right")

        # Escape închide în ambele variante (la fel ca și click pe OK).
        prev.bind("<Escape>", lambda _e: prev.destroy())

    def open_quick_catalog_dialog(self):
        """Dialog pentru catalogare rapidă: alege imagini + specie din taxonomy.db."""
        from core.catalog import quick_catalog_images, get_catalog_item
        from core.taxonomy import search_taxonomy_names

        if not self.image_paths:
            messagebox.showwarning(t("qc_title"), t("qc_no_images"), parent=self.root)
            return

        win = tk.Toplevel(self.root)
        win.title("Catalog rapid")
        win.geometry("820x640")
        win.minsize(520, 560)
        win.transient(self.root)
        win.grab_set()

        # Container principal: conținutul FIX, fără scroll global, ca să
        # nu dispară controalele (specie, categorie, butoane) când derulezi
        # pozele. Doar grid-ul de poze are scroll propriu, mai jos.
        scrollable = tk.Frame(win)
        scrollable.pack(fill="both", expand=True)

        _wrap_widgets = []

        def _on_win_resize(event):
            if event.widget is not win:
                return
            wrap = max(200, event.width - 40)
            for w in _wrap_widgets:
                try:
                    w.configure(wraplength=wrap)
                except tk.TclError:
                    pass

        win.bind("<Configure>", _on_win_resize)

        # --- Imaginea țintă (cea selectată în lista principală) ---
        # Default: catalogăm DOAR imaginea selectată, fără grid de poze.
        target_path = self.image_path
        if target_path is None and self.image_paths:
            target_path = self.image_paths[0]
        target_box = {"path": target_path}

        # Lista de lucru a DIALOGULUI: o copie a listei principale. "📥 Încarcă
        # imagini" și "📂 Folder întreg" modifică DOAR această listă, nu și
        # lista principală (self.image_paths) — altfel, după închiderea
        # dialogului, panoul Images rămânea cu pozele altui folder.
        qc_paths = list(self.image_paths)

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
        load_pics_btn = ttk.Button(pics_header, text=t("qc_load_images"))
        load_pics_btn.pack(side="right", padx=(0, 6))
        load_folder_btn = ttk.Button(pics_header, text=t("qc_load_folder"))
        load_folder_btn.pack(side="right", padx=(0, 6))
        toggle_pics_btn = ttk.Button(pics_header, text=t("qc_show_pics"))
        toggle_pics_btn.pack(side="right")
        pics_header.pack(fill="x", padx=10, pady=(6, 2))

        pics_box = tk.Frame(scrollable, bg="#fafafa")
        # NU pack la început — apare doar la cerere.
        # Grid-ul stă într-un canvas cu scroll propriu: pozele se derulează
        # independent, iar controalele de sus rămân pe loc.
        pics_canvas = tk.Canvas(pics_box, bg="#fafafa", highlightthickness=0)
        pics_vscroll = ttk.Scrollbar(pics_box, orient="vertical",
                                     command=pics_canvas.yview)
        pics_canvas.configure(yscrollcommand=pics_vscroll.set)
        pics_vscroll.pack(side="right", fill="y")
        pics_canvas.pack(side="left", fill="both", expand=True)

        def _on_pics_wheel(event):
            if getattr(event, "num", None) == 4:
                pics_canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                pics_canvas.yview_scroll(1, "units")
            else:
                pics_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

        for w in (pics_canvas, pics_box):
            w.bind("<MouseWheel>", _on_pics_wheel)
            w.bind("<Button-4>", _on_pics_wheel)
            w.bind("<Button-5>", _on_pics_wheel)

        THUMB = 80
        # Geometria unei celule, în pixeli. CELL_TOTAL = pantă reală pe
        # coloană (imagine + padx label + chenar + padx grid). Folosim o
        # lățime FIXĂ pe celulă (pack_propagate(False)) și coloane cu
        # weight=1, ca numele lung să nu lărgească celula și să nu mai
        # fie tăiată ultima poză de pe rând.
        GRID_PADX = 8      # padx=4 în grid × 2 părți
        CELL_BORDER = 2    # bd=1 × 2 părți
        IMG_PADX = 4       # padx=2 în jurul imagini × 2 părți
        CELL_W = THUMB + IMG_PADX + CELL_BORDER          # lățime celulă
        CELL_TOTAL = CELL_W + GRID_PADX                  # pantă per coloană
        CELL_H = THUMB + 24                              # imagine + legendă
        img_vars = []
        photo_refs = []
        all_cells = []  # refereță la toate celulele pentru reflow
        pics_state = {"built": False, "visible": False}
        grid_frame = None
        grid_window = None

        def _build_grid():
            """Construiește grid-ul de miniaturi (o singură dată, la cerere)."""
            nonlocal grid_frame, grid_window
            # Resetare pentru reconstrucție (la încărcarea unor poze noi).
            img_vars.clear()
            del all_cells[:]
            photo_refs.clear()
            if grid_frame is not None:
                grid_frame.destroy()
                grid_frame = None
            if grid_window is not None:
                pics_canvas.delete(grid_window)
                grid_window = None
            grid_frame = tk.Frame(pics_canvas, bg="#fafafa")
            grid_window = pics_canvas.create_window((0, 0), window=grid_frame,
                                                    anchor="nw")

            def _on_grid_configure(e):
                pics_canvas.configure(scrollregion=pics_canvas.bbox("all"))
                _reflow()

            grid_frame.bind("<Configure>", _on_grid_configure)
            pics_canvas.bind("<Configure>",
                             lambda e: pics_canvas.itemconfigure(grid_window,
                                                                 width=e.width))

            for idx, p in enumerate(qc_paths):
                already = get_catalog_item(p, OUTPUT_DIR) is not None
                var = tk.BooleanVar(value=False)
                img_vars.append((p, var))

                cell = tk.Frame(grid_frame, bd=1, relief="solid",
                                bg="#ddd" if already else "#fff",
                                width=CELL_W, height=CELL_H)
                # Lățime/înălțime FIXĂ: celula nu se mai mărește după numele
                # lung al fișierului, deci coloanele rămân egale cu CELL_TOTAL.
                cell.pack_propagate(False)
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
                        # Poză deja catalogată: deschide previzualizarea mare
                        # (în locul popup-ului mic cu doar text).
                        self._open_photo_preview(path, parent=win)
                        return
                    v.set(not v.get())
                    c.configure(bd=3 if v.get() else 1,
                                bg="#90caf9" if v.get() else "#ffffff")

                def _preview(event, path=p):
                    # Dublu-click sau click dreapta: previzualizare mare,
                    # indiferent dacă poza e catalogată sau nu.
                    self._open_photo_preview(path, parent=win)

                img_label.bind("<Button-1>", _toggle)
                img_label.bind("<Double-Button-1>", _preview)
                img_label.bind("<Button-3>", _preview)

                name_short = p.stem[:14] + ("…" if len(p.stem) > 14 else "")
                caption = name_short + (" ✓" if already else "")
                cap = tk.Label(cell, text=caption, font=("", 8),
                               bg="#ddd" if already else "#fff",
                               fg="#888" if already else "#333",
                               anchor="w")
                cap.pack(fill="x", padx=2, pady=(0, 2))
                cap.bind("<Button-1>", _toggle)
                cap.bind("<Double-Button-1>", _preview)
                cap.bind("<Button-3>", _preview)

                # Rotița mouse-ului peste o miniatură derulează zona de poze.
                for widget in (img_label, cap, cell):
                    widget.bind("<MouseWheel>", _on_pics_wheel)
                    widget.bind("<Button-4>", _on_pics_wheel)
                    widget.bind("<Button-5>", _on_pics_wheel)

        # Reflow la redimensionare (număr de coloane după lățimea reală)
        def _reflow(event=None):
            if grid_frame is None:
                return
            width = pics_canvas.winfo_width()
            # Împărțim pe pantă (CELL_TOTAL), nu pe lățimea simplă a celulei,
            # altfel se calculau prea multe coloane și ultima se tăia.
            cols = max(1, width // CELL_TOTAL)
            for col in range(cols):
                # weight=1: coloanele împart egal lățimea rămasă; uniform
                # forțează toate egale, deci nimic nu iese în afara rândului.
                grid_frame.columnconfigure(col, weight=1, uniform="qc")
            for i, (c, v, cat) in enumerate(all_cells):
                c.grid(row=i // cols, column=i % cols, padx=4, pady=4, sticky="nsew")

        def _toggle_pics():
            if not pics_state["built"]:
                _build_grid()
                pics_state["built"] = True
            if pics_state["visible"]:
                pics_box.pack_forget()
                toggle_pics_btn.configure(text=t("qc_show_pics"))
                pics_state["visible"] = False
            else:
                pics_box.pack(fill="both", expand=True, padx=0, pady=(2, 6))
                toggle_pics_btn.configure(text=t("qc_hide_pics"))
                pics_state["visible"] = True

        def _rebuild_grid():
            """Construiește/reconstruiește grid-ul și îl face vizibil."""
            _build_grid()
            pics_state["built"] = True
            _reflow()

        def _load_pics():
            """Încarcă imagini din orice folder și le adaugă în grid.

            Util când anumite poze din folderul curent s-au dovedit a fi alt
            organism și vor mutate în altă categorie (ex. ciuperci →
            protozoare): alegi pozele din alt folder fără a reîncărca analiza.
            """
            from tkinter import filedialog as _fd
            chosen = _fd.askopenfilenames(
                parent=win,
                title=t("qc_load_images_title"),
                initialdir=str(Path(self.image_paths[0]).parent)
                if self.image_paths else str(APP_ROOT / "input"),
                filetypes=[
                    ("Image files", "*.jpg *.jpeg *.png *.webp *.bmp *.tif *.tiff"),
                    ("All files", "*.*"),
                ],
            )
            if not chosen:
                return
            added = 0
            for raw in chosen:
                p = Path(raw)
                if p not in qc_paths:
                    qc_paths.append(p)
                    added += 1
            # Reîncarcă grid-ul ca să apară și noile poze.
            # Lista principală (self.image_paths) NU se atinge: dialogul are
            # propria listă de lucru (qc_paths).
            _rebuild_grid()
            if not pics_state["visible"]:
                pics_box.pack(fill="both", expand=True, padx=0, pady=(2, 6))
                toggle_pics_btn.configure(text=t("qc_hide_pics"))
                pics_state["visible"] = True

        def _load_folder():
            """Înlocuiește pozele din dialog cu un folder întreg.

            Modifică DOAR lista dialogului (qc_paths), nu lista principală
            de lucru — panoul Images rămâne neatins după închidere.
            Practic când organizezi pe categorii în foldere separate (ex.
            mushrooms/ vs protozoare/): alegi folderul nou și vezi pozele
            lui în grid.
            """
            from tkinter import filedialog as _fd
            start = (str(Path(self.image_paths[0]).parent)
                     if self.image_paths else str(APP_ROOT / "input"))
            folder = _fd.askdirectory(
                parent=win,
                title=t("qc_load_folder_title"),
                initialdir=start,
            )
            if not folder:
                return
            folder_path = Path(folder)
            extensions = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
            new_paths = sorted(
                p for p in folder_path.iterdir()
                if p.is_file() and p.suffix.lower() in extensions
            )
            if not new_paths:
                messagebox.showwarning(t("qc_title"), t("qc_folder_empty"),
                                       parent=win)
                return
            # Replace doar în dialog (qc_paths) — lista principală de lucru
            # NU se atinge, iar imaginea-țintă a ferestrei principale
            # (self.image_path) rămâne ce avea.
            qc_paths[:] = new_paths
            # Resetează imaginea-țintă pe prima din noul folder.
            target_box["path"] = new_paths[0]
            info_top.configure(
                text=t("qc_selected_image").format(name=Path(new_paths[0]).name))
            # Reîncarcă grid-ul cu noile poze.
            _rebuild_grid()
            if not pics_state["visible"]:
                pics_box.pack(fill="both", expand=True, padx=0, pady=(2, 6))
                toggle_pics_btn.configure(text=t("qc_hide_pics"))
                pics_state["visible"] = True

        toggle_pics_btn.configure(command=_toggle_pics)
        load_pics_btn.configure(command=_load_pics)
        load_folder_btn.configure(command=_load_folder)

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
        _wrap_widgets.extend([info_top, validated_label])

        # Info specie
        info_frame = ttk.LabelFrame(scrollable, text=t("qc_info_frame"))
        info_frame.pack(fill="x", padx=10, pady=6)

        cat_var = tk.StringVar(value="—")
        ro_var = tk.StringVar(value="—")
        en_var = tk.StringVar(value="—")

        ttk.Label(info_frame, text=t("lbl_category")).grid(row=0, column=0, sticky="w", padx=8, pady=2)
        cat_combo = ttk.Combobox(info_frame, textvariable=cat_var,
                                 values=["Plant", "Mushroom", "Insect", "Arachnid", "Mollusc",
                                         "Fish", "Reptile", "Amphibian", "Bird", "Mammal",
                                         "Animal", "Protozoa", "Other"],
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
            "plantae": "Plant",
            "animalia": "Animal",
            "fungi": "Mushroom",
            "protozoa": "Protozoa",
        }

        def _kingdom_to_category(kingdom: str) -> str:
            k = (kingdom or "").strip().casefold()
            for key, val in _KINGDOM_MAP.items():
                if key in k:
                    return val
            # fallback: detectează din pasare/bird
            if "pasar" in k or "bird" in k:
                return "Bird"
            return "Other"

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

            try:
                result = quick_catalog_images(selected, scientific, category, OUTPUT_DIR)
                copied = result["copied"]
                errors = result["errors"]
                from core.catalog import _category_folder as _qc_folder
                dest_folder = OUTPUT_DIR / _qc_folder(category) / scientific.replace(" ", "_").replace("×", "x").lower()
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
                        # Scoate din lista dialogului.
                        while p in qc_paths:
                            qc_paths.remove(p)
                        # Și din lista principală, DOAR dacă era acolo
                        # (pozele încărcate din alte foldere nu o ating).
                        try:
                            if p in self.image_paths:
                                self.image_paths.remove(p)
                        except ValueError:
                            pass
                        self.image_status.pop(p, None)
                        self.image_locations.pop(p, None)
                    if target_box["path"] in selected:
                        target_box["path"] = (
                            qc_paths[0] if qc_paths else None
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
        for section in ("start", "catalog", "archive", "books", "data"):
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
        # Textele pot avea lățimi diferite în engleză -> recalculează rândurile.
        self.button_bar.relayout(force=True)
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
            messagebox.showerror(t("title_admin"), t("err_no_zip"), parent=self.root)
            return
        if not messagebox.askyesno(t("title_admin"),
            t("confirm_rebuild_taxonomy").format(zip_path.name),
            parent=self.root):
            return
        self.admin_set_status("Reconstruiesc taxonomy.db...")
        def worker():
            try:
                stats = build_cache(zip_path)
                def _done():
                    text = f"Gata: {stats['accepted']:,} taxoni încărcați în taxonomy.db."
                    self.admin_set_status(text)
                    messagebox.showinfo(t("title_admin"), text, parent=self.root)
                self.root.after(0, _done)
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror(t("title_admin"), str(exc), parent=self.root))
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
            messagebox.showerror(t("title_admin"), t("err_no_zip"), parent=self.root)
            return
        if not messagebox.askyesno(t("title_admin"),
            t("confirm_import_vernacular"),
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
                        t("title_admin"),
                        t("info_import_col_done").format(stats['updated'], stats['en_count'], stats['ro_count']),
                        parent=self.root)))
            except Exception as exc:
                self.root.after(0, lambda: messagebox.showerror(
                    t("title_admin"), t("err_import_col").format(exc), parent=self.root))
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
            messagebox.showerror(t("title_admin"), t("err_cannot_prepare_rules").format(exc),
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
                messagebox.showerror(t("err_parse_rules_title"), t("err_json_invalid").format(exc),
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
                t("title_admin"),
                t("err_cannot_open_folder").format(COL_DIR, exc),
                parent=self.root,
            )

    def admin_purge_archived(self):
        """Șterge DEFINITIV pozele arhivate mai vechi de N zile (purge).

        Operațiune distructivă și ireversibilă: cere întâi fereastra de timp,
        apoi arată exact câte poze s-ar șterge și confirmarea explicita.
        Fișierele de pe disc și rândurile din catalog dispar împreună.
        """
        answer = simpledialog.askinteger(
            t("purge_title"),
            t("purge_ask_days"),
            parent=self.root,
            initialvalue=30,
            minvalue=1,
            maxvalue=3650,
        )
        if not answer:
            return

        try:
            victims = count_archived_older_than(OUTPUT_DIR, answer)
        except Exception as exc:
            messagebox.showerror(
                t("title_admin"), str(exc), parent=self.root
            )
            return

        if not victims:
            messagebox.showinfo(
                t("purge_title"),
                t("purge_none").format(answer),
                parent=self.root,
            )
            return

        count = len(victims)
        confirm = messagebox.askyesno(
            t("purge_title"),
            t("purge_confirm").format(count, answer),
            parent=self.root,
        )
        if not confirm:
            return

        try:
            deleted = purge_archived(OUTPUT_DIR, answer)
        except Exception as exc:
            messagebox.showerror(
                t("title_admin"), str(exc), parent=self.root
            )
            return

        self.set_status(t("purge_done").format(deleted))
        messagebox.showinfo(
            t("purge_title"),
            t("purge_done").format(deleted),
            parent=self.root,
        )


    def _lookup_vernacular(self, scientific_name: str):
        """(ro, en, other English variants) for *scientific_name*.

        Best effort by design: a missing cache or an unknown name yields
        empty values instead of an error, because both callers treat "not in
        the database" as "nothing to offer".
        """
        name = (scientific_name or "").strip()
        if not name:
            return "", "", []
        try:
            chain = taxonomy_lookup_chain(name) or {}
        except Exception:
            return "", "", []
        return self._split_vernacular(
            chain.get("ro_name"), chain.get("en_name")
        )

    @staticmethod
    def _split_vernacular(ro_value, en_value):
        """(ro, en, other English variants) from the database values.

        taxonomy.db joins the vernacular names with "; " — only the first one
        is the primary name, the rest are listed as alternatives.
        """
        def parts(value):
            if not value:
                return []
            return [part.strip() for part in str(value).split(";") if part.strip()]

        ro_parts = parts(ro_value)
        en_parts = parts(en_value)
        return (
            ro_parts[0] if ro_parts else "",
            en_parts[0] if en_parts else "",
            en_parts[1:],
        )

    def update_species_classification(self, prefill: dict | None = None):
        """
        Prezintă toate câmpurile de clasificare într-o singură
        fereastră cu formular — în loc de 4 ferestre simpledialog
        consecutive.

        Când mai multe imagini sunt selectate în lista „Images”, valorile
        introduse în formular se aplică la TOATE fotografiile selectate
        (aceeași specie, categoria și denumirile populare pentru fiecare).
        ``prefill`` (opțional) suprascrie câmpurile pre-completate din
        catalog — folosit de dialogul de verificare CoL.
        """
        if self.image_path is None:
            return

        # Alegeți imaginile țintă: toate rândurile selectate, cu fallback la
        # imaginea curentă când nu există selecție deliberată (backward-
        # compatibil, unică imagine).
        indices = self.selected_indices()
        if not indices:
            try:
                indices = [self.image_paths.index(self.image_path)]
            except (ValueError, AttributeError):
                return
        if not indices:
            return

        target_paths = [self.image_paths[i] for i in indices]

        # Prefill de la prima imagine selectată.
        first_item = get_catalog_item(target_paths[0], OUTPUT_DIR)
        if not first_item:
            messagebox.showinfo(
                t("title_catalog"),
                t("info_analyze_before_reclassify"),
                parent=self.root,
            )
            return

        # Formular într-o singură fereastră
        title = t("title_update_species")
        if len(target_paths) > 1:
            title = f"{title} ({len(target_paths)} imagini)"
        form = tk.Toplevel(self.root)
        form.title(title)
        form.transient(self.root)
        form.grab_set()
        form.geometry("480x340")
        form.resizable(True, True)

        values = {}

        # ttk.Entry has no trace_add(); watching a StringVar is the only
        # reliable way to notice the user typing.
        sci_var = tk.StringVar(master=form)

        for idx, (label, key) in enumerate([
            (t("form_category_hint"), "category"),
            (t("form_scientific_example"), "scientific_name"),
            (t("form_ro_name"), "ro_name"),
            (t("form_en_name"), "en_name"),
        ]):
            ttk.Label(form, text=label).pack(anchor="w", padx=12, pady=(8 if idx == 0 else 4, 0))
            initial = first_item.get(key, "") or ""
            if prefill and prefill.get(key):
                initial = prefill[key]
            if key == "category":
                initial = initial or "Plant"
            if key == "scientific_name":
                entry = ttk.Entry(form, width=50, textvariable=sci_var)
                sci_var.set(initial)
            else:
                entry = ttk.Entry(form, width=50)
                entry.insert(0, initial)
            entry.pack(padx=12, pady=2)
            values[key] = entry

        def set_entry(entry, text):
            entry.delete(0, tk.END)
            entry.insert(0, text)

        def on_names_from_col():
            """Fills the two common-name fields from taxonomy.db.

            Nothing is shown when the typed species has no vernacular names
            in the database: there is simply nothing to offer, and a popup
            would only get in the way. Otherwise the user confirms once and
            both fields are filled in.
            """
            name = sci_var.get().strip()
            ro_name, en_name, extra = self._lookup_vernacular(name)
            if not ro_name and not en_name:
                return

            if ro_name and en_name:
                text = t("form_names_confirm").format(
                    name=name, ro=ro_name, en=en_name
                )
            else:
                text = t("form_names_confirm_one").format(
                    name=name,
                    lang="română" if ro_name else "engleză",
                    value=ro_name or en_name,
                )
            if extra:
                text += t("form_names_extra").format(variants=", ".join(extra))
            if not messagebox.askyesno(
                t("btn_names_from_col"), text, parent=form
            ):
                return
            if ro_name:
                set_entry(values["ro_name"], ro_name)
            if en_name:
                set_entry(values["en_name"], en_name)

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

        # Packed after the OK/Cancel row, so it sits just above it. It stays
        # hidden until a scientific name is typed: before that there is nothing
        # to look up.
        names_row = ttk.Frame(form)
        names_row.pack(side="bottom", pady=(0, 8))
        names_button = ttk.Button(
            names_row,
            text=t("btn_names_from_col"),
            command=on_names_from_col,
        )
        # Not packed here on purpose: sync_names_button() decides, and it runs
        # once at the end.

        # Own flag instead of winfo_ismapped(): a packed child of a hidden
        # window still reports "not mapped", which would break the hiding.
        names_shown = [False]

        def sync_names_button(*_args):
            typed = sci_var.get().strip()
            wanted = bool(typed) and typed not in ("Necunoscut", "Unknown")
            if wanted == names_shown[0]:
                return
            names_shown[0] = wanted
            if wanted:
                names_button.pack()
            else:
                names_button.pack_forget()

        sci_var.trace_add("write", sync_names_button)
        sync_names_button()

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
                t("title_update_species"),
                t("warn_sci_name_required"),
                parent=self.root,
            )
            return

        # Aplică pe TOATE imaginile selectate, colectând succesurile și erorile.
        updated = []
        errors = []
        for path in target_paths:
            try:
                destination = update_catalog_classification(
                    path,
                    OUTPUT_DIR,
                    category,
                    scientific_name,
                    ro_name,
                    en_name,
                )
                updated.append((path, destination))
            except Exception as exc:
                errors.append((path, str(exc)))

        # Re-ancorează starea GUI-ului pentru fiecare imagine actualizată.
        for old_path, destination in updated:
            try:
                current_index = self.image_paths.index(old_path)
                self.image_paths[current_index] = destination
            except ValueError:
                pass
            if old_path in self.image_locations:
                self.image_locations[destination] = self.image_locations.pop(old_path)

        if updated:
            self.image_path = updated[0][1]
            self.update_catalog_buttons()
            self.refresh_image_list()
            self.load_saved_result(self.image_path)

        # Mesaj sumar
        total = len(target_paths)
        ok = len(updated)
        err = len(errors)

        if err == 0 and ok > 0:
            species_folder = Path(updated[0][1]).parent.name if updated else ""
            self.set_status(
                t("status_catalog_updated_name").format(
                    f"{ok} / {total} imagini → {species_folder}"
                )
            )
            if ok > 1:
                messagebox.showinfo(
                    t("title_catalog_updated"),
                    t("info_catalog_updated_multi").format(ok=ok, total=total),
                    parent=self.root,
                )
            else:
                # Comportament original (singură imagine): întreabă dacă
                # reîmprospătăm descrierea.
                if messagebox.askyesno(
                    t("title_catalog_updated"),
                    t("confirm_rerun_after_update"),
                    parent=self.root,
                ):
                    self.start_analysis()
                else:
                    self.set_status(
                        t("status_catalog_updated_pending").format(
                            Path(updated[0][1]).parent.name
                        )
                    )
        elif ok == 0 and err == 0:
            self.set_status(t("status_done").format(""))
        else:
            bad = "\n".join(
                f"• {Path(p).name}: {e}" for p, e in errors[:10]
            )
            messagebox.showerror(
                t("err_catalog_update"),
                t("err_catalog_update_multi").format(
                    ok=ok, total=total, errors=bad
                ),
                parent=self.root,
            )

    # În interiorul clasei MainWindow – adiacent altor metode ca get_catalog_item etc.
    # Sugerat: imediat după metoda update_species_classification()

    def on_thumbnail_click(self, image_path: str):
        """Afișează informațiile despre o poză catalogată când se dă click pe miniatură."""
        item = get_catalog_item(image_path, OUTPUT_DIR) or {}

        scientific = item.get("scientific_name", "Necunoscut")
        ro_name = item.get("ro_name", "")
        en_name = item.get("en_name", "")
        category_raw = item.get("category", "")
        # `map_category` nu a existat niciodată (NameError la click pe
        # miniatură); categoria din DB e deja forma-canonică de folder,
        # deci o normalizează cu același filtru folosit peste tot.
        from core.catalog import _category_folder
        category = _category_folder(category_raw)
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
            messagebox.showinfo(t("title_photo_info"), info_text)


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
                t("warn_nothing_to_save_title"),
                t("warn_no_result_to_save")
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
                t("err_save"),
                str(exc)
            )
            return

        if not updated:
            messagebox.showinfo(
                t("title_save"),
                t("info_not_in_catalog"),
                parent=self.root,
            )
            return

        self.set_output_text(
            response
        )

        self.set_status(
            t("status_saved_name").format(self.image_path.name)
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
                t("warn_nothing_to_save_title"),
                t("warn_no_profile_to_save")
            )
            return

        item = get_catalog_item(
            self.image_path,
            OUTPUT_DIR,
        )

        if not item:
            messagebox.showinfo(
                t("title_species_profile"),
                t("info_analyze_before_save_profile"),
                parent=self.root,
            )
            return

        scientific_name = (item.get("scientific_name") or "").strip()

        if not scientific_name or scientific_name in ("Necunoscut", "Unknown"):
            messagebox.showinfo(
                t("title_species_profile"),
                t("info_species_unknown"),
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
                t("err_save"),
                str(exc),
                parent=self.root,
            )
            return

        if not updated:
            messagebox.showinfo(
                t("title_species_profile"),
                t("info_no_profile_saved").format(scientific_name),
                parent=self.root,
            )
            return

        self.set_status(
            t("status_profile_saved").format(scientific_name)
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
                base_message = f"[{self.batch_current + 1}/{self.batch_total}] \"{self.batch_filename}\""
            else:
                base_message = f"Processing batch: {self.batch_current + 1} / {self.batch_total}"
            # Per-image timing: time on the current image, the running
            # average over the finished ones, and an ETA for the rest.
            img_start = getattr(self, "batch_image_start_time", None)
            if img_start is not None:
                cur = time.perf_counter() - img_start
                timing = [f"cur {self.format_seconds(cur)}"]
                # Average over the images actually analysed (cached items are
                # skipped: they finish in ~0s and would distort the estimate).
                analyzed = getattr(self, "batch_analyzed_count", 0)
                if analyzed > 0:
                    avg = self.batch_analyzed_time / analyzed
                    timing.append(f"avg {self.format_seconds(avg)}")
                    # Images still to process after the current one.
                    remaining = self.batch_total - 1 - self.batch_current
                    if remaining > 0:
                        timing.append(
                            f"left ~{self.format_seconds(avg * remaining)}"
                        )
                base_message += "  (" + " | ".join(timing) + ")"
        elif self.profile_start_time is not None:
            base_message = "Generating species profile"
        else:
            base_message = "Muse Glimmer is working"

        self.set_status(t("status_base_elapsed").format(base_message, elapsed_str))

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
                t("status_prompt_not_found").format(name)
            )
            return

        self.set_prompt_text(
            prompt
        )

        self.set_status(
            t("status_prompt_loaded").format(name)
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
                t("warn_no_prompt_selected_title"),
                t("info_select_prompt")
            )
            return

        prompt = self.get_prompt_text()

        if not prompt:
            messagebox.showwarning(
                t("warn_empty_prompt_title"),
                t("warn_prompt_empty_body")
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
                t("status_prompt_updated").format(name)
            )

        except Exception as exc:

            messagebox.showerror(
                t("err_update_prompt"),
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
                t("warn_invalid_name_title"),
                t("warn_prompt_name_empty")
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
                t("err_prompt_exists_title"),
                str(exc)
            )
            return

        except Exception as exc:

            messagebox.showerror(
                t("err_create_prompt"),
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
            t("status_prompt_created").format(name)
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
