"""Central bilingual dictionary (RO/EN) for user-visible GUI strings.

Usage:
    from core.lang import t, get_language, set_language, SUPPORTED_LANGUAGES

    t("btn_analyze")            # current global language
    t("btn_analyze", "en")      # explicit override

Persisted in config/settings.txt as LANGUAGE=ro|en. Switching from the
GUI dropdown saves the setting, so it survives restarts. Before a git
push, edit that single line to LANGUAGE=en for an English default.

NOT translated here (by design):
- Prompts sent to the LLM (prompts/*.txt, DENUMIRI blocks) — must stay
  Romanian so the model answers in Romanian.
- County names, species data, taxonomy content — data, not UI.
- Code comments / docstrings — invisible to the user.
"""
from __future__ import annotations

SUPPORTED_LANGUAGES = ("ro", "en")
DEFAULT_LANGUAGE = "ro"

_current = DEFAULT_LANGUAGE


def get_language() -> str:
    """Return the active UI language ('ro' or 'en')."""
    return _current


def set_language(code: str) -> str:
    """Set the active UI language; falls back to DEFAULT_LANGUAGE."""
    global _current
    code = (code or "").strip().lower()
    _current = code if code in SUPPORTED_LANGUAGES else DEFAULT_LANGUAGE
    return _current


# Key -> {"ro": ..., "en": ...}. Grouped by area.
STRINGS: dict[str, dict[str, str]] = {
    # ---- Status bar / administration ----
    "admin_title": {"ro": "Administrare", "en": "Admin"},
    "admin_show": {"ro": "▼ Administrare", "en": "▼ Admin"},
    "admin_hide": {"ro": "▲ Administrare", "en": "▲ Admin"},
    "btn_rebuild_taxonomy": {
        "ro": "🔄 Reconstruiește taxonomy.db",
        "en": "🔄 Rebuild taxonomy.db",
    },
    "btn_import_vernacular": {
        "ro": "🌐 Import denumiri din CoL",
        "en": "🌐 Import names from CoL",
    },
    "btn_extract_book": {
        "ro": "📄 Extract denumiri din PDF",
        "en": "📄 Extract names from PDF",
    },
    "btn_add_common_names": {
        "ro": "📥 Adaugă denumiri din Excel",
        "en": "📥 Add names from Excel",
    },
    "excel_import_file_title": {
        "ro": "Selectează fișierul Excel cu denumiri",
        "en": "Select the names Excel file",
    },
    "excel_import_need_openpyxl": {
        "ro": "openpyxl nu e instalat.\nRulează: pip install openpyxl",
        "en": "openpyxl is not installed.\nRun: pip install openpyxl",
    },
    "excel_import_read_error": {
        "ro": "Eroare citire:\n\n{exc}\n\nSe acceptă fișiere .xlsx; se citește doar foaia activă.",
        "en": "Read error:\n\n{exc}\n\nOnly the active sheet is read.",
    },
    "excel_import_none_offer": {
        "ro": ("Nicio intrare validă în foaia activă.\n\n"
               "Format așteptat:\n"
               "A = denumire științifică (ex: Rosa canina)\n"
               "B = denumire populară RO\n"
               "C = denumire populară EN (opțional)\n"
               "Prima linie poate fi antet.\n\n"
               "Vrei să creez un șablon .xlsx gol de completat?"),
        "en": ("No valid entries in the active sheet.\n\n"
               "Expected format:\n"
               "A = scientific name (e.g. Rosa canina)\n"
               "B = popular name RO\n"
               "C = popular name EN (optional)\n"
               "The first row may be a header.\n\n"
               "Create an empty .xlsx template to fill in?"),
    },
    "excel_import_missing_offer": {
        "ro": "Nu există fișiere .xlsx în:\n{folder}\n\nNu ai nevoie de o carte PDF pentru completarea manuală a denumirilor.\nFormat: A = denumire științifică, B = denumire RO, C = denumire EN. Completează cel puțin una dintre B/C.\n\nCreezi un șablon Excel?\nDa = creează șablon\nNu = caută un fișier în alt folder\nAnulare = închide",
        "en": "No .xlsx files found in:\n{folder}\n\nYou do not need a PDF book to enter common names manually.\nFormat: A = scientific name, B = Romanian name, C = English name. Fill at least one of B/C.\n\nCreate an Excel template?\nYes = create template\nNo = browse for a file elsewhere\nCancel = close",
    },
    "excel_import_template_title": {
        "ro": "Salvează șablonul Excel",
        "en": "Save Excel template",
    },
    "excel_import_template_saved": {
        "ro": "Șablon creat:\n{out}\n\nCompletează foaia „Import” și las-o activă, apoi importă fișierul de aici.",
        "en": "Template created:\n{out}\n\nFill the “Import” sheet (keep it active), then import it here.",
    },
    "excel_import_template_failed": {
        "ro": "Nu am putut crea șablonul:\n\n{exc}",
        "en": "Could not create the template:\n\n{exc}",
    },
    "excel_import_confirm": {
        "ro": "Am găsit {n} intrări.\n\nSe actualizează/adaugă în taxonomy.db. Continui?",
        "en": "Found {n} entries.\n\nThey will be updated/added in taxonomy.db. Continue?",
    },
    "excel_import_status": {
        "ro": "Adaug {n} intrări...",
        "en": "Adding {n} entries...",
    },
    "excel_import_done": {
        "ro": "Gata: {u} actualizate, {i} adăugate în taxonomy.db.",
        "en": "Done: {u} updated, {i} inserted in taxonomy.db.",
    },
    "excel_import_failed": {
        "ro": "Import eșuat:\n\n{exc}",
        "en": "Import failed:\n\n{exc}",
    },
    "btn_parse_rules": {"ro": "⚙ Reguli parsare", "en": "⚙ Parse rules"},
    "btn_open_col": {"ro": "📂 Deschide folder col/", "en": "📂 Open col/ folder"},
    "status_ready": {"ro": "Ready", "en": "Ready"},
    "lang_label": {"ro": "🌐 Limbă:", "en": "🌐 Language:"},
    # ---- In-app guide ----
    "help_button": {"ro": "? Ghid", "en": "? Guide"},
    "help_title": {"ro": "Specio Identify — Ghid", "en": "Specio Identify — User guide"},
    "help_font_size": {"ro": "Mărime text:", "en": "Text size:"},
    "help_font_smaller": {"ro": "A−", "en": "A−"},
    "help_font_larger": {"ro": "A+", "en": "A+"},
    "help_font_reset": {"ro": "Implicit", "en": "Reset"},
    "help_close": {"ro": "Închide", "en": "Close"},
    "help_intro": {
        "ro": "Alege o secțiune. Poți lăsa ghidul deschis în timp ce lucrezi în aplicație.",
        "en": "Choose a section. You can leave this guide open while working in the application.",
    },
    "help_start_title": {"ro": "Primii pași", "en": "Getting started"},
    "help_catalog_title": {"ro": "Catalogare", "en": "Cataloguing"},
    "help_books_title": {"ro": "Denumiri din PDF", "en": "Names from PDF"},
    "help_data_title": {"ro": "Date și siguranță", "en": "Data and safety"},
    "help_start_body": {
        "ro": (
            "SCOP\nIdentifică plante și animale din fotografii, verifică rezultatele și construiește un catalog cu profiluri de specie. Identificarea AI poate fi greșită; nu o folosi ca dovadă că o plantă sau ciupercă este comestibilă.\n\n"
            "1. PREGĂTIRE\nPornește Ollama și selectează un model compatibil cu imaginile. Pentru instalare, consultă README.md; pentru baza taxonomică, col/README.md și secțiunea Administrare.\n\n"
            "2. IDENTIFICARE\nAdaugă 2–3 fotografii, selectează una, alege un prompt potrivit și apasă ANALYZE. Citește Output Item și verifică specia. ANALYZE ALL procesează lista; începe cu un lot mic.\n\n"
            "3. SALVARE ȘI CORECTARE\nO analiză cu identificare recunoscută de parser este catalogată automat: imaginea se copiază în output, iar analiza se salvează în baza de date. SAVE păstrează editările textului pentru o înregistrare existentă. Pentru schimbarea clasificării folosește UPDATE SPECIES, nu doar editarea textului. Dacă identificarea nu a fost recunoscută, verifică întâi catalogarea; simpla afișare a unui răspuns nu garantează salvarea.\n\n"
            "4. PROFILUL SPECIEI\nDupă verificarea speciei, folosește GEN PROFILE. Verifică descrierea din Species Profile și folosește SAVE PROFILE pentru editările tale. LOAD PROFILE reîncarcă profilul salvat; salvează modificările înainte de reîncărcare. Profilul aparține speciei, nu unei singure fotografii."
        ),
        "en": (
            "PURPOSE\nIdentify plants and animals from photos, review the results and build a catalogue with species profiles. AI identification can be wrong; do not use it as proof that a plant or mushroom is edible.\n\n"
            "1. PREPARATION\nStart Ollama and select an image-capable model. See README.md for installation; for the taxonomy database, see col/README.md and the Admin section.\n\n"
            "2. IDENTIFICATION\nAdd 2–3 photos, select one, choose a suitable prompt and press ANALYZE. Read Output Item and verify the species. ANALYZE ALL processes the list; start with a small batch.\n\n"
            "3. SAVING AND CORRECTIONS\nAn analysis whose identification is recognised by the parser is catalogued automatically: the image is copied into output and the analysis is stored in the database. SAVE keeps text edits for an existing record. Use UPDATE SPECIES to change the classification, rather than only editing the text. If identification was not recognised, check cataloguing first; a displayed response does not guarantee that it was saved.\n\n"
            "4. SPECIES PROFILE\nAfter verifying the species, use GEN PROFILE. Review the description in Species Profile and use SAVE PROFILE to keep your edits. LOAD PROFILE reloads the saved profile; save changes before reloading. A profile belongs to the species, not just one photo."
        ),
    },
    "help_catalog_body": {
        "ro": "QUICK CATALOG\nFolosește această comandă pentru catalogare manuală, când cunoști deja specia. Alege fotografiile și verifică specia și categoria înainte de confirmare. Prin Arată imagini poți consulta fotografiile; un click afișează informațiile lor.\n\nVERIFICARE\nOPEN SPECIES FOLDER deschide folderul speciei pentru imaginea selectată. Verifică existența copiei înainte să ștergi originalul. Catalogarea fotografiilor nu garantează că există și un profil al speciei: acesta se generează separat.\n\nÎN TIMPUL LUCRULUI\nPoți reveni la ghid fără să închizi analiza. Folosește UPDATE SPECIES pentru clasificare și SAVE sau SAVE PROFILE pentru editările textului din panoul corespunzător.",
        "en": "QUICK CATALOG\nUse this command for manual cataloguing when you already know the species. Choose photos and check the species and category before confirming. Use Show images to browse photos; clicking a photo displays its information.\n\nCHECKING\nOPEN SPECIES FOLDER opens the species folder for the selected image. Check the copy exists before deleting the original. Cataloguing photos does not guarantee a species profile exists: generate it separately.\n\nWHILE WORKING\nReturn to this guide without closing the analysis. Use UPDATE SPECIES for classification and SAVE or SAVE PROFILE for text edits in the appropriate pane.",
    },
    "help_books_body": {
        "ro": "PDF → EXCEL → REVIZIE → IMPORT\nAcest flux opțional îmbogățește denumirile din taxonomie; nu trimite cartea către AI.\n\n1. În Administrare, alege Extract denumiri din PDF și sursa din name_sources. PDF-ul trebuie să conțină text extractabil; scanările necesită recunoașterea textului (OCR) în prealabil.\n\n2. Introdu un interval (8-200), o pagină sau lasă gol pentru tot documentul. Sunt paginile PDF, de la 1, nu neapărat numerele tipărite.\n\n3. Revizuiește tools/denumiri_draft.xlsx, foile Draft și Nepotrivite. Verifică inclusiv potrivirile exact; OCR indică o potrivire aproximativă. ro_name vine din PDF; en_name este copiat din baza existentă și poate fi editat manual. O nouă extragere poate suprascrie draftul; păstrează separat versiunile utile.\n\n4. Pentru extrageri greșite, deschide Reguli parsare. Regulile din tools/parse_rules.json sunt comune surselor. Fă o copie înainte de editare și testează pe o pagină. O regulă nouă poate necesita și integrare în cod.\n\n5. Fă backup la col/taxonomy.db. După revizie, folosește Adaugă denumiri din Excel, cu foaia Draft activă. Importul poate suprascrie ambele denumiri RO/EN; golirea unei celule poate șterge valoarea existentă. Importul modifică taxonomia, nu descrierile deja salvate ale fotografiilor.",
        "en": "PDF → EXCEL → REVIEW → IMPORT\nThis optional flow enriches taxonomy names; it does not send the book to AI.\n\n1. In Admin, choose Extract names from PDF and a source from name_sources. The PDF needs extractable text; scans require text recognition (OCR) beforehand.\n\n2. Enter a range (8-200), a single page, or leave blank for the whole document. These are PDF pages starting at 1, not necessarily printed numbers.\n\n3. Review tools/denumiri_draft.xlsx, both Draft and Nepotrivite sheets. Check even exact matches; OCR indicates an approximate match. ro_name comes from the PDF; en_name is copied from the existing database and can be edited manually. A new extraction can overwrite the draft; keep useful versions separately.\n\n4. For incorrect extraction, open Parse rules. Rules in tools/parse_rules.json are shared across sources. Back up before editing and test on one page. A new rule may also require code integration.\n\n5. Back up col/taxonomy.db. After review, use Add names from Excel with Draft as the active sheet. Import can overwrite both RO/EN names; clearing a cell can erase an existing value. Import updates taxonomy, not previously saved photo descriptions.",
    },
    "help_data_body": {
        "ro": "CE PĂSTREZI\noutput/catalog.db conține analizele și profilurile speciilor. Fotografiile catalogate sunt în folderele de specii din output. Păstrează baza și imaginile împreună în backup, preferabil cu aplicația închisă. col/taxonomy.db este o bază separată, pentru taxonomie și denumiri.\n\nINPUT ESTE TEMPORAR\nȘterge originalele numai după verificarea copiei din output și a înregistrării în catalog. Închide aplicația înainte de golirea Input, pentru a evita căile șterse din lista deschisă. Nu toate analizele produc o catalogare reușită.\n\nMUTARE ȘI EDITARE\nCăile imaginilor catalogate sunt relative la output; poți muta proiectul păstrând baza și structura folderelor împreună. Căile absolute vechi din structura standard sunt rezolvate în locația curentă. Modificarea conținutului imaginilor poate schimba hash-ul folosit la regăsire. Nu muta sau recomprima singura copie fără backup.\n\nSALVARE\nSAVE și SAVE PROFILE salvează în baza de date. Deschiderea ghidului nu modifică datele.",
        "en": "WHAT TO KEEP\noutput/catalog.db contains analyses and species profiles. Catalogued photos are in species folders under output. Back up the database and images together, preferably with the application closed. col/taxonomy.db is separate, for taxonomy and names.\n\nINPUT IS TEMPORARY\nDelete originals only after checking the output copy and its catalogue record. Close the application before clearing Input to avoid deleted paths in the open list. Not every analysis results in successful cataloguing.\n\nMOVING AND EDITING\nCatalogued image paths are relative to output; you can move the project while keeping the database and folder structure together. Legacy absolute paths in the standard layout are resolved at the current location. Changing image content can change the hash used for lookup. Do not move or recompress your only copy without a backup.\n\nSAVING\nSAVE and SAVE PROFILE write to the database. Opening this guide does not change your data.",
    },
    # ---- First-run setup check ----
    "setup_title": {"ro": "Configurare inițială", "en": "Initial setup"},
    "setup_welcome": {
        "ro": ("Bun venit! Pentru funcționalitate completă mai sunt necesari\n"
               "următorii pași (o singură dată):\n\n"),
        "en": ("Welcome! For full functionality the following one-time\n"
               "steps are still needed:\n\n"),
    },
    "setup_missing_zip": {
        "ro": ("• Descarcă arhiva Catalogue of Life de la\n"
               "  https://www.catalogueoflife.org/data/download\n"
               "  (fișierul .zip cu NameUsage.tsv) și copiaz-o în folderul col/\n"
               "  — vezi col/README.md pentru pași detaliați."),
        "en": ("• Download the Catalogue of Life archive from\n"
               "  https://www.catalogueoflife.org/data/download\n"
               "  (the .zip with NameUsage.tsv) and copy it into col/\n"
               "  — see col/README.md for detailed steps."),
    },
    "setup_missing_db": {
        "ro": ("• taxonomy.db lipsește — apasă 🔄 Reconstruiește taxonomy.db\n"
               "  din secțiunea Administrare (durează ~2-3 min)."),
        "en": ("• taxonomy.db is missing — press 🔄 Rebuild taxonomy.db\n"
               "  in the Administration section (takes ~2-3 min)."),
    },
    "setup_missing_names": {
        "ro": ("• Denumirile populare lipsesc — apasă\n"
               "  🌐 Import denumiri din CoL din secțiunea Administrare\n"
               "  (~2 min; umple en_name + ro_name)."),
        "en": ("• Common names are missing — press\n"
               "  🌐 Import names from CoL in the Administration section\n"
               "  (~2 min; fills en_name + ro_name)."),
    },
    "setup_readme_hint": {
        "ro": ("\n\nGhidul complet de instalare este în fișierul col/README.md.\n"
               "Vrei să-l deschid acum (împreună cu secțiunea Administrare)?"),
        "en": ("\n\nThe full installation guide is in col/README.md.\n"
               "Open it now (together with the Administration section)?"),
    },
    # ---- Main buttons (Item + Species Profile flows) ----
    # Politică i18n: butoanele principale de acțiune rămân în engleză în
    # AMBELE limbi (ancore invariante — memoria musculară nu se schimbă la
    # comutarea limbii). Se localizează doar textele descriptive (panouri,
    # status, administrare, dialoguri).
    "btn_analyze": {"ro": "ANALYZE", "en": "ANALYZE"},
    "btn_analyze_all": {"ro": "ANALYZE ALL", "en": "ANALYZE ALL"},
    "btn_save": {"ro": "SAVE", "en": "SAVE"},
    "btn_open_species_folder": {
        "ro": "📂 OPEN SPECIES FOLDER",
        "en": "📂 OPEN SPECIES FOLDER",
    },
    "btn_quick_catalog": {"ro": "📁 QUICK CATALOG", "en": "📁 QUICK CATALOG"},
    "btn_update_species": {"ro": "UPDATE SPECIES", "en": "UPDATE SPECIES"},
    "btn_gen_profile": {"ro": "GEN PROFILE", "en": "GEN PROFILE"},
    "btn_load_profile": {"ro": "LOAD PROFILE", "en": "LOAD PROFILE"},
    "btn_save_profile": {"ro": "SAVE PROFILE", "en": "SAVE PROFILE"},
    "btn_gen_profile_all": {"ro": "GEN PROFILE ALL", "en": "GEN PROFILE ALL"},
    "profile_model_label": {"ro": "MODEL:", "en": "MODEL:"},
    "batch_progress": {"ro": "Lot: {cur} / {total}", "en": "Batch: {cur} / {total}"},

    # ---- Common buttons / labels (reused across dialogs) ----
    "btn_cancel": {"ro": "Anulează", "en": "Cancel"},
    "btn_ok": {"ro": "OK", "en": "OK"},
    "btn_open_folder": {"ro": "📂 Deschide folder", "en": "📂 Open folder"},
    "lbl_species": {"ro": "Specie:", "en": "Species:"},
    "lbl_category": {"ro": "Categorie:", "en": "Category:"},
    "lbl_folder": {"ro": "Folder:", "en": "Folder:"},

    # ---- Extract names from PDF (book_source) dialog + results ----
    "extract_pdf_pages_title": {
        "ro": "Extract denumiri din PDF",
        "en": "Extract names from PDF",
    },
    "extract_pdf_pages_prompt": {
        "ro": "Interval de pagini PDF (1-based), ex: 8-200\n(gol = tot documentul):",
        "en": "PDF page range (1-based), e.g. 8-200\n(empty = whole document):",
    },
    "extract_pdf_all_pages": {
        "ro": "tot documentul",
        "en": "whole document",
    },
    "extract_pdf_preparing": {
        "ro": "Extrag denumiri din PDF; pregătesc potrivirea cu taxonomia...",
        "en": "Extracting names from PDF; preparing taxonomy matching...",
    },
    "extract_pdf_progress": {
        "ro": "Extract PDF: {done}/{total} intrări, {hits} potrivite...",
        "en": "Extract PDF: {done}/{total} entries, {hits} matched...",
    },
    "extract_pdf_done": {
        "ro": "Extract PDF gata.",
        "en": "PDF extraction complete.",
    },
    "extract_pdf_result_msg": {
        "ro": ("Draft generat:\n{out_path}\n\nPagini: {pages}\n"
               "Intrări detectate: {entries}\nPotrivite: {matched} "
               "(exact: {exact}, OCR: {ocr})\nNepotrivite: {unmatched}\n\n"
               "Revizuiește draftul, apoi apasă „Adaugă denumiri din Excel”."),
        "en": ("Draft generated:\n{out_path}\n\nPages: {pages}\n"
               "Detected entries: {entries}\nMatched: {matched} "
               "(exact: {exact}, OCR: {ocr})\nUnmatched: {unmatched}\n\n"
               "Review the draft, then press “Add names from Excel”."),
    },
    "extract_pdf_failed_status": {
        "ro": "Extract PDF eșuat.",
        "en": "PDF extraction failed.",
    },
    "extract_pdf_failed_msg": {
        "ro": "Extract PDF eșuat:\n\n{error}",
        "en": "PDF extraction failed:\n\n{error}",
    },
    "extract_pdf_choose_title": {
        "ro": "Selectează PDF-ul (dicționar/carte de denumiri populare)",
        "en": "Select the PDF (dictionary/book of common names)",
    },
    # ---- Quick Catalog dialog ----
    # Notă: valorile categoriei (Planta/Animal/Pasare/Ciuperca/Alta) NU se
    # traduc — sunt chei de date (mappate pe foldere + salvate în catalog).
    "qc_title": {"ro": "Catalog rapid", "en": "Quick Catalog"},
    "qc_no_images": {"ro": "Adaugă mai întâi imagini.", "en": "Add images first."},
    "qc_selected_image": {
        "ro": "Imagine selectată: {name}",
        "en": "Selected image: {name}",
    },
    "qc_no_selection": {"ro": "Nicio imagine selectată.", "en": "No image selected."},
    "qc_pics_header": {
        "ro": "Imagini (opțional, doar ca referință):",
        "en": "Images (optional, reference only):",
    },
    "qc_show_pics": {"ro": "🔎 Arată imagini", "en": "🔎 Show images"},
    "qc_hide_pics": {"ro": "🙈 Ascunde imagini", "en": "🙈 Hide images"},
    "qc_catalogued_photo": {"ro": "Poză catalogată", "en": "Catalogued photo"},
    "qc_sci_name_label": {
        "ro": "Numele științific (autocomplete din taxonomy.db):",
        "en": "Scientific name (autocomplete from taxonomy.db):",
    },
    "qc_info_frame": {"ro": "Informații specie", "en": "Species info"},
    "qc_type_to_search": {"ro": "Scrie pentru căutare...", "en": "Type to search..."},
    "qc_min_chars": {
        "ro": "Scrie minim 2 caractere...",
        "en": "Type at least 2 characters...",
    },
    "qc_no_results": {"ro": "(niciun rezultat)", "en": "(no results)"},
    "qc_validated": {
        "ro": "Specie selectată: {name}",
        "en": "Selected species: {name}",
    },
    "qc_valid": {
        "ro": "✓ Specie validă din taxonomy.db",
        "en": "✓ Valid species from taxonomy.db",
    },
    "qc_invalid": {
        "ro": "✗ Specia nu există în taxonomy.db",
        "en": "✗ Species not in taxonomy.db",
    },
    "qc_select_image": {
        "ro": "Selectează cel puțin o imagine.",
        "en": "Select at least one image.",
    },
    "qc_enter_name": {
        "ro": "Introdu numele științific.",
        "en": "Enter the scientific name.",
    },
    "qc_choose_category": {"ro": "Alege o categorie.", "en": "Choose a category."},
    "qc_result_title": {
        "ro": "Quick catalog - rezultat",
        "en": "Quick Catalog - result",
    },
    "qc_result_msg": {
        "ro": "Catalogate: {n} imagine(i).\n\nDestinație:\n{folder}",
        "en": "Catalogued: {n} image(s).\n\nDestination:\n{folder}",
    },
    "qc_catalog_btn": {"ro": "📁 Catalogează", "en": "📁 Catalogue"},
    # Formular Update Species
    "form_category_hint": {
        "ro": "Categorie (Planta, Animal, Pasare, Ciuperca sau Alta):",
        "en": "Category (Plant, Animal, Bird, Mushroom or Other):"
    },
    "form_scientific_example": {
        "ro": "Nume științific corect (ex. Duchesnea indica):",
        "en": "Correct scientific name (e.g. Duchesnea indica):"
    },
    "form_ro_name": {
        "ro": "Denumire populară în limba română:",
        "en": "Common name in Romanian:"
    },
    "form_en_name": {
        "ro": "Denumire populară în limba engleză:",
        "en": "Common name in English:"
    },

}


def t(key: str, lang: str | None = None) -> str:
    """Return the string for *key* in *lang* (or the active language)."""
    code = (lang or _current or DEFAULT_LANGUAGE).strip().lower()
    if code not in SUPPORTED_LANGUAGES:
        code = DEFAULT_LANGUAGE
    entry = STRINGS.get(key)
    if not entry:
        return key
    return entry.get(code) or entry.get(DEFAULT_LANGUAGE) or key

# --- i18n stage-2 auto-merge (panels only; main buttons live in STRINGS) ---
_EXTRA2 = {
    "panel_images": ("Imagini", "Images"),
    "panel_prompt": ("Prompt", "Prompt"),
    "panel_species_prompt": ("Prompt Specie", "Species Prompt"),
    "panel_output_item": ("Output Item", "Output Item"),
    "panel_output_profile": ("Output Species Profile", "Output Species Profile"),
}
for _k, (_ro, _en) in _EXTRA2.items():
    if _k in STRINGS:
        STRINGS[_k].setdefault("ro", _ro)
        STRINGS[_k].setdefault("en", _en)
    else:
        STRINGS[_k] = {"ro": _ro, "en": _en}
