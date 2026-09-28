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
            "OPRIREA UNUI LOT\nCât rulează ANALYZE ALL sau GEN PROFILE ALL, lângă bara de progres apare butonul ⏹ OPRESTE; scurtătura este Ctrl+. Butonul oprește lotul după imaginea în curs, fără să închidă aplicația: imaginile deja terminate rămân salvate, iar butoanele se reactivează imediat. Poți relua oricând. Închiderea ferestrei cu X oprește totul, dar pierde sesiunea.\n\n"
            "VERIFICARE WEB (AUXILIAR)\nCâmpul de lângă butoanele 🌐 CoL și 🌐 WIKIPEDIA se completează cu specia identificată și acceptă alternative scrise manual. 🌐 CoL deschide Catalogue of Life și copiază numele în clipboard — pe pagină e doar Ctrl+V. 🌐 WIKIPEDIA deschide articolul speciei: sinonimele sunt redirecționate la numele acceptat, util când CoL nu recunoaște numele dat de model. Aceleași butoane apar și în fereastra de atenționare a verificării CoL.\n\n"
            "3. SALVARE ȘI CORECTARE\nO analiză cu identificare recunoscută de parser este catalogată automat: imaginea se copiază în output, iar analiza se salvează în baza de date. SAVE păstrează editările textului pentru o înregistrare existentă. Pentru schimbarea clasificării folosește UPDATE SPECIES, nu doar editarea textului. Dacă identificarea nu a fost recunoscută, verifică întâi catalogarea; simpla afișare a unui răspuns nu garantează salvarea.\n\n"
            "4. PROFILUL SPECIEI\nDupă verificarea speciei, folosește GEN PROFILE. Verifică descrierea din Species Profile și folosește SAVE PROFILE pentru editările tale. LOAD PROFILE reîncarcă profilul salvat; salvează modificările înainte de reîncărcare. Profilul aparține speciei, nu unei singure fotografii."
        ),
        "en": (
            "PURPOSE\nIdentify plants and animals from photos, review the results and build a catalogue with species profiles. AI identification can be wrong; do not use it as proof that a plant or mushroom is edible.\n\n"
            "1. PREPARATION\nStart Ollama and select an image-capable model. See README.md for installation; for the taxonomy database, see col/README.md and the Admin section.\n\n"
            "2. IDENTIFICATION\nAdd 2–3 photos, select one, choose a suitable prompt and press ANALYZE. Read Output Item and verify the species. ANALYZE ALL processes the list; start with a small batch.\n\n"
            "STOPPING A BATCH\nWhile ANALYZE ALL or GEN PROFILE ALL is running, a ⏹ STOP button appears next to the progress bar; the shortcut is Ctrl+. The button stops the batch after the image in flight, without closing the app: images already finished stay saved and the buttons become usable again immediately. You can resume at any time. Closing the window with X stops everything but loses the session.\n\n"
            "WEB VERIFICATION (AUXILIARY)\nThe field next to the 🌐 CoL and 🌐 WIKIPEDIA buttons is pre-filled with the identified species and also accepts manually typed alternatives. 🌐 CoL opens the Catalogue of Life and copies the name to the clipboard — on the page it is just Ctrl+V. 🌐 WIKIPEDIA opens the species article: synonyms redirect to the accepted name, useful when CoL does not recognise the model's name. The same helpers appear in the Catalogue of Life warning dialog.\n\n"
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
    "confirm_selection_used": {
        "ro": "ATENȚIE: se va folosi doar selecția — {chosen} din {total} "
              "imagini încărcate.",
        "en": "NOTE: only the selection will be used — {chosen} of {total} "
              "loaded images.",
    },
    "status_selection_scope": {
        "ro": "{chosen} din {total} imagini selectate — loturile vor folosi "
              "doar selecția (Ctrl / Shift pentru a alege).",
        "en": "{chosen} of {total} images selected — batch commands will use "
              "the selection only (Ctrl / Shift to choose).",
    },
    "btn_stop": {"ro": "⏹ OPRESTE", "en": "⏹ STOP"},
    "btn_stopping": {"ro": "… se oprește", "en": "… stopping"},
    "status_stopping": {
        "ro": "Se oprește după imaginea curentă…",
        "en": "Stopping after the current image…",
    },
    "batch_stopped": {
        "ro": "Lot oprit: {done} din {total} imagini procesate.\n\n"
              "Poți relua când vrei; cele terminate rămân salvate.",
        "en": "Batch stopped: {done} of {total} images processed.\n\n"
              "You can resume at any time; finished ones stay saved.",
    },
    "status_stopped": {
        "ro": "Lot oprit: {done} din {total} — {duration}",
        "en": "Batch stopped: {done} of {total} — {duration}",
    },
    "btn_names_from_col": {"ro": "🌐 NUME DIN CoL", "en": "🌐 NAMES FROM CoL"},
    "btn_profiles_all": {"ro": "TOATE ({n})", "en": "ALL ({n})"},
    "btn_profiles_missing": {
        "ro": "DOAR FĂRĂ PROFIL ({n})",
        "en": "ONLY WITHOUT PROFILE ({n})",
    },
    "profile_scope_title": {
        "ro": "Generate all species profiles",
        "en": "Generate all species profiles",
    },
    "profile_scope_intro": {
        "ro": "{saved} din {total} specii identificate au deja un profil salvat.\n\n"
              "Ce alegi?",
        "en": "{saved} of {total} identified species already have a saved "
              "profile.\n\nWhat do you want to do?",
    },
    "profile_scope_note": {
        "ro": "Regenerarea suprascrie profilele existente, chiar dacă sunt "
              "corecte.\nModelul se poate schimba și din dropdown-ul MODEL. "
              "Procesul poate dura câteva minute.",
        "en": "Regenerating overwrites the existing profiles, even when they "
              "are fine.\nThe model can also be changed in the MODEL "
              "dropdown. The process may take a few minutes.",
    },
    "profiles_done_title": {
        "ro": "Profile de specie finalizate",
        "en": "Species profiles complete",
    },
    "profiles_errors_title": {
        "ro": "Profile de specie finalizate cu erori",
        "en": "Species profiles finished with errors",
    },
    "profiles_done": {
        "ro": "Toate cele {total} profile au fost generate cu succes.\n\n"
              "Durată totală: {duration}",
        "en": "All {total} species profiles were generated successfully.\n\n"
              "Total duration: {duration}",
    },
    "profiles_reserve": {
        "ro": "{n} specie au avut nevoie de bugetul de rezervă "
              "({tokens} tokeni). Ridică PROFILE_NUM_PREDICT la {tokens} "
              "ca să nu mai fie nevoie de reîncercare.",
        "en": "{n} species needed the reserve budget ({tokens} tokens). "
              "Raise PROFILE_NUM_PREDICT to {tokens} to avoid the retry.",
    },
    "batch_done_title": {"ro": "Analiză finalizată", "en": "Batch complete"},
    "batch_errors_title": {
        "ro": "Analiză finalizată cu erori",
        "en": "Batch finished with errors",
    },
    "batch_done": {
        "ro": "Toate cele {total} imagini au fost analizate cu succes.\n\n"
              "Durată totală: {duration}",
        "en": "All {total} images were analyzed successfully.\n\n"
              "Total duration: {duration}",
    },
    "summary_completed": {"ro": "Finalizate: {done}/{total}", "en": "Completed: {done}/{total}"},
    "summary_errors": {"ro": "Erori: {n}", "en": "Errors: {n}"},
    "summary_duration": {"ro": "Durată totală: {duration}", "en": "Total duration: {duration}"},
    "confirm_gen_profile": {"ro": "Vei genera profilul speciei:\n\n{0}\n\nProfilul se bazează pe cunoștințele modelului despre specia în general, nu pe fotografia curentă.\n\nModel (doar text): {1}\n\nContinui?", "en": "You will generate the profile of the species:\n\n{0}\n\nThe profile is based on the model's general knowledge of the species, not on the current photo.\n\nModel (text-only): {1}\n\nContinue?"},
    "confirm_gen_profiles_all": {"ro": "Vei genera profilul pentru {0} specii (din {1} identificate).\n\nPoți alege modelul și din dropdown-ul MODEL.\nProcesul poate dura câteva minute pentru un model local.\n\nContinui?", "en": "You will generate the profile for {0} species (out of {1} identified).\n\nYou can choose the model in the MODEL dropdown.\nThe process may take a few minutes for a local model.\n\nContinue?"},
    "confirm_import_vernacular": {"ro": "Se importă denumirile populare din arhiva CoL (eng + rom). Actualizează DOAR câmpurile goale, nu strică datele manuale. Continui?", "en": "The popular names are imported from the CoL archive (en + ro). Only empty fields are filled in; manual data is kept. Continue?"},
    "confirm_rebuild_taxonomy": {"ro": "Se reconstruiește taxonomy.db din:\n\n{0}\n\nOperația poate dura câteva minute. Continui?", "en": "taxonomy.db will be rebuilt from:\n\n{0}\n\nThis may take a few minutes. Continue?"},
    "confirm_rerun_after_update": {"ro": "Încadrarea a fost actualizată. Vrei să rulezi din nou promptul pentru a regenera descrierea?", "en": "The record was updated. Run the prompt again to regenerate the description?"},
    "err_cannot_open_folder": {"ro": "Nu pot deschide folderul:\n{0}\n\n{1}", "en": "Cannot open the folder:\n{0}\n\n{1}"},
    "err_cannot_open_readme": {"ro": "Nu pot deschide ghidul:\n{0}\n\n{1}", "en": "Cannot open the guide:\n{0}\n\n{1}"},
    "err_cannot_prepare_rules": {"ro": "Nu pot pregăti regulile:\n{0}", "en": "Cannot prepare the rules:\n{0}"},
    "err_import_col": {"ro": "Import CoL eșuat:\n\n{0}", "en": "CoL import failed:\n\n{0}"},
    "err_json_invalid": {"ro": "JSON invalid:\n\n{0}", "en": "Invalid JSON:\n\n{0}"},
    "info_analysis_done_body": {"ro": "Analiza a fost finalizată pentru:\n\n{0}\n\nDurată: {1}", "en": "Analysis finished for:\n\n{0}\n\nDuration: {1}"},
    "info_import_col_done": {"ro": "Import CoL gata: {0} taxoni actualizați (en: {1}, ro: {2}).", "en": "CoL import done: {0} taxa updated (en: {1}, ro: {2})."},
    "info_no_profile_saved": {"ro": "Nu există profil salvat pentru:\n\n{0}\n\nApasă GEN PROFILE pentru a genera unul.", "en": "No saved profile for:\n\n{0}\n\nPress GEN PROFILE to generate one."},
    "info_not_in_catalog": {"ro": "Imaginea nu este în catalog. Rulează ANALYZE mai întâi pentru a o identifica.", "en": "The image is not in the catalog. Run ANALYZE first to identify it."},
    "info_profile_generated": {"ro": "Profilul a fost generat pentru:\n\n{0}\n\nDurată: {1}", "en": "The profile was generated for:\n\n{0}\n\nDuration: {1}"},
    "info_species_unknown_run": {"ro": "Specia nu este încă identificată. Rulează ANALYZE pentru a o identifica.", "en": "The species is not identified yet. Run ANALYZE to identify it."},
    "status_added_images": {"ro": "S-au adăugat {0} imagine/imagini.", "en": "Added {0} image(s)."},
    "status_base_elapsed": {"ro": "{0} — {1}", "en": "{0} — {1}"},
    "status_batch_completed": {"ro": "Lot finalizat: {0}/{1} — {2}", "en": "Batch completed: {0}/{1} — {2}"},
    "status_batch_finished": {"ro": "Lot încheiat: {0}/{1}, erori: {2} — {3}", "en": "Batch finished: {0}/{1}, errors: {2} — {3}"},
    "status_catalog_updated_name": {"ro": "Catalog actualizat: {0}", "en": "Catalog updated: {0}"},
    "status_catalog_updated_pending": {"ro": "Catalog actualizat; analiză în așteptare: {0}", "en": "Catalog updated; analysis pending: {0}"},
    "status_done": {"ro": "Gata — {0}", "en": "Done — {0}"},
    "status_generating_profile": {"ro": "Generează profilul speciei: {0}...", "en": "Generating species profile: {0}..."},
    "status_generating_profile_n": {"ro": "Generează profilul speciei {0} / {1}: {2}...", "en": "Generating species profile {0} / {1}: {2}..."},
    "status_generating_profiles": {"ro": "Generează profilele de specie: 0 / {0}...", "en": "Generating species profiles: 0 / {0}..."},
    "status_loaded": {"ro": "Încărcat: {0}", "en": "Loaded: {0}"},
    "status_processing_batch": {"ro": "Procesare lot: 0 / {0}", "en": "Processing batch: 0 / {0}"},
    "status_profile_generated_short": {"ro": "Profil de specie generat: {0}", "en": "Species profile generated: {0}"},
    "status_profile_generated_time": {"ro": "Profil de specie generat: {0} — {1}", "en": "Species profile generated: {0} — {1}"},
    "status_profile_loaded": {"ro": "Profil de specie încărcat: {0}", "en": "Species profile loaded: {0}"},
    "status_profile_model": {"ro": "Model pentru profiluri: {0}", "en": "Profile model: {0}"},
    "status_profile_saved": {"ro": "Profil de specie salvat: {0}", "en": "Species profile saved: {0}"},
    "status_profiles_completed": {"ro": "Profile de specie finalizate: {0}/{1} — {2}", "en": "Species profiles completed: {0}/{1} — {2}"},
    "status_profiles_errors": {"ro": "Profile de specie: {0}/{1} finalizate, erori: {2} — {3}", "en": "Species profiles: {0}/{1} completed, errors: {2} — {3}"},
    "status_prompt_created": {"ro": "Prompt nou creat: {0}", "en": "New prompt created: {0}"},
    "status_prompt_loaded": {"ro": "Prompt încărcat: {0}", "en": "Prompt loaded: {0}"},
    "status_prompt_not_found": {"ro": "Prompt negăsit: {0}", "en": "Prompt not found: {0}"},
    "status_prompt_updated": {"ro": "Prompt actualizat: {0}", "en": "Prompt updated: {0}"},
    "status_result_not_saved": {"ro": "Rezultatul a fost afișat, dar nu a putut fi salvat: {0}", "en": "Result displayed, but could not be saved: {0}"},
    "status_saved_name": {"ro": "Salvat: {0}", "en": "Saved: {0}"},
    "status_selected": {"ro": "Selectat: {0}", "en": "Selected: {0}"},
    "status_step_completed": {"ro": "Finalizat {0} / {1}: {2}", "en": "Completed {0} / {1}: {2}"},
    "status_step_error": {"ro": "Eroare {0} / {1}: {2}", "en": "Error {0} / {1}: {2}"},
    "err_analysis_failed": {"ro": "Analiza a eșuat", "en": "Analysis failed"},
    "err_catalog_update": {"ro": "Actualizarea catalogului a eșuat", "en": "Catalog update failed"},
    "err_create_prompt": {"ro": "Eroare la crearea promptului", "en": "Create prompt error"},
    "err_image": {"ro": "Eroare imagine", "en": "Image error"},
    "err_initial_config": {"ro": "Configurare inițială", "en": "Initial configuration"},
    "err_no_zip": {"ro": "Nicio arhivă .zip în col/.", "en": "No .zip archive in col/."},
    "err_parse_rules_title": {"ro": "Reguli parsare", "en": "Parsing rules"},
    "err_prompt_exists_title": {"ro": "Promptul există deja", "en": "Prompt already exists"},
    "err_save": {"ro": "Eroare la salvare", "en": "Save error"},
    "err_species_profile": {"ro": "Eroare profil de specie", "en": "Species profile error"},
    "err_update_prompt": {"ro": "Eroare la actualizarea promptului", "en": "Update prompt error"},
    "info_analysis_done_title": {"ro": "Analiză finalizată", "en": "Analysis complete"},
    "info_analyze_before_gen_profile": {"ro": "Analizează imaginea înainte de a genera profilul speciei.", "en": "Analyze the image before generating the species profile."},
    "info_analyze_before_reclassify": {"ro": "Analizează imaginea înainte de reclasificare.", "en": "Analyze the image before reclassifying."},
    "info_analyze_before_save_profile": {"ro": "Analizează imaginea înainte de a salva profilul.", "en": "Analyze the image before saving the profile."},
    "info_enter_prompt": {"ro": "Introdu mai întâi un prompt.", "en": "Enter a prompt first."},
    "info_image_not_cataloged": {"ro": "Imaginea nu este încă în catalog.", "en": "The image is not in the catalog yet."},
    "info_profile_done_title": {"ro": "Profil finalizat", "en": "Profile complete"},
    "info_select_image": {"ro": "Selectează mai întâi o imagine.", "en": "Select an image first."},
    "info_select_prompt": {"ro": "Selectează mai întâi un prompt.", "en": "Select a prompt first."},
    "info_species_unknown": {"ro": "Specia nu este încă identificată.", "en": "The species is not identified yet."},
    "status_images_cleared": {"ro": "Imagini șterse", "en": "Images cleared"},
    "status_profile_failed": {"ro": "Generarea profilului a eșuat", "en": "Species profile generation failed"},
    "status_ready_idle": {"ro": "Gata", "en": "Ready"},
    "status_species_prompt_updated": {"ro": "Promptul speciei a fost actualizat", "en": "Species prompt updated"},
    "status_working": {"ro": "Se lucrează...", "en": "Muse Glimmer is working..."},
    "title_admin": {"ro": "Administrare", "en": "Administration"},
    "title_catalog": {"ro": "Catalog", "en": "Catalog"},
    "title_catalog_updated": {"ro": "Catalog actualizat", "en": "Catalog updated"},
    "title_error": {"ro": "Eroare", "en": "Error"},
    "title_gen_profile": {"ro": "Generează profilul speciei", "en": "Generate species profile"},
    "title_gen_profiles_all": {"ro": "Generează toate profilele de specie", "en": "Generate all species profiles"},
    "title_photo_info": {"ro": "Info poză", "en": "Photo info"},
    "title_save": {"ro": "Salvare", "en": "Save"},
    "title_species_profile": {"ro": "Profil de specie", "en": "Species profile"},
    "title_update_species": {"ro": "Actualizare specie", "en": "Update species"},
    "warn_add_image_first": {"ro": "Adaugă mai întâi cel puțin o imagine.", "en": "Add at least one image first."},
    "warn_catalog_folder_gone": {"ro": "Folderul catalogului nu mai există.", "en": "The catalog folder no longer exists."},
    "warn_empty_prompt_title": {"ro": "Prompt gol", "en": "Empty prompt"},
    "warn_invalid_name_title": {"ro": "Nume invalid", "en": "Invalid name"},
    "warn_no_image_title": {"ro": "Nicio imagine", "en": "No image"},
    "warn_no_images_title": {"ro": "Nicio imagine", "en": "No images"},
    "warn_no_profile_to_save": {"ro": "Nu există profil de salvat.", "en": "There is no profile to save."},
    "warn_no_prompt_selected_title": {"ro": "Niciun prompt selectat", "en": "No prompt selected"},
    "warn_no_prompt_title": {"ro": "Niciun prompt", "en": "No prompt"},
    "warn_no_result_to_save": {"ro": "Nu există rezultat de salvat.", "en": "There is no result to save."},
    "warn_nothing_to_save_title": {"ro": "Nimic de salvat", "en": "Nothing to save"},
    "warn_prompt_empty_body": {"ro": "Promptul nu poate fi gol.", "en": "The prompt cannot be empty."},
    "warn_prompt_name_empty": {"ro": "Numele promptului nu poate fi gol.", "en": "The prompt name cannot be empty."},
    "warn_sci_name_required": {"ro": "Numele științific este obligatoriu.", "en": "The scientific name is required."},
    "form_names_confirm": {
        "ro": "Baza de date conține denumiri populare pentru „{name}”:\n\n"
              "română: {ro}\nengleză: {en}\n\n"
              "Le completez în formular?",
        "en": "The database has common names for “{name}”:\n\n"
              "Romanian: {ro}\nEnglish: {en}\n\n"
              "Fill them in?",
    },
    "form_names_confirm_one": {
        "ro": "Baza de date conține denumirea populară pentru „{name}”:\n\n"
              "{lang}: {value}\n\nO completez în formular?",
        "en": "The database has a common name for “{name}”:\n\n"
              "{lang}: {value}\n\nFill it in?",
    },
    "form_names_extra": {
        "ro": "\n\nAlte variante: {variants}",
        "en": "\n\nOther variants: {variants}",
    },
    "btn_open_col": {"ro": "🌐 CoL", "en": "🌐 CoL"},
    "btn_open_wikipedia": {"ro": "🌐 WIKIPEDIA", "en": "🌐 WIKIPEDIA"},
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
    # Notă: valorile categoriei (Plant/Mushroom/Insect/Reptile/Amphibian/
    # Bird/Mammal/Animal/Other) NU se traduc — sunt chei de date
    # (mappate pe foldere + salvate în catalog).
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
        "ro": "Categorie (Plant, Mushroom, Insect, Arachnid, Mollusc, Fish, Reptile, Amphibian, Bird, Mammal, Animal sau Other):",
        "en": "Category (Plant, Mushroom, Insect, Arachnid, Mollusc, Fish, Reptile, Amphibian, Bird, Mammal, Animal or Other):"
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

    # ---- Dialog verificare CoL după identificare (validate_species) ----
    "col_title": {
        "ro": "Specie neverificată în Catalogue of Life",
        "en": "Species not verified in Catalogue of Life",
    },
    "col_intro": {
        "ro": "Specia identificată de model nu există în baza de date (Catalogue of Life):\n",
        "en": "The species identified by the model does not exist in the database (Catalogue of Life):\n",
    },
    "col_genus_list": {
        "ro": "• Genul {genus} conține {n} specii (primele {shown}): {names}",
        "en": "• Genus {genus} contains {n} species (first {shown}): {names}",
    },
    "col_closest_genus": {
        "ro": "• Cea mai apropiată specie din gen: {name}",
        "en": "• Closest species in the genus: {name}",
    },
    "col_accepted": {
        "ro": "• Specia acceptată: {name} — Genul acceptat: {genus} ({family})",
        "en": "• Accepted species: {name} — Accepted genus: {genus} ({family})",
    },
    "col_genus_changed": {
        "ro": "• Genul s-a schimbat {old} → {new}; epiteta „{epithet}” rămâne neschimbată.",
        "en": "• Genus changed {old} → {new}; epithet “{epithet}” stays the same.",
    },
    "col_k_closest": {
        "ro": "• Cea mai apropiată specie din date: {name} (nesigur — verifică manual)",
        "en": "• Closest species on record: {name} (uncertain — verify manually)",
    },
    "col_no_hint": {
        "ro": "• Nu s-a găsit nicio potrivire plauzibilă în CoL. Specia poate lipsi din snapshot.",
        "en": "• No plausible match found in CoL. The species may be missing from the snapshot.",
    },
    "col_btn_update": {
        "ro": "🌐 Actualizează în catalog...",
        "en": "🌐 Update in catalogue...",
    },
    "col_btn_later": {"ro": "Mai târziu", "en": "Later"},

    # ---- Raport ANALYZE ALL: specii absente din baza de date ----
    "col_batch_title": {
        "ro": "Specii absente din baza de date",
        "en": "Species missing from the database",
    },
    "col_batch_intro": {
        "ro": (
            "{n} specii identificate în acest lot nu au fost găsite în "
            "baza de date (Catalogue of Life).\nAlege o specie din listă "
            "pentru a-i deschide imaginile și a le verifica una câte una."
        ),
        "en": (
            "{n} species identified in this batch were not found in the "
            "database (Catalogue of Life).\nPick a species from the list "
            "to open its images and verify them one by one."
        ),
    },
    "col_batch_species": {"ro": "Specie", "en": "Species"},
    "col_batch_images": {"ro": "imagini", "en": "images"},
    "col_batch_next": {
        "ro": "Următoarea ▶",
        "en": "Next ▶",
    },
    "col_batch_done": {
        "ro": "Marchează ca verificată",
        "en": "Mark as checked",
    },
    "col_batch_close": {"ro": "Închide", "en": "Close"},

    # ---- Web lookups (butoanele 🌐 CoL / 🌐 WIKIPEDIA) ----
    "web_title": {"ro": "Wikipedia", "en": "Wikipedia"},
    "web_no_species": {
        "ro": ("Câmpul de căutare este gol. Introdu un nume de specie "
               "(ex. Lavatera trimestris) sau rulează ANALYZE mai întâi."),
        "en": ("The search field is empty. Enter a species name "
               "(e.g. Lavatera trimestris) or run ANALYZE first."),
    },
    "web_copied_clipboard": {
        "ro": "\"{name}\" copiat în clipboard — lipește (Ctrl+V) în căutarea CoL.",
        "en": "\"{name}\" copied to clipboard — paste (Ctrl+V) in the CoL search.",
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
