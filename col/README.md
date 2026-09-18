# Catalogue of Life — local data (NOT uploaded to GitHub)

This folder holds locally downloaded data required by the
app's taxonomic features (`col/taxonomy.db`).

## Why isn't it in the repo?

The CoL archive is ~1.2 GB and the built database ~900 MB.
GitHub does not accept files of that size
(100 MB / file limit), so these files are
excluded via `.gitignore` and every user
rebuilds them locally.

## Rebuild steps (~10 minutes total)

### 1. Download the CoL archive (manual, ~1.2 GB)

1. Go to: https://www.catalogueoflife.org/data/download
2. Under **"Monthly releases"** (or **"Latest release"**),
   download the **ColDP** archive — the `.zip` file containing
   `NameUsage.tsv` and `VernacularName.tsv`
   (names look like `01fdd380-aafc-4d74-88b4-84a3e35650a8.zip`).
3. Copy the `.zip` file into this folder (`col/`).
   Do not unzip or rename it — the app reads
   directly from the archive (streaming).

> Note: CoL links change with every monthly release,
> so there is no fixed URL for automatic download.
> Manual download guarantees you always get the current version.

### 2. Rebuild the database (from the app or CLI)

**From the app** (recommended) — the **Administration** section:
1. 🔄 **Rebuild taxonomy.db** (~2-3 min)
2. 🌐 **Import names from CoL** (~2 min, fills `en_name`
   + the `ron` names from `VernacularName.tsv`)

**From the command line** (equivalent):
```bash
python -c "from core.taxonomy import build_cache; build_cache()"
python tools/import_vernacular.py
```

### 3. (Optional) Romanian names from books

- 📄 **Extract names from PDF** — parses a dictionary
  selected by the user (for example from `name_sources/`) and generates
  `tools/denumiri_draft.xlsx`. The PDF must contain extractable text;
  scanned pages require OCR beforehand. Parsing rules are shared across sources
  and can be edited with **PDF parsing rules** in Administration.
- 📥 **Add names from Excel** — imports the reviewed draft
  into `taxonomy.db` (column B overwrites existing values)

## Folder contents after setup

| File | Source | On GitHub? |
|---|---|---|
| `01fdd380-....zip` | manually downloaded by you | ❌ NO (too large) |
| `taxonomy.db` | built locally (🔄) | ❌ NO (generated) |
| `README.md` | part of the repo | ✅ YES |

