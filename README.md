# Specio Identify - Species Identification and Cataloguing with AI

A Python application for analyzing images using local AI models (Ollama). Features both a GUI and CLI interface for batch and single image analysis with metadata extraction and customizable prompts.

## Features

- **GUI Interface**: User-friendly Tkinter-based interface for image analysis
- **Batch Processing**: Analyze multiple images in sequence
- **CLI Support**: Command-line interface for automation and scripting
- **Metadata Extraction**: Extract EXIF, XMP, and GPS metadata from images
- **Custom Prompts**: Create and manage custom analysis prompts
- **Preset Management**: Save and load prompt presets for different use cases
- **Real-time Status**: Live progress tracking with elapsed time display
- **Multiple Models**: Support for various Ollama models (muse-glimmer, qwen3-vl, etc.)
- **Rich Text Formatting**: Markdown-style formatting in prompt and output fields
- **Image Catalog**: Copies identified images into category/species folders and records their fields in SQLite

## Requirements

- Python 3.8+
- Ollama (local AI model server)
- Required Python packages (see `requirements.txt`)

## Installation

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Install and configure Ollama:
   - Download from [ollama.com](https://ollama.com)
   - Pull a vision model: `ollama pull muse-glimmer` or `ollama pull qwen3-vl:8b`
   - Start Ollama server: `ollama serve`

## Instalare de la zero (după `git clone`)

```bash
pip install -r requirements.txt   # Pillow, requests, piexif (+ openpyxl, pymupdf pentru Administrare)
python -m gui.main_window         # sau run.bat pe Windows
```

La prima pornire aplicația face singură pașii care nu necesită date externe:
`output/catalog.db` se creează automat, iar prompturile din `prompts/*.txt`
se importă în baza de date. Dacă lipsește ceva ce necesită descărcare
manuală (arhiva CoL, `taxonomy.db`, denumirile populare), apare **un singur
popup** „Configurare inițială" cu pașii exacți + opțiunea de a deschide
secțiunea **Administrare**. Detalii complete în [`col/README.md`](col/README.md).

Reconstrucție completă a datelor locale (ordinea din Administrare):
1. 🔄 **Reconstruiește taxonomy.db** (~2-3 min, necesită arhiva `.zip` în `col/`)
2. 🌐 **Import denumiri din CoL** (~2 min)
3. 📥 **Adaugă denumiri din Excel** (opțional, cu draftul revizuit)

Dacă distribuția include utilitarele de dezvoltare (`-IncludeDevTools` la export),
scriptul opțional de resetare poate elimina datele locale regenerabile.
**Faceți backup înainte: resetarea șterge inclusiv catalogul și fotografiile din Output.**
Scriptul nu este inclus în distribuția minimă și nu este necesar pentru utilizarea aplicației.
```bash
python tools/local_dev/reset_clean.py --dry-run   # previzualizare
python tools/local_dev/reset_clean.py             # ștergere cu confirmare
```

## Configuration

Edit `config/settings.txt` to configure:

- `OLLAMA_URL`: Ollama server URL (default: http://localhost:11434)
- `MODEL`: AI model to use (default: muse-glimmer)
- `NUM_CTX`: Context window size (default: 8192)
- `NUM_PREDICT`: Maximum tokens to generate (default: 3000)
- `THINKING`: Enable/disable thinking output (default: true)

## Usage

### GUI Mode

Run the application:
```bash
python -m gui.main_window
```

Or use the provided batch file (Windows):
```bash
run.bat
```

**Creating a Desktop Shortcut (Windows):**

To create a shortcut on your Desktop for easy access:

1. Right-click on your Desktop and select **New** → **Shortcut**
2. In the "Type the location of the item" field, enter:
   ```
   C:\path\to\appAI\run.bat
   ```
   (Replace `C:\path\to\appAI\` with the actual installation path. The application folder does not need to be renamed.)
3. Click **Next**
4. Name the shortcut "Specio Identify" and click **Finish**
5. (Optional) Right-click the shortcut → **Properties** → **Change Icon** to select a custom icon

**GUI Features:**
- Add single images or entire folders
- Select custom prompts and presets
- Enter manual location information
- View real-time analysis progress
- Save and edit analysis results
- Batch analyze multiple images

### CLI Mode

Analyze a single image:
```bash
python app.py input/image.jpg
```

Specify a custom prompt:
```bash
python app.py input/image.jpg prompts/custom.txt
```

## Project Structure

```
appAI/
├── app.py              # CLI interface
├── config/             # Configuration files
│   └── settings.txt    # Application settings
├── core/               # Core functionality
│   ├── images.py       # Image processing and metadata
│   ├── ollama.py       # Ollama API integration
│   ├── output.py       # Result saving
│   └── prompts.py      # Prompt management
├── gui/                # GUI interface
│   └── main_window.py  # Main window implementation
├── input/              # Input images directory
├── output/             # Analysis results directory
├── prompts/            # Custom prompts directory
└── logs/               # Application logs
```

## Prompts

Create custom prompts in the `prompts/` directory. Each prompt file should contain the analysis instructions in plain text or Markdown.

Example prompt:
```
Analyze this image and provide:
1. Main subject identification
2. Technical details (camera, settings)
3. Artistic composition analysis
4. Contextual information
```

## Output Format

Analysis results from `ANALYZE` are stored in `output/catalog.db`, containing:

- Image metadata (date, camera, GPS if available)
- Manual location information (if provided)
- Prompt used for analysis
- Model response
- Statistics (tokens, duration)
- Thinking output (if enabled)

The database is the source of truth for catalogued analysis data. Analysis,
`SAVE`, and `SAVE PROFILE` do not export output text files. `SAVE` updates an
existing image record in `output/catalog.db`; `SAVE PROFILE` updates an
existing species profile. Previously exported text files are left untouched.

Automatic cataloguing requires identification data recognized by the parser
(a structured data block or a supported scientific-name text format). If no
such data can be extracted, the response is not persisted by cataloguing and
`SAVE` cannot create a new image record.

## Image Catalog

Catalog image paths are stored relative to the directory containing `catalog.db`
(normally `output/`), using forward slashes. Move the database and species folders
together, with the application closed. The reader also resolves legacy absolute
`.../output/category/species/image` paths against the current output directory,
never falling back to another installation. Existing rows are not bulk-migrated.
External original-image paths remain provenance only. Nonstandard legacy layouts
require explicit migration; renaming individual species folders is not supported.

When the model returns the structured data block requested by the application,
the analyzed image is copied without being moved from `input/` into:

```text
output/
├── plante/
│   └── Rosa_canina/
│       └── Rosa_canina_0001.jpg
├── animale/
├── pasari/
├── ciuperci/
└── altele/
└── catalog.db
```

`catalog.db` stores the image hash, taxonomy fields, Romanian and English
popular names, confidence, and additional prompt-specific fields. The same
image is not copied twice. The catalog currently requires a valid JSON block
at the end of the model response; responses without that block remain usable
as normal text analyses.

## Keyboard Shortcuts (GUI)

- `Enter` in custom location field: Save location
- Click outside text fields: Render Markdown formatting

## Troubleshooting

**Ollama connection error:**
- Ensure Ollama server is running: `ollama serve`
- Check `OLLAMA_URL` in settings.txt

**No metadata extracted:**
- Some images may not have EXIF data
- Try different image formats (JPEG recommended)

**Slow analysis:**
- Reduce `NUM_PREDICT` in settings.txt
- Use a faster model (e.g., qwen3-vl:8b instead of muse-glimmer)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.
