# PDF-TIFF Converter
Converts PDF to TIFF
# PDF → Multipage TIFF Converter (Warehouse)

Windows-only utility that converts PDFs into a single multipage TIFF for document imaging workflows.

## Features
- Drag & drop: drop a PDF onto `converter.exe`
- Virtual-printer workflow: `converter.exe --latest`
- Output TIFF:
  - 300 DPI (PixelsPerInch metadata)
  - 1-bit bilevel (fax/OCR safe)
  - CCITT Group 4 compression (MODI compatible)
  - Multipage TIFF (single `.tif`)
- Opens the result in Microsoft Office Document Imaging (MODI) if installed, otherwise default Windows TIFF viewer
- No folder watchers; only processes files explicitly dropped or requested with `--latest`

## Folder Layout (Portable)
Place these next to the EXE:
converter.exe

- `magick\` = portable ImageMagick folder (must contain `magick.exe` and its supporting files)
- `gs\` = Ghostscript folder (must contain `bin\gswin64c.exe` and `lib\`)

## Workflow 1 — Drag & Drop
1. Drag a `.pdf` onto `converter.exe`
2. A multipage TIFF is written next to the PDF (same filename, `.tif`)
3. TIFF opens in MODI (or default viewer)

## Workflow 2 — Virtual Printer Drop Folder
1. Configure a virtual PDF printer (PDFCreator or Bullzip) to auto-save PDFs to:

   `C:\WarehouseDrop\PrintToTiff\`

2. Run:

   `converter.exe --latest`

This will:
- pick the newest PDF in the drop folder that does **not** already have a matching `.tif`
- wait briefly for the PDF to finish writing
- convert to TIFF
- open the TIFF

## Notes
- Windows “Microsoft Print to PDF” cannot auto-save; it will always prompt for a file name. It can still be used if you manually save into the drop folder.
- If `--latest` reports “No unprocessed PDFs found”, it means every PDF in the folder already has a `.tif` next to it.

## Build (Developer)
Create venv, install PyInstaller, build:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip pyinstaller
python -m PyInstaller --clean --noconfirm --onefile --windowed --name converter converter.py