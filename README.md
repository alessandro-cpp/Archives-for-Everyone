# Archives-for-Everyone — multilingual OCR and translation

**Author:** Alessandro Ceppetelli

A locally run, open-source workflow for extracting text from photographs and scans of machine-written archival documents and exporting **original transcription + translation** in parallel Word columns.

## Features
- Inputs: `.heic`, `.jpg`, `.jpeg`, `.png`, `.pdf` (text-based PDFs are parsed directly where possible).
- Source languages: choose one or more ISO language codes; language identification is restricted to the selected list. Default `auto` refers to **Dutch, French, English, German**, not all world languages.
- Target language: choose an ISO language code; translation uses the Argos Translate packages available for the requested language pair, with intermediate language routes where necessary.
- The original transcription is preserved beside the translated text in a landscape `.docx` file, one per input subfolder.
- Text remains local during OCR/translation; installing translation models initially requires downloading packages.

## Requirements

Python 3.10+, the [Tesseract OCR engine](https://github.com/tesseract-ocr/tesseract) and relevant traineddata language packs. For scanned PDFs, `pdf2image` also needs [Poppler](https://poppler.freedesktop.org/) installed and available in your PATH.

```bash
pip install -r requirements.txt
```

## Input folder example

```text
input/
  archive_collection_01/
    page_001.jpg
    page_002.heic
  archive_collection_02/
    source.pdf
```

## Usage

From the directory containing the script:

```bash
# Choose multiple source languages; translate to Italian
python archives_for_everyone_multilingual.py --source nl,fr,de,en --target it --input ./input --output ./output

# Single-source document: Italian to English
python archives_for_everyone_multilingual.py --source it --target en --input ./input

# Default sources (nl,fr,en,de), output in French
python archives_for_everyone_multilingual.py --source auto --target fr --input ./input

# Interactive language selection
python archives_for_everyone_multilingual.py --interactive --input ./input

# Display supported Tesseract source-language codes
python archives_for_everyone_multilingual.py --list-languages
```

The script automatically searches for `tesseract.exe` in PATH and common Windows install directories. If needed, use `--tesseract "C:\\...\\tesseract.exe"` or set `TESSERACT_CMD`.

**Important limitations:** The destination language must be reachable via published or installed Argos Translate language packages. Not all target languages or pairs have models. Multi-language OCR and automatic language detection can be unreliable on short pages, damaged material, mixed scripts or historical typewriting; limit the selected sources for better accuracy. Handwriting is not a validated use case. Photographs depicting non-text content are not the intended input. Review transcriptions and translations against the originals before research citation.

**Rights and privacy:** Publish only examples for which you have sharing rights; institutional archives can impose reproduction restrictions.

## Windows / PowerShell quick start

```powershell
python -m pip install -r requirements.txt
python archives_for_everyone_multilingual.py --source nl,fr,de,en --target it --input .\input --output .\output
```

The Tesseract executable and selected OCR language packs must be installed separately. For scanned PDF files, install Poppler as well.


## Argos Translate compatibility

The script detects installed direct translation models through `argostranslate.package.get_installed_packages()`. This avoids relying on unsupported `Language.translations` attributes and correctly distinguishes installed packages from composite translation paths.
