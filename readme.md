# ComixConvert

**ComixConvert** is a desktop app for converting comic archives (`.cbz`, `.cbr`, `.zip`, `.rar`) into **PDF** and **EPUB**.

7zip is required to use this ComixConvert, download it from here https://www.7-zip.org/download.html. 

Program is built with **PyQt6** and focuses on predictable output, batch processing, and simple control over image quality.

<p style="text-align: center;">
  <img src="docs/img.png" alt="CleanText" width="800">
</p>

## Features

- Convert comic archives to **PDF**
- Convert comic archives to **EPUB**
- Drag and drop files or folders
- Recursive folder scan for supported archives
- Adjustable JPEG quality
- EPUB cover support using the first image
- Option to skip duplicate cover page in EPUB
- Built-in progress and details panel
- Remembers last output folder and export settings between runs

## Requirements

- **Python 3.10+**
- **7-Zip** installed and available in `PATH`, or installed in the default Windows location

## Install Dependencies

```bash
pip install -r requirements
```

## Run From Source

```bash
py -3 main.py
```

## Build One EXE

```powershell
py -3 -m PyInstaller --noconfirm --clean --onefile --windowed --name ComixConvert --add-data "assets\app_icon.svg;assets" main.py
```

The built file will be created at:

```text
dist\ComixConvert.exe
```

## Python Dependencies

- `PyQt6`
- `Pillow`
- `img2pdf`

## Notes

- Archives are extracted with **7-Zip**
- Images are normalized to JPEG before export
- PDF and EPUB can be generated in the same batch run

## License

MIT
