# ComixConvert

**ComixConvert** is a lightweight tool for converting comic archives  
(**CBR / CBZ**) into **PDF and EPUB** formats.

![ComixConvert screenshot](docs/screenshot.png)

---

## Features

- Convert **CBR / CBZ → PDF**
- Convert **CBR / CBZ → EPUB**
- Drag & drop files and folders
- Recursive batch processing
- Adjustable JPEG quality (default: 85)
- EPUB support with **first image as cover** (optional)
- One image per page (no image splitting)
- Built-in log window
- File-level progress bar
- Simple, native GUI (Tkinter)

---

## Dependencies

### Runtime (for source usage)

- **Python 3.10+**
- **7-Zip** (required for CBR/CBZ extraction)

### Python packages

```bash
pip install pillow img2pdf tkinterdnd2
