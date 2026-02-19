import os
import sys
import shutil
import tempfile
import subprocess
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from PIL import Image
import img2pdf

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QSlider, QCheckBox, QPushButton, QListWidget, QProgressBar,
    QPlainTextEdit, QGroupBox, QFormLayout, QFileDialog, QMessageBox,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont

SUPPORTED_EXTS = {".cbz", ".cbr", ".zip", ".rar"}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}


def find_7z_exe() -> str | None:
    exe = shutil.which("7z")
    if exe:
        return exe

    candidates = [
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def run_7z_extract(seven_zip: str, archive_path: Path, out_dir: Path) -> None:
    cmd = [seven_zip, "x", "-y", str(archive_path), f"-o{out_dir}"]
    p = subprocess.run(cmd, capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError(
            "7-Zip extraction failed.\n\n"
            f"File: {archive_path}\n\nSTDOUT:\n{p.stdout}\n\nSTDERR:\n{p.stderr}"
        )


def collect_images(root: Path) -> list[Path]:
    imgs = []
    for p in root.rglob("*"):
        if p.is_file() and p.suffix.lower() in IMAGE_EXTS:
            imgs.append(p)
    imgs.sort(key=lambda x: str(x).lower())
    return imgs


def convert_to_jpegs(images: list[Path], out_dir: Path, quality: int) -> list[Path]:
    """
    Convert all input images to JPEG so the quality slider always has a consistent meaning.
    PNG with alpha -> composited on white background.
    """
    out_files = []
    for i, img_path in enumerate(images, start=1):
        out_name = f"{i:05d}.jpg"
        out_path = out_dir / out_name

        with Image.open(img_path) as im:
            has_alpha = (im.mode in ("RGBA", "LA")) or ("transparency" in im.info)

            if has_alpha:
                rgba = im.convert("RGBA")
                bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
                bg.alpha_composite(rgba)
                rgb = bg.convert("RGB")
            else:
                rgb = im.convert("RGB")

            rgb.save(out_path, "JPEG", quality=int(quality), optimize=True)

        out_files.append(out_path)

    return out_files


def images_to_pdf(jpegs: list[Path], out_pdf: Path) -> None:
    with open(out_pdf, "wb") as f:
        f.write(img2pdf.convert([str(p) for p in jpegs]))


def build_epub_from_images(jpegs: list[Path], out_epub: Path, title: str, use_first_image_as_cover: bool = True):
    """
    Build a simple, comic-friendly EPUB3:
    - optional cover.xhtml using first image as cover
    - one XHTML page per image
    - images embedded under OEBPS/images/
    - minimal nav.xhtml
    """
    if not jpegs:
        raise RuntimeError("EPUB build: no images provided.")

    with tempfile.TemporaryDirectory(prefix="epub_") as tmp:
        root = Path(tmp)

        meta_inf = root / "META-INF"
        oebps = root / "OEBPS"
        images_dir = oebps / "images"

        meta_inf.mkdir()
        images_dir.mkdir(parents=True)

        # mimetype (must be first & stored)
        (root / "mimetype").write_text("application/epub+zip", encoding="utf-8")

        # container.xml
        (meta_inf / "container.xml").write_text(
            """<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0"
 xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf"
     media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>
""",
            encoding="utf-8"
        )

        manifest_items = []
        spine_items = []

        safe_title = escape(title)

        # ---- Cover (first image)
        cover_item_id = None
        if use_first_image_as_cover:
            cover_src = jpegs[0]
            cover_name = "cover.jpg"
            shutil.copy(cover_src, images_dir / cover_name)

            cover_item_id = "coverimg"
            manifest_items.append(
                f'<item id="{cover_item_id}" href="images/{cover_name}" media-type="image/jpeg" properties="cover-image"/>'
            )

            # cover.xhtml (some readers like having it in spine)
            (oebps / "cover.xhtml").write_text(
                f"""<?xml version="1.0" encoding="utf-8"?>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>{safe_title} - Cover</title>
  <meta charset="utf-8"/>
  <style>
    body {{ margin:0; padding:0; }}
    img {{ width:100%; height:auto; display:block; }}
  </style>
</head>
<body>
  <img src="images/{cover_name}" alt="cover"/>
</body>
</html>
""",
                encoding="utf-8"
            )

            manifest_items.append(
                '<item id="coverpage" href="cover.xhtml" media-type="application/xhtml+xml"/>'
            )
            spine_items.append('<itemref idref="coverpage"/>')

        # ---- Pages + images
        for i, img in enumerate(jpegs, start=1):
            img_name = f"{i:05d}.jpg"
            page_name = f"page_{i:05d}.xhtml"

            shutil.copy(img, images_dir / img_name)

            (oebps / page_name).write_text(
                f"""<?xml version="1.0" encoding="utf-8"?>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>{safe_title}</title>
  <meta charset="utf-8"/>
  <style>
    body {{ margin:0; padding:0; }}
    img {{ width:100%; height:auto; display:block; }}
  </style>
</head>
<body>
  <img src="images/{img_name}" alt="page {i}"/>
</body>
</html>
""",
                encoding="utf-8"
            )

            manifest_items.append(
                f'<item id="img{i}" href="images/{img_name}" media-type="image/jpeg"/>'
            )
            manifest_items.append(
                f'<item id="page{i}" href="{page_name}" media-type="application/xhtml+xml"/>'
            )
            spine_items.append(f'<itemref idref="page{i}"/>')

        # nav.xhtml (minimal)
        start_href = "cover.xhtml" if use_first_image_as_cover else "page_00001.xhtml"
        (oebps / "nav.xhtml").write_text(
            f"""<?xml version="1.0" encoding="utf-8"?>
<html xmlns="http://www.w3.org/1999/xhtml"
      xmlns:epub="http://www.idpf.org/2007/ops">
<head>
  <meta charset="utf-8"/>
  <title>Navigation</title>
</head>
<body>
<nav epub:type="toc">
  <ol>
    <li><a href="{start_href}">Start</a></li>
  </ol>
</nav>
</body>
</html>
""",
            encoding="utf-8"
        )

        # content.opf
        cover_meta = f'\n    <meta name="cover" content="{cover_item_id}"/>' if cover_item_id else ""
        (oebps / "content.opf").write_text(
            f"""<?xml version="1.0" encoding="utf-8"?>
<package version="3.0"
 xmlns="http://www.idpf.org/2007/opf"
 unique-identifier="uid">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:title>{safe_title}</dc:title>
    <dc:language>en</dc:language>
    <dc:identifier id="uid">{safe_title}</dc:identifier>{cover_meta}
  </metadata>
  <manifest>
    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>
    {''.join(manifest_items)}
  </manifest>
  <spine>
    {''.join(spine_items)}
  </spine>
</package>
""",
            encoding="utf-8"
        )

        # zip as epub
        with zipfile.ZipFile(out_epub, "w") as z:
            z.write(root / "mimetype", "mimetype", compress_type=zipfile.ZIP_STORED)
            for p in root.rglob("*"):
                if p.name == "mimetype":
                    continue
                z.write(p, p.relative_to(root), compress_type=zipfile.ZIP_DEFLATED)


def find_archives_in_folder(folder: Path) -> list[Path]:
    items = []
    for p in folder.rglob("*"):
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS:
            items.append(p)
    items.sort(key=lambda x: str(x).lower())
    return items


class ConvertWorker(QThread):
    log_line = pyqtSignal(str)
    progress = pyqtSignal(int, int)   # current, total
    finished = pyqtSignal(int, int)   # ok, fail

    def __init__(self, files, out_dir, quality, export_pdf, export_epub,
                 epub_cover, seven_zip):
        super().__init__()
        self.files = files
        self.out_dir = out_dir
        self.quality = quality
        self.export_pdf = export_pdf
        self.export_epub = export_epub
        self.epub_cover = epub_cover
        self.seven_zip = seven_zip

    def run(self):
        ok = 0
        fail = 0
        failures = []
        total = len(self.files)

        self.log_line.emit("=" * 60)
        self.log_line.emit(f"Output: {self.out_dir}")
        self.log_line.emit(f"Quality: {self.quality}")
        self.log_line.emit(
            f"Export: PDF={self.export_pdf} EPUB={self.export_epub} "
            f"(cover={self.epub_cover})"
        )
        self.log_line.emit(f"Items: {total}")
        self.log_line.emit("=" * 60)

        for idx, archive in enumerate(self.files, start=1):
            self.progress.emit(idx - 1, total)
            try:
                self.log_line.emit(f"[{idx}/{total}] Extract: {archive.name}")
                with tempfile.TemporaryDirectory(prefix="comic2export_") as tmp:
                    tmp_path = Path(tmp)
                    extract_dir = tmp_path / "extracted"
                    jpeg_dir = tmp_path / "jpegs"
                    extract_dir.mkdir(parents=True, exist_ok=True)
                    jpeg_dir.mkdir(parents=True, exist_ok=True)

                    run_7z_extract(self.seven_zip, archive, extract_dir)

                    imgs = collect_images(extract_dir)
                    if not imgs:
                        raise RuntimeError("No images found after extraction.")

                    self.log_line.emit(
                        f"  Images: {len(imgs)} → JPEG (q={self.quality})"
                    )
                    jpegs = convert_to_jpegs(imgs, jpeg_dir, self.quality)

                    if self.export_pdf:
                        out_pdf = self.out_dir / (archive.stem + ".pdf")
                        self.log_line.emit(f"  Build PDF: {out_pdf.name}")
                        images_to_pdf(jpegs, out_pdf)

                    if self.export_epub:
                        out_epub = self.out_dir / (archive.stem + ".epub")
                        self.log_line.emit(f"  Build EPUB: {out_epub.name}")
                        build_epub_from_images(
                            jpegs=jpegs,
                            out_epub=out_epub,
                            title=archive.stem,
                            use_first_image_as_cover=self.epub_cover,
                        )

                self.log_line.emit("  OK")
                ok += 1
                self.progress.emit(idx, total)

            except Exception as e:
                fail += 1
                failures.append((archive.name, str(e)))
                self.log_line.emit(f"  FAIL: {e}")
                self.progress.emit(idx, total)

        self.log_line.emit("-" * 60)
        self.log_line.emit(f"Done. OK={ok}, FAIL={fail}")
        if failures:
            for name, err in failures[:10]:
                self.log_line.emit(f"  - {name}: {err}")
        self.log_line.emit("-" * 60)

        self.finished.emit(ok, fail)


class DropZone(QLabel):
    def __init__(self, on_drop, parent=None):
        super().__init__(parent)
        self._on_drop = on_drop
        self.setText("Drop CBZ / CBR files or folders here")
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setAcceptDrops(True)
        self.setMinimumHeight(80)
        self.setObjectName("DropZone")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setProperty("dragover", True)
            self.style().unpolish(self)
            self.style().polish(self)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setProperty("dragover", False)
        self.style().unpolish(self)
        self.style().polish(self)

    def dropEvent(self, event):
        self.setProperty("dragover", False)
        self.style().unpolish(self)
        self.style().polish(self)
        paths = [
            Path(url.toLocalFile())
            for url in event.mimeData().urls()
            if url.isLocalFile()
        ]
        if paths:
            self._on_drop(paths)
        event.acceptProposedAction()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ComixConvert")
        self.setMinimumSize(700, 600)
        self.resize(820, 680)

        self.seven_zip = find_7z_exe()
        self.files: list[Path] = []
        self._worker = None

        central = QWidget()
        self.setCentralWidget(central)
        self._layout = QVBoxLayout(central)
        self._layout.setContentsMargins(16, 16, 16, 16)
        self._layout.setSpacing(12)

        self._build_drop_zone()
        self._build_settings()
        self._build_buttons()
        self._build_queue()
        self._build_progress()
        self._build_log()

        self._apply_stylesheet()

        if not self.seven_zip:
            self._log("WARNING: 7z.exe not found. Install 7-Zip or add it to PATH.")
            self._btn_convert.setEnabled(False)
        else:
            self._log("Ready. Add files or drag & drop.")

    def _build_drop_zone(self):
        self._drop_zone = DropZone(on_drop=self.add_paths)
        self._layout.addWidget(self._drop_zone)

    def _build_settings(self):
        group = QGroupBox("Settings")
        form = QFormLayout(group)
        form.setSpacing(8)

        slider_row = QHBoxLayout()
        self._slider = QSlider(Qt.Orientation.Horizontal)
        self._slider.setRange(40, 100)
        self._slider.setValue(85)
        self._slider_label = QLabel("85")
        self._slider_label.setFixedWidth(28)
        self._slider.valueChanged.connect(
            lambda v: self._slider_label.setText(str(v))
        )
        slider_row.addWidget(self._slider)
        slider_row.addWidget(self._slider_label)
        form.addRow("JPEG quality:", slider_row)

        export_row = QHBoxLayout()
        self._chk_pdf = QCheckBox("PDF")
        self._chk_pdf.setChecked(True)
        self._chk_epub = QCheckBox("EPUB")
        export_row.addWidget(self._chk_pdf)
        export_row.addWidget(self._chk_epub)
        export_row.addStretch()
        form.addRow("Output:", export_row)

        self._chk_cover = QCheckBox("Use first image as cover")
        self._chk_cover.setChecked(True)
        form.addRow("EPUB option:", self._chk_cover)

        self._layout.addWidget(group)

    def _build_buttons(self):
        row = QHBoxLayout()
        btn_files = QPushButton("Select files…")
        btn_folder = QPushButton("Select folder…")
        btn_clear = QPushButton("Clear")
        self._btn_convert = QPushButton("Convert →")
        self._btn_convert.setObjectName("ConvertBtn")

        btn_files.clicked.connect(self.select_files)
        btn_folder.clicked.connect(self.select_folder)
        btn_clear.clicked.connect(self.clear_list)
        self._btn_convert.clicked.connect(self.start_convert)

        row.addWidget(btn_files)
        row.addWidget(btn_folder)
        row.addWidget(btn_clear)
        row.addStretch()
        row.addWidget(self._btn_convert)
        self._layout.addLayout(row)

    def _build_queue(self):
        self._queue_label = QLabel("Queue (0 files)")
        self._layout.addWidget(self._queue_label)
        self._queue_list = QListWidget()
        self._queue_list.setFixedHeight(110)
        self._layout.addWidget(self._queue_list)

    def _build_progress(self):
        self._progress = QProgressBar()
        self._progress.setTextVisible(True)
        self._progress.setFormat("%v / %m")
        self._progress.setValue(0)
        self._progress.setMaximum(0)
        self._layout.addWidget(self._progress)

    def _build_log(self):
        self._layout.addWidget(QLabel("Log"))
        self._log_box = QPlainTextEdit()
        self._log_box.setReadOnly(True)
        self._log_box.setFont(QFont("Consolas", 9))
        self._layout.addWidget(self._log_box)

    # --- Task 5: File management ---

    def _log(self, text: str):
        self._log_box.appendPlainText(text)

    def _refresh_queue(self):
        self._queue_list.clear()
        for p in self.files:
            self._queue_list.addItem(str(p))
        self._queue_label.setText(f"Queue ({len(self.files)} files)")

    def add_paths(self, paths: list[Path]):
        added = 0
        for p in paths:
            if p.is_dir():
                for a in find_archives_in_folder(p):
                    if a not in self.files:
                        self.files.append(a)
                        added += 1
            elif p.suffix.lower() in SUPPORTED_EXTS and p not in self.files:
                self.files.append(p)
                added += 1
        self.files.sort(key=lambda x: str(x).lower())
        self._refresh_queue()
        self._log(
            f"Added {added} item(s)."
            if added else "Nothing new added (duplicates/unsupported)."
        )

    def clear_list(self):
        self.files = []
        self._refresh_queue()
        self._log("Queue cleared.")

    def select_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select .cbr / .cbz files",
            "",
            "Comic archives (*.cbr *.cbz *.rar *.zip);;All files (*.*)",
        )
        if paths:
            self.add_paths([Path(p) for p in paths])

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select folder (recursive search for CBZ/CBR)"
        )
        if folder:
            self.add_paths([Path(folder)])

    # --- Task 6: ConvertWorker wiring ---

    def start_convert(self):
        if not self.files:
            QMessageBox.warning(self, "No files", "Add CBZ/CBR files first.")
            return
        if not self._chk_pdf.isChecked() and not self._chk_epub.isChecked():
            QMessageBox.warning(self, "Nothing selected", "Select PDF and/or EPUB export.")
            return
        if not self.seven_zip:
            QMessageBox.critical(self, "7-Zip missing", "7z.exe not found.")
            return

        out_dir = QFileDialog.getExistingDirectory(self, "Choose output folder")
        if not out_dir:
            return

        self._progress.setMaximum(len(self.files))
        self._progress.setValue(0)
        self._btn_convert.setEnabled(False)

        self._worker = ConvertWorker(
            files=list(self.files),
            out_dir=Path(out_dir),
            quality=self._slider.value(),
            export_pdf=self._chk_pdf.isChecked(),
            export_epub=self._chk_epub.isChecked(),
            epub_cover=self._chk_cover.isChecked(),
            seven_zip=self.seven_zip,
        )
        self._worker.log_line.connect(self._log)
        self._worker.progress.connect(
            lambda cur, _total: self._progress.setValue(cur)
        )
        self._worker.finished.connect(self._on_convert_finished)
        self._worker.start()

    def _on_convert_finished(self, ok: int, fail: int):
        self._btn_convert.setEnabled(True)
        self._progress.setValue(self._progress.maximum())
        if fail == 0:
            QMessageBox.information(self, "Done", f"Converted {ok} file(s).")
        else:
            QMessageBox.warning(
                self, "Partial success", f"OK={ok}, FAIL={fail}\nCheck Log."
            )

    # --- Task 7: QSS Stylesheet ---

    def _apply_stylesheet(self):
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #f5f5f5;
                color: #1a1a1a;
                font-family: Segoe UI, Arial, sans-serif;
                font-size: 10pt;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                margin-top: 6px;
                padding-top: 8px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 4px;
                color: #444;
            }
            QPushButton {
                background-color: #ffffff;
                border: 1px solid #c0c0c0;
                border-radius: 4px;
                padding: 5px 14px;
            }
            QPushButton:hover {
                background-color: #e8e8e8;
                border-color: #999;
            }
            QPushButton#ConvertBtn {
                background-color: #0078d4;
                color: #ffffff;
                border: none;
                font-weight: bold;
                padding: 5px 20px;
            }
            QPushButton#ConvertBtn:hover {
                background-color: #106ebe;
            }
            QPushButton#ConvertBtn:disabled {
                background-color: #a0c4e8;
            }
            QLabel#DropZone {
                border: 2px dashed #b0b0b0;
                border-radius: 6px;
                color: #888;
                font-size: 11pt;
                padding: 20px;
                background-color: #fafafa;
            }
            QLabel#DropZone[dragover="true"] {
                border-color: #0078d4;
                color: #0078d4;
                background-color: #e8f2fc;
            }
            QProgressBar {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                background-color: #e8e8e8;
                height: 18px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #0078d4;
                border-radius: 3px;
            }
            QListWidget, QPlainTextEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                background-color: #ffffff;
            }
        """)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
