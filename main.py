import os
import sys
import re
import shutil
import tempfile
import subprocess
import zipfile
import uuid
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

_NSRE = re.compile(r"(\d+)")


def natural_key(p: Path):
    s = p.name.lower()
    return [int(t) if t.isdigit() else t for t in _NSRE.split(s)]


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
    # -y = assume Yes on all queries
    # -aoa = overwrite all existing files
    # -bd = disable progress indicator (cleaner output)
    cmd = [seven_zip, "x", "-y", "-aoa", "-bd", str(archive_path), f"-o{out_dir}"]
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
    imgs.sort(key=natural_key)
    return imgs


def convert_to_jpegs(
    images: list[Path],
    out_dir: Path,
    quality: int,
    on_step=None,  # callback: on_step(cur:int, total:int, name:str)
) -> list[Path]:
    """
    Convert all input images to JPEG so the quality slider always has a consistent meaning.
    PNG with alpha -> composited on white background.
    """
    out_files = []
    total = len(images)

    for i, img_path in enumerate(images, start=1):
        if on_step:
            on_step(i - 1, total, img_path.name)

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

            # NOTE: optimize=True can be slow on huge books; keep it for now.
            rgb.save(out_path, "JPEG", quality=int(quality), optimize=True)

        out_files.append(out_path)

        if on_step:
            on_step(i, total, img_path.name)

    return out_files


def images_to_pdf(jpegs: list[Path], out_pdf: Path) -> None:
    with open(out_pdf, "wb") as f:
        f.write(
            img2pdf.convert(
                [str(p) for p in jpegs],
                layout_fun=img2pdf.default_layout_fun,
            )
        )


def build_epub_from_images(
    jpegs: list[Path],
    out_epub: Path,
    title: str,
    use_first_image_as_cover: bool = True,
    skip_cover_in_pages: bool = False,
):
    """
    Build a simple, comic-friendly EPUB3:
    - optional cover.xhtml using first image as cover
    - one XHTML page per image (optionally skipping cover image as a page)
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
            encoding="utf-8",
        )

        manifest_items = []
        spine_items = []

        safe_title = escape(title)
        book_id = str(uuid.uuid4())

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

            # cover.xhtml
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
                encoding="utf-8",
            )

            manifest_items.append(
                '<item id="coverpage" href="cover.xhtml" media-type="application/xhtml+xml"/>'
            )
            spine_items.append('<itemref idref="coverpage"/>')

        # ---- Pages + images
        start_idx = 2 if (use_first_image_as_cover and skip_cover_in_pages) else 1
        page_no = 1

        for src_idx in range(start_idx, len(jpegs) + 1):
            img = jpegs[src_idx - 1]
            img_name = f"{page_no:05d}.jpg"
            page_name = f"page_{page_no:05d}.xhtml"

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
  <img src="images/{img_name}" alt="page {page_no}"/>
</body>
</html>
""",
                encoding="utf-8",
            )

            manifest_items.append(
                f'<item id="img{page_no}" href="images/{img_name}" media-type="image/jpeg"/>'
            )
            manifest_items.append(
                f'<item id="page{page_no}" href="{page_name}" media-type="application/xhtml+xml"/>'
            )
            spine_items.append(f'<itemref idref="page{page_no}"/>')

            page_no += 1

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
            encoding="utf-8",
        )

        # content.opf
        cover_meta = (
            f'\n    <meta name="cover" content="{cover_item_id}"/>'
            if cover_item_id
            else ""
        )
        (oebps / "content.opf").write_text(
            f"""<?xml version="1.0" encoding="utf-8"?>
<package version="3.0"
 xmlns="http://www.idpf.org/2007/opf"
 unique-identifier="uid">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:title>{safe_title}</dc:title>
    <dc:language>en</dc:language>
    <dc:identifier id="uid">{book_id}</dc:identifier>{cover_meta}
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
            encoding="utf-8",
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
    items.sort(key=natural_key)
    return items


class ConvertWorker(QThread):
    log_line = pyqtSignal(str)

    progress = pyqtSignal(int, int)       # current_file, total_files
    subprogress = pyqtSignal(int, int)    # current_img, total_imgs
    status = pyqtSignal(str)             # status text

    finished = pyqtSignal(int, int, str)  # ok, fail, out_dir

    def __init__(
        self,
        files,
        out_dir,
        quality,
        export_pdf,
        export_epub,
        epub_cover,
        epub_skip_cover_page,
        seven_zip,
    ):
        super().__init__()
        self.files = files
        self.out_dir = out_dir
        self.quality = quality
        self.export_pdf = export_pdf
        self.export_epub = export_epub
        self.epub_cover = epub_cover
        self.epub_skip_cover_page = epub_skip_cover_page
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
            f"(cover={self.epub_cover}, skip-cover-page={self.epub_skip_cover_page})"
        )
        self.log_line.emit(f"Items: {total}")
        self.log_line.emit("=" * 60)

        self.subprogress.emit(0, 0)
        self.status.emit("Starting…")

        for idx, archive in enumerate(self.files, start=1):
            self.progress.emit(idx - 1, total)

            try:
                self.status.emit(f"[{idx}/{total}] Extracting…")
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

                    self.log_line.emit(f"  Images: {len(imgs)} → JPEG (q={self.quality})")
                    self.status.emit(f"[{idx}/{total}] Converting images to JPEG…")
                    self.subprogress.emit(0, len(imgs))

                    def step(cur, tot, name):
                        self.subprogress.emit(cur, tot)
                        # cur može biti 0 na početku; zato prikazujemo "cur/tot"
                        self.status.emit(f"[{idx}/{total}] JPEG {cur}/{tot} — {name}")

                    jpegs = convert_to_jpegs(imgs, jpeg_dir, self.quality, on_step=step)

                    self.subprogress.emit(0, 0)

                    if self.export_pdf:
                        out_pdf = self.out_dir / (archive.stem + ".pdf")
                        self.status.emit(f"[{idx}/{total}] Building PDF…")
                        self.log_line.emit(f"  Build PDF: {out_pdf.name}")
                        images_to_pdf(jpegs, out_pdf)

                    if self.export_epub:
                        out_epub = self.out_dir / (archive.stem + ".epub")
                        self.status.emit(f"[{idx}/{total}] Building EPUB…")
                        self.log_line.emit(f"  Build EPUB: {out_epub.name}")
                        build_epub_from_images(
                            jpegs=jpegs,
                            out_epub=out_epub,
                            title=archive.stem,
                            use_first_image_as_cover=self.epub_cover,
                            skip_cover_in_pages=self.epub_skip_cover_page,
                        )

                self.status.emit(f"[{idx}/{total}] Done.")
                self.log_line.emit("  OK")
                ok += 1
                self.progress.emit(idx, total)

            except Exception as e:
                fail += 1
                failures.append((archive.name, str(e)))
                self.status.emit(f"[{idx}/{total}] Failed.")
                self.log_line.emit(f"  FAIL: {e}")
                self.progress.emit(idx, total)

        self.log_line.emit("-" * 60)
        self.log_line.emit(f"Done. OK={ok}, FAIL={fail}")
        if failures:
            for name, err in failures[:10]:
                self.log_line.emit(f"  - {name}: {err}")
        self.log_line.emit("-" * 60)

        self.status.emit("Ready.")
        self.subprogress.emit(0, 0)
        self.finished.emit(ok, fail, str(self.out_dir))


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
        self._last_out_dir: str | None = None

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
        self._slider.valueChanged.connect(lambda v: self._slider_label.setText(str(v)))
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

        self._chk_skip_cover_page = QCheckBox("Do not duplicate cover as page 1")
        self._chk_skip_cover_page.setChecked(True)
        form.addRow("", self._chk_skip_cover_page)

        self._layout.addWidget(group)

    def _build_buttons(self):
        row = QHBoxLayout()

        self._btn_files = QPushButton("Select files…")
        self._btn_folder = QPushButton("Select folder…")
        self._btn_remove_selected = QPushButton("Remove selected")
        self._btn_clear = QPushButton("Clear")
        self._btn_open_out = QPushButton("Open output folder")
        self._btn_open_out.setEnabled(False)

        self._btn_convert = QPushButton("Convert →")
        self._btn_convert.setObjectName("ConvertBtn")

        self._btn_files.clicked.connect(self.select_files)
        self._btn_folder.clicked.connect(self.select_folder)
        self._btn_remove_selected.clicked.connect(self.remove_selected)
        self._btn_clear.clicked.connect(self.clear_list)
        self._btn_open_out.clicked.connect(self.open_output_folder)
        self._btn_convert.clicked.connect(self.start_convert)

        row.addWidget(self._btn_files)
        row.addWidget(self._btn_folder)
        row.addWidget(self._btn_remove_selected)
        row.addWidget(self._btn_clear)
        row.addStretch()
        row.addWidget(self._btn_open_out)
        row.addWidget(self._btn_convert)
        self._layout.addLayout(row)

    def _build_queue(self):
        self._queue_label = QLabel("Queue (0 files)")
        self._layout.addWidget(self._queue_label)
        self._queue_list = QListWidget()
        self._queue_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self._queue_list.setFixedHeight(110)
        self._layout.addWidget(self._queue_list)

    def _build_progress(self):
        # Status line (what is happening now)
        self._status_label = QLabel("")
        self._status_label.setObjectName("StatusLabel")
        self._layout.addWidget(self._status_label)

        # Main progress (files)
        self._progress = QProgressBar()
        self._progress.setTextVisible(True)
        self._progress.setFormat("Files: %v / %m")
        self._progress.setRange(0, 1)
        self._progress.setValue(0)
        self._progress.hide()
        self._layout.addWidget(self._progress)

        # Sub progress (images in current file)
        self._subprogress = QProgressBar()
        self._subprogress.setTextVisible(True)
        self._subprogress.setFormat("Images: %v / %m")
        self._subprogress.setRange(0, 1)
        self._subprogress.setValue(0)
        self._subprogress.hide()
        self._layout.addWidget(self._subprogress)

    def _build_log(self):
        self._layout.addWidget(QLabel("Log"))
        self._log_box = QPlainTextEdit()
        self._log_box.setReadOnly(True)
        self._log_box.setFont(QFont("Consolas", 9))
        self._layout.addWidget(self._log_box)

    def _log(self, text: str):
        self._log_box.appendPlainText(text)

    def _set_busy(self, busy: bool):
        self._btn_convert.setEnabled(not busy)
        self._btn_files.setEnabled(not busy)
        self._btn_folder.setEnabled(not busy)
        self._btn_remove_selected.setEnabled(not busy)
        self._btn_clear.setEnabled(not busy)
        self._chk_pdf.setEnabled(not busy)
        self._chk_epub.setEnabled(not busy)
        self._chk_cover.setEnabled(not busy)
        self._chk_skip_cover_page.setEnabled(not busy)
        self._slider.setEnabled(not busy)

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
        self.files.sort(key=natural_key)
        self._refresh_queue()
        self._log(
            f"Added {added} item(s)."
            if added
            else "Nothing new added (duplicates/unsupported)."
        )

    def remove_selected(self):
        selected = self._queue_list.selectedItems()
        if not selected:
            self._log("No selection to remove.")
            return
        remove_set = {Path(it.text()) for it in selected}
        before = len(self.files)
        self.files = [p for p in self.files if p not in remove_set]
        removed = before - len(self.files)
        self._refresh_queue()
        self._log(f"Removed {removed} item(s).")

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

    def start_convert(self):
        if not self.files:
            QMessageBox.warning(self, "No files", "Add CBZ/CBR files first.")
            return
        if not self._chk_pdf.isChecked() and not self._chk_epub.isChecked():
            QMessageBox.warning(
                self, "Nothing selected", "Select PDF and/or EPUB export."
            )
            return
        if not self.seven_zip:
            QMessageBox.critical(self, "7-Zip missing", "7z.exe not found.")
            return

        out_dir = QFileDialog.getExistingDirectory(self, "Choose output folder")
        if not out_dir:
            return

        self._last_out_dir = out_dir
        self._btn_open_out.setEnabled(True)

        # show progress UI only while working
        self._status_label.setText("Starting…")

        self._progress.show()
        self._progress.setRange(0, len(self.files))
        self._progress.setValue(0)

        self._subprogress.hide()
        self._subprogress.setRange(0, 1)
        self._subprogress.setValue(0)

        self._set_busy(True)

        # If cover is OFF, skip-cover-page should effectively be OFF
        epub_skip = self._chk_cover.isChecked() and self._chk_skip_cover_page.isChecked()

        self._worker = ConvertWorker(
            files=list(self.files),
            out_dir=Path(out_dir),
            quality=self._slider.value(),
            export_pdf=self._chk_pdf.isChecked(),
            export_epub=self._chk_epub.isChecked(),
            epub_cover=self._chk_cover.isChecked(),
            epub_skip_cover_page=epub_skip,
            seven_zip=self.seven_zip,
        )

        self._worker.log_line.connect(self._log)

        self._worker.progress.connect(lambda cur, _total: self._progress.setValue(cur))
        self._worker.status.connect(self._status_label.setText)

        def on_sub(cur: int, tot: int):
            if tot <= 0:
                self._subprogress.hide()
                return
            if not self._subprogress.isVisible():
                self._subprogress.show()
            self._subprogress.setRange(0, max(tot, 1))
            self._subprogress.setValue(cur)

        self._worker.subprogress.connect(on_sub)

        self._worker.finished.connect(self._on_convert_finished)
        self._worker.finished.connect(self._worker.deleteLater)
        self._worker.start()

    def _on_convert_finished(self, ok: int, fail: int, out_dir: str):
        self._set_busy(False)

        # hide progress UI after work
        self._progress.hide()
        self._subprogress.hide()
        self._status_label.setText("Ready.")

        self._last_out_dir = out_dir
        self._btn_open_out.setEnabled(True)

        if fail == 0:
            QMessageBox.information(self, "Done", f"Converted {ok} file(s).")
        else:
            QMessageBox.warning(
                self, "Partial success", f"OK={ok}, FAIL={fail}\nCheck Log."
            )

    def open_output_folder(self):
        if not self._last_out_dir:
            QMessageBox.information(self, "No output folder", "No output folder yet.")
            return
        try:
            os.startfile(self._last_out_dir)  # Windows
        except Exception as e:
            QMessageBox.warning(self, "Open failed", str(e))

    def _apply_stylesheet(self):
        self.setStyleSheet(
            """
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
            QLabel#StatusLabel {
                color: #444;
                padding: 2px 0;
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
        """
        )


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
