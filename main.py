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
    QPlainTextEdit, QFileDialog, QMessageBox, QFrame, QLineEdit, QListWidgetItem,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize, QSettings
from PyQt6.QtGui import QFont, QIcon

SUPPORTED_EXTS = {".cbz", ".cbr", ".zip", ".rar"}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
APP_ICON_PATH = Path(__file__).resolve().parent / "assets" / "app_icon.svg"

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
        if APP_ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(APP_ICON_PATH)))
        self.setMinimumSize(1120, 760)
        self.resize(1320, 860)

        self._settings = QSettings("ComixConvert", "ComixConvert")
        self.seven_zip = find_7z_exe()
        self.files: list[Path] = []
        self._worker = None
        self._last_out_dir: str | None = None

        central = QWidget()
        central.setObjectName("Root")
        self.setCentralWidget(central)
        self._layout = QVBoxLayout(central)
        self._layout.setContentsMargins(20, 18, 20, 20)
        self._layout.setSpacing(14)

        self._build_header()
        self._build_main_area()
        self._apply_stylesheet()
        self._restore_settings()
        self._sync_epub_options()
        self._refresh_queue()
        self._set_status("Ready")

        if not self.seven_zip:
            self._warning_badge.setText("7-Zip missing")
            self._warning_badge.setObjectName("WarningBadge")
            self._warning_badge.style().unpolish(self._warning_badge)
            self._warning_badge.style().polish(self._warning_badge)
            self._log("WARNING: 7z.exe not found. Install 7-Zip or add it to PATH.")
            self._btn_convert.setEnabled(False)
        else:
            self._log("Ready. Add files or drag & drop.")

    def _make_card(self, title: str, object_name: str = "Card"):
        card = QFrame()
        card.setObjectName(object_name)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(10)

        title_label = QLabel(title)
        title_label.setObjectName("CardTitle")
        layout.addWidget(title_label)
        return card, layout

    def _build_header(self):
        row = QHBoxLayout()
        row.setSpacing(10)

        title_col = QVBoxLayout()
        title_col.setSpacing(2)

        title = QLabel("ComixConvert")
        title.setObjectName("AppTitle")
        subtitle = QLabel("Convert comic archives to clean PDF and EPUB editions")
        subtitle.setObjectName("AppSubtitle")
        title_col.addWidget(title)
        title_col.addWidget(subtitle)

        row.addLayout(title_col)
        row.addStretch()

        self._warning_badge = QLabel("7-Zip detected")
        self._warning_badge.setObjectName("Badge")
        row.addWidget(self._warning_badge)

        self._layout.addLayout(row)

    def _build_main_area(self):
        split = QHBoxLayout()
        split.setSpacing(14)

        left = QVBoxLayout()
        left.setSpacing(14)
        right = QVBoxLayout()
        right.setSpacing(14)

        split.addLayout(left, 5)
        split.addLayout(right, 3)
        self._layout.addLayout(split, 1)

        self._build_import_card(left)
        self._build_queue_card(left)
        self._build_bottom_panel(left)
        self._build_sidebar(right)

    def _build_import_card(self, parent_layout):
        card, layout = self._make_card("Import")
        self._drop_zone = DropZone(on_drop=self.add_paths)
        self._drop_zone.setMinimumHeight(88)
        layout.addWidget(self._drop_zone)

        row = QHBoxLayout()
        row.setSpacing(10)
        self._btn_files = QPushButton("Add files +")
        self._btn_folder = QPushButton("Add folder +")
        self._btn_files.clicked.connect(self.select_files)
        self._btn_folder.clicked.connect(self.select_folder)
        row.addWidget(self._btn_files)
        row.addWidget(self._btn_folder)
        row.addStretch()
        layout.addLayout(row)

        parent_layout.addWidget(card)

    def _build_queue_card(self, parent_layout):
        card = QFrame()
        card.setObjectName("Card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(10)

        header = QHBoxLayout()
        header.setSpacing(10)

        title = QLabel("Queue")
        title.setObjectName("CardTitle")
        header.addWidget(title)

        self._queue_badge = QLabel("0 items")
        self._queue_badge.setObjectName("CountBadge")
        header.addWidget(self._queue_badge)
        header.addStretch()

        self._btn_remove_selected = QPushButton("Remove")
        self._btn_clear = QPushButton("Clear all")
        self._btn_remove_selected.clicked.connect(self.remove_selected)
        self._btn_clear.clicked.connect(self.clear_list)
        header.addWidget(self._btn_remove_selected)
        header.addWidget(self._btn_clear)
        layout.addLayout(header)

        self._queue_list = QListWidget()
        self._queue_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self._queue_list.setObjectName("QueueList")
        layout.addWidget(self._queue_list, 1)

        self._queue_hint = QLabel("Drop archives here or use Add files / Add folder to build your queue.")
        self._queue_hint.setObjectName("HintText")
        layout.addWidget(self._queue_hint)

        parent_layout.addWidget(card, 1)

    def _build_sidebar(self, parent_layout):
        settings_card, settings_layout = self._make_card("Export Settings")

        format_label = QLabel("Format")
        format_label.setObjectName("SectionLabel")
        settings_layout.addWidget(format_label)

        format_row = QHBoxLayout()
        format_row.setSpacing(10)
        self._chk_pdf = QCheckBox("PDF")
        self._chk_pdf.setChecked(True)
        self._chk_pdf.setObjectName("PillCheck")
        self._chk_epub = QCheckBox("EPUB")
        self._chk_epub.setChecked(True)
        self._chk_epub.setObjectName("PillCheck")
        self._chk_pdf.toggled.connect(self._update_summary)
        self._chk_pdf.toggled.connect(self._save_settings)
        self._chk_epub.toggled.connect(self._sync_epub_options)
        self._chk_epub.toggled.connect(self._update_summary)
        self._chk_epub.toggled.connect(self._save_settings)
        format_row.addWidget(self._chk_pdf)
        format_row.addWidget(self._chk_epub)
        format_row.addStretch()
        settings_layout.addLayout(format_row)

        quality_label = QLabel("JPEG Quality")
        quality_label.setObjectName("SectionLabel")
        settings_layout.addWidget(quality_label)

        slider_row = QHBoxLayout()
        slider_row.setSpacing(10)
        self._slider = QSlider(Qt.Orientation.Horizontal)
        self._slider.setRange(40, 100)
        self._slider.setValue(85)
        self._slider_label = QLabel("85")
        self._slider_label.setObjectName("ValueLabel")
        self._slider.valueChanged.connect(lambda v: self._slider_label.setText(str(v)))
        self._slider.valueChanged.connect(self._update_summary)
        self._slider.valueChanged.connect(self._save_settings)
        slider_row.addWidget(self._slider, 1)
        slider_row.addWidget(self._slider_label)
        settings_layout.addLayout(slider_row)

        self._epub_options_label = QLabel("EPUB Options")
        self._epub_options_label.setObjectName("SectionLabel")
        settings_layout.addWidget(self._epub_options_label)

        self._chk_cover = QCheckBox("Use first image as cover")
        self._chk_cover.setChecked(True)
        self._chk_skip_cover_page = QCheckBox("Skip duplicate page 1")
        self._chk_skip_cover_page.setChecked(True)
        self._chk_cover.toggled.connect(self._sync_epub_options)
        self._chk_skip_cover_page.toggled.connect(self._update_summary)
        self._chk_skip_cover_page.toggled.connect(self._save_settings)
        self._chk_cover.toggled.connect(self._update_summary)
        self._chk_cover.toggled.connect(self._save_settings)
        settings_layout.addWidget(self._chk_cover)
        settings_layout.addWidget(self._chk_skip_cover_page)

        parent_layout.addWidget(settings_card, 3)

        output_card, output_layout = self._make_card("Output Folder")
        path_row = QHBoxLayout()
        path_row.setSpacing(8)
        self._out_dir_input = QLineEdit()
        self._out_dir_input.setPlaceholderText("Choose where converted files will be written")
        self._out_dir_input.textChanged.connect(self._update_summary)
        self._out_dir_input.textChanged.connect(self._update_convert_button)
        self._out_dir_input.textChanged.connect(self._save_settings)
        self._btn_browse_out = QPushButton("Browse...")
        self._btn_browse_out.clicked.connect(self.browse_output_folder)
        path_row.addWidget(self._out_dir_input, 1)
        path_row.addWidget(self._btn_browse_out)
        output_layout.addLayout(path_row)

        self._btn_open_out = QPushButton("Open output")
        self._btn_open_out.setEnabled(False)
        self._btn_open_out.clicked.connect(self.open_output_folder)
        output_layout.addWidget(self._btn_open_out)
        parent_layout.addWidget(output_card, 2)

        summary_card, summary_layout = self._make_card("Ready to Convert")
        summary_card.setObjectName("AccentCard")

        summary_row = QHBoxLayout()
        summary_row.setSpacing(12)
        self._summary_files = QLabel()
        self._summary_files.setObjectName("SummaryMetric")
        self._summary_format = QLabel()
        self._summary_format.setObjectName("SummaryMetric")
        self._summary_quality = QLabel()
        self._summary_quality.setObjectName("SummaryMetric")
        summary_row.addWidget(self._summary_files)
        summary_row.addWidget(self._summary_format)
        summary_row.addWidget(self._summary_quality)
        summary_row.addStretch()
        summary_layout.addLayout(summary_row)

        self._summary_note = QLabel()
        self._summary_note.setObjectName("HintText")
        summary_layout.addWidget(self._summary_note)

        self._btn_convert = QPushButton("Convert")
        self._btn_convert.setObjectName("ConvertBtn")
        self._btn_convert.clicked.connect(self.start_convert)
        summary_layout.addWidget(self._btn_convert)
        parent_layout.addWidget(summary_card, 2)

    def _build_bottom_panel(self, parent_layout):
        panel = QFrame()
        panel.setObjectName("BottomPanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(10)

        top = QHBoxLayout()
        top.setSpacing(10)
        self._status_label = QLabel()
        self._status_label.setObjectName("BottomTitle")
        self._toggle_log_btn = QPushButton("Details")
        self._toggle_log_btn.setCheckable(True)
        self._toggle_log_btn.clicked.connect(self._toggle_log_panel)
        top.addWidget(self._status_label)
        top.addStretch()
        top.addWidget(self._toggle_log_btn)
        layout.addLayout(top)

        self._progress_info = QLabel("Files: 0 / 0")
        self._progress_info.setObjectName("HintText")
        layout.addWidget(self._progress_info)

        self._progress = QProgressBar()
        self._progress.setTextVisible(False)
        self._progress.setRange(0, 1)
        self._progress.setValue(0)
        layout.addWidget(self._progress)

        self._subprogress_info = QLabel("Waiting for conversion to start")
        self._subprogress_info.setObjectName("HintText")
        layout.addWidget(self._subprogress_info)

        self._subprogress = QProgressBar()
        self._subprogress.setTextVisible(False)
        self._subprogress.setRange(0, 1)
        self._subprogress.setValue(0)
        layout.addWidget(self._subprogress)

        self._log_box = QPlainTextEdit()
        self._log_box.setObjectName("LogBox")
        self._log_box.setReadOnly(True)
        self._log_box.setFont(QFont("Consolas", 9))
        self._log_box.hide()
        layout.addWidget(self._log_box)

        parent_layout.addWidget(panel, 1)

    def _restore_settings(self):
        out_dir = self._settings.value("output_dir", "", str)
        if out_dir:
            self._out_dir_input.setText(out_dir)
            self._last_out_dir = out_dir
            self._btn_open_out.setEnabled(True)

        self._chk_pdf.setChecked(self._settings.value("export_pdf", True, bool))
        self._chk_epub.setChecked(self._settings.value("export_epub", True, bool))
        self._slider.setValue(self._settings.value("jpeg_quality", 85, int))
        self._chk_cover.setChecked(self._settings.value("epub_cover", True, bool))
        self._chk_skip_cover_page.setChecked(self._settings.value("epub_skip_cover_page", True, bool))

    def _save_settings(self, *_args):
        self._settings.setValue("output_dir", self._out_dir_input.text().strip())
        self._settings.setValue("export_pdf", self._chk_pdf.isChecked())
        self._settings.setValue("export_epub", self._chk_epub.isChecked())
        self._settings.setValue("jpeg_quality", self._slider.value())
        self._settings.setValue("epub_cover", self._chk_cover.isChecked())
        self._settings.setValue("epub_skip_cover_page", self._chk_skip_cover_page.isChecked())

    def _toggle_log_panel(self, checked: bool):
        self._log_box.setVisible(checked)
        self._toggle_log_btn.setText("Hide details" if checked else "Details")

    def _set_status(self, text: str):
        self._status_label.setText(text)

    def _update_convert_button(self):
        count = len(self.files)
        self._btn_convert.setText(f"Convert {count} file{'s' if count != 1 else ''}")

    def _update_summary(self):
        count = len(self.files)
        self._summary_files.setText(f"{count} file{'s' if count != 1 else ''}")

        formats = []
        if self._chk_pdf.isChecked():
            formats.append("PDF")
        if self._chk_epub.isChecked():
            formats.append("EPUB")
        self._summary_format.setText(" + ".join(formats) if formats else "No format")
        self._summary_quality.setText(f"Quality {self._slider.value()}")

        if self._out_dir_input.text().strip():
            self._summary_note.setText("Output folder is ready. Conversion will write files directly there.")
        else:
            self._summary_note.setText("Choose an output folder before starting conversion.")

        self._update_convert_button()

    def _sync_epub_options(self):
        epub_enabled = self._chk_epub.isChecked()
        self._epub_options_label.setVisible(epub_enabled)
        self._chk_cover.setVisible(epub_enabled)
        self._chk_skip_cover_page.setVisible(epub_enabled)
        self._chk_cover.setEnabled(epub_enabled)
        self._chk_skip_cover_page.setEnabled(epub_enabled and self._chk_cover.isChecked())
        self._update_summary()

    def _log(self, text: str):
        self._log_box.appendPlainText(text)

    def _set_busy(self, busy: bool):
        self._btn_convert.setEnabled((not busy) and bool(self.seven_zip))
        self._btn_files.setEnabled(not busy)
        self._btn_folder.setEnabled(not busy)
        self._btn_remove_selected.setEnabled(not busy)
        self._btn_clear.setEnabled(not busy)
        self._btn_browse_out.setEnabled(not busy)
        self._out_dir_input.setEnabled(not busy)
        self._chk_pdf.setEnabled(not busy)
        self._chk_epub.setEnabled(not busy)
        self._chk_cover.setEnabled((not busy) and self._chk_epub.isChecked())
        self._chk_skip_cover_page.setEnabled((not busy) and self._chk_epub.isChecked() and self._chk_cover.isChecked())
        self._slider.setEnabled(not busy)
        self._btn_open_out.setEnabled((not busy) and bool(self._last_out_dir))

    def _make_queue_item_widget(self, path: Path) -> QWidget:
        row = QFrame()
        row.setObjectName("QueueRow")
        layout = QHBoxLayout(row)
        layout.setContentsMargins(12, 9, 12, 9)
        layout.setSpacing(10)

        badge = QLabel(path.suffix.upper().replace('.', ''))
        badge.setObjectName("TypeBadge")
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        badge.setFixedWidth(44)
        layout.addWidget(badge, 0, Qt.AlignmentFlag.AlignTop)

        text_col = QVBoxLayout()
        text_col.setSpacing(2)

        name = QLabel(path.name)
        name.setObjectName("QueueItemName")
        folder = QLabel(str(path.parent))
        folder.setObjectName("QueueItemPath")
        text_col.addWidget(name)
        text_col.addWidget(folder)
        layout.addLayout(text_col, 1)
        return row

    def _refresh_queue(self):
        self._queue_list.clear()
        for p in self.files:
            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, str(p))
            item.setSizeHint(QSize(0, 60))
            self._queue_list.addItem(item)
            self._queue_list.setItemWidget(item, self._make_queue_item_widget(p))
        self._queue_badge.setText(f"{len(self.files)} item{'s' if len(self.files) != 1 else ''}")
        self._queue_hint.setVisible(len(self.files) == 0)
        self._update_summary()

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
        remove_set = {Path(it.data(Qt.ItemDataRole.UserRole)) for it in selected}
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

    def browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose output folder")
        if folder:
            self._out_dir_input.setText(folder)

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

        out_dir = self._out_dir_input.text().strip()
        if not out_dir:
            QMessageBox.warning(self, "No output folder", "Choose an output folder first.")
            return
        if not Path(out_dir).exists():
            QMessageBox.warning(self, "Output folder missing", "The selected output folder does not exist.")
            return

        self._last_out_dir = out_dir
        self._btn_open_out.setEnabled(True)
        self._set_status("Starting conversion")
        self._progress_info.setText(f"Files: 0 / {len(self.files)}")
        self._progress.setRange(0, len(self.files))
        self._progress.setValue(0)
        self._subprogress_info.setText("Preparing archive")
        self._subprogress.setRange(0, 1)
        self._subprogress.setValue(0)
        self._set_busy(True)

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
        self._worker.progress.connect(self._on_file_progress)
        self._worker.status.connect(self._on_status)
        self._worker.subprogress.connect(self._on_subprogress)
        self._worker.finished.connect(self._on_convert_finished)
        self._worker.finished.connect(self._worker.deleteLater)
        self._worker.start()

    def _on_file_progress(self, cur: int, total: int):
        self._progress.setRange(0, max(total, 1))
        self._progress.setValue(cur)
        self._progress_info.setText(f"Files: {cur} / {total}")

    def _on_status(self, text: str):
        clean = text.replace("â€¦", "...").replace("â€”", "-")
        self._set_status(clean)
        self._subprogress_info.setText(clean)

    def _on_subprogress(self, cur: int, tot: int):
        if tot <= 0:
            self._subprogress.setRange(0, 1)
            self._subprogress.setValue(0)
            return
        self._subprogress.setRange(0, max(tot, 1))
        self._subprogress.setValue(cur)

    def _on_convert_finished(self, ok: int, fail: int, out_dir: str):
        self._set_busy(False)
        self._set_status("Ready")
        self._progress_info.setText(f"Files: {ok + fail} / {ok + fail}")
        self._subprogress_info.setText("Conversion finished")
        self._subprogress.setRange(0, 1)
        self._subprogress.setValue(0)

        self._last_out_dir = out_dir
        self._btn_open_out.setEnabled(True)

        if fail == 0:
            QMessageBox.information(self, "Done", f"Converted {ok} file(s).")
        else:
            QMessageBox.warning(
                self, "Partial success", f"OK={ok}, FAIL={fail}\nCheck details for errors."
            )

    def open_output_folder(self):
        if not self._last_out_dir:
            QMessageBox.information(self, "No output folder", "No output folder yet.")
            return
        try:
            os.startfile(self._last_out_dir)
        except Exception as e:
            QMessageBox.warning(self, "Open failed", str(e))

    def _apply_stylesheet(self):
        self.setStyleSheet(
            """
            QMainWindow, QWidget#Root {
                background-color: #f3f4f5;
                color: #2b2f33;
                font-family: Segoe UI, Arial, sans-serif;
                font-size: 10pt;
            }
            QLabel {
                background: transparent;
            }
            QLabel#AppTitle {
                font-size: 19pt;
                font-weight: 700;
                color: #2b2f33;
            }
            QLabel#AppSubtitle {
                font-size: 9.5pt;
                color: #565b60;
            }
            QLabel#Badge, QLabel#WarningBadge, QLabel#CountBadge {
                border-radius: 12px;
                padding: 5px 10px;
                font-weight: 600;
            }
            QLabel#Badge, QLabel#CountBadge {
                background-color: #eceff1;
                color: #3a3a3a;
            }
            QLabel#WarningBadge {
                background-color: #3a3a3a;
                color: white;
            }
            QFrame#Card, QFrame#AccentCard, QFrame#BottomPanel {
                background-color: #ffffff;
                border: 1px solid #d9dde1;
                border-radius: 13px;
            }
            QFrame#AccentCard {
                background-color: #f6f7f8;
                border-color: #cfd4d9;
            }
            QLabel#CardTitle, QLabel#BottomTitle {
                font-size: 12.5pt;
                font-weight: 700;
                color: #2b2f33;
            }
            QLabel#SectionLabel, QLabel#HintText {
                color: #565b60;
            }
            QLabel#ValueLabel {
                min-width: 28px;
                font-size: 11pt;
                font-weight: 700;
                color: #2b2f33;
            }
            QLabel#SummaryMetric {
                font-size: 10.5pt;
                font-weight: 700;
                color: #2b2f33;
            }
            QPushButton {
                background-color: #eceff1;
                color: #2b2f33;
                border: none;
                border-radius: 11px;
                padding: 8px 12px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #dfe3e6;
            }
            QPushButton#ConvertBtn {
                background-color: #3a3a3a;
                color: white;
                min-height: 38px;
                font-size: 10.5pt;
            }
            QPushButton#ConvertBtn:hover {
                background-color: #2f2f2f;
            }
            QPushButton#ConvertBtn:disabled {
                background-color: #a7a7a7;
            }
            QLabel#DropZone {
                border: 2px dashed #cfd4d9;
                border-radius: 13px;
                color: #565b60;
                font-size: 11pt;
                padding: 18px;
                background-color: #f6f7f8;
            }
            QLabel#DropZone[dragover="true"] {
                border-color: #3a3a3a;
                color: #3a3a3a;
                background-color: #eceff1;
            }
            QListWidget#QueueList {
                border: 1px solid #d9dde1;
                border-radius: 12px;
                background-color: #fafafa;
                padding: 4px;
                outline: none;
            }
            QListWidget#QueueList::item {
                border: none;
                padding: 2px;
                margin: 1px 0;
            }
            QListWidget#QueueList::item:selected {
                background-color: transparent;
                color: #2b2f33;
            }
            QFrame#QueueRow {
                background-color: #f5f6f7;
                border: 1px solid #e4e7ea;
                border-radius: 11px;
            }
            QListWidget#QueueList::item:selected QFrame#QueueRow {
                background-color: #eceff1;
                border-color: #cfd4d9;
            }
            QLabel#TypeBadge {
                background-color: #e1e4e8;
                color: #3a3a3a;
                border-radius: 9px;
                padding: 4px 0;
                font-size: 8.5pt;
                font-weight: 700;
            }
            QLabel#QueueItemName {
                color: #2b2f33;
                font-size: 10pt;
                font-weight: 700;
            }
            QLabel#QueueItemPath {
                color: #4e5358;
                font-size: 8.8pt;
            }
            QLineEdit, QPlainTextEdit {
                background-color: #f6f7f8;
                border: 1px solid #d9dde1;
                border-radius: 11px;
                padding: 8px 10px;
                color: #2b2f33;
            }
            QProgressBar {
                border: none;
                border-radius: 10px;
                background-color: #e4e7ea;
                min-height: 16px;
            }
            QProgressBar::chunk {
                background-color: #3a3a3a;
                border-radius: 10px;
            }
            QSlider::groove:horizontal {
                height: 6px;
                background: #d9dde1;
                border-radius: 3px;
            }
            QSlider::sub-page:horizontal {
                background: #3a3a3a;
                border-radius: 3px;
            }
            QSlider::handle:horizontal {
                background: #3a3a3a;
                width: 20px;
                margin: -7px 0;
                border-radius: 10px;
            }
            QCheckBox {
                spacing: 8px;
                color: #2b2f33;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 6px;
                border: 1px solid #cfd4d9;
                background: #f6f7f8;
            }
            QCheckBox::indicator:checked {
                background: #3a3a3a;
                border-color: #3a3a3a;
            }
            """
        )

    def closeEvent(self, event):
        self._save_settings()
        super().closeEvent(event)

def main():
    app = QApplication(sys.argv)
    if APP_ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(APP_ICON_PATH)))
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
