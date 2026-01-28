import os
import shutil
import tempfile
import subprocess
import threading
import zipfile
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from xml.sax.saxutils import escape

from PIL import Image
import img2pdf

# Drag & Drop
from tkinterdnd2 import DND_FILES, TkinterDnD

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


def split_dnd_list(data: str) -> list[str]:
    """
    TkinterDnD gives a string like:
      '{C:/path/file 1.cbz} {C:/path/file2.cbr}'
    or plain 'C:/path/file.cbz'
    """
    data = data.strip()
    if not data:
        return []
    out = []
    cur = ""
    in_brace = False
    for ch in data:
        if ch == "{":
            in_brace = True
            cur = ""
        elif ch == "}":
            in_brace = False
            if cur:
                out.append(cur)
                cur = ""
        elif ch == " " and not in_brace:
            if cur:
                out.append(cur)
                cur = ""
        else:
            cur += ch
    if cur:
        out.append(cur)
    return out


class App:
    def __init__(self, root: TkinterDnD.Tk):
        self.root = root
        root.title("CBR/CBZ → PDF/EPUB (Quality, Drag&Drop, Batch, Log)")
        root.geometry("900x650")

        self.seven_zip = find_7z_exe()
        self.files: list[Path] = []
        self.is_running = False

        main = ttk.Frame(root, padding=12)
        main.pack(fill="both", expand=True)

        left = ttk.Frame(main)
        left.pack(side="left", fill="y")

        right = ttk.Frame(main)
        right.pack(side="right", fill="both", expand=True, padx=(12, 0))

        # Quality
        ttk.Label(left, text="JPEG quality (default 85):").pack(anchor="w")
        self.quality = ttk.Scale(left, from_=40, to=100, orient="horizontal")
        self.quality.set(85)
        self.quality.pack(fill="x", pady=(4, 8))

        self.q_label = ttk.Label(left, text="85")
        self.q_label.pack(anchor="w", pady=(0, 10))
        self.quality.configure(command=lambda _v: self.q_label.config(text=str(int(self.quality.get()))))

        # Export options
        self.export_pdf = tk.BooleanVar(value=True)
        self.export_epub = tk.BooleanVar(value=False)

        ttk.Checkbutton(left, text="Export PDF", variable=self.export_pdf).pack(anchor="w")
        ttk.Checkbutton(left, text="Export EPUB", variable=self.export_epub).pack(anchor="w", pady=(0, 10))

        # Cover option for EPUB
        self.epub_cover = tk.BooleanVar(value=True)
        ttk.Checkbutton(left, text="EPUB: first image as cover", variable=self.epub_cover).pack(anchor="w", pady=(0, 10))

        # Buttons
        ttk.Button(left, text="Select files…", command=self.select_files).pack(fill="x", pady=3)
        ttk.Button(left, text="Select folder (recursive)…", command=self.select_folder).pack(fill="x", pady=3)
        ttk.Button(left, text="Clear list", command=self.clear_list).pack(fill="x", pady=(3, 10))

        self.out_btn = ttk.Button(left, text="Convert →", command=self.start_convert_thread)
        self.out_btn.pack(fill="x", pady=3)

        # Progress bar (per-file)
        self.progress = ttk.Progressbar(left, orient="horizontal", mode="determinate", length=240)
        self.progress.pack(fill="x", pady=(10, 5))
        self.progress_label = ttk.Label(left, text="Progress: 0 / 0")
        self.progress_label.pack(anchor="w")

        # Status
        self.info = ttk.Label(left, text="", wraplength=280, foreground="#444")
        self.info.pack(fill="x", pady=(10, 0))

        # Drop zone
        self.drop = ttk.Label(
            right,
            text="⬇️ Drag & drop CBZ/CBR files or folders here",
            anchor="center",
            relief="ridge",
            padding=18
        )
        self.drop.pack(fill="x")

        self.drop.drop_target_register(DND_FILES)
        self.drop.dnd_bind("<<Drop>>", self.on_drop)

        # Queue list
        ttk.Label(right, text="Queue:").pack(anchor="w", pady=(10, 3))
        self.queue_box = ScrolledText(right, height=8)
        self.queue_box.pack(fill="x")
        self.queue_box.configure(state="disabled")

        # Log
        ttk.Label(right, text="Log:").pack(anchor="w", pady=(10, 3))
        self.log_box = ScrolledText(right)
        self.log_box.pack(fill="both", expand=True)

        if not self.seven_zip:
            self.set_info(
                "⚠️ 7z.exe not found.\n"
                "Install 7-Zip or add it to PATH.\n"
                "Expected: C:\\Program Files\\7-Zip\\7z.exe"
            )
            self.out_btn.state(["disabled"])
        else:
            self.set_info("Ready. Add files or a folder. Drag&drop works.")

    def set_info(self, text: str):
        self.info.config(text=text)

    def log(self, text: str):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.root.update_idletasks()

    def refresh_queue(self):
        self.queue_box.configure(state="normal")
        self.queue_box.delete("1.0", "end")
        for p in self.files:
            self.queue_box.insert("end", str(p) + "\n")
        self.queue_box.configure(state="disabled")
        self.set_info(
            f"Queue: {len(self.files)} item(s). "
            f"Quality: {int(self.quality.get())}. "
            f"PDF={self.export_pdf.get()} EPUB={self.export_epub.get()}"
        )

    def add_paths(self, paths: list[Path]):
        added = 0
        for p in paths:
            if p.is_dir():
                archives = find_archives_in_folder(p)
                for a in archives:
                    if a not in self.files:
                        self.files.append(a)
                        added += 1
            else:
                if p.suffix.lower() in SUPPORTED_EXTS and p not in self.files:
                    self.files.append(p)
                    added += 1

        self.files.sort(key=lambda x: str(x).lower())
        self.refresh_queue()
        self.log(f"Added {added} item(s)." if added else "Nothing new added (duplicates/unsupported).")

    def clear_list(self):
        self.files = []
        self.refresh_queue()
        self.log("Queue cleared.")

    def select_files(self):
        picked = filedialog.askopenfilenames(
            title="Select .cbr / .cbz files",
            filetypes=[("Comic archives", "*.cbr *.cbz *.rar *.zip"), ("All files", "*.*")]
        )
        if picked:
            self.add_paths([Path(p) for p in picked])

    def select_folder(self):
        folder = filedialog.askdirectory(title="Select folder (recursive search for CBZ/CBR)")
        if folder:
            self.add_paths([Path(folder)])

    def on_drop(self, event):
        items = split_dnd_list(event.data)
        self.add_paths([Path(i) for i in items])

    def start_convert_thread(self):
        if self.is_running:
            messagebox.showinfo("Running", "Conversion is already running.")
            return

        if not self.files:
            messagebox.showwarning("No files", "Add CBZ/CBR files (or a folder) first.")
            return

        if not (self.export_pdf.get() or self.export_epub.get()):
            messagebox.showwarning("Nothing selected", "Select PDF and/or EPUB export.")
            return

        if not self.seven_zip:
            messagebox.showerror("7-Zip missing", "7z.exe not found.")
            return

        out_dir = filedialog.askdirectory(title="Choose output folder for exports")
        if not out_dir:
            return

        # Reset progress
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.files)
        self.progress_label.config(text=f"Progress: 0 / {len(self.files)}")

        self.is_running = True
        self.out_btn.state(["disabled"])

        t = threading.Thread(target=self.convert_all, args=(Path(out_dir),), daemon=True)
        t.start()

    def convert_all(self, out_dir: Path):
        quality = int(self.quality.get())
        ok = 0
        fail = 0
        failures: list[tuple[str, str]] = []

        self.log("=" * 82)
        self.log(f"Start. Output: {out_dir}")
        self.log(f"Quality: {quality}")
        self.log(f"Export: PDF={self.export_pdf.get()} EPUB={self.export_epub.get()} (cover={self.epub_cover.get()})")
        self.log(f"Items: {len(self.files)}")
        self.log("=" * 82)

        for idx, archive in enumerate(self.files, start=1):
            # Update UI progress (start of item)
            self.root.after(
                0,
                lambda i=idx: (
                    self.progress.config(value=i - 1),
                    self.progress_label.config(text=f"Progress: {i - 1} / {len(self.files)}")
                )
            )

            try:
                self.log(f"[{idx}/{len(self.files)}] Extract: {archive.name}")
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

                    self.log(f"  Images: {len(imgs)} → JPEG (q={quality})")
                    jpegs = convert_to_jpegs(imgs, jpeg_dir, quality)

                    if self.export_pdf.get():
                        out_pdf = out_dir / (archive.stem + ".pdf")
                        self.log(f"  Build PDF: {out_pdf.name}")
                        images_to_pdf(jpegs, out_pdf)

                    if self.export_epub.get():
                        out_epub = out_dir / (archive.stem + ".epub")
                        self.log(f"  Build EPUB: {out_epub.name}")
                        build_epub_from_images(
                            jpegs=jpegs,
                            out_epub=out_epub,
                            title=archive.stem,
                            use_first_image_as_cover=self.epub_cover.get()
                        )

                self.log("  ✅ OK")
                ok += 1

                # Update UI progress (end of item)
                self.root.after(
                    0,
                    lambda i=idx: (
                        self.progress.config(value=i),
                        self.progress_label.config(text=f"Progress: {i} / {len(self.files)}")
                    )
                )

            except Exception as e:
                fail += 1
                failures.append((archive.name, str(e)))
                self.log(f"  ❌ FAIL: {e}")

                self.root.after(
                    0,
                    lambda i=idx: (
                        self.progress.config(value=i),
                        self.progress_label.config(text=f"Progress: {i} / {len(self.files)}")
                    )
                )

        self.log("-" * 82)
        self.log(f"Done. OK={ok}, FAIL={fail}")
        if fail:
            self.log("Failures summary:")
            for name, err in failures[:10]:
                self.log(f" - {name}: {err}")
            if len(failures) > 10:
                self.log(f" (+{len(failures) - 10} more)")
        self.log("-" * 82)

        def finish():
            self.is_running = False
            self.out_btn.state(["!disabled"])
            # Force 100%
            self.progress["value"] = self.progress["maximum"]
            self.progress_label.config(text=f"Progress: {self.progress['maximum']} / {self.progress['maximum']}")

            if fail == 0:
                messagebox.showinfo("Done", f"Converted {ok} file(s).\nOutput: {out_dir}")
            else:
                messagebox.showwarning("Partial success", f"OK={ok}, FAIL={fail}\nCheck Log for details.")
            self.set_info(f"Done. OK={ok}, FAIL={fail}. Quality: {quality}")

        self.root.after(0, finish)


def main():
    root = TkinterDnD.Tk()

    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
