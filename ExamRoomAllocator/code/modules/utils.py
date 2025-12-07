# modules/utils.py
import os
import zipfile
import shutil
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
import io
import logging

LOG = logging.getLogger(__name__)

MODULE_ROOT = Path(__file__).resolve().parent
ASSETS_DIR = MODULE_ROOT / "assets"
PLACEHOLDER = ASSETS_DIR / "nopic.jpg"


def ensure_assets_placeholder():
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    if PLACEHOLDER.exists():
        return str(PLACEHOLDER)

    try:
        img = Image.new("RGB", (400, 400), color=(245, 245, 245))
        draw = ImageDraw.Draw(img)
        try:
            fnt = ImageFont.load_default()
        except Exception:
            fnt = None
        text = "No Image Available"
        # Use textbbox for Pillow >=10
        try:
            bbox = draw.textbbox((0, 0), text, font=fnt)
            w = bbox[2] - bbox[0]
            h = bbox[3] - bbox[1]
        except Exception:
            # fallback
            w, h = draw.textsize(text, font=fnt) if hasattr(draw, "textsize") else (120, 14)
        draw.text(((400 - w) / 2, (400 - h) / 2), text, fill=(90, 90, 90), font=fnt)
        img.save(PLACEHOLDER, format="JPEG", quality=85)
        LOG.info("Created placeholder image at %s", PLACEHOLDER)
    except Exception:
        # If PIL operations fail, ensure an empty file exists so other code can reference it
        try:
            with open(PLACEHOLDER, "wb") as f:
                f.write(b"")
            LOG.warning("Created empty placeholder file at %s", PLACEHOLDER)
        except Exception:
            LOG.exception("Failed to create placeholder image; please create %s manually.", PLACEHOLDER)
    return str(PLACEHOLDER)


def safe_extract_photos_zip(uploaded_zip_file, dest_dir):
    Path(dest_dir).mkdir(parents=True, exist_ok=True)
    # read bytes
    if hasattr(uploaded_zip_file, "read"):
        data = uploaded_zip_file.read()
        bio = io.BytesIO(data)
    elif isinstance(uploaded_zip_file, (bytes, bytearray)):
        bio = io.BytesIO(uploaded_zip_file)
    else:
        bio = None

    if bio is None:
        # treat uploaded_zip_file as path
        with zipfile.ZipFile(uploaded_zip_file) as z:
            for member in z.namelist():
                p = Path(member)
                if p.is_absolute() or ".." in p.parts or member.endswith("/"):
                    continue
                outp = Path(dest_dir) / p.name
                with z.open(member) as src, open(outp, "wb") as dst:
                    shutil.copyfileobj(src, dst)
        return

    with zipfile.ZipFile(bio) as z:
        for member in z.namelist():
            p = Path(member)
            if p.is_absolute() or ".." in p.parts or member.endswith("/"):
                continue
            outp = Path(dest_dir) / p.name
            try:
                with z.open(member) as src, open(outp, "wb") as dst:
                    shutil.copyfileobj(src, dst)
            except RuntimeError:
                LOG.exception("Skipping member %s", member)


def setup_output_dir(tmpdir: str) -> str:
    out = Path(tmpdir) / "Output_Zip"
    out.mkdir(parents=True, exist_ok=True)
    return str(out)


def create_final_zip(output_root: str, output_name: str) -> str:
    out_root = Path(output_root)
    out_zip_path = out_root.parent / f"{output_name}.zip"
    shutil.make_archive(str(out_zip_path.with_suffix("")), 'zip', str(out_root))
    return str(out_zip_path)


def get_photo_for_roll(photos_dir: str, roll: str):
    # fallback to project-level /photos directory if photos_dir is empty
    if not photos_dir:
        MODULE_ROOT = Path(__file__).resolve().parent.parent   # go from modules/ → code/
        photos_dir = MODULE_ROOT.parent / "photos"

    placeholder = ensure_assets_placeholder()
    base = (str(roll) or "").strip()
    if not base:
        return placeholder

    photos_path = Path(photos_dir) if photos_dir else Path(".")
    # direct candidates
    for ext in ("jpg", "jpeg", "png"):
        p = photos_path / f"{base}.{ext}"
        if p.exists():
            return str(p)
        p_up = photos_path / f"{base}.{ext.upper()}"
        if p_up.exists():
            return str(p_up)

    # normalized search
    norm = base.replace(" ", "").replace(".", "").lower()
    for p in photos_path.glob("*"):
        if p.is_file() and p.suffix.lower() in {".jpg", ".jpeg", ".png"}:
            stem = p.stem.lower()
            if norm == stem or base.lower() == stem or base.lower() in p.name.lower():
                return str(p)

    return placeholder
