"""
Slide Image Renderer
====================

Rendert PPTX-Folien als Bilder für Vision-LLM Analyse.

Nutzt LibreOffice im Headless-Modus für hochwertige Konvertierung.
Fallback: Extrahiert Thumbnails aus der PPTX wenn vorhanden.
"""

import subprocess
import tempfile
import shutil
from pathlib import Path
from typing import Optional
import zipfile
import io

from .models import SlideModel, Slide


def is_libreoffice_available() -> bool:
    """Prüft ob LibreOffice installiert ist."""
    # macOS
    if Path("/Applications/LibreOffice.app").exists():
        return True
    # Linux
    if shutil.which("libreoffice") or shutil.which("soffice"):
        return True
    return False


def get_libreoffice_command() -> Optional[str]:
    """Gibt den LibreOffice-Befehl zurück."""
    # macOS
    macos_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if Path(macos_path).exists():
        return macos_path
    # Linux
    if shutil.which("libreoffice"):
        return "libreoffice"
    if shutil.which("soffice"):
        return "soffice"
    return None


def render_slides_to_images(
    pptx_path: Path,
    output_dir: Optional[Path] = None,
    dpi: int = 150,
    format: str = "png"
) -> list[Path]:
    """
    Rendert alle Folien einer PPTX als Bilder.

    Args:
        pptx_path: Pfad zur PPTX-Datei
        output_dir: Ausgabeverzeichnis (oder temp)
        dpi: Auflösung (150 ist gut für Vision-LLMs)
        format: Bildformat (png, jpg)

    Returns:
        Liste der Bildpfade, sortiert nach Foliennummer
    """
    soffice = get_libreoffice_command()
    if not soffice:
        print("   ⚠️  LibreOffice nicht gefunden - Folienbilder nicht verfügbar")
        return []

    # Temporäres Verzeichnis wenn keins angegeben
    if output_dir is None:
        output_dir = Path(tempfile.mkdtemp(prefix="pptx2ua_slides_"))

    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        # LibreOffice Headless-Konvertierung
        # Konvertiert zu PDF erst, dann zu Bildern
        cmd = [
            soffice,
            "--headless",
            "--convert-to", format,
            "--outdir", str(output_dir),
            str(pptx_path)
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120  # 2 Minuten Timeout
        )

        if result.returncode != 0:
            print(f"   ⚠️  LibreOffice Fehler: {result.stderr}")
            return []

        # Finde generierte Bilder
        # LibreOffice benennt sie: presentation-1.png, presentation-2.png, etc.
        # Oder einfach: presentation.png für einzelne Folie

        image_files = sorted(
            output_dir.glob(f"*.{format}"),
            key=lambda p: _extract_slide_number(p.name)
        )

        return image_files

    except subprocess.TimeoutExpired:
        print("   ⚠️  LibreOffice Timeout")
        return []
    except Exception as e:
        print(f"   ⚠️  Render-Fehler: {e}")
        return []


def render_pptx_via_pdf(
    pptx_path: Path,
    output_dir: Optional[Path] = None,
    dpi: int = 150
) -> list[Path]:
    """
    Alternative Methode: PPTX → PDF → PNG.

    Nützlich wenn direkte PNG-Konvertierung nicht funktioniert.
    """
    soffice = get_libreoffice_command()
    if not soffice:
        return []

    if output_dir is None:
        output_dir = Path(tempfile.mkdtemp(prefix="pptx2ua_slides_"))

    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        # Schritt 1: PPTX → PDF
        pdf_cmd = [
            soffice,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(pptx_path)
        ]

        subprocess.run(pdf_cmd, capture_output=True, timeout=120)

        pdf_path = output_dir / f"{pptx_path.stem}.pdf"
        if not pdf_path.exists():
            return []

        # Schritt 2: PDF → PNGs mit pdftoppm (wenn verfügbar)
        if shutil.which("pdftoppm"):
            img_prefix = output_dir / "slide"
            png_cmd = [
                "pdftoppm",
                "-png",
                "-r", str(dpi),
                str(pdf_path),
                str(img_prefix)
            ]
            subprocess.run(png_cmd, capture_output=True, timeout=60)

            return sorted(output_dir.glob("slide-*.png"))

        # Fallback: pdf2image Python-Bibliothek
        try:
            from pdf2image import convert_from_path
            images = convert_from_path(pdf_path, dpi=dpi)

            image_paths = []
            for i, img in enumerate(images, 1):
                img_path = output_dir / f"slide-{i:03d}.png"
                img.save(img_path, "PNG")
                image_paths.append(img_path)

            return image_paths
        except ImportError:
            pass

        return []

    except Exception as e:
        print(f"   ⚠️  PDF-Render-Fehler: {e}")
        return []


def extract_pptx_thumbnails(pptx_path: Path) -> list[bytes]:
    """
    Extrahiert eingebettete Thumbnails aus der PPTX.

    PowerPoint speichert manchmal Slide-Thumbnails in:
    - docProps/thumbnail.jpeg (nur Titelbild)
    - ppt/media/ (alle Medien)

    Returns:
        Liste von Bild-Bytes (oft leer, da PPTX keine Slide-Thumbnails speichert)
    """
    thumbnails = []

    try:
        with zipfile.ZipFile(pptx_path, 'r') as zf:
            # Titelbild
            if 'docProps/thumbnail.jpeg' in zf.namelist():
                thumbnails.append(zf.read('docProps/thumbnail.jpeg'))

    except Exception as e:
        print(f"   Thumbnail-Extraktion fehlgeschlagen: {e}")

    return thumbnails


def populate_slide_images(model: SlideModel, pptx_path: Path) -> bool:
    """
    Rendert alle Folien und speichert die Bilder im SlideModel.

    Args:
        model: Das SlideModel das erweitert wird
        pptx_path: Pfad zur Original-PPTX

    Returns:
        True wenn erfolgreich
    """
    # Versuche LibreOffice Rendering
    with tempfile.TemporaryDirectory(prefix="pptx2ua_") as tmpdir:
        output_dir = Path(tmpdir)

        # Methode 1: Direkte PNG-Konvertierung
        image_paths = render_slides_to_images(pptx_path, output_dir)

        # Methode 2: Über PDF wenn direkt nicht klappt
        if not image_paths:
            image_paths = render_pptx_via_pdf(pptx_path, output_dir)

        if not image_paths:
            return False

        # Ordne Bilder den Folien zu
        for slide in model.slides:
            if slide.number <= len(image_paths):
                img_path = image_paths[slide.number - 1]
                slide.slide_image = img_path.read_bytes()

        return True


def _extract_slide_number(filename: str) -> int:
    """Extrahiert Foliennummer aus Dateinamen."""
    import re
    match = re.search(r'(\d+)', filename)
    if match:
        return int(match.group(1))
    return 0


# === Convenience für einzelne Folien ===

def render_single_slide(
    pptx_path: Path,
    slide_number: int,
    dpi: int = 150
) -> Optional[bytes]:
    """
    Rendert eine einzelne Folie als PNG.

    Args:
        pptx_path: Pfad zur PPTX
        slide_number: 1-basierte Foliennummer
        dpi: Auflösung

    Returns:
        PNG-Bytes oder None
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        output_dir = Path(tmpdir)
        images = render_pptx_via_pdf(pptx_path, output_dir, dpi)

        if slide_number <= len(images):
            return images[slide_number - 1].read_bytes()

    return None
