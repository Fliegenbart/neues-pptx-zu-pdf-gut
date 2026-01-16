"""
Docling Integration
===================
Integration von IBM's Docling für erweiterte Dokumentanalyse.

Features:
- VLM-basierte Alt-Text-Generierung (GraniteDocling)
- Reading Order Detection
- Tabellen-Strukturerkennung
- Layout-Analyse

DSGVO-konform: Alles läuft lokal.

Projekt: https://github.com/docling-project/docling
"""

import io
import logging
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Any
from enum import Enum

from .models import (
    SlideModel, Slide, Block, BlockType,
    Figure, Table, TableCell, Paragraph, TextRun
)

# Logger konfigurieren
logger = logging.getLogger(__name__)


# Lazy imports für optionale Docling-Abhängigkeit
_docling_available: Optional[bool] = None
_docling_version: Optional[str] = None


def _check_docling() -> bool:
    """Prüft ob Docling verfügbar ist."""
    global _docling_available, _docling_version
    if _docling_available is None:
        try:
            from docling.document_converter import DocumentConverter
            import docling
            _docling_available = True
            _docling_version = getattr(docling, "__version__", "unknown")
            logger.info(f"Docling {_docling_version} verfügbar")
        except ImportError as e:
            _docling_available = False
            logger.debug(f"Docling nicht verfügbar: {e}")
    return _docling_available


def get_docling_version() -> Optional[str]:
    """Gibt die Docling-Version zurück."""
    _check_docling()
    return _docling_version


class DoclingBackend(Enum):
    """Verfügbare Docling-Backends für Bildanalyse."""
    STANDARD = "standard"           # Docling Standard-Pipeline
    VLM = "vlm"                     # Visual Language Model (GraniteDocling)
    HYBRID = "hybrid"               # Standard + VLM für komplexe Bilder


@dataclass
class DoclingConfig:
    """Konfiguration für Docling-Integration."""
    # Backend-Auswahl
    backend: DoclingBackend = DoclingBackend.VLM

    # VLM-Einstellungen
    vlm_model: str = "granite_docling"  # oder: custom model path

    # Sprache für Ausgabe
    language: str = "de"

    # Feature Flags
    use_ocr: bool = True
    use_table_structure: bool = True
    use_reading_order: bool = True

    # Performance
    batch_size: int = 4
    timeout: int = 120


@dataclass
class DoclingAnalysisResult:
    """Ergebnis einer Docling-Analyse."""
    # Reading Order
    reading_order: list[dict] = field(default_factory=list)

    # Tabellen mit erkannter Struktur
    tables: list[dict] = field(default_factory=list)

    # Alt-Texte für Bilder
    image_descriptions: dict[str, str] = field(default_factory=dict)

    # Layout-Informationen
    layout_elements: list[dict] = field(default_factory=list)

    # Rohes DoclingDocument (für Debugging)
    raw_document: Optional[Any] = None


class DoclingAnalyzer:
    """
    Analysiert Dokumente mit Docling.

    Usage:
        analyzer = DoclingAnalyzer()
        if analyzer.is_available:
            result = analyzer.analyze_pptx("presentation.pptx")
    """

    def __init__(self, config: Optional[DoclingConfig] = None):
        self.config = config or DoclingConfig()
        self._converter = None
        self._vlm_pipeline = None

    @property
    def is_available(self) -> bool:
        """Prüft ob Docling verfügbar ist."""
        return _check_docling()

    def _get_converter(self, with_picture_description: bool = False):
        """Lazy-Load des DocumentConverter."""
        if self._converter is None and self.is_available:
            from docling.document_converter import DocumentConverter
            from docling.datamodel.pipeline_options import PipelineOptions

            # Pipeline-Optionen konfigurieren
            pipeline_options = PipelineOptions()
            pipeline_options.do_ocr = self.config.use_ocr
            pipeline_options.do_table_structure = self.config.use_table_structure

            # Bildbeschreibung aktivieren wenn gewünscht
            if with_picture_description:
                pipeline_options.do_picture_description = True
                # Prompt für barrierefreie Beschreibungen anpassen
                if self.config.language == "de":
                    prompt = "Beschreibe dieses Bild kurz für sehbehinderte Menschen. Maximal 2 Sätze."
                else:
                    prompt = "Describe this image briefly for visually impaired people. Maximum 2 sentences."
                pipeline_options.picture_description_options.prompt = prompt
                pipeline_options.picture_description_options.batch_size = self.config.batch_size
                logger.info("Bildbeschreibung aktiviert mit SmolVLM")

            self._converter = DocumentConverter(pipeline_options=pipeline_options)
            logger.info("DocumentConverter initialisiert")

        return self._converter

    def analyze_pptx(self, pptx_path: Path | str) -> Optional[DoclingAnalysisResult]:
        """
        Analysiert eine PPTX-Datei mit Docling.

        Args:
            pptx_path: Pfad zur PPTX-Datei

        Returns:
            DoclingAnalysisResult oder None bei Fehler
        """
        if not self.is_available:
            print("⚠️  Docling nicht installiert. Installiere mit: pip install pptx2ua[docling]")
            return None

        converter = self._get_converter()
        if converter is None:
            return None

        try:
            # Konvertiere PPTX zu DoclingDocument
            result = converter.convert(str(pptx_path))
            doc = result.document

            analysis = DoclingAnalysisResult(raw_document=doc)

            # Extrahiere Reading Order
            if self.config.use_reading_order:
                analysis.reading_order = self._extract_reading_order(doc)

            # Extrahiere Tabellen-Struktur
            if self.config.use_table_structure:
                analysis.tables = self._extract_tables(doc)

            # Extrahiere Layout-Elemente
            analysis.layout_elements = self._extract_layout(doc)

            return analysis

        except Exception as e:
            print(f"⚠️  Docling-Analyse fehlgeschlagen: {e}")
            return None

    def generate_alt_text(
        self,
        image_data: bytes,
        context: Optional[str] = None
    ) -> Optional[str]:
        """
        Generiert Alt-Text für ein Bild.

        HINWEIS: Docling ist primär für Dokumentanalyse optimiert,
        nicht für einzelne Bildbeschreibungen. Für beste Alt-Text-Ergebnisse
        wird Ollama mit llava/qwen2-vl empfohlen.

        Diese Methode ist ein Fallback wenn Ollama nicht verfügbar ist.

        Args:
            image_data: Bild als Bytes
            context: Optionaler Kontext (z.B. Folientitel)

        Returns:
            Alt-Text oder None bei Fehler
        """
        if not self.is_available:
            return None

        if self.config.backend == DoclingBackend.STANDARD:
            return None

        try:
            return self._vlm_describe_image(image_data, context)
        except Exception as e:
            logger.debug(f"Docling Alt-Text fehlgeschlagen: {e}")
            return None

    def _vlm_describe_image(
        self,
        image_data: bytes,
        context: Optional[str] = None
    ) -> Optional[str]:
        """
        Nutzt Docling VLM für Bildbeschreibung.

        Docling verwendet SmolVLM für Bildbeschreibungen, das bei der
        Dokumentkonvertierung automatisch angewendet wird.
        """
        try:
            from docling.document_converter import DocumentConverter
            from docling.datamodel.pipeline_options import VlmPipelineOptions
            from docling.datamodel.base_models import InputFormat

            # Temp-Datei für das Bild erstellen
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
                f.write(image_data)
                temp_path = Path(f.name)

            try:
                # VLM Pipeline-Optionen mit Bildbeschreibung
                vlm_options = VlmPipelineOptions()
                vlm_options.do_picture_description = True

                # Prompt für barrierefreie Beschreibungen
                if self.config.language == "de":
                    prompt = "Beschreibe dieses Bild kurz für sehbehinderte Menschen. Maximal 2 Sätze, beginne direkt mit dem Inhalt."
                else:
                    prompt = "Describe this image briefly for visually impaired people. Maximum 2 sentences, start directly with the content."

                if context:
                    prompt += f" Kontext: {context}" if self.config.language == "de" else f" Context: {context}"

                vlm_options.picture_description_options.prompt = prompt

                # Converter mit VLM-Optionen
                converter = DocumentConverter(
                    format_options={
                        InputFormat.IMAGE: vlm_options
                    }
                )

                # Bild als "Dokument" konvertieren
                result = converter.convert(str(temp_path))

                # Beschreibung aus dem Ergebnis extrahieren
                if result and result.document:
                    md = result.document.export_to_markdown()
                    if md and len(md.strip()) > 10:
                        description = self._polish_description(md.strip())
                        logger.debug(f"Alt-Text generiert: {description[:50]}...")
                        return description

            finally:
                # Aufräumen
                temp_path.unlink(missing_ok=True)

            return None

        except Exception as e:
            logger.warning(f"VLM Beschreibung fehlgeschlagen: {e}")
            return None

    def _get_german_prompt(self, context: Optional[str] = None) -> str:
        """Deutscher Prompt für Alt-Text-Generierung."""
        base = """Beschreibe dieses Bild für eine sehbehinderte Person.

Regeln:
- Maximal 2 Sätze
- Beschreibe WAS zu sehen ist und die Kernaussage
- Bei Diagrammen: Nenne Typ und Hauptaussage
- Bei Fotos: Beschreibe Motiv und Kontext
- Beginne direkt mit dem Inhalt, ohne "Das Bild zeigt"
"""
        if context:
            base += f"\nKontext: Dieses Bild ist Teil von: {context}"

        return base

    def _get_english_prompt(self, context: Optional[str] = None) -> str:
        """English prompt for alt-text generation."""
        base = """Describe this image for a visually impaired person.

Rules:
- Maximum 2 sentences
- Describe WHAT is shown and the key message
- For diagrams: State the type and main takeaway
- For photos: Describe subject and context
- Start directly with content, without "The image shows"
"""
        if context:
            base += f"\nContext: This image is part of: {context}"

        return base

    def _polish_description(self, text: str) -> str:
        """Bereinigt die generierte Beschreibung."""
        if not text:
            return ""

        text = text.strip()

        # Entferne typische Präfixe
        prefixes = [
            "Das Bild zeigt ", "Zu sehen ist ", "Die Abbildung zeigt ",
            "The image shows ", "This shows ", "We can see "
        ]

        for prefix in prefixes:
            if text.lower().startswith(prefix.lower()):
                text = text[len(prefix):]
                text = text[0].upper() + text[1:] if text else ""
                break

        # Punkt am Ende
        if text and text[-1] not in '.!?':
            text += '.'

        return text

    def _fallback_image_description(self, image_data: bytes) -> Optional[str]:
        """Fallback: Speichere Bild temporär und analysiere mit Docling."""
        try:
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
                f.write(image_data)
                temp_path = f.name

            converter = self._get_converter()
            if converter:
                result = converter.convert(temp_path)
                # Extrahiere Text/Beschreibung aus dem Ergebnis
                text = result.document.export_to_markdown()
                if text and len(text) > 10:
                    return text[:200]  # Beschränke auf 200 Zeichen

            Path(temp_path).unlink(missing_ok=True)
            return None

        except Exception:
            return None

    def _extract_reading_order(self, doc) -> list[dict]:
        """Extrahiert Lesereihenfolge aus DoclingDocument."""
        reading_order = []

        try:
            # DoclingDocument hat eine strukturierte Hierarchie
            for idx, element in enumerate(doc.iterate_elements()):
                item = {
                    "index": idx,
                    "type": element.element_type.value if hasattr(element, 'element_type') else "unknown",
                    "text_preview": str(element)[:100] if element else "",
                }

                # Bounding Box wenn verfügbar
                if hasattr(element, 'bbox') and element.bbox:
                    item["bbox"] = {
                        "x": element.bbox.x,
                        "y": element.bbox.y,
                        "width": element.bbox.width,
                        "height": element.bbox.height,
                    }

                reading_order.append(item)

        except Exception as e:
            print(f"⚠️  Reading Order Extraktion: {e}")

        return reading_order

    def _extract_tables(self, doc) -> list[dict]:
        """Extrahiert Tabellen mit Strukturinformationen."""
        tables = []

        try:
            for element in doc.iterate_elements():
                if hasattr(element, 'element_type'):
                    elem_type = str(element.element_type.value).lower()
                    if 'table' in elem_type:
                        table_info = {
                            "rows": [],
                            "has_header": False,
                            "caption": None,
                        }

                        # Versuche Tabellenstruktur zu extrahieren
                        if hasattr(element, 'table_data'):
                            td = element.table_data
                            if hasattr(td, 'rows'):
                                for row_idx, row in enumerate(td.rows):
                                    row_cells = []
                                    for cell in row.cells:
                                        cell_info = {
                                            "text": str(cell.text) if hasattr(cell, 'text') else "",
                                            "is_header": row_idx == 0 or getattr(cell, 'is_header', False),
                                            "colspan": getattr(cell, 'colspan', 1),
                                            "rowspan": getattr(cell, 'rowspan', 1),
                                        }
                                        row_cells.append(cell_info)
                                    table_info["rows"].append(row_cells)

                            # Header-Erkennung
                            if table_info["rows"]:
                                table_info["has_header"] = any(
                                    c.get("is_header", False)
                                    for c in table_info["rows"][0]
                                )

                        tables.append(table_info)

        except Exception as e:
            print(f"⚠️  Tabellen-Extraktion: {e}")

        return tables

    def _extract_layout(self, doc) -> list[dict]:
        """Extrahiert Layout-Informationen."""
        layout = []

        try:
            for element in doc.iterate_elements():
                elem_info = {
                    "type": str(element.element_type.value) if hasattr(element, 'element_type') else "unknown",
                }

                if hasattr(element, 'bbox') and element.bbox:
                    elem_info["bbox"] = {
                        "x": element.bbox.x,
                        "y": element.bbox.y,
                        "width": element.bbox.width,
                        "height": element.bbox.height,
                    }

                layout.append(elem_info)

        except Exception:
            pass

        return layout


class DoclingEnricher:
    """
    Alt-Text-Enricher basierend auf Docling.

    Drop-in Ersatz für den Ollama-basierten Enricher.
    """

    def __init__(self, config: Optional[DoclingConfig] = None):
        self.config = config or DoclingConfig()
        self.analyzer = DoclingAnalyzer(config)

        # Stats
        self.stats = {
            "processed": 0,
            "generated": 0,
            "failed": 0,
        }

    @property
    def is_available(self) -> bool:
        """Prüft ob Docling VLM verfügbar ist."""
        return self.analyzer.is_available

    def enrich(self, model: SlideModel, verbose: bool = True) -> SlideModel:
        """
        Reichert SlideModel mit Docling-generierten Alt-Texten an.
        """
        figures_to_process = model.figures_needing_alt_text

        if not figures_to_process:
            if verbose:
                print("   Alle Bilder haben bereits Alt-Texte")
            return model

        if verbose:
            print(f"   Generiere Alt-Texte für {len(figures_to_process)} Bilder mit Docling VLM...")

        for slide_num, figure in figures_to_process:
            self.stats["processed"] += 1

            # Kontext aus Folie holen
            slide = next((s for s in model.slides if s.number == slide_num), None)
            context = slide.title if slide else None

            # Alt-Text generieren
            alt_text = self.analyzer.generate_alt_text(
                figure.image_data,
                context=context
            )

            if alt_text:
                figure.alt_text = alt_text
                figure.needs_alt_text = False
                figure.alt_text_confidence = 0.85  # Docling VLM ist sehr zuverlässig
                self.stats["generated"] += 1

                if verbose:
                    preview = alt_text[:60] + "..." if len(alt_text) > 60 else alt_text
                    print(f"      Folie {slide_num}: \"{preview}\"")
            else:
                self.stats["failed"] += 1
                if verbose:
                    print(f"      Folie {slide_num}: Fehlgeschlagen")

        if verbose:
            self._print_stats()

        return model

    def _print_stats(self):
        """Gibt Statistiken aus."""
        print(f"\n      Docling VLM Statistik:")
        print(f"         Verarbeitet: {self.stats['processed']}")
        print(f"         Generiert: {self.stats['generated']}")
        print(f"         Fehlgeschlagen: {self.stats['failed']}")


def apply_docling_reading_order(
    model: SlideModel,
    analysis: DoclingAnalysisResult
) -> SlideModel:
    """
    Wendet Docling's Reading Order auf das SlideModel an.

    Args:
        model: Das zu aktualisierende SlideModel
        analysis: Docling-Analyseergebnis

    Returns:
        Aktualisiertes SlideModel
    """
    if not analysis.reading_order:
        return model

    # Mapping von Docling-Elementen zu SlideModel-Blöcken
    # basierend auf Position (BoundingBox)
    for slide in model.slides:
        docling_order = [
            item for item in analysis.reading_order
            # Filter relevante Elemente für diese Folie
            # (würde in Produktion über Page-Nummer gemacht)
        ]

        # Sortiere Blöcke nach Docling-Reihenfolge
        for idx, block in enumerate(slide.blocks):
            # Finde entsprechendes Docling-Element
            matching = _find_matching_docling_element(block, docling_order)
            if matching:
                block.reading_order = matching.get("index", idx)

    return model


def apply_docling_table_structure(
    model: SlideModel,
    analysis: DoclingAnalysisResult
) -> SlideModel:
    """
    Wendet Docling's Tabellen-Strukturerkennung an.

    Args:
        model: Das zu aktualisierende SlideModel
        analysis: Docling-Analyseergebnis

    Returns:
        Aktualisiertes SlideModel
    """
    if not analysis.tables:
        return model

    table_idx = 0
    for slide in model.slides:
        for block in slide.blocks:
            if block.table and table_idx < len(analysis.tables):
                docling_table = analysis.tables[table_idx]

                # Header-Erkennung übernehmen
                if docling_table.get("has_header") and block.table.rows:
                    for cell in block.table.rows[0]:
                        cell.is_header = True

                # Colspan/Rowspan übernehmen wenn verfügbar
                for row_idx, row in enumerate(docling_table.get("rows", [])):
                    if row_idx < len(block.table.rows):
                        for cell_idx, cell_info in enumerate(row):
                            if cell_idx < len(block.table.rows[row_idx]):
                                cell = block.table.rows[row_idx][cell_idx]
                                cell.colspan = cell_info.get("colspan", 1)
                                cell.rowspan = cell_info.get("rowspan", 1)

                table_idx += 1

    return model


def _find_matching_docling_element(
    block: Block,
    docling_elements: list[dict]
) -> Optional[dict]:
    """Findet passendes Docling-Element basierend auf BoundingBox."""
    if not block.bbox:
        return None

    best_match = None
    best_overlap = 0

    for elem in docling_elements:
        if "bbox" not in elem:
            continue

        # Berechne Überlappung
        overlap = _calculate_bbox_overlap(
            block.bbox,
            elem["bbox"]
        )

        if overlap > best_overlap:
            best_overlap = overlap
            best_match = elem

    return best_match if best_overlap > 0.5 else None


def _calculate_bbox_overlap(bbox1, bbox2: dict) -> float:
    """Berechnet Überlappungsanteil zweier Bounding Boxes."""
    # bbox1 ist BoundingBox-Objekt, bbox2 ist dict
    x1 = max(bbox1.x, bbox2["x"])
    y1 = max(bbox1.y, bbox2["y"])
    x2 = min(bbox1.x + bbox1.width, bbox2["x"] + bbox2["width"])
    y2 = min(bbox1.y + bbox1.height, bbox2["y"] + bbox2["height"])

    if x2 <= x1 or y2 <= y1:
        return 0.0

    intersection = (x2 - x1) * (y2 - y1)
    area1 = bbox1.width * bbox1.height

    if area1 == 0:
        return 0.0

    return intersection / area1


# === Convenience Functions ===

def is_docling_available() -> bool:
    """Prüft ob Docling verfügbar ist."""
    return _check_docling()


def analyze_with_docling(
    pptx_path: Path | str,
    config: Optional[DoclingConfig] = None
) -> Optional[DoclingAnalysisResult]:
    """
    Convenience-Funktion für Docling-Analyse.

    Usage:
        result = analyze_with_docling("presentation.pptx")
        if result:
            print(f"Reading order: {len(result.reading_order)} elements")
    """
    analyzer = DoclingAnalyzer(config)
    return analyzer.analyze_pptx(pptx_path)


def enrich_with_docling(
    model: SlideModel,
    config: Optional[DoclingConfig] = None,
    verbose: bool = True
) -> SlideModel:
    """
    Convenience-Funktion für Docling-basierte Alt-Text-Generierung.

    Usage:
        model = parser.parse("slides.pptx")
        model = enrich_with_docling(model)
    """
    enricher = DoclingEnricher(config)

    if not enricher.is_available:
        if verbose:
            print("⚠️  Docling nicht verfügbar, überspringe Alt-Text-Generierung")
        return model

    return enricher.enrich(model, verbose=verbose)
