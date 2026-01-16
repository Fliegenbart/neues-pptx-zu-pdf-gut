"""
Semantische Datenmodelle für Slides.

Das SlideModel ist das Herzstück der Pipeline:
PPTX → SlideModel → HTML → PDF/UA

Alle Strukturinformationen werden hier normalisiert,
bevor sie zu HTML/PDF gerendert werden.
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional
from pathlib import Path


class BlockType(Enum):
    """Semantische Block-Typen nach PDF/UA."""
    HEADING = "heading"
    PARAGRAPH = "paragraph"
    LIST = "list"
    LIST_ITEM = "list_item"
    TABLE = "table"
    FIGURE = "figure"
    QUOTE = "quote"
    CODE = "code"


class ListStyle(Enum):
    """Listen-Stile."""
    BULLET = "bullet"
    NUMBERED = "numbered"
    NONE = "none"


@dataclass
class BoundingBox:
    """Position und Größe eines Elements."""
    x: float  # in mm
    y: float
    width: float
    height: float
    
    def __post_init__(self):
        # Normalisiere negative Werte
        self.width = abs(self.width)
        self.height = abs(self.height)


@dataclass
class TextRun:
    """Ein Textabschnitt mit einheitlicher Formatierung."""
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    font_size: Optional[float] = None  # in pt
    font_name: Optional[str] = None
    color: Optional[str] = None  # Hex: "FF0000"
    hyperlink: Optional[str] = None


@dataclass
class Paragraph:
    """Ein Absatz bestehend aus TextRuns."""
    runs: list[TextRun] = field(default_factory=list)
    alignment: str = "left"  # left, center, right, justify
    level: int = 0  # Einrückungsebene (für Listen)
    
    @property
    def text(self) -> str:
        """Gesamter Text ohne Formatierung."""
        return "".join(run.text for run in self.runs)
    
    @property
    def is_empty(self) -> bool:
        return not self.text.strip()


@dataclass
class TableCell:
    """Eine Tabellenzelle."""
    paragraphs: list[Paragraph] = field(default_factory=list)
    colspan: int = 1
    rowspan: int = 1
    is_header: bool = False
    
    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.paragraphs)


@dataclass
class Table:
    """Eine Tabelle mit Zeilen und Zellen."""
    rows: list[list[TableCell]] = field(default_factory=list)
    caption: Optional[str] = None
    
    @property
    def has_header(self) -> bool:
        """Prüft ob erste Zeile Header-Zellen enthält."""
        if not self.rows:
            return False
        return any(cell.is_header for cell in self.rows[0])
    
    @property
    def column_count(self) -> int:
        if not self.rows:
            return 0
        return max(
            sum(cell.colspan for cell in row)
            for row in self.rows
        )


@dataclass
class Figure:
    """Eine Abbildung (Bild, Chart, Diagramm)."""
    image_path: Optional[Path] = None
    image_data: Optional[bytes] = None
    mime_type: str = "image/png"
    alt_text: Optional[str] = None
    long_description: Optional[str] = None
    caption: Optional[str] = None
    
    # Metadaten für KI-Verarbeitung
    needs_alt_text: bool = True
    alt_text_confidence: float = 0.0
    image_hash: Optional[str] = None  # Für Caching


@dataclass
class Block:
    """
    Ein semantischer Block auf einer Folie.
    
    Kann verschiedene Inhaltstypen haben:
    - Text (paragraphs)
    - Tabelle (table)
    - Abbildung (figure)
    """
    block_type: BlockType
    reading_order: int
    bbox: Optional[BoundingBox] = None
    
    # Inhalt (je nach block_type)
    paragraphs: list[Paragraph] = field(default_factory=list)
    table: Optional[Table] = None
    figure: Optional[Figure] = None
    
    # Struktur
    heading_level: int = 1  # 1-6 für Überschriften
    list_style: ListStyle = ListStyle.NONE
    
    # Metadaten
    source_shape_id: Optional[str] = None
    confidence: float = 1.0  # Wie sicher ist die Typ-Erkennung?
    
    @property
    def text(self) -> str:
        """Gesamter Text des Blocks."""
        if self.paragraphs:
            return "\n".join(p.text for p in self.paragraphs)
        if self.table:
            return self.table.caption or ""
        if self.figure:
            return self.figure.alt_text or ""
        return ""
    
    @property
    def is_empty(self) -> bool:
        if self.paragraphs:
            return all(p.is_empty for p in self.paragraphs)
        if self.table:
            return len(self.table.rows) == 0
        if self.figure:
            return self.figure.image_data is None and self.figure.image_path is None
        return True


@dataclass
class Slide:
    """Eine einzelne Folie."""
    number: int
    blocks: list[Block] = field(default_factory=list)
    notes: Optional[str] = None
    
    # Layout-Metadaten
    width_mm: float = 254.0   # 16:9 Standard
    height_mm: float = 142.9
    background_color: str = "FFFFFF"
    
    @property
    def title(self) -> Optional[str]:
        """Findet den Titel der Folie (erstes Heading)."""
        for block in self.sorted_blocks:
            if block.block_type == BlockType.HEADING:
                return block.text
        return None
    
    @property
    def sorted_blocks(self) -> list[Block]:
        """Blöcke in Lesereihenfolge."""
        return sorted(self.blocks, key=lambda b: b.reading_order)
    
    @property
    def figures(self) -> list[Figure]:
        """Alle Abbildungen auf der Folie."""
        return [
            block.figure for block in self.blocks 
            if block.figure is not None
        ]
    
    @property
    def figures_without_alt(self) -> list[Figure]:
        """Abbildungen die noch Alt-Text brauchen."""
        return [
            fig for fig in self.figures
            if fig.needs_alt_text and not fig.alt_text
        ]


@dataclass 
class SlideModel:
    """
    Komplettes Präsentationsmodell.
    
    Das ist das zentrale Datenformat das durch die 
    gesamte Pipeline fließt:
    
    PPTX → [Parser] → SlideModel → [Enricher] → SlideModel → [Renderer] → PDF
    """
    slides: list[Slide] = field(default_factory=list)
    
    # Dokument-Metadaten
    title: Optional[str] = None
    author: Optional[str] = None
    language: str = "de"
    subject: Optional[str] = None
    keywords: list[str] = field(default_factory=list)
    
    # Verarbeitungs-Metadaten
    source_file: Optional[Path] = None
    created_at: Optional[str] = None
    
    @property
    def slide_count(self) -> int:
        return len(self.slides)
    
    @property
    def all_figures(self) -> list[Figure]:
        """Alle Abbildungen im Dokument."""
        figures = []
        for slide in self.slides:
            figures.extend(slide.figures)
        return figures
    
    @property
    def figures_needing_alt_text(self) -> list[tuple[int, Figure]]:
        """Abbildungen ohne Alt-Text mit Foliennummer."""
        result = []
        for slide in self.slides:
            for fig in slide.figures_without_alt:
                result.append((slide.number, fig))
        return result
    
    def to_dict(self) -> dict:
        """Serialisiert zu Dictionary (für JSON-Export)."""
        # TODO: Implementieren für Debug/Inspection
        pass
    
    @classmethod
    def from_dict(cls, data: dict) -> "SlideModel":
        """Deserialisiert von Dictionary."""
        # TODO: Implementieren für Import
        pass
