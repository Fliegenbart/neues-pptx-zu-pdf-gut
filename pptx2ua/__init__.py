"""
PPTX to PDF/UA Converter
========================
DSGVO-konforme Konvertierung von PowerPoint zu barrierefreien PDFs.

Architektur:
    PPTX → SlideModel (JSON) → Semantic HTML → WeasyPrint → PDF/UA → veraPDF

Module:
    - parser: PPTX → SlideModel
    - enricher: KI-basierte Alt-Texte (Ollama oder Docling)
    - docling_integration: IBM Docling für Reading Order & Tabellen
    - accessibility_optimizer: Screenreader-UX-Optimierung
    - renderer: SlideModel → HTML → PDF
    - validator: PDF/UA Validierung mit veraPDF

Quick Start:
    >>> from pptx2ua import PPTXParser, PDFUARenderer, AccessibilityOptimizer
    >>> model = PPTXParser().parse("slides.pptx")
    >>> model = AccessibilityOptimizer().optimize(model)
    >>> PDFUARenderer().render(model, "output.pdf")

Mit Docling (optional):
    >>> from pptx2ua import Enricher, EnricherConfig, EnricherBackend
    >>> config = EnricherConfig(backend=EnricherBackend.DOCLING)
    >>> enricher = Enricher(config)
    >>> model = enricher.enrich(model)
"""

__version__ = "0.1.0"

from .models import (
    SlideModel,
    Slide,
    Block,
    BlockType,
    Paragraph,
    TextRun,
    Figure,
    Table,
    TableCell,
)
from .parser import PPTXParser
from .renderer import PDFUARenderer, RendererConfig, HTMLGenerator
from .validator import PDFUAValidator, ValidationResult
from .enricher import Enricher, EnricherConfig, EnricherBackend
from .accessibility_optimizer import (
    AccessibilityOptimizer,
    AccessibilityConfig,
    ElementRole,
    optimize_for_screenreader,
)

# Optionale Docling-Integration (nur wenn installiert)
try:
    from .docling_integration import (
        DoclingAnalyzer,
        DoclingConfig,
        DoclingEnricher,
        is_docling_available,
    )
    _docling_available = True
except ImportError:
    _docling_available = False

__all__ = [
    # Version
    "__version__",

    # Models
    "SlideModel",
    "Slide",
    "Block",
    "BlockType",
    "Paragraph",
    "TextRun",
    "Figure",
    "Table",
    "TableCell",

    # Parser
    "PPTXParser",

    # Enricher
    "Enricher",
    "EnricherConfig",
    "EnricherBackend",

    # Accessibility
    "AccessibilityOptimizer",
    "AccessibilityConfig",
    "ElementRole",
    "optimize_for_screenreader",

    # Renderer
    "PDFUARenderer",
    "RendererConfig",
    "HTMLGenerator",

    # Validator
    "PDFUAValidator",
    "ValidationResult",
]

# Docling-Exports nur wenn verfügbar
if _docling_available:
    __all__.extend([
        "DoclingAnalyzer",
        "DoclingConfig",
        "DoclingEnricher",
        "is_docling_available",
    ])
