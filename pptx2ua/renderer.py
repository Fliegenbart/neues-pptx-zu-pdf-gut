"""
PDF/UA Renderer
===============
Konvertiert SlideModel → HTML → PDF/UA

Pipeline:
1. SlideModel → Semantisches HTML
2. HTML → WeasyPrint → Tagged PDF
3. pikepdf → PDF/UA Compliance patchen
"""

import base64
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from datetime import datetime
import html

from weasyprint import HTML, CSS
from weasyprint.text.fonts import FontConfiguration

from .models import (
    SlideModel, Slide, Block, BlockType,
    Paragraph, TextRun, Table, Figure, ListStyle
)


@dataclass
class RendererConfig:
    """Konfiguration für den Renderer."""
    # Seitenformat
    page_width_mm: float = 297.0   # A4 Landscape
    page_height_mm: float = 210.0
    
    # Margins
    margin_top_mm: float = 15.0
    margin_bottom_mm: float = 15.0
    margin_left_mm: float = 20.0
    margin_right_mm: float = 20.0
    
    # Styling
    font_family: str = "Liberation Sans, Arial, sans-serif"
    base_font_size_pt: float = 11.0
    heading_scale: float = 1.4
    
    # PDF Optionen
    pdf_version: str = "1.7"
    embed_fonts: bool = True


class HTMLGenerator:
    """
    Generiert semantisches HTML aus SlideModel.
    
    Das HTML ist optimiert für:
    - Screenreader (ARIA, semantische Tags)
    - WeasyPrint Rendering
    - PDF/UA Tag-Mapping
    """
    
    def __init__(self, config: RendererConfig):
        self.config = config
    
    def generate(self, model: SlideModel) -> str:
        """Generiert komplettes HTML-Dokument."""
        slides_html = "\n".join(
            self._render_slide(slide) 
            for slide in model.slides
        )
        
        return f"""<!DOCTYPE html>
<html lang="{model.language}">
<head>
    <meta charset="UTF-8">
    <title>{html.escape(model.title or 'Präsentation')}</title>
    <meta name="author" content="{html.escape(model.author or '')}">
    <meta name="subject" content="{html.escape(model.subject or '')}">
    <meta name="generator" content="pptx2ua">
    <style>
{self._generate_css()}
    </style>
</head>
<body>
{slides_html}
</body>
</html>"""
    
    def _generate_css(self) -> str:
        """Generiert CSS für PDF-Rendering."""
        cfg = self.config
        
        return f"""
/* Page Setup */
@page {{
    size: {cfg.page_width_mm}mm {cfg.page_height_mm}mm;
    margin: {cfg.margin_top_mm}mm {cfg.margin_right_mm}mm {cfg.margin_bottom_mm}mm {cfg.margin_left_mm}mm;
    
    @bottom-center {{
        content: "Seite " counter(page) " von " counter(pages);
        font-size: 9pt;
        color: #666;
    }}
}}

/* Slide als Section */
section.slide {{
    page-break-after: always;
    page-break-inside: avoid;
}}

section.slide:last-child {{
    page-break-after: auto;
}}

/* Base Typography */
body {{
    font-family: {cfg.font_family};
    font-size: {cfg.base_font_size_pt}pt;
    line-height: 1.5;
    color: #1a1a1a;
}}

/* Headings */
h1 {{
    font-size: {cfg.base_font_size_pt * cfg.heading_scale ** 3}pt;
    font-weight: bold;
    margin: 0 0 0.5em 0;
    color: #003366;
}}

h2 {{
    font-size: {cfg.base_font_size_pt * cfg.heading_scale ** 2}pt;
    font-weight: bold;
    margin: 1em 0 0.4em 0;
    color: #004080;
}}

h3 {{
    font-size: {cfg.base_font_size_pt * cfg.heading_scale}pt;
    font-weight: bold;
    margin: 0.8em 0 0.3em 0;
}}

h4, h5, h6 {{
    font-size: {cfg.base_font_size_pt}pt;
    font-weight: bold;
    margin: 0.6em 0 0.2em 0;
}}

/* Paragraphs */
p {{
    margin: 0 0 0.8em 0;
    orphans: 2;
    widows: 2;
}}

/* Lists */
ul, ol {{
    margin: 0 0 1em 0;
    padding-left: 1.5em;
}}

li {{
    margin-bottom: 0.3em;
}}

/* Tables */
table {{
    width: 100%;
    border-collapse: collapse;
    margin: 1em 0;
    font-size: {cfg.base_font_size_pt * 0.9}pt;
}}

th, td {{
    border: 1px solid #ccc;
    padding: 0.5em 0.8em;
    text-align: left;
    vertical-align: top;
}}

th {{
    background-color: #f0f0f0;
    font-weight: bold;
}}

caption {{
    caption-side: top;
    font-weight: bold;
    margin-bottom: 0.5em;
    text-align: left;
}}

/* Figures */
figure {{
    margin: 1em 0;
    page-break-inside: avoid;
}}

figure img {{
    max-width: 100%;
    height: auto;
}}

figcaption {{
    font-size: {cfg.base_font_size_pt * 0.85}pt;
    color: #555;
    margin-top: 0.5em;
    font-style: italic;
}}

/* Text Formatting */
strong {{ font-weight: bold; }}
em {{ font-style: italic; }}
u {{ text-decoration: underline; }}

/* Links */
a {{
    color: #0066cc;
    text-decoration: underline;
}}

/* Slide Number */
.slide-number {{
    font-size: 9pt;
    color: #888;
    text-align: right;
    margin-bottom: 0.5em;
}}

/* Accessibility: Skip Links (hidden but accessible) */
.sr-only {{
    position: absolute;
    width: 1px;
    height: 1px;
    padding: 0;
    margin: -1px;
    overflow: hidden;
    clip: rect(0, 0, 0, 0);
    border: 0;
}}

/* Screenreader-Zusammenfassung für Tabellen */
.sr-summary {{
    font-style: italic;
    color: #555;
    margin-bottom: 0.5em;
    padding: 0.5em;
    background: #f9f9f9;
    border-left: 3px solid #0066cc;
}}

/* Dekorative Elemente ausblenden */
.decorative {{
    display: none;
}}

/* Accessible Table Container */
.accessible-table {{
    margin: 1em 0;
}}
"""
    
    def _render_slide(self, slide: Slide) -> str:
        """Rendert eine Folie als HTML-Section."""
        blocks_html = "\n".join(
            self._render_block(block)
            for block in slide.sorted_blocks
        )
        
        # ARIA-Label für Screenreader
        slide_label = f"Folie {slide.number}"
        if slide.title:
            slide_label += f": {slide.title}"
        
        return f"""
<section class="slide" role="region" aria-label="{html.escape(slide_label)}">
    <div class="slide-number" aria-hidden="true">Folie {slide.number}</div>
{blocks_html}
</section>"""
    
    def _render_block(self, block: Block) -> str:
        """Rendert einen Block basierend auf seinem Typ."""
        
        # Prüfe Accessibility-Annotationen
        if hasattr(block, 'a11y'):
            from .accessibility_optimizer import ElementRole
            
            # Dekorative Elemente: aria-hidden
            if block.a11y.role == ElementRole.DECORATIVE:
                return f'    <div aria-hidden="true" class="decorative"><!-- {block.a11y.skip_reason or "Dekorativ"} --></div>'
            
            # Redundante Elemente: überspringen
            if block.a11y.role == ElementRole.REDUNDANT:
                return f'    <!-- Redundant: {block.a11y.skip_reason or "Bereits vorgelesen"} -->'
            
            # Optimierter Screenreader-Text vorhanden?
            if block.a11y.screen_reader_text:
                sr_text = html.escape(block.a11y.screen_reader_text)
                
                # Bei Tabellen: Sowohl visuelle Tabelle als auch SR-Text
                if block.block_type == BlockType.TABLE and block.table:
                    visual_table = self._render_table(block.table)
                    return f'''    <div class="accessible-table">
        <p class="sr-summary">{sr_text}</p>
{visual_table}
    </div>'''
                
                # Bei Figures: SR-Text als verbesserter Alt-Text
                if block.block_type == BlockType.FIGURE and block.figure:
                    block.figure.alt_text = block.a11y.screen_reader_text
        
        # Standard-Rendering
        if block.block_type == BlockType.HEADING:
            return self._render_heading(block)
        
        if block.block_type == BlockType.PARAGRAPH:
            return self._render_paragraphs(block.paragraphs)
        
        if block.block_type == BlockType.LIST:
            return self._render_list(block)
        
        if block.block_type == BlockType.TABLE and block.table:
            return self._render_table(block.table)
        
        if block.block_type == BlockType.FIGURE and block.figure:
            return self._render_figure(block.figure)
        
        # Fallback
        return self._render_paragraphs(block.paragraphs)
    
    def _render_heading(self, block: Block) -> str:
        """Rendert eine Überschrift."""
        level = min(block.heading_level, 6)
        text = html.escape(block.text)
        return f"    <h{level}>{text}</h{level}>"
    
    def _render_paragraphs(self, paragraphs: list[Paragraph]) -> str:
        """Rendert Absätze."""
        result = []
        for para in paragraphs:
            if para.is_empty:
                continue
            
            runs_html = "".join(self._render_run(run) for run in para.runs)
            
            # Alignment als Style
            style = ""
            if para.alignment != "left":
                style = f' style="text-align: {para.alignment}"'
            
            result.append(f"    <p{style}>{runs_html}</p>")
        
        return "\n".join(result)
    
    def _render_run(self, run: TextRun) -> str:
        """Rendert einen TextRun mit Formatierung."""
        text = html.escape(run.text)
        
        # Formatierungen verschachteln
        if run.bold:
            text = f"<strong>{text}</strong>"
        if run.italic:
            text = f"<em>{text}</em>"
        if run.underline:
            text = f"<u>{text}</u>"
        if run.hyperlink:
            text = f'<a href="{html.escape(run.hyperlink)}">{text}</a>'
        
        return text
    
    def _render_list(self, block: Block) -> str:
        """Rendert eine Liste."""
        tag = "ul" if block.list_style == ListStyle.BULLET else "ol"
        
        items = []
        for para in block.paragraphs:
            if para.is_empty:
                continue
            runs_html = "".join(self._render_run(run) for run in para.runs)
            items.append(f"        <li>{runs_html}</li>")
        
        items_html = "\n".join(items)
        return f"    <{tag}>\n{items_html}\n    </{tag}>"
    
    def _render_table(self, table: Table) -> str:
        """Rendert eine Tabelle."""
        rows_html = []
        
        for row_idx, row in enumerate(table.rows):
            cells_html = []
            
            for cell in row:
                tag = "th" if cell.is_header else "td"
                
                # Zell-Inhalt
                content = ""
                for para in cell.paragraphs:
                    runs = "".join(self._render_run(run) for run in para.runs)
                    content += runs
                
                # Colspan/Rowspan
                attrs = ""
                if cell.colspan > 1:
                    attrs += f' colspan="{cell.colspan}"'
                if cell.rowspan > 1:
                    attrs += f' rowspan="{cell.rowspan}"'
                
                # Scope für Header
                if cell.is_header:
                    attrs += ' scope="col"'
                
                cells_html.append(f"            <{tag}{attrs}>{content}</{tag}>")
            
            row_html = "\n".join(cells_html)
            rows_html.append(f"        <tr>\n{row_html}\n        </tr>")
        
        # Thead/Tbody Trennung
        if table.has_header:
            thead = f"        <thead>\n{rows_html[0]}\n        </thead>"
            tbody = "        <tbody>\n" + "\n".join(rows_html[1:]) + "\n        </tbody>"
            body = f"{thead}\n{tbody}"
        else:
            body = "\n".join(rows_html)
        
        # Caption
        caption = ""
        if table.caption:
            caption = f"        <caption>{html.escape(table.caption)}</caption>\n"
        
        return f"""    <table>
{caption}{body}
    </table>"""
    
    def _render_figure(self, figure: Figure) -> str:
        """Rendert eine Abbildung."""
        # Alt-Text (Pflicht für Barrierefreiheit!)
        alt = html.escape(figure.alt_text or "Bild ohne Beschreibung")
        
        # Bild als Base64 Data-URI
        if figure.image_data:
            b64 = base64.b64encode(figure.image_data).decode('utf-8')
            src = f"data:{figure.mime_type};base64,{b64}"
        elif figure.image_path:
            # Externe Referenz
            src = str(figure.image_path)
        else:
            # Placeholder
            src = ""
        
        # Figcaption
        caption_html = ""
        if figure.caption:
            caption_html = f"\n        <figcaption>{html.escape(figure.caption)}</figcaption>"
        
        # Long Description als aria-describedby
        # (könnte auch als versteckter Text implementiert werden)
        
        return f"""    <figure>
        <img src="{src}" alt="{alt}" role="img">{caption_html}
    </figure>"""


class PDFUAPatcher:
    """
    Patcht PDF für PDF/UA-1 Compliance.
    
    WeasyPrint generiert Tagged PDFs, aber nicht alle
    PDF/UA-Anforderungen werden erfüllt. Dieser Patcher
    fügt fehlende Metadaten hinzu.
    """
    
    def __init__(self):
        self._pikepdf_available = self._check_pikepdf()
    
    def _check_pikepdf(self) -> bool:
        """Prüft ob pikepdf verfügbar ist."""
        try:
            import pikepdf
            return True
        except ImportError:
            return False
    
    def patch(self, input_pdf: Path, output_pdf: Path, model: SlideModel):
        """
        Patcht PDF für PDF/UA Compliance.
        
        Fügt hinzu:
        - /MarkInfo mit /Marked true
        - /Lang für Dokumentsprache
        - /ViewerPreferences
        - Metadaten (Title, Author, etc.)
        """
        if not self._pikepdf_available:
            # Fallback: Einfach kopieren
            import shutil
            shutil.copy(input_pdf, output_pdf)
            print("⚠️  pikepdf nicht installiert - PDF/UA-Patching übersprungen")
            return
        
        import pikepdf
        
        with pikepdf.open(input_pdf) as pdf:
            # 1. MarkInfo (zeigt an dass PDF getaggt ist)
            pdf.Root.MarkInfo = pikepdf.Dictionary({
                "/Marked": True
            })

            # 2. Dokumentsprache
            pdf.Root.Lang = model.language

            # 3. ViewerPreferences
            pdf.Root.ViewerPreferences = pikepdf.Dictionary({
                "/DisplayDocTitle": True
            })
            
            # 4. Metadaten via Info Dictionary
            pdf.docinfo['/Title'] = model.title or "Präsentation"
            pdf.docinfo['/Author'] = model.author or ""
            pdf.docinfo['/Subject'] = model.subject or ""
            pdf.docinfo['/Creator'] = "pptx2ua"
            pdf.docinfo['/Producer'] = "WeasyPrint + pptx2ua"
            
            # 5. XMP Metadaten (für PDF/UA Part Identifier)
            # TODO: PDF/UA-1 Identifier hinzufügen
            # Dies erfordert tieferes XMP-Handling
            
            pdf.save(output_pdf)
    
    def add_pdfua_identifier(self, pdf_path: Path):
        """
        Fügt PDF/UA-1 Identifier hinzu.
        
        Dies ist der technische Marker der besagt dass das
        Dokument PDF/UA-1 konform sein soll.
        
        TODO: XMP Metadata Manipulation
        """
        pass


class PDFUARenderer:
    """
    Hauptklasse für PDF/UA Rendering.
    
    Usage:
        renderer = PDFUARenderer()
        renderer.render(model, "output.pdf")
    """
    
    def __init__(self, config: Optional[RendererConfig] = None):
        self.config = config or RendererConfig()
        self.html_generator = HTMLGenerator(self.config)
        self.patcher = PDFUAPatcher()
    
    def render(
        self, 
        model: SlideModel, 
        output_path: Path | str,
        verbose: bool = True
    ) -> bool:
        """
        Rendert SlideModel zu PDF/UA.
        
        Args:
            model: Das zu rendernde SlideModel
            output_path: Pfad für das Ausgabe-PDF
            verbose: Fortschrittsausgabe
            
        Returns:
            True bei Erfolg
        """
        output_path = Path(output_path)
        
        try:
            if verbose:
                print("📝 Generiere HTML...")
            
            # 1. HTML generieren
            html_content = self.html_generator.generate(model)
            
            # Debug: HTML speichern
            debug_html = output_path.with_suffix('.html')
            debug_html.write_text(html_content, encoding='utf-8')
            
            if verbose:
                print("🖨️  Rendere PDF mit WeasyPrint...")
            
            # 2. PDF rendern
            font_config = FontConfiguration()
            
            # WeasyPrint Optionen für besseres Tagging
            html_doc = HTML(string=html_content, base_url=str(output_path.parent))
            
            # Temporäres PDF
            temp_pdf = output_path.with_suffix('.tmp.pdf')
            
            html_doc.write_pdf(
                temp_pdf,
                font_config=font_config,
                # PDF/UA relevante Optionen
                pdf_variant='pdf/ua-1',  # Experimentell!
            )
            
            if verbose:
                print("🔧 Patche PDF/UA Metadaten...")
            
            # 3. PDF/UA patchen
            self.patcher.patch(temp_pdf, output_path, model)
            
            # Cleanup
            temp_pdf.unlink(missing_ok=True)
            
            if verbose:
                print(f"✅ PDF erstellt: {output_path}")
            
            return True
            
        except Exception as e:
            print(f"❌ Rendering fehlgeschlagen: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def render_html_only(self, model: SlideModel, output_path: Path | str) -> str:
        """
        Rendert nur HTML (für Debugging/Preview).
        
        Returns:
            Der generierte HTML-String
        """
        output_path = Path(output_path)
        html_content = self.html_generator.generate(model)
        output_path.write_text(html_content, encoding='utf-8')
        return html_content
