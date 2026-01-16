"""
Tests für pptx2ua
=================

Ausführen mit: pytest
"""

import pytest
from pathlib import Path

from pptx2ua.models import (
    SlideModel, Slide, Block, BlockType,
    Paragraph, TextRun, Figure, Table, TableCell
)


class TestModels:
    """Tests für Datenmodelle."""
    
    def test_slide_model_creation(self):
        """SlideModel kann erstellt werden."""
        model = SlideModel(language="de", title="Test")
        assert model.language == "de"
        assert model.title == "Test"
        assert model.slide_count == 0
    
    def test_slide_with_blocks(self):
        """Slide kann Blöcke enthalten."""
        slide = Slide(number=1)
        
        block = Block(
            block_type=BlockType.HEADING,
            reading_order=1,
            paragraphs=[Paragraph(runs=[TextRun(text="Titel")])]
        )
        slide.blocks.append(block)
        
        assert slide.title == "Titel"
        assert len(slide.blocks) == 1
    
    def test_paragraph_text(self):
        """Paragraph.text kombiniert alle Runs."""
        para = Paragraph(runs=[
            TextRun(text="Hello "),
            TextRun(text="World", bold=True)
        ])
        
        assert para.text == "Hello World"
        assert not para.is_empty
    
    def test_empty_paragraph(self):
        """Leere Paragraphen werden erkannt."""
        para = Paragraph(runs=[TextRun(text="  ")])
        assert para.is_empty
    
    def test_figure_needs_alt_text(self):
        """Figure erkennt fehlenden Alt-Text."""
        fig_without = Figure(image_data=b"test")
        fig_with = Figure(image_data=b"test", alt_text="Beschreibung")
        
        assert fig_without.needs_alt_text
        # Nach Setzen von alt_text manuell
        fig_with.needs_alt_text = False
        assert not fig_with.needs_alt_text
    
    def test_table_has_header(self):
        """Table erkennt Header-Zeile."""
        table = Table(rows=[
            [TableCell(is_header=True), TableCell(is_header=True)],
            [TableCell(), TableCell()]
        ])
        
        assert table.has_header
        assert table.column_count == 2
    
    def test_reading_order(self):
        """Blöcke werden nach Lesereihenfolge sortiert."""
        slide = Slide(number=1)
        slide.blocks = [
            Block(block_type=BlockType.PARAGRAPH, reading_order=3),
            Block(block_type=BlockType.HEADING, reading_order=1),
            Block(block_type=BlockType.LIST, reading_order=2),
        ]
        
        sorted_blocks = slide.sorted_blocks
        assert sorted_blocks[0].reading_order == 1
        assert sorted_blocks[1].reading_order == 2
        assert sorted_blocks[2].reading_order == 3


class TestBlockTypes:
    """Tests für Block-Typ Erkennung."""
    
    def test_heading_block(self):
        """Heading-Blöcke haben Level."""
        block = Block(
            block_type=BlockType.HEADING,
            reading_order=1,
            heading_level=2
        )
        assert block.heading_level == 2
    
    def test_figure_block(self):
        """Figure-Blöcke enthalten Bild-Daten."""
        figure = Figure(
            image_data=b"PNG...",
            mime_type="image/png",
            alt_text="Ein Bild"
        )
        block = Block(
            block_type=BlockType.FIGURE,
            reading_order=1,
            figure=figure
        )
        
        assert block.figure is not None
        assert block.figure.alt_text == "Ein Bild"


class TestSlideModel:
    """Tests für SlideModel Aggregation."""
    
    def test_all_figures(self):
        """SlideModel sammelt alle Figures."""
        model = SlideModel()
        
        slide1 = Slide(number=1)
        slide1.blocks.append(Block(
            block_type=BlockType.FIGURE,
            reading_order=1,
            figure=Figure(image_data=b"1")
        ))
        
        slide2 = Slide(number=2)
        slide2.blocks.append(Block(
            block_type=BlockType.FIGURE,
            reading_order=1,
            figure=Figure(image_data=b"2")
        ))
        slide2.blocks.append(Block(
            block_type=BlockType.FIGURE,
            reading_order=2,
            figure=Figure(image_data=b"3")
        ))
        
        model.slides = [slide1, slide2]
        
        assert len(model.all_figures) == 3
    
    def test_figures_needing_alt_text(self):
        """Erkennt Figures ohne Alt-Text."""
        model = SlideModel()
        
        slide = Slide(number=1)
        slide.blocks.append(Block(
            block_type=BlockType.FIGURE,
            reading_order=1,
            figure=Figure(image_data=b"1", alt_text="Hat Alt")
        ))
        slide.blocks.append(Block(
            block_type=BlockType.FIGURE,
            reading_order=2,
            figure=Figure(image_data=b"2", needs_alt_text=True)
        ))
        
        # Manuell needs_alt_text setzen
        slide.blocks[0].figure.needs_alt_text = False
        
        model.slides = [slide]
        
        needing = model.figures_needing_alt_text
        assert len(needing) == 1
        assert needing[0][0] == 1  # Slide number


# Integration Tests (benötigen echte Dateien)

@pytest.mark.skip(reason="Benötigt Test-PPTX")
class TestParser:
    """Parser-Tests mit echten Dateien."""
    
    def test_parse_simple_pptx(self, tmp_path):
        """Kann einfache PPTX parsen."""
        from pptx2ua.parser import PPTXParser
        
        # Hier würde eine Test-PPTX erstellt/geladen
        pass


@pytest.mark.skip(reason="Benötigt WeasyPrint")
class TestRenderer:
    """Renderer-Tests."""
    
    def test_render_simple_model(self, tmp_path):
        """Kann einfaches Model rendern."""
        from pptx2ua.renderer import PDFUARenderer
        
        model = SlideModel(
            title="Test",
            language="de",
            slides=[
                Slide(
                    number=1,
                    blocks=[
                        Block(
                            block_type=BlockType.HEADING,
                            reading_order=1,
                            paragraphs=[Paragraph(runs=[TextRun(text="Titel")])]
                        )
                    ]
                )
            ]
        )
        
        renderer = PDFUARenderer()
        output = tmp_path / "test.pdf"
        
        success = renderer.render(model, output, verbose=False)
        assert success
        assert output.exists()
