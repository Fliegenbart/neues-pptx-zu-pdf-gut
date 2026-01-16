"""
Accessibility Optimizer
=======================
KI-gestützte Optimierung für Screenreader-UX.

Ziel: Blinde Menschen sollen den INHALT so gut wie möglich
verstehen - nicht das Layout nachvollziehen.

Optimierungen:
1. Dekorative Elemente erkennen und ausblenden
2. Fußnoten inline auflösen
3. Speaker Notes als Kontext integrieren
4. Redundanzen entfernen (wiederkehrende Logos etc.)
5. Optimale Verständnis-Reihenfolge
6. Tabellen/Charts in natürliche Sprache
7. Beziehungen zwischen Elementen erkennen
"""

import re
import hashlib
from dataclasses import dataclass, field
from typing import Optional
from enum import Enum
import base64
import requests

from .models import (
    SlideModel, Slide, Block, BlockType,
    Paragraph, TextRun, Figure, Table
)


class ElementRole(Enum):
    """Rolle eines Elements für Barrierefreiheit."""
    ESSENTIAL = "essential"      # Muss vorgelesen werden
    CONTEXTUAL = "contextual"    # Hilfreich für Kontext
    DECORATIVE = "decorative"    # Nicht vorlesen
    REDUNDANT = "redundant"      # Schon vorgelesen (z.B. Logo)
    NAVIGATION = "navigation"    # Foliennummer etc.
    BOILERPLATE = "boilerplate"  # Copyright, Impressum, etc.
    PLACEHOLDER = "placeholder"  # Leere Template-Texte


class ComplexSlideType(Enum):
    """Typen von visuell komplexen Folien."""
    TIMELINE = "timeline"           # Zeitachsen, Roadmaps
    FLOWCHART = "flowchart"         # Flussdiagramme, Prozesse
    ORG_CHART = "org_chart"         # Organigramme
    COMPARISON = "comparison"       # Vergleichstabellen, vs-Layouts
    INFOGRAPHIC = "infographic"     # Komplexe Infografiken
    DIAGRAM = "diagram"             # Sonstige Diagramme
    SIMPLE = "simple"               # Einfache Folie, keine Spezialbehandlung


@dataclass
class AccessibilityConfig:
    """Konfiguration für Accessibility-Optimierung."""
    ollama_url: str = "http://localhost:11434"
    model: str = "llama3.2:3b"  # Schnelles Modell für Analyse
    vision_model: str = "qwen2.5vl:7b"  # Bestes lokales Vision-Modell für Diagramme
    language: str = "de"

    # Docling Integration
    use_docling: bool = True  # Docling für Reading Order & Tabellen nutzen
    docling_analysis: Optional[any] = None  # Gecachte Docling-Analyse

    # Feature Flags
    inline_footnotes: bool = True
    use_speaker_notes: bool = True
    remove_decorative: bool = True
    remove_redundant: bool = True
    optimize_reading_order: bool = True
    summarize_complex_slides: bool = True
    naturalize_tables: bool = True
    naturalize_charts: bool = True

    # NEU: Vision-LLM für komplexe Folien
    use_vision_for_complex_slides: bool = True  # Ganze Folie analysieren
    complex_slide_types: list = None  # Welche Typen erkennen

    # Schwellwerte
    complex_slide_threshold: int = 6  # Blöcke ab denen Zusammenfassung
    redundancy_hash_threshold: int = 2  # Ab wann Element als redundant gilt

    # Neue Filter für unnötige Informationen
    remove_page_numbers: bool = True       # Foliennummern entfernen
    remove_boilerplate: bool = True        # Copyright, Impressum, etc.
    remove_empty_placeholders: bool = True # Leere "Titel eingeben" Texte
    remove_navigation_hints: bool = True   # "Klicken Sie hier", etc.
    simplify_contact_info: bool = True     # Kontaktdaten nur einmal

    def __post_init__(self):
        if self.complex_slide_types is None:
            self.complex_slide_types = [
                ComplexSlideType.TIMELINE,
                ComplexSlideType.FLOWCHART,
                ComplexSlideType.ORG_CHART,
                ComplexSlideType.COMPARISON,
                ComplexSlideType.INFOGRAPHIC,
            ]


@dataclass
class AccessibilityAnnotation:
    """Annotation für ein Element mit A11y-Infos."""
    role: ElementRole
    screen_reader_text: Optional[str] = None  # Optimierter Text
    skip_reason: Optional[str] = None  # Warum übersprungen
    relationships: list[str] = field(default_factory=list)  # Verknüpfte Elemente
    context_from_notes: Optional[str] = None  # Kontext aus Speaker Notes


class AccessibilityOptimizer:
    """
    Optimiert SlideModel für Screenreader-UX.

    Unterstützt optionale Docling-Integration für:
    - Verbesserte Reading Order Detection
    - Tabellen-Strukturerkennung

    Usage:
        optimizer = AccessibilityOptimizer()
        optimized_model = optimizer.optimize(model)
    """

    def __init__(self, config: Optional[AccessibilityConfig] = None):
        self.config = config or AccessibilityConfig()
        self._seen_hashes: dict[str, int] = {}  # Für Redundanz-Erkennung
        self._footnotes: dict[str, str] = {}  # Fußnoten-Sammlung
        self._llm_available = self._check_llm()
        self._docling_available = self._check_docling()
        self._docling_analysis = None

    def _check_llm(self) -> bool:
        """Prüft ob LLM verfügbar."""
        try:
            r = requests.get(f"{self.config.ollama_url}/api/tags", timeout=3)
            return r.status_code == 200
        except:
            return False

    def _check_docling(self) -> bool:
        """Prüft ob Docling verfügbar ist."""
        if not self.config.use_docling:
            return False
        try:
            from .docling_integration import is_docling_available
            return is_docling_available()
        except ImportError:
            return False

    def analyze_with_docling(self, pptx_path) -> bool:
        """
        Führt Docling-Analyse für die PPTX durch.

        Args:
            pptx_path: Pfad zur Original-PPTX

        Returns:
            True wenn erfolgreich
        """
        if not self._docling_available:
            return False

        try:
            from .docling_integration import analyze_with_docling, DoclingConfig

            docling_config = DoclingConfig(
                language=self.config.language,
                use_table_structure=True,
                use_reading_order=True,
            )

            self._docling_analysis = analyze_with_docling(pptx_path, docling_config)
            return self._docling_analysis is not None

        except Exception as e:
            print(f"   Docling-Analyse fehlgeschlagen: {e}")
            return False
    
    def optimize(self, model: SlideModel, verbose: bool = True) -> SlideModel:
        """
        Führt alle Accessibility-Optimierungen durch.
        
        Args:
            model: Das zu optimierende SlideModel
            verbose: Fortschrittsausgabe
            
        Returns:
            Optimiertes SlideModel (in-place modifiziert)
        """
        if verbose:
            print("\n♿ Accessibility-Optimierung...")

        # Zähle ursprüngliche Elemente für Statistik
        self._original_block_count = sum(len(s.blocks) for s in model.slides)

        # Phase 1: Analyse (sammelt Informationen)
        if verbose:
            print("   📊 Phase 1: Analyse...")
        self._analyze_document(model)
        
        # Phase 2: Dekorative Elemente markieren
        if self.config.remove_decorative:
            if verbose:
                print("   🎨 Phase 2: Dekorative Elemente erkennen...")
            self._mark_decorative_elements(model)
        
        # Phase 3: Redundanzen erkennen
        if self.config.remove_redundant:
            if verbose:
                print("   🔄 Phase 3: Redundanzen erkennen...")
            self._mark_redundant_elements(model)

        # Phase 3b: Unnötige Informationen entfernen
        if verbose:
            print("   🧹 Phase 3b: Unnötige Informationen filtern...")
        self._remove_unnecessary_info(model)

        # Phase 3c: Komplexe Folien mit Vision-LLM analysieren
        if self.config.use_vision_for_complex_slides and self._llm_available:
            if verbose:
                print("   🔍 Phase 3c: Komplexe Folien analysieren (Vision-LLM)...")
            self._analyze_complex_slides_with_vision(model)

        # Phase 4: Fußnoten inline auflösen
        if self.config.inline_footnotes:
            if verbose:
                print("   📝 Phase 4: Fußnoten inline auflösen...")
            self._inline_footnotes(model)
        
        # Phase 5: Speaker Notes integrieren
        if self.config.use_speaker_notes:
            if verbose:
                print("   🎤 Phase 5: Speaker Notes als Kontext...")
            self._integrate_speaker_notes(model)
        
        # Phase 6: Lesereihenfolge optimieren
        if self.config.optimize_reading_order:
            if verbose:
                print("   📖 Phase 6: Lesereihenfolge optimieren...")
            self._optimize_reading_order(model)
        
        # Phase 7: Tabellen in natürliche Sprache
        if self.config.naturalize_tables:
            if verbose:
                print("   📊 Phase 7: Tabellen optimieren...")
            self._naturalize_tables(model)
        
        # Phase 8: Charts beschreiben
        if self.config.naturalize_charts:
            if verbose:
                print("   📈 Phase 8: Charts beschreiben...")
            self._describe_charts(model)
        
        # Phase 9: Komplexe Folien zusammenfassen
        if self.config.summarize_complex_slides:
            if verbose:
                print("   📋 Phase 9: Komplexe Folien zusammenfassen...")
            self._add_slide_summaries(model)
        
        # Phase 10: Finale Bereinigung
        if verbose:
            print("   ✨ Phase 10: Finale Bereinigung...")
        self._final_cleanup(model)
        
        if verbose:
            self._print_stats(model)
        
        return model
    
    # === Phase 1: Analyse ===
    
    def _analyze_document(self, model: SlideModel):
        """Analysiert das Dokument und sammelt Metadaten."""
        # Fußnoten sammeln
        for slide in model.slides:
            self._extract_footnotes(slide)
        
        # Element-Hashes für Redundanz-Erkennung
        for slide in model.slides:
            for block in slide.blocks:
                hash_val = self._compute_content_hash(block)
                if hash_val:
                    self._seen_hashes[hash_val] = self._seen_hashes.get(hash_val, 0) + 1
    
    def _extract_footnotes(self, slide: Slide):
        """Extrahiert Fußnoten aus einer Folie."""
        for block in slide.blocks:
            text = block.text
            
            # Pattern: Fußnote am unteren Rand mit Nummer
            # z.B. "¹ Quelle: BMI 2024" oder "1) Siehe Anhang"
            footnote_patterns = [
                r'^[¹²³⁴⁵⁶⁷⁸⁹⁰]+\s*(.+)$',  # Hochgestellte Zahlen
                r'^\[(\d+)\]\s*(.+)$',  # [1] Format
                r'^(\d+)\)\s*(.+)$',  # 1) Format
                r'^\*+\s*(.+)$',  # Sternchen
            ]
            
            for pattern in footnote_patterns:
                match = re.match(pattern, text.strip(), re.MULTILINE)
                if match:
                    # Extrahiere Marker und Text
                    groups = match.groups()
                    if len(groups) == 2:
                        marker, content = groups
                    else:
                        marker = "1"
                        content = groups[0]
                    
                    self._footnotes[marker] = content.strip()
    
    def _compute_content_hash(self, block: Block) -> Optional[str]:
        """Berechnet Hash für Inhaltsvergleich."""
        if block.figure and block.figure.image_data:
            return hashlib.md5(block.figure.image_data).hexdigest()
        
        if block.text:
            # Normalisiere Text für Vergleich
            normalized = block.text.lower().strip()
            if len(normalized) > 10:  # Nur für substantiellen Content
                return hashlib.md5(normalized.encode()).hexdigest()
        
        return None
    
    # === Phase 2: Dekorative Elemente ===
    
    def _mark_decorative_elements(self, model: SlideModel):
        """Markiert dekorative Elemente die nicht vorgelesen werden sollen."""
        for slide in model.slides:
            for block in slide.blocks:
                if self._is_decorative(block, slide):
                    if not hasattr(block, 'a11y'):
                        block.a11y = AccessibilityAnnotation(role=ElementRole.DECORATIVE)
                    else:
                        block.a11y.role = ElementRole.DECORATIVE
                    block.a11y.skip_reason = self._get_decorative_reason(block)
    
    def _is_decorative(self, block: Block, slide: Slide) -> bool:
        """Prüft ob Element dekorativ ist."""
        # Bilder ohne Alt-Text und ohne Kontext
        if block.figure:
            fig = block.figure
            
            # Sehr kleine Bilder sind oft Icons/Bullets
            if block.bbox:
                if block.bbox.width < 20 and block.bbox.height < 20:
                    return True
            
            # Hintergrundbilder (volle Foliengröße)
            if block.bbox:
                if (block.bbox.width > slide.width_mm * 0.9 and 
                    block.bbox.height > slide.height_mm * 0.9):
                    # Könnte Hintergrund sein - KI fragen wenn verfügbar
                    if self._llm_available:
                        return self._ask_if_decorative(fig)
                    return True
        
        # Linien und Formen ohne Text
        if block.block_type == BlockType.PARAGRAPH:
            text = block.text.strip()
            # Nur Striche, Punkte, Leerzeichen
            if re.match(r'^[\s\-–—_\.•·│|]+$', text):
                return True
        
        return False
    
    def _ask_if_decorative(self, figure: Figure) -> bool:
        """Fragt KI ob Bild dekorativ ist."""
        if not figure.image_data:
            return False
        
        prompt = """Analysiere dieses Bild. Ist es:
A) DEKORATIV - Hintergrund, Muster, Rahmen, rein ästhetisch
B) INHALTLICH - Trägt Information bei (Foto, Diagramm, Screenshot)

Antworte nur mit A oder B."""

        try:
            image_b64 = base64.b64encode(figure.image_data).decode('utf-8')
            response = requests.post(
                f"{self.config.ollama_url}/api/generate",
                json={
                    "model": self.config.vision_model,
                    "prompt": prompt,
                    "images": [image_b64],
                    "stream": False,
                    "options": {"temperature": 0.1, "num_predict": 10}
                },
                timeout=30
            )
            
            if response.status_code == 200:
                answer = response.json().get("response", "").strip().upper()
                return answer.startswith("A")
        except:
            pass
        
        return False
    
    def _get_decorative_reason(self, block: Block) -> str:
        """Gibt Grund für Dekoration zurück."""
        if block.figure:
            if block.bbox and block.bbox.width < 20:
                return "Kleines Icon/Aufzählungszeichen"
            return "Dekoratives Hintergrundbild"
        return "Dekoratives Element (Linie/Form)"
    
    # === Phase 3: Redundanzen ===
    
    def _mark_redundant_elements(self, model: SlideModel):
        """Markiert redundante Elemente (z.B. Logo auf jeder Folie)."""
        seen_on_slides: dict[str, list[int]] = {}
        
        # Sammle wo jedes Element vorkommt
        for slide in model.slides:
            for block in slide.blocks:
                hash_val = self._compute_content_hash(block)
                if hash_val:
                    if hash_val not in seen_on_slides:
                        seen_on_slides[hash_val] = []
                    seen_on_slides[hash_val].append(slide.number)
        
        # Markiere Elemente die auf mehreren Folien vorkommen
        for slide in model.slides:
            for block in slide.blocks:
                hash_val = self._compute_content_hash(block)
                if hash_val and hash_val in seen_on_slides:
                    occurrences = seen_on_slides[hash_val]
                    
                    if len(occurrences) >= self.config.redundancy_hash_threshold:
                        # Nur auf erster Folie vorlesen
                        if slide.number != min(occurrences):
                            if not hasattr(block, 'a11y'):
                                block.a11y = AccessibilityAnnotation(
                                    role=ElementRole.REDUNDANT,
                                    skip_reason=f"Bereits auf Folie {min(occurrences)} vorgelesen"
                                )
                            else:
                                block.a11y.role = ElementRole.REDUNDANT
    
    # === Phase 3b: Unnötige Informationen filtern ===

    def _remove_unnecessary_info(self, model: SlideModel):
        """Entfernt verschiedene Arten unnötiger Informationen."""
        # Pattern für unnötige Texte
        boilerplate_patterns = [
            # Copyright & Legal
            r'^©.*$', r'^copyright.*$', r'^\(c\).*$',
            r'^alle rechte vorbehalten.*$', r'^all rights reserved.*$',
            r'^impressum.*$', r'^datenschutz.*$', r'^privacy.*$',
            r'^vertraulich.*$', r'^confidential.*$', r'^intern.*$',
            # Navigation
            r'^klicken sie hier.*$', r'^click here.*$',
            r'^weiter$', r'^zurück$', r'^next$', r'^back$',
            r'^mehr erfahren.*$', r'^learn more.*$',
            # Leere Platzhalter
            r'^titel.*eingeben.*$', r'^text.*eingeben.*$',
            r'^add title.*$', r'^add text.*$', r'^click to add.*$',
            r'^untertitel.*$', r'^subtitle.*$',
            # Formatierung
            r'^folie \d+.*$', r'^slide \d+.*$',
            r'^\d+\s*/\s*\d+$',  # "5 / 10" Seitenzahlen
            r'^seite \d+.*$', r'^page \d+.*$',
        ]

        # Kontaktdaten-Pattern (für Deduplizierung)
        contact_patterns = [
            r'[\w\.-]+@[\w\.-]+\.\w+',  # E-Mail
            r'\+?\d[\d\s\-/]{8,}',       # Telefon
            r'www\.[\w\.-]+\.\w+',       # Website
        ]

        seen_contacts = set()

        for slide in model.slides:
            for block in slide.blocks:
                text = block.text.lower().strip() if block.text else ""

                if not text:
                    continue

                # Seitenzahlen entfernen
                if self.config.remove_page_numbers:
                    if self._is_page_number(text, slide):
                        self._mark_as_skip(block, ElementRole.NAVIGATION, "Seitenzahl")
                        continue

                # Boilerplate entfernen
                if self.config.remove_boilerplate:
                    for pattern in boilerplate_patterns:
                        if re.match(pattern, text, re.IGNORECASE):
                            self._mark_as_skip(block, ElementRole.BOILERPLATE, "Standardtext")
                            break

                # Leere Platzhalter
                if self.config.remove_empty_placeholders:
                    if self._is_placeholder_text(text):
                        self._mark_as_skip(block, ElementRole.PLACEHOLDER, "Platzhaltertext")
                        continue

                # Kontaktdaten deduplizieren
                if self.config.simplify_contact_info:
                    for pattern in contact_patterns:
                        matches = re.findall(pattern, text, re.IGNORECASE)
                        for match in matches:
                            normalized = match.lower().strip()
                            if normalized in seen_contacts:
                                self._mark_as_skip(
                                    block, ElementRole.REDUNDANT,
                                    "Kontaktdaten bereits genannt"
                                )
                                break
                            seen_contacts.add(normalized)

    def _is_page_number(self, text: str, slide: Slide) -> bool:
        """Erkennt Seitenzahlen/Foliennummern."""
        text = text.strip()

        # Reine Zahlen die zur Foliennummer passen
        if text.isdigit():
            num = int(text)
            if 1 <= num <= 100:  # Plausible Foliennummer
                return True

        # "Folie X", "Seite X", "X von Y"
        patterns = [
            r'^folie\s*\d+$',
            r'^slide\s*\d+$',
            r'^seite\s*\d+$',
            r'^page\s*\d+$',
            r'^\d+\s*/\s*\d+$',
            r'^\d+\s*von\s*\d+$',
            r'^\d+\s*of\s*\d+$',
        ]

        for pattern in patterns:
            if re.match(pattern, text, re.IGNORECASE):
                return True

        return False

    def _is_placeholder_text(self, text: str) -> bool:
        """Erkennt leere Template-Platzhalter."""
        placeholders = [
            'titel eingeben', 'text eingeben', 'inhalt einfügen',
            'add title', 'add text', 'add content', 'click to add',
            'titel hinzufügen', 'text hinzufügen',
            'enter title', 'enter text',
            'lorem ipsum',
        ]

        text_lower = text.lower()
        return any(p in text_lower for p in placeholders)

    def _mark_as_skip(self, block: Block, role: ElementRole, reason: str):
        """Markiert Block zum Überspringen."""
        if not hasattr(block, 'a11y') or block.a11y is None:
            block.a11y = AccessibilityAnnotation(role=role, skip_reason=reason)
        else:
            block.a11y.role = role
            block.a11y.skip_reason = reason

    # === Phase 3c: Komplexe Folien mit Vision-LLM ===

    def _analyze_complex_slides_with_vision(self, model: SlideModel):
        """Analysiert komplexe Folien mit Vision-LLM für bessere Beschreibungen."""
        for slide in model.slides:
            slide_type = self._detect_slide_type(slide)

            if slide_type != ComplexSlideType.SIMPLE:
                # Komplexe Folie gefunden - Vision-LLM Analyse
                narrative = self._generate_slide_narrative(slide, slide_type)

                if narrative:
                    # Markiere alle bestehenden Blöcke als "ersetzt"
                    for block in slide.blocks:
                        if not hasattr(block, 'a11y') or block.a11y is None:
                            block.a11y = AccessibilityAnnotation(
                                role=ElementRole.REDUNDANT,
                                skip_reason="Ersetzt durch Folien-Narrative"
                            )
                        else:
                            block.a11y.role = ElementRole.REDUNDANT

                    # Füge narrative Beschreibung als neuen Block ein
                    narrative_block = Block(
                        block_type=BlockType.PARAGRAPH,
                        reading_order=1,
                        paragraphs=[Paragraph(runs=[TextRun(text=narrative)])]
                    )
                    narrative_block.a11y = AccessibilityAnnotation(
                        role=ElementRole.ESSENTIAL,
                        screen_reader_text=narrative
                    )

                    # Behalte nur Titel + Narrative
                    title_block = None
                    for block in slide.blocks:
                        if block.block_type == BlockType.HEADING and block.heading_level == 1:
                            title_block = block
                            title_block.a11y = AccessibilityAnnotation(role=ElementRole.ESSENTIAL)
                            title_block.reading_order = 0
                            break

                    if title_block:
                        slide.blocks = [title_block, narrative_block]
                    else:
                        slide.blocks = [narrative_block]

    def _detect_slide_type(self, slide: Slide) -> ComplexSlideType:
        """Erkennt den Typ einer komplexen Folie basierend auf Indikatoren."""
        all_text = " ".join(b.text.lower() for b in slide.blocks if b.text)
        block_count = len(slide.blocks)

        # Timeline/Roadmap Indikatoren
        timeline_indicators = [
            r'\b(20\d{2})\b',  # Jahreszahlen 2000-2099
            r'\b(q[1-4])\b',   # Q1, Q2, etc.
            r'\b(phase\s*\d)\b',
            r'\b(schritt\s*\d)\b',
            r'\b(step\s*\d)\b',
            r'(→|➔|➜|▶|►)',   # Pfeile
            r'\b(roadmap|timeline|zeitachse|meilenstein)\b',
        ]

        timeline_score = 0
        for pattern in timeline_indicators:
            if re.search(pattern, all_text, re.IGNORECASE):
                timeline_score += 1

        # Viele Jahreszahlen = sehr wahrscheinlich Timeline
        year_matches = re.findall(r'\b(20\d{2})\b', all_text)
        if len(year_matches) >= 3:
            timeline_score += 3

        if timeline_score >= 3:
            return ComplexSlideType.TIMELINE

        # Flowchart Indikatoren
        flowchart_indicators = [
            r'\b(wenn|dann|sonst|if|then|else)\b',
            r'\b(entscheidung|decision)\b',
            r'\b(prozess|process|ablauf)\b',
            r'\b(start|ende|end)\b',
            r'(→|➔|➜|▶|►|↓|↑)',
        ]

        flowchart_score = 0
        for pattern in flowchart_indicators:
            if re.search(pattern, all_text, re.IGNORECASE):
                flowchart_score += 1

        if flowchart_score >= 3:
            return ComplexSlideType.FLOWCHART

        # Organigramm Indikatoren
        org_indicators = [
            r'\b(ceo|cto|cfo|coo)\b',
            r'\b(leiter|manager|direktor|vorstand)\b',
            r'\b(abteilung|team|bereich)\b',
            r'\b(organisation|struktur)\b',
        ]

        org_score = 0
        for pattern in org_indicators:
            if re.search(pattern, all_text, re.IGNORECASE):
                org_score += 1

        if org_score >= 2:
            return ComplexSlideType.ORG_CHART

        # Vergleich Indikatoren
        comparison_indicators = [
            r'\b(vs\.?|versus|vergleich|compared)\b',
            r'\b(vorher|nachher|before|after)\b',
            r'\b(alt|neu|old|new)\b',
            r'\b(pro|contra|vorteil|nachteil)\b',
        ]

        comparison_score = 0
        for pattern in comparison_indicators:
            if re.search(pattern, all_text, re.IGNORECASE):
                comparison_score += 1

        if comparison_score >= 2:
            return ComplexSlideType.COMPARISON

        # Viele verstreute Blöcke = wahrscheinlich komplexes Layout
        if block_count >= 8:
            # Prüfe ob Blöcke visuell verstreut sind
            if self._has_scattered_layout(slide):
                return ComplexSlideType.INFOGRAPHIC

        return ComplexSlideType.SIMPLE

    def _has_scattered_layout(self, slide: Slide) -> bool:
        """Prüft ob Blöcke über die Folie verstreut sind (nicht linear)."""
        blocks_with_bbox = [b for b in slide.blocks if b.bbox]

        if len(blocks_with_bbox) < 4:
            return False

        # Prüfe Y-Positionen - wenn viele auf ähnlicher Höhe aber unterschiedlichem X
        y_positions = sorted([b.bbox.y for b in blocks_with_bbox])
        x_positions = sorted([b.bbox.x for b in blocks_with_bbox])

        # Wenn X-Varianz hoch und mehrere Y-Cluster = verstreut
        x_range = max(x_positions) - min(x_positions) if x_positions else 0
        y_clusters = len(set(round(y / 50) for y in y_positions))  # ~50mm Cluster

        return x_range > 150 and y_clusters >= 3

    def _generate_slide_narrative(self, slide: Slide, slide_type: ComplexSlideType) -> Optional[str]:
        """Generiert narrative Beschreibung für komplexe Folie mit Vision-LLM."""
        # Typ-spezifische Prompts
        type_prompts = {
            ComplexSlideType.TIMELINE: """Analysiere diese Folie. Es ist eine TIMELINE/ROADMAP.

Beschreibe für einen blinden Menschen:
1. Den Gesamtzeitraum (von wann bis wann)
2. Die Phasen/Meilensteine in CHRONOLOGISCHER Reihenfolge
3. Was bereits abgeschlossen ist (Häkchen = erledigt)
4. Was aktuell läuft und was geplant ist

Wichtig: Folge den Pfeilen/der visuellen Zeitachse, NICHT der Textposition!
Maximal 200 Wörter, Fließtext.""",

            ComplexSlideType.FLOWCHART: """Analysiere diese Folie. Es ist ein PROZESS/FLUSSDIAGRAMM.

Beschreibe für einen blinden Menschen:
1. Den Startpunkt des Prozesses
2. Die Schritte in der RICHTIGEN Ablauf-Reihenfolge (folge den Pfeilen!)
3. Verzweigungen und Entscheidungspunkte
4. Das Endergebnis

Wichtig: Folge dem visuellen Fluss, NICHT der Textposition!
Maximal 200 Wörter, Fließtext.""",

            ComplexSlideType.ORG_CHART: """Analysiere diese Folie. Es ist ein ORGANIGRAMM.

Beschreibe für einen blinden Menschen:
1. Die oberste Ebene (Leitung/Chef)
2. Die Struktur darunter (wer berichtet an wen)
3. Die verschiedenen Abteilungen/Bereiche

Wichtig: Beschreibe die Hierarchie von oben nach unten!
Maximal 200 Wörter, Fließtext.""",

            ComplexSlideType.COMPARISON: """Analysiere diese Folie. Es ist ein VERGLEICH.

Beschreibe für einen blinden Menschen:
1. Was wird verglichen (welche Optionen/Varianten)
2. Die wichtigsten Unterschiede
3. Vor- und Nachteile jeder Option
4. Falls vorhanden: Empfehlung oder Fazit

Maximal 200 Wörter, Fließtext.""",

            ComplexSlideType.INFOGRAPHIC: """Analysiere diese Folie. Es ist eine KOMPLEXE INFOGRAFIK.

Beschreibe für einen blinden Menschen:
1. Das Hauptthema der Folie
2. Die wichtigsten Informationen und Zahlen
3. Wie die Elemente zusammenhängen
4. Die Kernaussage

Wichtig: Ordne die Informationen LOGISCH, nicht nach Position!
Maximal 200 Wörter, Fließtext.""",
        }

        type_instruction = type_prompts.get(slide_type, type_prompts[ComplexSlideType.INFOGRAPHIC])

        # Wenn Folienbild verfügbar: Vision-LLM mit Bild
        if slide.slide_image:
            return self._analyze_slide_with_vision(slide, type_instruction)

        # Fallback: Text-basierte Analyse
        return self._analyze_slide_with_text(slide, slide_type, type_instruction)

    def _analyze_slide_with_vision(self, slide: Slide, instruction: str) -> Optional[str]:
        """Analysiert Folie mit Vision-LLM und echtem Bild."""
        prompt = f"""Du bist ein Experte für Barrierefreiheit. Deine Aufgabe: Beschreibe diese
Präsentationsfolie so, dass ein blinder Mensch den INHALT vollständig versteht.

Folientitel: "{slide.title or 'Ohne Titel'}"

ANALYSIERE DAS BILD SCHRITT FÜR SCHRITT:

1. STRUKTUR: Was für ein Diagramm/Layout siehst du?
   (Timeline? Prozess? Organigramm? Vergleich? Liste?)

2. VISUELLE HINWEISE beachten:
   - Pfeile → zeigen Richtung/Ablauf
   - Häkchen ✓ → abgeschlossen/erledigt
   - Farben → Gruppierungen oder Status
   - Position → zeitliche oder hierarchische Ordnung

3. INHALT in LOGISCHER Reihenfolge:
{instruction}

AUSGABE-FORMAT:
- Beginne DIREKT mit dem Inhalt (NICHT "Diese Folie zeigt...")
- Schreibe einen zusammenhängenden Fließtext (KEINE Aufzählungspunkte)
- Beschreibe in der Reihenfolge, die SINN ergibt (chronologisch, hierarchisch, etc.)
- Maximal 250 Wörter
- Auf Deutsch

Deine Beschreibung:"""

        try:
            image_b64 = base64.b64encode(slide.slide_image).decode('utf-8')

            response = requests.post(
                f"{self.config.ollama_url}/api/generate",
                json={
                    "model": self.config.vision_model,
                    "prompt": prompt,
                    "images": [image_b64],
                    "stream": False,
                    "options": {"temperature": 0.3, "num_predict": 500}
                },
                timeout=120  # Vision braucht länger
            )

            if response.status_code == 200:
                narrative = response.json().get("response", "").strip()
                # Bereinige typische Einleitungen
                narrative = re.sub(
                    r'^(diese folie zeigt|die folie zeigt|das bild zeigt|hier sehen wir|'
                    r'auf dieser folie|die präsentation zeigt|ich sehe)[:\s]*',
                    '', narrative, flags=re.IGNORECASE
                ).strip()
                if narrative:
                    return narrative

        except Exception as e:
            print(f"      Vision-Analyse Fehler: {e}")

        return None

    def _analyze_slide_with_text(self, slide: Slide, slide_type: ComplexSlideType, instruction: str) -> Optional[str]:
        """Fallback: Analysiert Folie nur mit extrahiertem Text."""
        # Sammle alle Textinhalte mit Position
        content_items = []
        for block in slide.blocks:
            if block.text and block.bbox:
                content_items.append({
                    "text": block.text.strip(),
                    "x": block.bbox.x,
                    "y": block.bbox.y,
                    "type": block.block_type.value
                })

        if not content_items:
            return None

        # Sortiere nach Position für Kontext
        content_items.sort(key=lambda x: (x["y"], x["x"]))
        content_text = "\n".join(f"- {item['text']}" for item in content_items[:20])

        prompt = f"""Du bist ein Accessibility-Experte. Du hilfst blinden Menschen,
visuelle Präsentationsfolien zu verstehen.

Folientitel: "{slide.title or 'Ohne Titel'}"

{instruction}

Die Folie enthält folgende Textelemente (NICHT in der richtigen Lesereihenfolge!):
{content_text}

WICHTIG:
- Ordne die Informationen LOGISCH, nicht nach Textposition
- Für Timelines: CHRONOLOGISCH ordnen
- Für Prozesse: Nach ABLAUF ordnen
- Beginne direkt mit dem Inhalt, nicht mit "Diese Folie zeigt..."

Deine Beschreibung:"""

        try:
            response = requests.post(
                f"{self.config.ollama_url}/api/generate",
                json={
                    "model": self.config.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {"temperature": 0.3, "num_predict": 400}
                },
                timeout=60
            )

            if response.status_code == 200:
                narrative = response.json().get("response", "").strip()
                narrative = re.sub(r'^(diese folie zeigt|die folie zeigt|hier sehen wir)',
                                   '', narrative, flags=re.IGNORECASE).strip()
                if narrative:
                    return narrative

        except Exception as e:
            print(f"      Text-Analyse Fehler: {e}")

        return None

    # === Phase 4: Fußnoten Inline ===

    def _inline_footnotes(self, model: SlideModel):
        """Löst Fußnoten inline auf."""
        for slide in model.slides:
            for block in slide.blocks:
                if block.paragraphs:
                    for para in block.paragraphs:
                        for run in para.runs:
                            run.text = self._replace_footnote_markers(run.text)
    
    def _replace_footnote_markers(self, text: str) -> str:
        """Ersetzt Fußnotenmarker durch Inline-Text."""
        if not self._footnotes:
            return text
        
        # Ersetze hochgestellte Zahlen
        def replace_superscript(match):
            marker = match.group(1)
            # Konvertiere Unicode-Hochzahlen zu normalen Zahlen
            superscript_map = {'¹': '1', '²': '2', '³': '3', '⁴': '4', 
                             '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9', '⁰': '0'}
            normal_marker = ''.join(superscript_map.get(c, c) for c in marker)
            
            if normal_marker in self._footnotes:
                return f" ({self._footnotes[normal_marker]})"
            return match.group(0)
        
        # Pattern für Fußnotenmarker
        text = re.sub(r'([¹²³⁴⁵⁶⁷⁸⁹⁰]+)', replace_superscript, text)
        text = re.sub(r'\[(\d+)\]', lambda m: f" ({self._footnotes.get(m.group(1), m.group(0))})", text)
        
        return text
    
    # === Phase 5: Speaker Notes ===
    
    def _integrate_speaker_notes(self, model: SlideModel):
        """Integriert Speaker Notes als Kontext."""
        for slide in model.slides:
            if not slide.notes:
                continue
            
            notes = slide.notes.strip()
            if len(notes) < 20:  # Zu kurz für sinnvollen Kontext
                continue
            
            # Analysiere Notes für Kontext
            context = self._extract_context_from_notes(notes, slide)
            
            if context:
                # Füge Kontext-Block am Anfang der Folie ein
                context_block = Block(
                    block_type=BlockType.PARAGRAPH,
                    reading_order=0,  # Ganz am Anfang
                    paragraphs=[Paragraph(runs=[TextRun(text=context)])]
                )
                
                if not hasattr(context_block, 'a11y'):
                    context_block.a11y = AccessibilityAnnotation(
                        role=ElementRole.CONTEXTUAL,
                        context_from_notes=notes
                    )
                
                # Füge vor allen anderen Blöcken ein
                slide.blocks.insert(0, context_block)
                
                # Reading Order neu nummerieren
                for i, block in enumerate(slide.blocks):
                    block.reading_order = i + 1
    
    def _extract_context_from_notes(self, notes: str, slide: Slide) -> Optional[str]:
        """Extrahiert relevanten Kontext aus Speaker Notes."""
        if not self._llm_available:
            # Fallback: Ersten Satz nehmen wenn er Kontext gibt
            first_sentence = notes.split('.')[0].strip()
            if len(first_sentence) > 30 and len(first_sentence) < 200:
                return f"Kontext: {first_sentence}."
            return None
        
        prompt = f"""Du bist ein Accessibility-Experte. 

Die folgende Folie hat den Titel: "{slide.title or 'Ohne Titel'}"

Der Vortragende hat diese Notizen:
"{notes[:500]}"

Aufgabe: Extrahiere in 1-2 Sätzen den wichtigsten KONTEXT, den ein blinder 
Zuhörer VOR dem Lesen der Folie wissen sollte. Das könnte sein:
- Warum diese Folie wichtig ist
- Wie sie mit dem Vorherigen zusammenhängt
- Was das Hauptargument ist

Wenn die Notizen keinen nützlichen Kontext bieten, antworte mit: KEIN_KONTEXT

Antworte direkt mit dem Kontextsatz oder KEIN_KONTEXT."""

        try:
            response = requests.post(
                f"{self.config.ollama_url}/api/generate",
                json={
                    "model": self.config.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {"temperature": 0.3, "num_predict": 150}
                },
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json().get("response", "").strip()
                if result and "KEIN_KONTEXT" not in result.upper():
                    return result
        except:
            pass
        
        return None
    
    # === Phase 6: Lesereihenfolge ===

    def _optimize_reading_order(self, model: SlideModel):
        """Optimiert Lesereihenfolge für Verständnis statt Layout."""
        # Wenn Docling-Analyse verfügbar: nutze diese als Basis
        if self._docling_analysis and self._docling_analysis.reading_order:
            self._apply_docling_reading_order(model)
            return

        # Fallback: Eigene Heuristik
        for slide in model.slides:
            # Filtere nicht zu lesende Elemente
            readable_blocks = [
                b for b in slide.blocks 
                if not (hasattr(b, 'a11y') and b.a11y.role in 
                       [ElementRole.DECORATIVE, ElementRole.REDUNDANT])
            ]
            
            if not readable_blocks:
                continue
            
            # Sortierlogik:
            # 1. Titel/Hauptüberschrift
            # 2. Kontext (aus Speaker Notes)
            # 3. Inhalt nach logischer Reihenfolge
            # 4. Ergänzende Infos
            
            def reading_priority(block: Block) -> tuple:
                # Niedrigere Zahl = früher lesen
                
                # Titel hat höchste Priorität
                if block.block_type == BlockType.HEADING and block.heading_level == 1:
                    return (0, 0, 0)
                
                # Kontext aus Notes
                if hasattr(block, 'a11y') and block.a11y.role == ElementRole.CONTEXTUAL:
                    return (1, 0, 0)
                
                # Andere Überschriften
                if block.block_type == BlockType.HEADING:
                    return (2, block.heading_level, 0)
                
                # Text vor Bildern (meist erklärt Text das Bild)
                if block.block_type == BlockType.PARAGRAPH:
                    return (3, 0, block.bbox.y if block.bbox else 0)
                
                if block.block_type == BlockType.LIST:
                    return (3, 1, block.bbox.y if block.bbox else 0)
                
                # Tabellen
                if block.block_type == BlockType.TABLE:
                    return (4, 0, block.bbox.y if block.bbox else 0)
                
                # Bilder zuletzt (nachdem Kontext gegeben wurde)
                if block.block_type == BlockType.FIGURE:
                    return (5, 0, block.bbox.y if block.bbox else 0)
                
                return (6, 0, 0)
            
            readable_blocks.sort(key=reading_priority)
            
            # Reading Order aktualisieren
            for i, block in enumerate(readable_blocks):
                block.reading_order = i + 1
            
            # Nicht lesbare Blöcke ans Ende (mit hoher reading_order)
            for block in slide.blocks:
                if block not in readable_blocks:
                    block.reading_order = 999

    def _apply_docling_reading_order(self, model: SlideModel):
        """Wendet Docling's Reading Order auf das Model an."""
        try:
            from .docling_integration import apply_docling_reading_order
            apply_docling_reading_order(model, self._docling_analysis)
        except Exception as e:
            print(f"   Docling Reading Order fehlgeschlagen: {e}")
            # Fallback auf eigene Heuristik
            for slide in model.slides:
                readable_blocks = [
                    b for b in slide.blocks
                    if not (hasattr(b, 'a11y') and b.a11y.role in
                           [ElementRole.DECORATIVE, ElementRole.REDUNDANT])
                ]
                for i, block in enumerate(readable_blocks):
                    block.reading_order = i + 1

    # === Phase 7: Tabellen Naturalisieren ===

    def _naturalize_tables(self, model: SlideModel):
        """Wandelt Tabellen in natürliche Sprache um."""
        # Wenn Docling-Analyse verfügbar: verbessere Tabellenstruktur zuerst
        if self._docling_analysis and self._docling_analysis.tables:
            self._apply_docling_table_structure(model)

        for slide in model.slides:
            for block in slide.blocks:
                if block.block_type != BlockType.TABLE or not block.table:
                    continue
                
                table = block.table
                natural_text = self._table_to_natural_language(table, slide)
                
                if natural_text:
                    if not hasattr(block, 'a11y'):
                        block.a11y = AccessibilityAnnotation(role=ElementRole.ESSENTIAL)
                    block.a11y.screen_reader_text = natural_text

    def _apply_docling_table_structure(self, model: SlideModel):
        """Wendet Docling's Tabellen-Strukturerkennung an."""
        try:
            from .docling_integration import apply_docling_table_structure
            apply_docling_table_structure(model, self._docling_analysis)
        except Exception as e:
            print(f"   Docling Tabellen-Struktur fehlgeschlagen: {e}")

    def _table_to_natural_language(self, table: Table, slide: Slide) -> Optional[str]:
        """Konvertiert Tabelle in natürliche Sprache."""
        if not table.rows:
            return None
        
        # Einfache Tabellen regelbasiert
        if len(table.rows) <= 4 and table.column_count <= 3:
            return self._simple_table_to_text(table)
        
        # Komplexe Tabellen mit KI
        if self._llm_available:
            return self._complex_table_to_text(table, slide)
        
        return self._simple_table_to_text(table)
    
    def _simple_table_to_text(self, table: Table) -> str:
        """Regelbasierte Tabellen-zu-Text Konvertierung."""
        lines = []
        
        if table.caption:
            lines.append(f"Tabelle: {table.caption}")
        
        headers = []
        if table.has_header and table.rows:
            headers = [cell.text.strip() for cell in table.rows[0]]
            lines.append(f"Spalten: {', '.join(headers)}")
        
        # Datenzeilen
        data_rows = table.rows[1:] if table.has_header else table.rows
        
        for row in data_rows:
            if headers:
                pairs = []
                for i, cell in enumerate(row):
                    header = headers[i] if i < len(headers) else f"Spalte {i+1}"
                    value = cell.text.strip()
                    if value:
                        pairs.append(f"{header}: {value}")
                if pairs:
                    lines.append("; ".join(pairs))
            else:
                values = [cell.text.strip() for cell in row if cell.text.strip()]
                if values:
                    lines.append(", ".join(values))
        
        return " | ".join(lines)
    
    def _complex_table_to_text(self, table: Table, slide: Slide) -> Optional[str]:
        """KI-basierte Tabellen-zu-Text Konvertierung."""
        # Tabelle als Text aufbereiten
        table_text = []
        for i, row in enumerate(table.rows):
            row_texts = [cell.text.strip() for cell in row]
            table_text.append(" | ".join(row_texts))
        
        table_str = "\n".join(table_text)
        
        prompt = f"""Du bist ein Accessibility-Experte.

Folientitel: "{slide.title or 'Ohne Titel'}"

Diese Tabelle soll für blinde Menschen vorgelesen werden:
{table_str}

Aufgabe: Fasse die KERNAUSSAGE der Tabelle in 2-3 natürlichen Sätzen zusammen.
Nenne die wichtigsten Datenpunkte. Vermeide "Die Tabelle zeigt...".

Beispiel guter Ausgabe: "Der Umsatz stieg von 5 Mio auf 8 Mio Euro zwischen Q1 und Q4. 
Der stärkste Monat war Oktober mit 2,3 Mio."

Deine Zusammenfassung:"""

        try:
            response = requests.post(
                f"{self.config.ollama_url}/api/generate",
                json={
                    "model": self.config.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {"temperature": 0.3, "num_predict": 200}
                },
                timeout=30
            )
            
            if response.status_code == 200:
                return response.json().get("response", "").strip()
        except:
            pass
        
        return None
    
    # === Phase 8: Charts ===
    
    def _describe_charts(self, model: SlideModel):
        """Beschreibt Charts/Diagramme für Screenreader."""
        for slide in model.slides:
            for block in slide.blocks:
                if block.block_type != BlockType.FIGURE or not block.figure:
                    continue
                
                fig = block.figure
                
                # Prüfe ob es ein Chart sein könnte
                if fig.alt_text and any(word in fig.alt_text.lower() 
                    for word in ['diagramm', 'chart', 'grafik', 'graph']):
                    
                    enhanced = self._enhance_chart_description(fig, slide)
                    if enhanced:
                        fig.alt_text = enhanced
    
    def _enhance_chart_description(self, figure: Figure, slide: Slide) -> Optional[str]:
        """Verbessert Chart-Beschreibung mit KI."""
        if not self._llm_available or not figure.image_data:
            return None
        
        prompt = f"""Du bist ein Accessibility-Experte für Datenvisualisierung.

Folientitel: "{slide.title or 'Ohne Titel'}"
Bisherige Beschreibung: "{figure.alt_text or 'Keine'}"

Analysiere dieses Diagramm/Chart und erstelle eine Beschreibung die:
1. Den Diagrammtyp nennt (Balken, Linie, Kreis, etc.)
2. Die KERNAUSSAGE in einem Satz formuliert
3. Die wichtigsten 2-3 Datenpunkte konkret nennt
4. Trends oder Vergleiche beschreibt

Format: Maximal 3 Sätze, direkt und informativ.

Beispiel: "Balkendiagramm zum Quartalsumsatz. Q4 erreichte mit 2,8 Mio den Höchstwert, 
ein Plus von 40% gegenüber Q1. Der Trend zeigt stetiges Wachstum."

Deine Beschreibung:"""

        try:
            image_b64 = base64.b64encode(figure.image_data).decode('utf-8')
            response = requests.post(
                f"{self.config.ollama_url}/api/generate",
                json={
                    "model": self.config.vision_model,
                    "prompt": prompt,
                    "images": [image_b64],
                    "stream": False,
                    "options": {"temperature": 0.3, "num_predict": 200}
                },
                timeout=60
            )
            
            if response.status_code == 200:
                return response.json().get("response", "").strip()
        except:
            pass
        
        return None
    
    # === Phase 9: Folien-Zusammenfassungen ===
    
    def _add_slide_summaries(self, model: SlideModel):
        """Fügt Zusammenfassungen für komplexe Folien hinzu."""
        for slide in model.slides:
            # Zähle "lesbare" Blöcke
            readable = [
                b for b in slide.blocks
                if not (hasattr(b, 'a11y') and b.a11y.role in 
                       [ElementRole.DECORATIVE, ElementRole.REDUNDANT])
            ]
            
            if len(readable) < self.config.complex_slide_threshold:
                continue
            
            summary = self._generate_slide_summary(slide, readable)
            
            if summary:
                # Füge Summary als ersten Block ein
                summary_block = Block(
                    block_type=BlockType.PARAGRAPH,
                    reading_order=0,
                    paragraphs=[Paragraph(runs=[
                        TextRun(text=f"Zusammenfassung dieser Folie: {summary}")
                    ])]
                )
                
                if not hasattr(summary_block, 'a11y'):
                    summary_block.a11y = AccessibilityAnnotation(
                        role=ElementRole.CONTEXTUAL
                    )
                
                slide.blocks.insert(0, summary_block)
    
    def _generate_slide_summary(self, slide: Slide, blocks: list[Block]) -> Optional[str]:
        """Generiert Zusammenfassung einer komplexen Folie."""
        if not self._llm_available:
            return None
        
        # Sammle Inhalte
        contents = []
        for block in blocks[:10]:  # Max 10 für Prompt-Länge
            if block.text:
                contents.append(block.text[:200])
        
        if not contents:
            return None
        
        content_str = "\n- ".join(contents)
        
        prompt = f"""Du bist ein Accessibility-Experte.

Diese Folie hat viele Elemente. Erstelle eine kurze Orientierung für blinde Nutzer.

Folientitel: "{slide.title or 'Ohne Titel'}"

Inhalte:
- {content_str}

Aufgabe: Fasse in EINEM Satz zusammen, worum es auf dieser Folie geht.
Der Satz soll helfen, die folgenden Details einzuordnen.

Beispiel: "Diese Folie vergleicht drei Produktvarianten nach Preis und Leistung."

Dein Satz:"""

        try:
            response = requests.post(
                f"{self.config.ollama_url}/api/generate",
                json={
                    "model": self.config.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {"temperature": 0.3, "num_predict": 100}
                },
                timeout=20
            )
            
            if response.status_code == 200:
                return response.json().get("response", "").strip()
        except:
            pass
        
        return None
    
    # === Phase 10: Finale Bereinigung ===
    
    def _final_cleanup(self, model: SlideModel):
        """Finale Bereinigung und Konsistenzprüfung."""
        # Rollen die komplett entfernt werden sollen
        skip_roles = {
            ElementRole.DECORATIVE,
            ElementRole.REDUNDANT,
            ElementRole.BOILERPLATE,
            ElementRole.PLACEHOLDER,
            ElementRole.NAVIGATION,
        }

        for slide in model.slides:
            # Entferne markierte und leere Blöcke
            slide.blocks = [
                b for b in slide.blocks
                if not (hasattr(b, 'a11y') and b.a11y and b.a11y.role in skip_roles)
                and (not b.is_empty or (hasattr(b, 'a11y') and b.a11y and b.a11y.screen_reader_text))
            ]

            # Sortiere nach Reading Order
            slide.blocks.sort(key=lambda b: b.reading_order)

            # Renummeriere
            for i, block in enumerate(slide.blocks):
                block.reading_order = i + 1
    
    def _print_stats(self, model: SlideModel):
        """Gibt Optimierungs-Statistiken aus."""
        stats = {
            "total_original": self._original_block_count if hasattr(self, '_original_block_count') else 0,
            "total_remaining": 0,
            "removed": {
                "decorative": 0,
                "redundant": 0,
                "boilerplate": 0,
                "placeholder": 0,
                "navigation": 0,
            },
            "enhanced": 0,
        }

        for slide in model.slides:
            stats["total_remaining"] += len(slide.blocks)
            for block in slide.blocks:
                if hasattr(block, 'a11y') and block.a11y:
                    if block.a11y.screen_reader_text:
                        stats["enhanced"] += 1

        removed_total = stats["total_original"] - stats["total_remaining"] if stats["total_original"] > 0 else 0

        print(f"\n   📊 Optimierungs-Statistik:")
        print(f"      Ursprüngliche Elemente: {stats['total_original']}")
        print(f"      Entfernt (unnötig): {removed_total}")
        print(f"      Verbleibende Elemente: {stats['total_remaining']}")
        print(f"      Mit optimiertem Text: {stats['enhanced']}")


# === Convenience Function ===

def optimize_for_screenreader(
    model: SlideModel,
    ollama_url: str = "http://localhost:11434",
    verbose: bool = True
) -> SlideModel:
    """
    Convenience-Funktion für Accessibility-Optimierung.
    
    Usage:
        model = parser.parse("slides.pptx")
        model = optimize_for_screenreader(model)
        renderer.render(model, "output.pdf")
    """
    config = AccessibilityConfig(ollama_url=ollama_url)
    optimizer = AccessibilityOptimizer(config)
    return optimizer.optimize(model, verbose=verbose)
