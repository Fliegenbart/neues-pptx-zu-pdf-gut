"""
KI-Enricher für Alt-Texte
=========================
Zwei-Stufen-System für qualitativ hochwertige Alt-Texte:

1. Vision-LLM: Erstellt Draft-Beschreibung
2. Text-LLM: Kürzt und poliert nach Barrierefreiheits-Regeln

Backends:
- Ollama (lokal, DSGVO-konform)
- Docling VLM (lokal, IBM GraniteDocling)

DSGVO-konform: Alles läuft lokal.
"""

import base64
import hashlib
import json
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Optional
import requests

from .models import SlideModel, Figure


class EnricherBackend(Enum):
    """Verfügbare Backends für Alt-Text-Generierung."""
    OLLAMA = "ollama"       # Ollama mit llava/qwen2-vl
    DOCLING = "docling"     # IBM Docling VLM
    AUTO = "auto"           # Automatische Auswahl (Docling > Ollama)


@dataclass
class EnricherConfig:
    """Konfiguration für den KI-Enricher."""
    # Backend-Auswahl
    backend: EnricherBackend = EnricherBackend.AUTO

    # Ollama-Einstellungen
    ollama_url: str = "http://localhost:11434"

    # Vision-Modell für Bildbeschreibung (Ollama)
    vision_model: str = "llava:13b"  # oder: qwen2-vl, bakllava

    # Text-Modell für Kürzung/Normalisierung (Ollama)
    text_model: str = "llama3.2:3b"  # oder: mistral, qwen2.5

    # Docling-Einstellungen
    docling_vlm_model: str = "granite_docling"

    # Sprache
    language: str = "de"

    # Cache-Verzeichnis für Bild-Hashes
    cache_dir: Optional[Path] = None

    # Timeouts
    vision_timeout: int = 120
    text_timeout: int = 30


class AltTextCache:
    """
    Cache für generierte Alt-Texte basierend auf Bild-Hash.
    
    Spart KI-Calls wenn identische Bilder mehrfach vorkommen
    (sehr häufig in Präsentationen: Logos, wiederkehrende Grafiken).
    """
    
    def __init__(self, cache_dir: Optional[Path] = None):
        self.cache_dir = cache_dir
        self._memory_cache: dict[str, str] = {}
        
        if cache_dir:
            cache_dir.mkdir(parents=True, exist_ok=True)
            self._load_disk_cache()
    
    def _load_disk_cache(self):
        """Lädt Cache von Disk."""
        if not self.cache_dir:
            return
        
        cache_file = self.cache_dir / "alt_text_cache.json"
        if cache_file.exists():
            try:
                self._memory_cache = json.loads(cache_file.read_text())
            except:
                pass
    
    def _save_disk_cache(self):
        """Speichert Cache auf Disk."""
        if not self.cache_dir:
            return
        
        cache_file = self.cache_dir / "alt_text_cache.json"
        cache_file.write_text(json.dumps(self._memory_cache, ensure_ascii=False, indent=2))
    
    def get(self, image_hash: str) -> Optional[str]:
        """Holt Alt-Text aus Cache."""
        return self._memory_cache.get(image_hash)
    
    def set(self, image_hash: str, alt_text: str):
        """Speichert Alt-Text im Cache."""
        self._memory_cache[image_hash] = alt_text
        self._save_disk_cache()
    
    @staticmethod
    def compute_hash(image_data: bytes) -> str:
        """Berechnet Hash für Bild-Daten."""
        return hashlib.md5(image_data).hexdigest()


class VisionLLM:
    """
    Vision-LLM Wrapper für Ollama.
    
    Unterstützte Modelle:
    - llava:7b / llava:13b (Standard, gut getestet)
    - bakllava (schneller, etwas weniger genau)
    - llama3.2-vision (neu, sehr gut)
    - qwen2-vl (sehr gut für Dokumente/Charts)
    """
    
    VISION_PROMPT_DE = """Beschreibe dieses Bild für sehbehinderte Menschen.

Regeln:
- Maximal 2 Sätze
- Beschreibe WAS zu sehen ist und WARUM es relevant ist
- Bei Diagrammen/Charts: Nenne den Typ und die Kernaussage
- Bei Fotos: Beschreibe Motiv und Kontext
- Bei Logos/Icons: Nenne was es darstellt
- WICHTIG: Wenn du etwas nicht sicher erkennen kannst, sage es ehrlich

Antworte NUR mit der Beschreibung, ohne Einleitung."""

    VISION_PROMPT_EN = """Describe this image for visually impaired users.

Rules:
- Maximum 2 sentences
- Describe WHAT is shown and WHY it's relevant
- For charts/diagrams: State the type and key message
- For photos: Describe subject and context
- For logos/icons: State what it represents
- IMPORTANT: If uncertain about something, say so honestly

Reply ONLY with the description, no introduction."""

    def __init__(self, config: EnricherConfig):
        self.config = config
        self.available = self._check_availability()
    
    def _check_availability(self) -> bool:
        """Prüft ob Ollama erreichbar ist und Modell geladen."""
        try:
            response = requests.get(
                f"{self.config.ollama_url}/api/tags", 
                timeout=5
            )
            if response.status_code != 200:
                return False
            
            # Prüfe ob Vision-Modell verfügbar
            models = response.json().get("models", [])
            model_names = [m.get("name", "") for m in models]
            
            # Check mit und ohne Tag
            base_model = self.config.vision_model.split(":")[0]
            return any(
                base_model in name 
                for name in model_names
            )
        except Exception:
            return False
    
    def generate_description(self, image_data: bytes) -> Optional[str]:
        """
        Generiert Bildbeschreibung via Vision-LLM.
        
        Args:
            image_data: Rohe Bild-Bytes
            
        Returns:
            Beschreibung oder None bei Fehler
        """
        if not self.available:
            return None
        
        # Bild zu Base64
        image_b64 = base64.b64encode(image_data).decode('utf-8')
        
        # Prompt basierend auf Sprache
        prompt = (
            self.VISION_PROMPT_DE 
            if self.config.language == "de" 
            else self.VISION_PROMPT_EN
        )
        
        payload = {
            "model": self.config.vision_model,
            "prompt": prompt,
            "images": [image_b64],
            "stream": False,
            "options": {
                "temperature": 0.3,  # Niedrig für Konsistenz
                "num_predict": 200,  # Kurze Antworten
            }
        }
        
        try:
            response = requests.post(
                f"{self.config.ollama_url}/api/generate",
                json=payload,
                timeout=self.config.vision_timeout
            )
            
            if response.status_code == 200:
                result = response.json()
                return result.get("response", "").strip()
                
        except requests.Timeout:
            print(f"⚠️  Vision-LLM Timeout ({self.config.vision_timeout}s)")
        except Exception as e:
            print(f"⚠️  Vision-LLM Fehler: {e}")
        
        return None


class TextLLM:
    """
    Text-LLM für Nachbearbeitung der Alt-Texte.
    
    Aufgaben:
    - Kürzen auf optimale Länge
    - Entfernen von KI-typischen Phrasen
    - Konsistente Formatierung
    - Qualitätsprüfung
    """
    
    POLISH_PROMPT_DE = """Du bist ein Experte für barrierefreie Alt-Texte.

Aufgabe: Optimiere diese Bildbeschreibung für Screenreader.

Original: "{draft}"

Regeln:
1. Maximal 1-2 Sätze (unter 125 Zeichen ideal)
2. Entferne Phrasen wie "Das Bild zeigt", "Zu sehen ist", "Dies ist"
3. Beginne direkt mit dem Inhalt
4. Behalte die Kernaussage
5. Wenn das Original "unsicher" oder "unklar" enthält: behalte diese Ehrlichkeit

Antworte NUR mit dem optimierten Alt-Text."""

    POLISH_PROMPT_EN = """You are an expert in accessible alt texts.

Task: Optimize this image description for screen readers.

Original: "{draft}"

Rules:
1. Maximum 1-2 sentences (under 125 characters ideal)
2. Remove phrases like "The image shows", "This is", "We can see"
3. Start directly with the content
4. Keep the key message
5. If original contains "unclear" or "uncertain": keep that honesty

Reply ONLY with the optimized alt text."""

    def __init__(self, config: EnricherConfig):
        self.config = config
        self.available = self._check_availability()
    
    def _check_availability(self) -> bool:
        """Prüft ob Text-Modell verfügbar."""
        try:
            response = requests.get(
                f"{self.config.ollama_url}/api/tags", 
                timeout=5
            )
            if response.status_code != 200:
                return False
            
            models = response.json().get("models", [])
            model_names = [m.get("name", "") for m in models]
            base_model = self.config.text_model.split(":")[0]
            
            return any(base_model in name for name in model_names)
        except:
            return False
    
    def polish(self, draft: str) -> str:
        """
        Poliert einen Draft-Alt-Text.
        
        Falls LLM nicht verfügbar: Regelbasierte Bereinigung.
        """
        if not draft:
            return ""
        
        # Versuche LLM-Polishing
        if self.available:
            polished = self._llm_polish(draft)
            if polished:
                return polished
        
        # Fallback: Regelbasiert
        return self._rule_based_polish(draft)
    
    def _llm_polish(self, draft: str) -> Optional[str]:
        """LLM-basiertes Polishing."""
        prompt_template = (
            self.POLISH_PROMPT_DE 
            if self.config.language == "de" 
            else self.POLISH_PROMPT_EN
        )
        
        payload = {
            "model": self.config.text_model,
            "prompt": prompt_template.format(draft=draft),
            "stream": False,
            "options": {
                "temperature": 0.2,
                "num_predict": 150,
            }
        }
        
        try:
            response = requests.post(
                f"{self.config.ollama_url}/api/generate",
                json=payload,
                timeout=self.config.text_timeout
            )
            
            if response.status_code == 200:
                result = response.json()
                polished = result.get("response", "").strip()
                
                # Sanity Check
                if polished and len(polished) > 5:
                    return polished
                    
        except Exception as e:
            print(f"⚠️  Text-LLM Fehler: {e}")
        
        return None
    
    def _rule_based_polish(self, draft: str) -> str:
        """Regelbasierte Bereinigung ohne LLM."""
        text = draft.strip()
        
        # Entferne typische Präfixe
        prefixes_de = [
            "Das Bild zeigt ",
            "Auf dem Bild ist ",
            "Zu sehen ist ",
            "Dieses Bild zeigt ",
            "Die Abbildung zeigt ",
            "Es ist ",
            "Es zeigt ",
        ]
        prefixes_en = [
            "The image shows ",
            "This image shows ",
            "The picture shows ",
            "We can see ",
            "This is ",
            "It shows ",
        ]
        
        prefixes = prefixes_de + prefixes_en
        
        for prefix in prefixes:
            if text.lower().startswith(prefix.lower()):
                text = text[len(prefix):]
                break
        
        # Erster Buchstabe groß
        if text:
            text = text[0].upper() + text[1:]
        
        # Punkt am Ende falls keiner
        if text and text[-1] not in '.!?':
            text += '.'
        
        return text


class Enricher:
    """
    Hauptklasse für KI-basierte Anreicherung.

    Unterstützt zwei Backends:
    - Ollama (llava, qwen2-vl)
    - Docling VLM (GraniteDocling)

    Usage:
        enricher = Enricher()
        enricher.enrich(slide_model)
    """

    def __init__(self, config: Optional[EnricherConfig] = None):
        self.config = config or EnricherConfig()

        self.cache = AltTextCache(self.config.cache_dir)

        # Backend initialisieren basierend auf Konfiguration
        self._active_backend: Optional[str] = None
        self._docling_enricher = None
        self.vision = None
        self.text = None

        self._init_backend()

        # Stats
        self.stats = {
            "processed": 0,
            "from_cache": 0,
            "generated": 0,
            "failed": 0,
            "backend": self._active_backend,
        }

    def _init_backend(self):
        """Initialisiert das passende Backend."""
        backend = self.config.backend

        if backend == EnricherBackend.AUTO:
            # Versuche zuerst Docling, dann Ollama
            if self._try_init_docling():
                return
            self._try_init_ollama()

        elif backend == EnricherBackend.DOCLING:
            if not self._try_init_docling():
                print("⚠️  Docling nicht verfügbar, falle zurück auf Ollama")
                self._try_init_ollama()

        elif backend == EnricherBackend.OLLAMA:
            self._try_init_ollama()

    def _try_init_docling(self) -> bool:
        """Versucht Docling zu initialisieren."""
        try:
            from .docling_integration import DoclingEnricher, DoclingConfig, is_docling_available

            if not is_docling_available():
                return False

            docling_config = DoclingConfig(
                vlm_model=self.config.docling_vlm_model,
                language=self.config.language,
            )
            self._docling_enricher = DoclingEnricher(docling_config)

            if self._docling_enricher.is_available:
                self._active_backend = "docling"
                return True

        except ImportError:
            pass

        return False

    def _try_init_ollama(self) -> bool:
        """Initialisiert Ollama Backend."""
        self.vision = VisionLLM(self.config)
        self.text = TextLLM(self.config)

        if self.vision.available:
            self._active_backend = "ollama"
            return True

        return False

    @property
    def is_available(self) -> bool:
        """Prüft ob KI-Enrichment verfügbar ist."""
        if self._active_backend == "docling":
            return self._docling_enricher is not None and self._docling_enricher.is_available
        elif self._active_backend == "ollama":
            return self.vision is not None and self.vision.available
        return False

    @property
    def active_backend(self) -> Optional[str]:
        """Gibt das aktive Backend zurück."""
        return self._active_backend
    
    def enrich(
        self,
        model: SlideModel,
        verbose: bool = True
    ) -> SlideModel:
        """
        Reichert SlideModel mit KI-generierten Alt-Texten an.

        Args:
            model: Das zu bearbeitende SlideModel
            verbose: Fortschrittsausgabe

        Returns:
            Das angereicherte SlideModel (in-place modifiziert)
        """
        # Bei Docling-Backend: Delegiere komplett
        if self._active_backend == "docling" and self._docling_enricher:
            return self._docling_enricher.enrich(model, verbose=verbose)

        # Ollama-Backend
        figures_to_process = model.figures_needing_alt_text

        if not figures_to_process:
            if verbose:
                print("   Alle Bilder haben bereits Alt-Texte")
            return model

        backend_info = f" (Backend: {self._active_backend or 'keins'})"
        if verbose:
            print(f"🤖 Generiere Alt-Texte für {len(figures_to_process)} Bilder...{backend_info}")

        for slide_num, figure in figures_to_process:
            self.stats["processed"] += 1

            # Kontext für bessere Beschreibungen
            slide = next((s for s in model.slides if s.number == slide_num), None)
            context = slide.title if slide else None

            alt_text = self._generate_alt_text(figure, context=context)

            if alt_text:
                figure.alt_text = alt_text
                figure.needs_alt_text = False

                if verbose:
                    preview = alt_text[:60] + "..." if len(alt_text) > 60 else alt_text
                    print(f"   Folie {slide_num}: \"{preview}\"")
            else:
                self.stats["failed"] += 1
                if verbose:
                    print(f"   Folie {slide_num}: Fehlgeschlagen")

        if verbose:
            self._print_stats()

        return model
    
    def _generate_alt_text(
        self,
        figure: Figure,
        context: Optional[str] = None
    ) -> Optional[str]:
        """Generiert Alt-Text für eine Figure mit Ollama."""
        if not figure.image_data:
            return None

        if not self.vision:
            return None

        # 1. Cache prüfen
        if figure.image_hash:
            cached = self.cache.get(figure.image_hash)
            if cached:
                self.stats["from_cache"] += 1
                return cached
        else:
            figure.image_hash = AltTextCache.compute_hash(figure.image_data)
            cached = self.cache.get(figure.image_hash)
            if cached:
                self.stats["from_cache"] += 1
                return cached

        # 2. Vision-LLM: Draft generieren
        draft = self.vision.generate_description(figure.image_data)

        if not draft:
            return None

        # 3. Text-LLM: Polieren
        polished = self.text.polish(draft) if self.text else draft

        if polished:
            self.stats["generated"] += 1
            # Im Cache speichern
            self.cache.set(figure.image_hash, polished)
            figure.alt_text_confidence = 0.8  # Gute Confidence bei 2-Stufen-Prozess
            return polished
        
        return draft  # Fallback auf unpolierten Draft
    
    def _print_stats(self):
        """Gibt Statistiken aus."""
        print(f"\n   📊 Alt-Text Statistik:")
        print(f"      Verarbeitet: {self.stats['processed']}")
        print(f"      Aus Cache: {self.stats['from_cache']}")
        print(f"      Neu generiert: {self.stats['generated']}")
        print(f"      Fehlgeschlagen: {self.stats['failed']}")


# === Convenience Functions ===

def enrich_model(
    model: SlideModel,
    ollama_url: str = "http://localhost:11434",
    vision_model: str = "llava:13b",
    language: str = "de",
    verbose: bool = True
) -> SlideModel:
    """
    Convenience-Funktion für Alt-Text-Generierung.
    
    Usage:
        model = parser.parse("slides.pptx")
        model = enrich_model(model)
    """
    config = EnricherConfig(
        ollama_url=ollama_url,
        vision_model=vision_model,
        language=language,
    )
    
    enricher = Enricher(config)
    return enricher.enrich(model, verbose=verbose)
