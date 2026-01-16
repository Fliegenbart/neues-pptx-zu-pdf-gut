# pptx2ua - PPTX zu PDF/UA Konverter

> DSGVO-konforme Konvertierung von PowerPoint zu barrierefreien PDFs nach PDF/UA-1 Standard.
> **Optimiert für Screenreader-UX** - nicht nur technische Compliance.

## 🎯 Philosophie

> "Blinde Menschen sollen den **INHALT** verstehen - nicht das Layout nachvollziehen."

Dieses Tool geht über technische PDF/UA-Compliance hinaus und optimiert aktiv für Screenreader-Nutzererlebnis:

| Standard-Ansatz | Unser Ansatz |
|-----------------|--------------|
| Fußnote¹ → "siehe Ende" | Fußnote inline: "...(Quelle: BMI 2024)" |
| Tabelle zeilenweise | "Umsatz stieg von 5 auf 8 Mio, Trend: positiv" |
| Jedes Bild beschreiben | Dekorative Bilder ausblenden |
| Logo auf jeder Folie | Nur einmal erwähnen |
| Layout-Reihenfolge | Verständnis-Reihenfolge |

## 🏗️ Architektur

```
┌─────────────────────────────────────────────────────────────────────┐
│                         Pipeline                                     │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  PPTX ──▶ Parser ──▶ SlideModel ──▶ Enricher ──▶ SlideModel        │
│              │                          │                           │
│              │                    ┌─────┴─────┐                     │
│              │                    │ Backends: │                     │
│              │                    │ • Ollama  │                     │
│              │                    │ • Docling │                     │
│              │                    └─────┬─────┘                     │
│              │                          │                           │
│              ▼                          ▼                           │
│         [Docling] ───────▶ Accessibility ──▶ SlideModel            │
│         (optional)         Optimizer                                │
│         • Reading Order        │                                    │
│         • Tabellen-Struktur    │                                    │
│                  ┌─────────────┴──────────────┐                     │
│                  │  • Dekoratives ausblenden   │                    │
│                  │  • Fußnoten inline          │                    │
│                  │  • Speaker Notes nutzen     │                    │
│                  │  • Redundanzen entfernen    │                    │
│                  │  • Tabellen naturalisieren  │                    │
│                  │  • Charts beschreiben       │                    │
│                  │  • Lesereihenfolge          │                    │
│                  └─────────────┬──────────────┘                     │
│                                │                                    │
│                                ▼                                    │
│                 SlideModel ──▶ Renderer ──▶ PDF/UA                  │
│                                    │                                │
│                                    ▼                                │
│                          PDF ──▶ Validator ──▶ Report               │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

## 📦 Module

| Modul | Datei | Beschreibung |
|-------|-------|--------------|
| **models** | `models.py` | Datenstrukturen (SlideModel, Block, Figure, etc.) |
| **parser** | `parser.py` | PPTX → SlideModel Extraktion |
| **enricher** | `enricher.py` | KI Alt-Text (Ollama oder Docling) |
| **docling_integration** | `docling_integration.py` | IBM Docling Integration (VLM, Reading Order, Tabellen) |
| **accessibility_optimizer** | `accessibility_optimizer.py` | **Screenreader-UX-Optimierung** |
| **renderer** | `renderer.py` | SlideModel → HTML → PDF/UA |
| **validator** | `validator.py` | PDF/UA Validierung mit veraPDF |
| **cli** | `cli.py` | Command Line Interface |

## ♿ Accessibility-Optimierungen

### 1. Dekorative Elemente erkennen
```python
# KI analysiert: Ist das Bild inhaltlich relevant?
# Hintergrundbilder, Linien, Icons → aria-hidden="true"
```

### 2. Fußnoten inline auflösen
```
Vorher:  "Laut Studie¹ steigt der Umsatz."
         ...
         "¹ BMI Jahresbericht 2024, S. 42"

Nachher: "Laut Studie (BMI Jahresbericht 2024, S. 42) steigt der Umsatz."
```

### 3. Speaker Notes als Kontext
```
Vorher:  [Folie wird ohne Kontext vorgelesen]

Nachher: "Kontext: Diese Folie zeigt warum Q4 entscheidend war."
         [Dann Folieninhalt]
```

### 4. Redundanzen entfernen
```
Vorher:  "Firmenlogo" (auf jeder der 30 Folien)

Nachher: "Firmenlogo" (nur auf Folie 1)
```

### 5. Tabellen in natürliche Sprache
```
Vorher:  "Zeile 1: Q1, 5 Mio. Zeile 2: Q2, 6 Mio. Zeile 3: Q3, 7 Mio..."

Nachher: "Der Umsatz stieg kontinuierlich von 5 Mio in Q1 auf 8 Mio in Q4,
          ein Wachstum von 60%. Der stärkste Sprung war zwischen Q3 und Q4."
```

### 6. Charts beschreiben
```
Vorher:  "Balkendiagramm"

Nachher: "Balkendiagramm zum Quartalsumsatz. Q4 erreichte mit 2,8 Mio 
          den Höchstwert, ein Plus von 40% gegenüber Q1."
```

### 7. Lesereihenfolge für Verständnis
```
Vorher:  Layout-Reihenfolge (links-oben → rechts-unten)

Nachher: 1. Titel
         2. Kontext (aus Speaker Notes)
         3. Erklärungstext
         4. Dann erst Bilder/Tabellen (mit Kontext)
```

## 🚀 Installation

```bash
# Repository klonen
git clone https://github.com/drv-rvevolution/pptx2ua.git
cd pptx2ua

# Virtual Environment (Python 3.10+)
python -m venv .venv
source .venv/bin/activate

# Basis-Installation
pip install -e ".[dev]"

# Mit Docling (empfohlen für beste Ergebnisse)
pip install -e ".[dev,docling]"
```

### Optionale Abhängigkeiten

| Komponente | Installation | Funktion |
|------------|--------------|----------|
| **Docling** | `pip install pptx2ua[docling]` | IBM GraniteDocling VLM, Reading Order, Tabellen-Struktur |
| **Ollama** | [ollama.ai](https://ollama.ai) | Lokale LLMs (llava, qwen2-vl) |
| **veraPDF** | [verapdf.org](https://verapdf.org) | PDF/UA Validierung |

### KI-Backend Vergleich

| Feature | Ollama | Docling |
|---------|--------|---------|
| Alt-Text Generierung | ✅ llava, qwen2-vl (empfohlen) | ⚠️ Experimentell |
| Reading Order | ❌ | ✅ (empfohlen) |
| Tabellen-Struktur | ❌ | ✅ (empfohlen) |
| Layout-Analyse | ❌ | ✅ |
| Setup-Aufwand | Mittel | Gering (pip) |
| Modellgröße | ~4GB | ~500MB-2GB |
| GPU empfohlen | Ja | Optional |
| DSGVO-konform | ✅ Lokal | ✅ Lokal |

**Empfehlung:** Beide kombinieren - Ollama für Alt-Texte, Docling für Dokumentstruktur.

## 🔧 Nutzung

### CLI

```bash
# Vollständige Pipeline (nutzt Docling wenn verfügbar, sonst Ollama)
pptx2ua convert presentation.pptx

# Nur Ollama verwenden (Docling deaktivieren)
pptx2ua convert presentation.pptx --no-docling

# Ohne KI (regelbasierte Optimierungen)
pptx2ua convert presentation.pptx --no-ai

# Ausgabe-Datei angeben
pptx2ua convert presentation.pptx -o barrierefreie_version.pdf

# Struktur inspizieren
pptx2ua inspect presentation.pptx

# PDF validieren
pptx2ua validate document.pdf

# JSON-Output für Automation
pptx2ua convert presentation.pptx --json
```

### Python API

```python
from pptx2ua import (
    PPTXParser,
    Enricher,
    EnricherConfig,
    EnricherBackend,
    AccessibilityOptimizer,
    PDFUARenderer
)

# 1. Parse
model = PPTXParser().parse("presentation.pptx")

# 2. Alt-Texte generieren (AUTO: Docling > Ollama)
enricher = Enricher()  # Wählt automatisch das beste Backend
if enricher.is_available:
    print(f"Nutze Backend: {enricher.active_backend}")
    model = enricher.enrich(model)

# 3. Accessibility optimieren (das Herzstück!)
optimizer = AccessibilityOptimizer()
model = optimizer.optimize(model)

# 4. Rendern
PDFUARenderer().render(model, "output.pdf")
```

### Backend explizit wählen

```python
from pptx2ua import Enricher, EnricherConfig, EnricherBackend

# Nur Docling
config = EnricherConfig(backend=EnricherBackend.DOCLING)
enricher = Enricher(config)

# Nur Ollama
config = EnricherConfig(backend=EnricherBackend.OLLAMA)
enricher = Enricher(config)

# Automatisch (Docling wenn verfügbar, sonst Ollama)
config = EnricherConfig(backend=EnricherBackend.AUTO)
enricher = Enricher(config)
```

### Docling direkt nutzen

```python
from pptx2ua.docling_integration import (
    DoclingAnalyzer,
    DoclingConfig,
    is_docling_available
)

if is_docling_available():
    analyzer = DoclingAnalyzer()

    # PPTX analysieren
    result = analyzer.analyze_pptx("presentation.pptx")

    # Reading Order
    print(f"Elemente: {len(result.reading_order)}")

    # Tabellen-Struktur
    print(f"Tabellen: {len(result.tables)}")

    # Alt-Text für einzelnes Bild
    with open("image.png", "rb") as f:
        alt_text = analyzer.generate_alt_text(f.read())
        print(f"Alt-Text: {alt_text}")
```

### Nur Accessibility-Optimierung

```python
from pptx2ua import optimize_for_screenreader

model = parser.parse("slides.pptx")
model = optimize_for_screenreader(model)  # Convenience-Funktion
```

## 🤖 KI-Einsatz

KI wird **gezielt** eingesetzt, nicht flächendeckend:

| Aufgabe | KI? | Warum |
|---------|-----|-------|
| Alt-Text für Fotos | ✅ | Nur KI kann "Was zeigt das Bild?" beantworten |
| Chart-Analyse | ✅ | Kernaussage aus Visualisierung extrahieren |
| Tabellen-Summary | ✅ | Trends und Muster erkennen |
| Dekorativ ja/nein? | ✅ | Bei Grenzfällen (Hintergrundbilder) |
| Speaker Notes Kontext | ✅ | Relevanten Kontext extrahieren |
| Heading-Erkennung | ❌ | Font-Size + Placeholder reichen |
| Listen-Erkennung | ❌ | PPTX hat explizite Marker |
| Fußnoten-Parsing | ❌ | Regex-Patterns reichen |

## 🧪 Entwicklung mit Claude Code

### Typische Aufgaben

```
"Verbessere die Tabellen-zu-Text Konvertierung für Pivot-Tabellen"
→ accessibility_optimizer.py: _table_to_natural_language()

"Erkenne SmartArt-Grafiken und beschreibe ihre Struktur"
→ parser.py: neuer _parse_smartart()
→ accessibility_optimizer.py: SmartArt-Handler

"Speaker Notes werden nicht richtig geparst"
→ parser.py: _parse_slide() Notes-Extraktion
```

## 🔒 DSGVO

- ✅ Alle KI lokal (Ollama + Docling)
- ✅ Keine Cloud-Dienste
- ✅ Keine Telemetrie
- ✅ Temp-Dateien gelöscht
- ✅ Docling von IBM Research, MIT-Lizenz
- ✅ Modelle werden lokal gespeichert (~/.cache/huggingface)

## 📄 Lizenz

MIT License
