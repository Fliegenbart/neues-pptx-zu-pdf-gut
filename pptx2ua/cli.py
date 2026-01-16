#!/usr/bin/env python3
"""
PPTX to PDF/UA Converter - CLI
==============================

Usage:
    pptx2ua convert presentation.pptx
    pptx2ua convert presentation.pptx -o output.pdf
    pptx2ua convert presentation.pptx --no-ai
    pptx2ua validate document.pdf
    pptx2ua inspect presentation.pptx
"""

import argparse
import json
import sys
from pathlib import Path
from datetime import datetime

from .parser import PPTXParser
from .enricher import Enricher, EnricherConfig, EnricherBackend
from .renderer import PDFUARenderer, RendererConfig
from .validator import PDFUAValidator
from .accessibility_optimizer import AccessibilityOptimizer, AccessibilityConfig
from .models import SlideModel
from .slide_renderer import populate_slide_images, is_libreoffice_available


class Pipeline:
    """
    Hauptpipeline für die Konvertierung.

    PPTX → Parse → Enrich → Optimize → Render → Validate

    Unterstützt zwei AI-Backends:
    - Ollama (llava, qwen2-vl) - Standard
    - Docling (GraniteDocling) - Wenn installiert
    """

    def __init__(
        self,
        language: str = "de",
        ollama_url: str = "http://localhost:11434",
        vision_model: str = "llava:13b",
        enable_ai: bool = True,
        optimize_accessibility: bool = True,
        use_docling: bool = True,  # Docling bevorzugen wenn verfügbar
        verbose: bool = True
    ):
        self.language = language
        self.enable_ai = enable_ai
        self.optimize_accessibility = optimize_accessibility
        self.use_docling = use_docling
        self.verbose = verbose

        # Components
        self.parser = PPTXParser()

        if enable_ai:
            # Backend-Auswahl
            backend = EnricherBackend.AUTO if use_docling else EnricherBackend.OLLAMA

            self.enricher = Enricher(EnricherConfig(
                backend=backend,
                ollama_url=ollama_url,
                vision_model=vision_model,
                language=language,
            ))

            self.optimizer = AccessibilityOptimizer(AccessibilityConfig(
                ollama_url=ollama_url,
                vision_model=vision_model,
                language=language,
                use_docling=use_docling,
            ))
        else:
            self.enricher = None
            self.optimizer = None

        self.renderer = PDFUARenderer(RendererConfig())
        self.validator = PDFUAValidator()
    
    def convert(
        self, 
        input_pptx: Path, 
        output_pdf: Path,
        validate: bool = True
    ) -> dict:
        """
        Führt die komplette Konvertierung durch.
        
        Returns:
            Dictionary mit Ergebnis und Statistiken
        """
        result = {
            "success": False,
            "input": str(input_pptx),
            "output": str(output_pdf),
            "timestamp": datetime.now().isoformat(),
            "stats": {},
            "validation": None,
            "errors": []
        }
        
        try:
            # 1. Parse
            if self.verbose:
                print(f"\n{'='*60}")
                print("🔄 PPTX → PDF/UA Konverter")
                print(f"{'='*60}")
                print(f"\n📂 Eingabe: {input_pptx.name}")
                print("\n📊 Schritt 1/4: Parsing PPTX...")
            
            model = self.parser.parse(input_pptx)
            model.language = self.language
            
            result["stats"]["slides"] = model.slide_count
            result["stats"]["figures"] = len(model.all_figures)
            
            if self.verbose:
                print(f"   ✓ {model.slide_count} Folien gefunden")
                print(f"   ✓ {len(model.all_figures)} Abbildungen extrahiert")
            
            # 2. Enrich (AI Alt-Texts)
            if self.verbose:
                print("\n🤖 Schritt 2/5: KI Alt-Text-Generierung...")
            
            if self.enricher and self.enricher.is_available:
                model = self.enricher.enrich(model, verbose=self.verbose)
                result["stats"]["alt_texts_generated"] = self.enricher.stats["generated"]
                result["stats"]["alt_texts_cached"] = self.enricher.stats["from_cache"]
            elif self.enable_ai:
                if self.verbose:
                    print("   ⚠️  Ollama nicht verfügbar - überspringe KI")
                result["stats"]["alt_texts_generated"] = 0
            else:
                if self.verbose:
                    print("   ⏭️  KI deaktiviert")
            
            # 2b. Folienbilder für Vision-Analyse rendern
            if self.verbose:
                print("\n🖼️  Schritt 2b: Folienbilder rendern...")

            if is_libreoffice_available():
                success = populate_slide_images(model, input_pptx)
                if success and self.verbose:
                    slides_with_images = sum(1 for s in model.slides if s.slide_image)
                    print(f"   ✓ {slides_with_images} Folienbilder für Vision-Analyse")
            elif self.verbose:
                print("   ⚠️  LibreOffice nicht installiert - Vision-Analyse nur text-basiert")

            # 3. Accessibility Optimization
            if self.verbose:
                print("\n♿ Schritt 3/5: Accessibility-Optimierung...")

            if self.optimizer and self.optimize_accessibility:
                model = self.optimizer.optimize(model, verbose=self.verbose)
            elif self.optimize_accessibility and not self.optimizer:
                if self.verbose:
                    print("   ⚠️  Optimizer nicht verfügbar")
            else:
                if self.verbose:
                    print("   ⏭️  Optimierung deaktiviert")
            
            # 4. Render
            if self.verbose:
                print("\n🖨️  Schritt 4/5: PDF/UA Rendering...")
            
            success = self.renderer.render(model, output_pdf, verbose=self.verbose)
            
            if not success:
                result["errors"].append("PDF Rendering fehlgeschlagen")
                return result
            
            # 4. Validate
            if validate:
                if self.verbose:
                    print("\n✅ Schritt 5/5: PDF/UA Validierung...")
                
                val_result = self.validator.validate(output_pdf)
                
                result["validation"] = {
                    "compliant": val_result.is_compliant,
                    "errors": val_result.errors,
                    "warnings": val_result.warnings,
                    "is_tagged": val_result.is_tagged,
                    "has_language": val_result.has_language
                }
                
                if self.verbose:
                    self.validator.print_report(val_result, verbose=False)
            
            result["success"] = True
            
            if self.verbose:
                print(f"\n{'='*60}")
                print(f"✅ Konvertierung abgeschlossen!")
                print(f"📁 Ausgabe: {output_pdf}")
                print(f"{'='*60}\n")
            
        except Exception as e:
            result["errors"].append(str(e))
            if self.verbose:
                print(f"\n❌ Fehler: {e}")
                import traceback
                traceback.print_exc()
        
        return result


def cmd_convert(args):
    """Convert-Befehl."""
    input_path = Path(args.input)
    
    if not input_path.exists():
        print(f"❌ Datei nicht gefunden: {input_path}")
        return 1
    
    output_path = Path(args.output) if args.output else input_path.with_suffix('.pdf')
    
    pipeline = Pipeline(
        language=args.lang,
        ollama_url=args.ollama_url,
        vision_model=args.model,
        enable_ai=not args.no_ai,
        use_docling=not args.no_docling,
        verbose=not args.quiet
    )
    
    result = pipeline.convert(
        input_path,
        output_path,
        validate=not args.skip_validation
    )
    
    if args.json:
        print(json.dumps(result, indent=2, ensure_ascii=False))
    
    return 0 if result["success"] else 1


def cmd_validate(args):
    """Validate-Befehl."""
    pdf_path = Path(args.input)
    
    if not pdf_path.exists():
        print(f"❌ Datei nicht gefunden: {pdf_path}")
        return 1
    
    validator = PDFUAValidator()
    
    if not validator.available:
        print("⚠️  veraPDF nicht gefunden - eingeschränkte Validierung")
    
    result = validator.validate(pdf_path)
    validator.print_report(result, verbose=args.verbose)
    
    if args.json:
        # JSON Output
        data = {
            "compliant": result.is_compliant,
            "errors": result.errors,
            "warnings": result.warnings,
            "issues": [
                {
                    "rule_id": i.rule_id,
                    "severity": i.severity,
                    "message": i.message,
                    "clause": i.clause
                }
                for i in result.issues
            ]
        }
        print(json.dumps(data, indent=2, ensure_ascii=False))
    
    return 0 if result.is_compliant else 1


def cmd_inspect(args):
    """Inspect-Befehl - zeigt Struktur einer PPTX."""
    input_path = Path(args.input)
    
    if not input_path.exists():
        print(f"❌ Datei nicht gefunden: {input_path}")
        return 1
    
    parser = PPTXParser()
    model = parser.parse(input_path)
    
    print(f"\n📊 PPTX Struktur: {input_path.name}")
    print("="*60)
    print(f"Titel: {model.title or '(kein Titel)'}")
    print(f"Autor: {model.author or '(kein Autor)'}")
    print(f"Folien: {model.slide_count}")
    print(f"Abbildungen: {len(model.all_figures)}")
    print(f"  - Ohne Alt-Text: {len(model.figures_needing_alt_text)}")
    
    print("\n📑 Folien:")
    for slide in model.slides:
        print(f"\n  Folie {slide.number}: {slide.title or '(ohne Titel)'}")
        for block in slide.sorted_blocks:
            icon = {
                "heading": "📌",
                "paragraph": "📝",
                "list": "📋",
                "table": "📊",
                "figure": "🖼️",
            }.get(block.block_type.value, "▪️")
            
            preview = block.text[:50] + "..." if len(block.text) > 50 else block.text
            preview = preview.replace('\n', ' ')
            
            print(f"    {icon} {block.block_type.value}: {preview}")
    
    print("\n" + "="*60)
    return 0


def main():
    """CLI Haupteinstiegspunkt."""
    parser = argparse.ArgumentParser(
        prog="pptx2ua",
        description="PPTX zu PDF/UA Konverter (DSGVO-konform)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Beispiele:
  pptx2ua convert praesentation.pptx
  pptx2ua convert praesentation.pptx -o output.pdf --lang de
  pptx2ua convert praesentation.pptx --no-ai
  pptx2ua validate dokument.pdf
  pptx2ua inspect praesentation.pptx

Mehr Info: https://github.com/your-org/pptx2ua
        """
    )
    
    subparsers = parser.add_subparsers(dest="command", help="Verfügbare Befehle")
    
    # Convert
    convert_parser = subparsers.add_parser("convert", help="PPTX zu PDF/UA konvertieren")
    convert_parser.add_argument("input", help="Eingabe PPTX-Datei")
    convert_parser.add_argument("-o", "--output", help="Ausgabe PDF-Datei")
    convert_parser.add_argument("--lang", default="de", help="Dokumentsprache (default: de)")
    convert_parser.add_argument("--no-ai", action="store_true", help="Keine KI Alt-Text-Generierung")
    convert_parser.add_argument("--model", default="llava:13b", help="Vision-Modell (Ollama)")
    convert_parser.add_argument("--ollama-url", default="http://localhost:11434", help="Ollama URL")
    convert_parser.add_argument("--no-docling", action="store_true", help="Docling deaktivieren (nur Ollama)")
    convert_parser.add_argument("--skip-validation", action="store_true", help="Validierung überspringen")
    convert_parser.add_argument("-q", "--quiet", action="store_true", help="Keine Ausgabe")
    convert_parser.add_argument("--json", action="store_true", help="JSON Output")
    
    # Validate
    validate_parser = subparsers.add_parser("validate", help="PDF/UA validieren")
    validate_parser.add_argument("input", help="PDF-Datei")
    validate_parser.add_argument("-v", "--verbose", action="store_true", help="Ausführliche Ausgabe")
    validate_parser.add_argument("--json", action="store_true", help="JSON Output")
    
    # Inspect
    inspect_parser = subparsers.add_parser("inspect", help="PPTX Struktur anzeigen")
    inspect_parser.add_argument("input", help="PPTX-Datei")

    # Serve (Web Server)
    serve_parser = subparsers.add_parser("serve", help="Web-Server starten")
    serve_parser.add_argument("--port", type=int, default=3003, help="Port (default: 3003)")
    serve_parser.add_argument("--host", default="0.0.0.0", help="Host (default: 0.0.0.0)")

    args = parser.parse_args()

    if args.command == "convert":
        return cmd_convert(args)
    elif args.command == "validate":
        return cmd_validate(args)
    elif args.command == "inspect":
        return cmd_inspect(args)
    elif args.command == "serve":
        return cmd_serve(args)
    else:
        parser.print_help()
        return 0


def cmd_serve(args):
    """Serve-Befehl - startet Web-Server."""
    try:
        from .server import run_server
        run_server(host=args.host, port=args.port)
        return 0
    except ImportError as e:
        print(f"Fehler: Web-Server Dependencies nicht installiert: {e}")
        print("Installiere mit: pip install fastapi uvicorn python-multipart")
        return 1


if __name__ == "__main__":
    sys.exit(main())
