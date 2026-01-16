"""
PDF/UA Validator
================
Validiert PDFs gegen PDF/UA-1 Standard mit veraPDF.

veraPDF ist der de-facto Standard für PDF/A und PDF/UA Validierung.
Open Source und von der PDF Association empfohlen.
"""

import json
import subprocess
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET


@dataclass
class ValidationIssue:
    """Ein einzelnes Validierungsproblem."""
    rule_id: str
    severity: str  # error, warning, info
    message: str
    specification: str  # z.B. "ISO 14289-1:2014"
    clause: str  # z.B. "7.1"
    test: str
    location: Optional[str] = None
    
    @property
    def is_error(self) -> bool:
        return self.severity.lower() == "error"
    
    @property
    def is_warning(self) -> bool:
        return self.severity.lower() == "warning"


@dataclass
class ValidationResult:
    """Ergebnis einer PDF/UA Validierung."""
    is_valid: bool
    is_compliant: bool  # Alle Regeln erfüllt
    
    issues: list[ValidationIssue] = field(default_factory=list)
    
    # Statistiken
    errors: int = 0
    warnings: int = 0
    passed_rules: int = 0
    failed_rules: int = 0
    
    # Metadaten
    pdf_version: str = ""
    is_tagged: bool = False
    has_language: bool = False
    has_title: bool = False
    
    # Rohdaten
    raw_output: str = ""
    
    @property
    def error_issues(self) -> list[ValidationIssue]:
        return [i for i in self.issues if i.is_error]
    
    @property
    def warning_issues(self) -> list[ValidationIssue]:
        return [i for i in self.issues if i.is_warning]
    
    def summary(self) -> str:
        """Kurze Zusammenfassung."""
        status = "✅ VALIDE" if self.is_compliant else "❌ NICHT VALIDE"
        return f"{status} | Fehler: {self.errors} | Warnungen: {self.warnings}"


class PDFUAValidator:
    """
    PDF/UA Validator mit veraPDF Backend.
    
    Installation veraPDF:
        # Download von https://verapdf.org/software/
        # Oder via Package Manager
        
    Usage:
        validator = PDFUAValidator()
        result = validator.validate("document.pdf")
        if not result.is_compliant:
            for issue in result.error_issues:
                print(f"- {issue.message}")
    """
    
    def __init__(self, verapdf_path: Optional[str] = None):
        """
        Args:
            verapdf_path: Pfad zur veraPDF CLI (oder None für Auto-Detect)
        """
        self.verapdf_path = verapdf_path or self._find_verapdf()
        self.available = self.verapdf_path is not None
    
    def _find_verapdf(self) -> Optional[str]:
        """Sucht veraPDF im System."""
        # Mögliche Namen
        names = ["verapdf", "verapdf.bat", "verapdf.sh"]
        
        for name in names:
            path = shutil.which(name)
            if path:
                return path
        
        # Typische Installationspfade
        common_paths = [
            "/opt/verapdf/verapdf",
            "/usr/local/bin/verapdf",
            "~/.local/bin/verapdf",
            "C:/Program Files/veraPDF/verapdf.bat",
        ]
        
        for p in common_paths:
            expanded = Path(p).expanduser()
            if expanded.exists():
                return str(expanded)
        
        return None
    
    def validate(self, pdf_path: Path | str, profile: str = "ua1") -> ValidationResult:
        """
        Validiert ein PDF gegen PDF/UA-1.
        
        Args:
            pdf_path: Pfad zum PDF
            profile: Validierungsprofil (ua1 = PDF/UA-1)
            
        Returns:
            ValidationResult mit allen Findings
        """
        pdf_path = Path(pdf_path)
        
        if not self.available:
            return self._fallback_validation(pdf_path)
        
        try:
            # veraPDF aufrufen
            result = subprocess.run(
                [
                    self.verapdf_path,
                    "--format", "mrr",  # Machine-Readable Report (XML)
                    "--profile", profile,
                    str(pdf_path)
                ],
                capture_output=True,
                text=True,
                timeout=120
            )
            
            return self._parse_verapdf_output(result.stdout, result.returncode)
            
        except subprocess.TimeoutExpired:
            return ValidationResult(
                is_valid=False,
                is_compliant=False,
                raw_output="Timeout bei Validierung"
            )
        except Exception as e:
            return ValidationResult(
                is_valid=False,
                is_compliant=False,
                raw_output=f"Fehler: {e}"
            )
    
    def _parse_verapdf_output(self, xml_output: str, return_code: int) -> ValidationResult:
        """Parst veraPDF XML Output."""
        result = ValidationResult(
            is_valid=True,
            is_compliant=(return_code == 0),
            raw_output=xml_output
        )
        
        try:
            root = ET.fromstring(xml_output)
            
            # Namespace handling
            ns = {
                'vera': 'http://www.verapdf.org/MachineReadableReport'
            }
            
            # Job-Ergebnis finden
            job = root.find('.//vera:job', ns) or root.find('.//job')
            if job is None:
                # Versuche ohne Namespace
                job = root
            
            # Validation Result
            val_result = job.find('.//vera:validationResult', ns) or job.find('.//validationResult')
            
            if val_result is not None:
                result.is_compliant = val_result.get('isCompliant', '').lower() == 'true'
                
                # Regeln zählen
                passed = val_result.find('.//vera:passedRules', ns) or val_result.find('.//passedRules')
                failed = val_result.find('.//vera:failedRules', ns) or val_result.find('.//failedRules')
                
                if passed is not None:
                    result.passed_rules = int(passed.text or 0)
                if failed is not None:
                    result.failed_rules = int(failed.text or 0)
            
            # Einzelne Issues extrahieren
            for assertion in root.iter():
                if 'assertion' in assertion.tag.lower() or 'rule' in assertion.tag.lower():
                    issue = self._parse_assertion(assertion, ns)
                    if issue:
                        result.issues.append(issue)
                        if issue.is_error:
                            result.errors += 1
                        elif issue.is_warning:
                            result.warnings += 1
            
            # Metadaten extrahieren
            self._extract_metadata(root, result, ns)
            
        except ET.ParseError as e:
            result.raw_output += f"\n\nXML Parse Error: {e}"
        
        return result
    
    def _parse_assertion(self, elem, ns: dict) -> Optional[ValidationIssue]:
        """Parst eine einzelne Assertion/Rule."""
        try:
            # Status prüfen
            status = elem.get('status', '') or elem.get('outcome', '')
            if status.lower() in ('passed', 'pass'):
                return None
            
            rule_id = elem.get('ruleId', '') or elem.get('id', '')
            
            # Details aus Kindelementen
            message = ""
            clause = ""
            test = ""
            
            for child in elem:
                tag = child.tag.lower()
                if 'message' in tag or 'description' in tag:
                    message = child.text or ""
                elif 'clause' in tag:
                    clause = child.text or ""
                elif 'test' in tag:
                    test = child.text or ""
            
            if not message:
                message = elem.get('message', '') or rule_id
            
            return ValidationIssue(
                rule_id=rule_id,
                severity="error" if status.lower() == 'failed' else "warning",
                message=message,
                specification="ISO 14289-1:2014",  # PDF/UA-1
                clause=clause,
                test=test
            )
            
        except Exception:
            return None
    
    def _extract_metadata(self, root, result: ValidationResult, ns: dict):
        """Extrahiert PDF Metadaten aus veraPDF Output."""
        # Suche nach Features/Metadata Info
        for elem in root.iter():
            tag = elem.tag.lower()
            text = (elem.text or "").lower()
            
            if 'tagged' in tag:
                result.is_tagged = text in ('true', 'yes', '1')
            elif 'language' in tag or 'lang' in tag:
                result.has_language = bool(elem.text and len(elem.text) > 0)
            elif 'title' in tag:
                result.has_title = bool(elem.text and len(elem.text) > 0)
            elif 'version' in tag and 'pdf' in tag:
                result.pdf_version = elem.text or ""
    
    def _fallback_validation(self, pdf_path: Path) -> ValidationResult:
        """
        Fallback-Validierung wenn veraPDF nicht verfügbar.
        
        Prüft grundlegende PDF/UA Anforderungen mit pikepdf.
        """
        result = ValidationResult(
            is_valid=True,
            is_compliant=False,  # Kann ohne veraPDF nicht garantiert werden
            raw_output="veraPDF nicht verfügbar - eingeschränkte Prüfung"
        )
        
        try:
            import pikepdf
            
            with pikepdf.open(pdf_path) as pdf:
                # 1. Getaggt?
                if '/MarkInfo' in pdf.Root:
                    mark_info = pdf.Root.MarkInfo
                    result.is_tagged = mark_info.get('/Marked', False)
                
                if not result.is_tagged:
                    result.issues.append(ValidationIssue(
                        rule_id="PDFUA-7.1",
                        severity="error",
                        message="PDF ist nicht als getaggt markiert",
                        specification="ISO 14289-1:2014",
                        clause="7.1",
                        test="MarkInfo.Marked == true"
                    ))
                    result.errors += 1
                
                # 2. Sprache?
                if '/Lang' in pdf.Root:
                    result.has_language = True
                else:
                    result.issues.append(ValidationIssue(
                        rule_id="PDFUA-7.2",
                        severity="error",
                        message="Dokumentsprache nicht definiert",
                        specification="ISO 14289-1:2014",
                        clause="7.2",
                        test="Document.Lang exists"
                    ))
                    result.errors += 1
                
                # 3. Titel?
                if pdf.docinfo.get('/Title'):
                    result.has_title = True
                else:
                    result.issues.append(ValidationIssue(
                        rule_id="PDFUA-7.3",
                        severity="warning",
                        message="Dokumenttitel nicht definiert",
                        specification="ISO 14289-1:2014",
                        clause="7.3",
                        test="Info.Title exists"
                    ))
                    result.warnings += 1
                
                # PDF Version
                result.pdf_version = f"PDF {pdf.pdf_version}"
                
                # Compliant wenn keine Errors
                result.is_compliant = (result.errors == 0)
                
        except ImportError:
            result.raw_output += "\npikepdf nicht installiert"
        except Exception as e:
            result.raw_output += f"\nFehler bei Fallback-Validierung: {e}"
        
        return result
    
    def print_report(self, result: ValidationResult, verbose: bool = True):
        """Gibt einen formatierten Report aus."""
        print("\n" + "="*60)
        print("PDF/UA Validierungsbericht")
        print("="*60)
        
        print(f"\nStatus: {result.summary()}")
        
        print(f"\nMetadaten:")
        print(f"  PDF Version: {result.pdf_version}")
        print(f"  Getaggt: {'✓' if result.is_tagged else '✗'}")
        print(f"  Sprache: {'✓' if result.has_language else '✗'}")
        print(f"  Titel: {'✓' if result.has_title else '✗'}")
        
        if result.error_issues:
            print(f"\n❌ Fehler ({result.errors}):")
            for issue in result.error_issues[:10]:  # Max 10
                print(f"  [{issue.rule_id}] {issue.message}")
                if verbose and issue.clause:
                    print(f"      Klausel: {issue.clause}")
        
        if result.warning_issues and verbose:
            print(f"\n⚠️  Warnungen ({result.warnings}):")
            for issue in result.warning_issues[:5]:  # Max 5
                print(f"  [{issue.rule_id}] {issue.message}")
        
        print("\n" + "="*60)


# === Installation Helper ===

def install_verapdf():
    """
    Hilft bei der veraPDF Installation.
    
    Gibt Instruktionen aus wie veraPDF installiert werden kann.
    """
    instructions = """
veraPDF Installation
====================

Option 1: Download (empfohlen)
------------------------------
1. https://verapdf.org/software/ besuchen
2. "veraPDF Installer" downloaden
3. Installer ausführen
4. Pfad zur PATH Variable hinzufügen

Option 2: Docker
----------------
docker pull verapdf/verapdf
docker run -v $(pwd):/data verapdf/verapdf /data/document.pdf

Option 3: Linux Package (Ubuntu/Debian)
---------------------------------------
# Snap
sudo snap install verapdf

# Oder manuell
wget https://software.verapdf.org/releases/verapdf-installer.zip
unzip verapdf-installer.zip
./verapdf-*/verapdf-install

Nach der Installation testen:
  verapdf --version
"""
    print(instructions)
