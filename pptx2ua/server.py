"""
PPTX2UA Web Server
==================
Lokaler Webserver für PPTX zu PDF/UA Konvertierung.

Usage:
    python -m pptx2ua.server --port 3003

Oder:
    pptx2ua serve --port 3003
"""

import asyncio
import tempfile
import shutil
from pathlib import Path
from typing import Optional
import logging

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from .parser import PPTXParser
from .enricher import Enricher, EnricherConfig, EnricherBackend
from .renderer import PDFUARenderer, RendererConfig
from .validator import PDFUAValidator
from .accessibility_optimizer import AccessibilityOptimizer, AccessibilityConfig
from .slide_renderer import populate_slide_images, is_libreoffice_available

# Logger
logger = logging.getLogger(__name__)

# FastAPI App
app = FastAPI(
    title="PPTX2UA Converter",
    description="DSGVO-konforme Konvertierung von PowerPoint zu barrierefreien PDFs",
    version="0.1.0",
)

# CORS für lokale Entwicklung
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Temp-Verzeichnis für Uploads
UPLOAD_DIR = Path(tempfile.gettempdir()) / "pptx2ua_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)


# === HTML UI ===

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PPTX2UA Converter</title>
    <style>
        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            min-height: 100vh;
            color: #e0e0e0;
            padding: 2rem;
        }

        .container {
            max-width: 800px;
            margin: 0 auto;
        }

        h1 {
            text-align: center;
            margin-bottom: 0.5rem;
            font-size: 2.5rem;
            background: linear-gradient(90deg, #00d4ff, #7b2cbf);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .subtitle {
            text-align: center;
            color: #888;
            margin-bottom: 2rem;
        }

        .card {
            background: rgba(255, 255, 255, 0.05);
            border-radius: 16px;
            padding: 2rem;
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.1);
            margin-bottom: 1.5rem;
        }

        .upload-zone {
            border: 2px dashed rgba(255, 255, 255, 0.2);
            border-radius: 12px;
            padding: 3rem;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
        }

        .upload-zone:hover, .upload-zone.dragover {
            border-color: #00d4ff;
            background: rgba(0, 212, 255, 0.05);
        }

        .upload-zone svg {
            width: 64px;
            height: 64px;
            margin-bottom: 1rem;
            opacity: 0.5;
        }

        .upload-zone p {
            color: #888;
        }

        .upload-zone .filename {
            color: #00d4ff;
            font-weight: bold;
            margin-top: 1rem;
        }

        input[type="file"] {
            display: none;
        }

        .options {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1rem;
            margin: 1.5rem 0;
        }

        .option {
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .option input[type="checkbox"] {
            width: 20px;
            height: 20px;
            accent-color: #00d4ff;
        }

        .option label {
            cursor: pointer;
        }

        select {
            background: rgba(255, 255, 255, 0.1);
            border: 1px solid rgba(255, 255, 255, 0.2);
            color: #e0e0e0;
            padding: 0.5rem 1rem;
            border-radius: 8px;
            width: 100%;
        }

        button {
            background: linear-gradient(90deg, #00d4ff, #7b2cbf);
            color: white;
            border: none;
            padding: 1rem 2rem;
            border-radius: 8px;
            font-size: 1.1rem;
            cursor: pointer;
            width: 100%;
            transition: transform 0.2s, box-shadow 0.2s;
        }

        button:hover:not(:disabled) {
            transform: translateY(-2px);
            box-shadow: 0 10px 30px rgba(0, 212, 255, 0.3);
        }

        button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        .progress {
            display: none;
            margin-top: 1.5rem;
        }

        .progress.active {
            display: block;
        }

        .progress-bar {
            height: 8px;
            background: rgba(255, 255, 255, 0.1);
            border-radius: 4px;
            overflow: hidden;
        }

        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #00d4ff, #7b2cbf);
            width: 0%;
            transition: width 0.3s;
            animation: pulse 1.5s ease-in-out infinite;
        }

        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.7; }
        }

        .progress-text {
            text-align: center;
            margin-top: 0.5rem;
            color: #888;
        }

        .result {
            display: none;
            text-align: center;
            padding: 2rem;
        }

        .result.active {
            display: block;
        }

        .result.success {
            color: #4ade80;
        }

        .result.error {
            color: #f87171;
        }

        .result svg {
            width: 64px;
            height: 64px;
            margin-bottom: 1rem;
        }

        .download-btn {
            display: inline-block;
            margin-top: 1rem;
            padding: 0.75rem 1.5rem;
            background: #4ade80;
            color: #1a1a2e;
            text-decoration: none;
            border-radius: 8px;
            font-weight: bold;
        }

        .stats {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
            gap: 1rem;
            margin-top: 1.5rem;
        }

        .stat {
            text-align: center;
            padding: 1rem;
            background: rgba(255, 255, 255, 0.05);
            border-radius: 8px;
        }

        .stat-value {
            font-size: 1.5rem;
            font-weight: bold;
            color: #00d4ff;
        }

        .stat-label {
            font-size: 0.8rem;
            color: #888;
        }

        .footer {
            text-align: center;
            margin-top: 2rem;
            color: #666;
            font-size: 0.9rem;
        }

        .footer a {
            color: #00d4ff;
            text-decoration: none;
        }

        .badge {
            display: inline-block;
            padding: 0.25rem 0.5rem;
            background: rgba(0, 212, 255, 0.2);
            color: #00d4ff;
            border-radius: 4px;
            font-size: 0.8rem;
            margin-left: 0.5rem;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>PPTX2UA</h1>
        <p class="subtitle">PowerPoint zu barrierefreiem PDF/UA konvertieren</p>

        <div class="card">
            <form id="convertForm" enctype="multipart/form-data">
                <div class="upload-zone" id="dropZone">
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                    </svg>
                    <p>PPTX-Datei hierher ziehen oder klicken</p>
                    <p class="filename" id="filename"></p>
                    <input type="file" id="fileInput" name="file" accept=".pptx">
                </div>

                <div class="options">
                    <div class="option">
                        <input type="checkbox" id="enableAi" name="enable_ai" checked>
                        <label for="enableAi">KI Alt-Texte generieren</label>
                    </div>
                    <div class="option">
                        <input type="checkbox" id="useDocling" name="use_docling" checked>
                        <label for="useDocling">Docling nutzen <span class="badge">Empfohlen</span></label>
                    </div>
                    <div class="option">
                        <input type="checkbox" id="validate" name="validate" checked>
                        <label for="validate">PDF/UA validieren</label>
                    </div>
                </div>

                <div class="option" style="margin-bottom: 1.5rem;">
                    <label for="language" style="margin-right: 1rem;">Sprache:</label>
                    <select id="language" name="language" style="width: auto;">
                        <option value="de" selected>Deutsch</option>
                        <option value="en">English</option>
                    </select>
                </div>

                <button type="submit" id="convertBtn" disabled>Konvertieren</button>
            </form>

            <div class="progress" id="progress">
                <div class="progress-bar">
                    <div class="progress-fill" id="progressFill"></div>
                </div>
                <p class="progress-text" id="progressText">Wird verarbeitet...</p>
            </div>

            <div class="result" id="result">
                <div id="resultContent"></div>
            </div>
        </div>

        <div class="card">
            <h3 style="margin-bottom: 1rem;">Status</h3>
            <div class="stats" id="statusStats">
                <div class="stat">
                    <div class="stat-value" id="doclingStatus">-</div>
                    <div class="stat-label">Docling</div>
                </div>
                <div class="stat">
                    <div class="stat-value" id="ollamaStatus">-</div>
                    <div class="stat-label">Ollama</div>
                </div>
                <div class="stat">
                    <div class="stat-value" id="verapdfStatus">-</div>
                    <div class="stat-label">veraPDF</div>
                </div>
            </div>
        </div>

        <p class="footer">
            DSGVO-konform - Alle Daten werden lokal verarbeitet<br>
            <a href="https://github.com/drv-rvevolution/pptx2ua" target="_blank">GitHub</a>
        </p>
    </div>

    <script>
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const filename = document.getElementById('filename');
        const convertBtn = document.getElementById('convertBtn');
        const convertForm = document.getElementById('convertForm');
        const progress = document.getElementById('progress');
        const progressFill = document.getElementById('progressFill');
        const progressText = document.getElementById('progressText');
        const result = document.getElementById('result');
        const resultContent = document.getElementById('resultContent');

        // Check status on load
        fetch('/api/status')
            .then(r => r.json())
            .then(data => {
                document.getElementById('doclingStatus').textContent = data.docling ? '✓' : '✗';
                document.getElementById('doclingStatus').style.color = data.docling ? '#4ade80' : '#f87171';
                document.getElementById('ollamaStatus').textContent = data.ollama ? '✓' : '✗';
                document.getElementById('ollamaStatus').style.color = data.ollama ? '#4ade80' : '#f87171';
                document.getElementById('verapdfStatus').textContent = data.verapdf ? '✓' : '✗';
                document.getElementById('verapdfStatus').style.color = data.verapdf ? '#4ade80' : '#f87171';
            });

        // Drag & Drop
        dropZone.addEventListener('click', () => fileInput.click());

        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('dragover');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0 && files[0].name.endsWith('.pptx')) {
                fileInput.files = files;
                updateFilename(files[0].name);
            }
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                updateFilename(e.target.files[0].name);
            }
        });

        function updateFilename(name) {
            filename.textContent = name;
            convertBtn.disabled = false;
        }

        // Form submit
        convertForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const formData = new FormData(convertForm);

            // Reset UI
            result.classList.remove('active', 'success', 'error');
            progress.classList.add('active');
            progressFill.style.width = '10%';
            progressText.textContent = 'Datei wird hochgeladen...';
            convertBtn.disabled = true;

            try {
                // Simulate progress
                let progressValue = 10;
                const progressInterval = setInterval(() => {
                    if (progressValue < 90) {
                        progressValue += Math.random() * 10;
                        progressFill.style.width = progressValue + '%';

                        if (progressValue < 30) progressText.textContent = 'Parsing PPTX...';
                        else if (progressValue < 50) progressText.textContent = 'Generiere Alt-Texte...';
                        else if (progressValue < 70) progressText.textContent = 'Accessibility-Optimierung...';
                        else progressText.textContent = 'Rendere PDF/UA...';
                    }
                }, 500);

                const response = await fetch('/api/convert', {
                    method: 'POST',
                    body: formData
                });

                clearInterval(progressInterval);
                progressFill.style.width = '100%';

                const data = await response.json();

                progress.classList.remove('active');
                result.classList.add('active');

                if (data.success) {
                    result.classList.add('success');
                    resultContent.innerHTML = `
                        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                        </svg>
                        <h3>Konvertierung erfolgreich!</h3>
                        <p>PDF/UA Status: ${data.validation?.compliant ? '✓ Valide' : 'Nicht validiert'}</p>
                        <div class="stats">
                            <div class="stat">
                                <div class="stat-value">${data.stats?.slides || '-'}</div>
                                <div class="stat-label">Folien</div>
                            </div>
                            <div class="stat">
                                <div class="stat-value">${data.stats?.figures || '-'}</div>
                                <div class="stat-label">Bilder</div>
                            </div>
                        </div>
                        <a href="/api/download/${data.download_id}" class="download-btn" download>PDF herunterladen</a>
                    `;
                } else {
                    result.classList.add('error');
                    resultContent.innerHTML = `
                        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                        </svg>
                        <h3>Fehler bei der Konvertierung</h3>
                        <p>${data.error || 'Unbekannter Fehler'}</p>
                    `;
                }
            } catch (err) {
                progress.classList.remove('active');
                result.classList.add('active', 'error');
                resultContent.innerHTML = `
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                    <h3>Verbindungsfehler</h3>
                    <p>${err.message}</p>
                `;
            }

            convertBtn.disabled = false;
        });
    </script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def index():
    """Startseite mit Upload-UI."""
    return HTML_TEMPLATE


@app.get("/api/status")
async def get_status():
    """Gibt den Status der verfügbaren Backends zurück."""
    # Check Docling
    docling_available = False
    try:
        from .docling_integration import is_docling_available
        docling_available = is_docling_available()
    except:
        pass

    # Check Ollama
    ollama_available = False
    try:
        import requests
        r = requests.get("http://localhost:11434/api/tags", timeout=2)
        ollama_available = r.status_code == 200
    except:
        pass

    # Check veraPDF
    verapdf_available = False
    try:
        validator = PDFUAValidator()
        verapdf_available = validator.available
    except:
        pass

    return {
        "docling": docling_available,
        "ollama": ollama_available,
        "verapdf": verapdf_available,
    }


@app.post("/api/convert")
async def convert_pptx(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    enable_ai: bool = Form(True),
    use_docling: bool = Form(True),
    validate: bool = Form(True),
    language: str = Form("de"),
):
    """Konvertiert eine PPTX-Datei zu PDF/UA."""

    # Validierung
    if not file.filename.endswith('.pptx'):
        raise HTTPException(status_code=400, detail="Nur PPTX-Dateien erlaubt")

    # Temp-Dateien
    import uuid
    job_id = str(uuid.uuid4())[:8]
    input_path = UPLOAD_DIR / f"{job_id}_input.pptx"
    output_path = UPLOAD_DIR / f"{job_id}_output.pdf"

    try:
        # Datei speichern
        with open(input_path, "wb") as f:
            content = await file.read()
            f.write(content)

        # Pipeline ausführen
        result = run_conversion(
            input_path=input_path,
            output_path=output_path,
            enable_ai=enable_ai,
            use_docling=use_docling,
            validate=validate,
            language=language,
        )

        # Aufräumen der Input-Datei nach 5 Minuten
        background_tasks.add_task(cleanup_file, input_path, delay=300)

        # Output-Datei nach 1 Stunde aufräumen
        background_tasks.add_task(cleanup_file, output_path, delay=3600)

        return {
            "success": True,
            "download_id": job_id,
            "stats": result.get("stats", {}),
            "validation": result.get("validation", {}),
        }

    except Exception as e:
        logger.exception("Konvertierung fehlgeschlagen")
        # Aufräumen bei Fehler
        input_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)
        return {
            "success": False,
            "error": str(e),
        }


@app.get("/api/download/{job_id}")
async def download_pdf(job_id: str):
    """Lädt die konvertierte PDF herunter."""
    output_path = UPLOAD_DIR / f"{job_id}_output.pdf"

    if not output_path.exists():
        raise HTTPException(status_code=404, detail="Datei nicht gefunden")

    return FileResponse(
        path=output_path,
        media_type="application/pdf",
        filename=f"converted_{job_id}.pdf",
    )


def run_conversion(
    input_path: Path,
    output_path: Path,
    enable_ai: bool,
    use_docling: bool,
    validate: bool,
    language: str,
) -> dict:
    """Führt die Konvertierung durch."""

    result = {
        "stats": {},
        "validation": {},
    }

    # 1. Parse
    parser = PPTXParser()
    model = parser.parse(input_path)

    result["stats"]["slides"] = model.slide_count
    result["stats"]["figures"] = len(model.all_figures)

    # 2. Enrich (Alt-Texte)
    if enable_ai:
        backend = EnricherBackend.AUTO if use_docling else EnricherBackend.OLLAMA
        enricher = Enricher(EnricherConfig(
            backend=backend,
            language=language,
        ))
        if enricher.is_available:
            model = enricher.enrich(model, verbose=False)

    # 2b. Folienbilder für Vision-Analyse rendern
    if enable_ai and is_libreoffice_available():
        try:
            populate_slide_images(model, input_path)
            result["stats"]["slide_images"] = sum(1 for s in model.slides if s.slide_image)
        except Exception as e:
            logger.warning(f"Slide rendering failed: {e}")

    # 3. Accessibility Optimize
    if enable_ai:
        optimizer = AccessibilityOptimizer(AccessibilityConfig(
            language=language,
            use_docling=use_docling,
        ))
        model = optimizer.optimize(model, verbose=False)

    # 4. Render
    renderer = PDFUARenderer(RendererConfig())
    renderer.render(model, output_path)

    # 5. Validate
    if validate:
        validator = PDFUAValidator()
        validation_result = validator.validate(output_path)
        result["validation"] = {
            "compliant": validation_result.is_compliant,
            "errors": validation_result.errors,
            "warnings": validation_result.warnings,
        }

    return result


async def cleanup_file(path: Path, delay: int = 0):
    """Löscht eine Datei nach einer Verzögerung."""
    if delay > 0:
        await asyncio.sleep(delay)
    path.unlink(missing_ok=True)


def run_server(host: str = "0.0.0.0", port: int = 3003):
    """Startet den Server."""
    import uvicorn

    print(f"""
╔══════════════════════════════════════════════════════════════╗
║                     PPTX2UA Server                           ║
╠══════════════════════════════════════════════════════════════╣
║  URL: http://localhost:{port}                                  ║
║  API: http://localhost:{port}/api/status                       ║
║                                                              ║
║  DSGVO-konform: Alle Daten werden lokal verarbeitet          ║
╚══════════════════════════════════════════════════════════════╝
    """)

    uvicorn.run(app, host=host, port=port, log_level="info")


if __name__ == "__main__":
    run_server()
