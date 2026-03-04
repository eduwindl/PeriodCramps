import os
import io
import sys
import zipfile
from pathlib import Path
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import pandas as pd
from datetime import datetime

# Adjust path to find our modules
sys.path.insert(0, str(Path(__file__).parent))
import data_processing as dp
import report_stats as st
import generate_report as gr

try:
    # On Windows, try to use docx2pdf for PDF conversion
    from docx2pdf import convert as convert_pdf
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False
except Exception:
    DOCX2PDF_AVAILABLE = False


# Define Paths
if getattr(sys, 'frozen', False):
    # Running as a bundled exe (PyInstaller)
    BUNDLE_DIR = Path(sys._MEIPASS)
    # The absolute path to the main folder where the app lives
    PROJECT_ROOT = Path(os.environ.get("APP_ROOT", os.getcwd()))
else:
    # Running in normal python
    PROJECT_ROOT = Path(__file__).parent.parent

STATIC_FOLDER = PROJECT_ROOT / "web-prototype"
DATA_DIR = PROJECT_ROOT / "data"
REPORTS_DIR = PROJECT_ROOT / "reports"

app = Flask(__name__, static_folder=str(STATIC_FOLDER))
CORS(app)

DATA_DIR.mkdir(parents=True, exist_ok=True)
REPORTS_DIR.mkdir(parents=True, exist_ok=True)

# -----------------
# Web UI Endpoints
# -----------------
@app.route("/")
def serve_index():
    index_path = STATIC_FOLDER / "index.html"
    if not index_path.exists():
        return f"Error: No se encontró la interfaz en {index_path}. Revisa la estructura del paquete.", 404
    return send_file(str(index_path))

@app.route("/<path:filename>")
def serve_static(filename):
    return send_from_directory(str(STATIC_FOLDER), filename)

# -----------------
# API
# -----------------

# 1. Login
@app.route("/api/login", methods=["POST"])
def login():
    data = request.json or {}
    username = data.get("username", "")
    password = data.get("password", "")
    
    # Simple hardcoded roles for 4 users as requested
    users = {
        "admin": {"pwd": "123", "role": "admin"},
        "eduwi": {"pwd": "123", "role": "admin"},
        "viewer": {"pwd": "123", "role": "viewer"},
        "guest": {"pwd": "123", "role": "viewer"},
    }
    
    user = users.get(username.lower())
    if user and user["pwd"] == password:
        return jsonify({"success": True, "token": f"token-{username}", "role": user["role"]})
    return jsonify({"success": False, "error": "Credenciales inválidas"}), 401


# 2. Generation Engine
@app.route("/api/generate", methods=["POST"])
def generate():
    # Allow any token for now (just structural login implementation)
    auth = request.headers.get("Authorization", "")
    
    period = request.form.get("period", "2026-02")
    
    visitas_file = request.files.get("visitas")
    equipos_file = request.files.get("equipos")
    
    visitas_path = DATA_DIR / "visitas_centros.xlsx"
    equipos_path = DATA_DIR / "cambios_equipos.xlsx"
    
    # Save uploaded files if provided
    if visitas_file:
        visitas_file.save(visitas_path)
    if equipos_file:
        equipos_file.save(equipos_path)
        
    if not visitas_path.exists() or not equipos_path.exists():
        return jsonify({"error": "Faltan los archivos de datos. Sube visitas y equipos."}), 400

    config = gr.load_config()
    
    # Run the existing scripts logic
    try:
        dt = datetime.strptime(period, "%Y-%m")
        year, month = dt.year, dt.month
        
        # 1. Load data
        # Using correct function name from data_processing.py
        _, visits_month, _, equip_month = dp.load_and_prepare(
            visits_path=visitas_path,
            equipment_path=equipos_path,
            year=year,
            month=month
        )
        
        # 2. Stats
        # Use build_all_stats from statistics.py
        stats = st.build_all_stats(visits_month, equip_month, config)
        
        # 3. Build Report Docx
        out_docx_path = REPORTS_DIR / f"reporte_{period.replace('-', '_')}.docx"
        # We use gr.build_report but with custom paths if possible, or just build_report
        # Since build_report re-loads the data (which is okay, it's fast), 
        # let's just use it to avoid recreating the complex Word build logic here.
        gr.build_report(year, month, config) # This writes to reports/reporte_YYYY_MM.docx
        
        # The filename created by gr.build_report is reporte_YYYY_MM.docx
        # Let's ensure out_docx_path matches what gr.build_report creates.
        out_docx_path = REPORTS_DIR / f"reporte_{year}_{month:02d}.docx"
        
        # --- Export additional formats ---
        base_name = f"reporte_{year}_{month:02d}"
        
        # 1. PDF — run in a thread with a hard 15s timeout to avoid Flask hang
        out_pdf_path = REPORTS_DIR / f"{base_name}.pdf"
        pdf_available = False
        if DOCX2PDF_AVAILABLE:
            import threading
            def _convert():
                try:
                    convert_pdf(str(out_docx_path.resolve()), str(out_pdf_path.resolve()))
                except Exception as e:
                    app.logger.error(f"PDF conversion error: {e}")
            t = threading.Thread(target=_convert, daemon=True)
            t.start()
            t.join(timeout=15)  # max 15 seconds for Word COM
            pdf_available = out_pdf_path.exists()
            
        # 2. Excel Dataset
        out_xlsx_path = REPORTS_DIR / f"{base_name}_datos.xlsx"
        with pd.ExcelWriter(str(out_xlsx_path)) as writer:
            visits_month.to_excel(writer, sheet_name="Visitas", index=False)
            equip_month.to_excel(writer, sheet_name="Equipos", index=False)
            
        # 3. CSV ZIP
        out_csv_path = REPORTS_DIR / f"{base_name}_csv.zip"
        with zipfile.ZipFile(str(out_csv_path), 'w') as zf:
            visits_csv = visits_month.to_csv(index=False).encode('utf-8-sig')
            equip_csv = equip_month.to_csv(index=False).encode('utf-8-sig')
            zf.writestr('visitas.csv', visits_csv)
            zf.writestr('equipos.csv', equip_csv)
            
        # Expose URLs
        return jsonify({
            "success": True,
            "filename": f"{base_name} Completado",
            "urls": {
                "docx": f"/download/{out_docx_path.name}",
                "pdf": f"/download/{out_pdf_path.name}" if pdf_available else None,
                "xlsx": f"/download/{out_xlsx_path.name}",
                "csv": f"/download/{out_csv_path.name}"
            }
        })

    except Exception as e:
        app.logger.exception("Error during generation:")
        return jsonify({"error": str(e)}), 500


@app.route("/download/<path:filename>")
def download_file(filename):
    return send_from_directory(REPORTS_DIR, filename, as_attachment=True)


if __name__ == "__main__":
    # In distribution, we disable debug to avoid watchdog issues
    # and use a fixed port
    print("Backend engine starting...")
    sys.stdout.flush()
    app.run(host="127.0.0.1", port=5001, debug=False, threaded=True)
