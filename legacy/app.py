"""
app.py  –  ABB Saving Calculations Filler  (Flask web backend)

Run:
    python app.py
Open:
    http://localhost:5000
"""

import os
import re
import sys
import glob
import json
import tempfile
import traceback
import subprocess
import zipfile as zipf
from flask import Flask, request, send_file, render_template, jsonify
from fill_saving_calculations import run_fill, run_fill_multi, generate_multi_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500 MB

BASE_DIR       = os.path.dirname(os.path.abspath(__file__))
EXCEL_TMPL_DIR = os.path.join(BASE_DIR, "excel_templates")

# Maps template key → Excel file, label, flag, default assumptions, report script
BUNDLED = {
    "standard": {
        "file":          "1_0_Region_Site_Name_Saving_Calculations.xlsx",
        "label":         "Standard (Global)",
        "flag":          "🌐",
        "report_script": "generate_report_standard.py",
        "defaults":      {"currency": "INR", "tariff": 8,    "co2": 0.54,  "tax": 34.9, "discount": 6.5},
    },
    "france": {
        "file":          "2_0_France_Site_Name_Saving_Calculations.xlsx",
        "label":         "France / CEE",
        "flag":          "🇫🇷",
        "report_script": "generate_report_cee.py",
        "defaults":      {"currency": "EUR", "tariff": 0.10, "co2": 0.053, "tax": 25,   "discount": 6.5},
    },
    "poland": {
        "file":          "3_0_Poland_Site_Name_Saving_Calculations.xlsx",
        "label":         "Poland",
        "flag":          "🇵🇱",
        "report_script": "generate_report_poland.py",
        "defaults":      {"currency": "EUR", "tariff": 0.14, "co2": 0.75,  "tax": 19,   "discount": 6.5},
    },
}

# Report-type → script name (for custom / future templates)
REPORT_SCRIPT_MAP = {
    "standard": "generate_report_standard.py",
    "cee":      "generate_report_cee.py",
    "poland":   "generate_report_poland.py",
    "xlatam":   "generate_report_xlatam.py",
}


@app.route("/")
def index():
    return render_template("index.html", bundled=BUNDLED)


@app.route("/defaults/<key>")
def get_defaults(key):
    if key in BUNDLED:
        return jsonify(BUNDLED[key]["defaults"])
    return jsonify({}), 404


@app.route("/generate", methods=["POST"])
def generate():
    tmp_dir = tempfile.mkdtemp()
    try:
        # ── Template ────────────────────────────────────────────────────────
        template_key = request.form.get("template_key", "")
        if template_key in BUNDLED:
            template_path = os.path.join(EXCEL_TMPL_DIR, BUNDLED[template_key]["file"])
            default_script = BUNDLED[template_key]["report_script"]
        else:
            f = request.files.get("template_file")
            if not f or not f.filename:
                return jsonify({"error": "No template selected."}), 400
            template_path = os.path.join(tmp_dir, f.filename)
            f.save(template_path)
            default_script = "generate_report_standard.py"

        # ── Zip files ───────────────────────────────────────────────────────
        zips = request.files.getlist("zip_files")
        if not zips or all(z.filename == "" for z in zips):
            return jsonify({"error": "At least one zip file is required."}), 400

        zip_paths = []
        for z in zips:
            if z.filename:
                p = os.path.join(tmp_dir, z.filename)
                z.save(p)
                zip_paths.append(p)

        # ── Assumptions ─────────────────────────────────────────────────────
        def opt_float(key):
            v = request.form.get(key, "").strip()
            return float(v) if v else None

        currency = request.form.get("currency", "").strip() or None
        tariff   = opt_float("tariff")
        co2      = opt_float("co2")
        tax      = opt_float("tax")
        discount = opt_float("discount")

        # ── Output name ─────────────────────────────────────────────────────
        out_name = request.form.get("output_name", "").strip()
        if not out_name:
            base = os.path.splitext(
                BUNDLED[template_key]["file"] if template_key in BUNDLED
                else os.path.basename(template_path)
            )[0]
            out_name = f"{base}_filled"
        if not out_name.lower().endswith(".xlsx"):
            out_name += ".xlsx"

        excel_out = os.path.join(tmp_dir, out_name)

        # ── Fill Excel ──────────────────────────────────────────────────────
        asset_count, sheet_names = run_fill(
            template_path, zip_paths, excel_out,
            currency=currency, tariff=tariff, co2=co2,
            tax=tax, discount=discount,
        )

        # ── Report generation (optional) ────────────────────────────────────
        customer   = request.form.get("customer", "").strip()
        plant      = request.form.get("plant", "").strip()
        rpt_date   = request.form.get("rpt_date", "").strip()
        gen_report = request.form.get("gen_report", "0") == "1"

        # Override report script if user explicitly chose a region
        report_type = request.form.get("report_type", "").strip()
        if report_type in REPORT_SCRIPT_MAP:
            script_name = REPORT_SCRIPT_MAP[report_type]
        else:
            script_name = default_script

        if gen_report and customer and plant:
            script_path = os.path.join(BASE_DIR, script_name)
            if not os.path.exists(script_path):
                return jsonify({"error": f"Report script not found: {script_name}"}), 500

            cmd = [sys.executable, script_path, excel_out, customer, plant]
            if rpt_date:
                cmd.append(rpt_date)

            # Force UTF-8 I/O so Unicode chars (→ arrows etc.) don't crash on Windows cp1252
            _env = os.environ.copy()
            _env['PYTHONIOENCODING'] = 'utf-8'
            result = subprocess.run(
                cmd, cwd=BASE_DIR,
                capture_output=True, text=True, timeout=180,
                encoding='utf-8', env=_env,
            )

            if result.returncode != 0:
                err_msg = result.stderr.strip() or result.stdout.strip() or "Unknown error"
                return jsonify({"error": f"Report generation failed:\n{err_msg}"}), 500

            # Find generated .docx (script saves to BASE_DIR)
            safe_c = re.sub(r'[\\/:*?"<>|]', '_', customer)
            safe_p = re.sub(r'[\\/:*?"<>|]', '_', plant)
            pattern = os.path.join(BASE_DIR, f"{safe_c}_{safe_p}_EA_Report*.docx")
            found = sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)

            if found:
                docx_path  = found[0]
                docx_fname = os.path.basename(docx_path)

                # Bundle Excel + Report into a ZIP
                zip_name = f"EA_{safe_c}_{safe_p}.zip"
                zip_out  = os.path.join(tmp_dir, zip_name)
                with zipf.ZipFile(zip_out, "w", zipf.ZIP_DEFLATED) as zout:
                    zout.write(excel_out, out_name)
                    zout.write(docx_path, docx_fname)

                try:
                    os.remove(docx_path)   # clean up from BASE_DIR
                except OSError:
                    pass

                return send_file(
                    zip_out,
                    as_attachment=True,
                    download_name=zip_name,
                    mimetype="application/zip",
                )
            else:
                # Script ran OK but docx not found — return just the Excel with a warning
                print(f"WARNING: report script succeeded but no .docx found. stdout:\n{result.stdout}")

        # ── Excel only ──────────────────────────────────────────────────────
        return send_file(
            excel_out,
            as_attachment=True,
            download_name=out_name,
            mimetype=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/generate_multi", methods=["POST"])
def generate_multi():
    """Multi-sheet endpoint: one Saving Calculations tab per sheet spec."""
    tmp_dir = tempfile.mkdtemp()
    try:
        # ── Template ─────────────────────────────────────────────────────
        template_key = request.form.get("template_key", "")
        if template_key in BUNDLED:
            template_path = os.path.join(EXCEL_TMPL_DIR,
                                         BUNDLED[template_key]["file"])
        else:
            f = request.files.get("template_file")
            if not f or not f.filename:
                return jsonify({"error": "No template selected."}), 400
            template_path = os.path.join(tmp_dir, f.filename)
            f.save(template_path)

        # ── ZIP pool — save once per filename, shared across all sheets ──
        zip_pool = {}
        for z in request.files.getlist("zip_files"):
            if z.filename and z.filename not in zip_pool:
                p = os.path.join(tmp_dir, z.filename)
                z.save(p)
                zip_pool[z.filename] = p

        # ── Parse sheet specs ─────────────────────────────────────────────
        raw = request.form.get("sheet_specs", "").strip()
        if not raw:
            return jsonify({"error": "sheet_specs is required."}), 400
        try:
            specs_json = json.loads(raw)
        except json.JSONDecodeError as exc:
            return jsonify({"error": f"Invalid sheet_specs JSON: {exc}"}), 400

        if not isinstance(specs_json, list) or not specs_json:
            return jsonify({"error": "sheet_specs must be a non-empty JSON array."}), 400

        # ── Validate + resolve ZIP basenames → saved absolute paths ───────
        def _opt_float(d, key):
            v = d.get(key)
            try:
                return float(v) if v is not None and str(v).strip() != "" else None
            except (TypeError, ValueError):
                return None

        seen_names  = set()
        sheet_specs = []

        for i, s in enumerate(specs_json, start=1):
            name = (s.get("name") or "").strip()
            if not name:
                return jsonify({"error": f"Sheet {i}: name is empty."}), 400
            if len(name) > 31:
                return jsonify(
                    {"error": f"Sheet {i} ('{name}'): name exceeds 31 characters."}
                ), 400
            if name in seen_names:
                return jsonify({"error": f"Duplicate sheet name: '{name}'."}), 400
            seen_names.add(name)

            zip_names = s.get("zips") or []
            if not zip_names:
                return jsonify({"error": f"Sheet '{name}': no ZIPs assigned."}), 400

            zip_paths = []
            for zname in zip_names:
                if zname not in zip_pool:
                    return jsonify(
                        {"error": f"Sheet '{name}': ZIP '{zname}' was not uploaded."}
                    ), 400
                zip_paths.append(zip_pool[zname])

            sheet_specs.append({
                "name":     name,
                "zips":     zip_paths,
                "currency": (s.get("currency") or "").strip() or None,
                "tariff":   _opt_float(s, "tariff"),
                "co2":      _opt_float(s, "co2"),
                "tax":      _opt_float(s, "tax"),
                "discount": _opt_float(s, "discount"),
            })

        # ── Output filename ───────────────────────────────────────────────
        out_name = (request.form.get("output_name") or "").strip()
        if not out_name:
            base = os.path.splitext(
                BUNDLED[template_key]["file"] if template_key in BUNDLED
                else os.path.basename(template_path)
            )[0]
            out_name = f"{base}_filled"
        if not out_name.lower().endswith(".xlsx"):
            out_name += ".xlsx"

        excel_out = os.path.join(tmp_dir, out_name)

        # ── Fill ──────────────────────────────────────────────────────────
        sheet_names, asset_counts = generate_multi_workbook(
            template_path, sheet_specs, excel_out,
        )

        print(f"\n{'='*60}")
        print(f"Multi-sheet output : {out_name}")
        for sname, cnt in asset_counts.items():
            print(f"  '{sname}' : {cnt} assets")

        return send_file(
            excel_out,
            as_attachment=True,
            download_name=out_name,
            mimetype=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("\n  ABB Energy Appraisal Tool")
    print("  Open in browser:  http://localhost:5000\n")
    app.run(debug=True, port=5000)
