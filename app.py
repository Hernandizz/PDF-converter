"""
HD Word to PDF Converter - Web Application Server
=================================================
Flask web server untuk menyediakan antarmuka drag-and-drop modern,
batch conversion, inspeksi kualitas gambar DOCX, dan unduh PDF HD.
"""

import os
import sys
import shutil
import zipfile
import subprocess
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from converter import HDWordToPdfConverter, WordConverterError

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {".docx", ".doc"}


def is_allowed_file(filename: str) -> bool:
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/inspect", methods=["POST"])
def inspect_file():
    """
    Inspeksi dokumen sebelum konversi untuk melihat gambar di dalamnya.
    """
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file yang diunggah"}), 400

    file = request.files["file"]
    if not file or file.filename == "":
        return jsonify({"error": "Nama file tidak valid"}), 400

    filename = secure_filename(file.filename)
    if not filename.lower().endswith(".docx"):
        return jsonify({"error": "Hanya file .docx yang dapat diinspeksi gambarnya"}), 400

    temp_path = os.path.join(UPLOAD_FOLDER, f"inspect_{filename}")
    try:
        file.save(temp_path)
        stats = HDWordToPdfConverter.inspect_docx(temp_path)
        return jsonify(stats)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


@app.route("/api/convert", methods=["POST"])
def convert_files():
    """
    Menerima satu atau banyak file Word, melakukan konversi HD, dan mengembalikan status/URL unduh.
    """
    files = request.files.getlist("files")
    if not files or len(files) == 0 or (len(files) == 1 and files[0].filename == ""):
        return jsonify({"error": "Harap pilih minimal satu file Word (.docx / .doc)"}), 400

    # Opsi konversi dari formulir
    optimize_print = request.form.get("optimize_print", "true").lower() == "true"
    preserve_fonts = request.form.get("preserve_fonts", "true").lower() == "true"
    create_bookmarks = request.form.get("create_bookmarks", "true").lower() == "true"

    converter = HDWordToPdfConverter(
        optimize_for_print=optimize_print,
        preserve_fonts=preserve_fonts,
        create_bookmarks=create_bookmarks
    )

    converted_files = []
    errors = []

    for file in files:
        raw_name = file.filename
        if not raw_name or not is_allowed_file(raw_name):
            errors.append({"file": raw_name, "error": "Format file tidak didukung (harus .docx atau .doc)"})
            continue

        clean_name = secure_filename(raw_name)
        if not clean_name:
            clean_name = "document.docx"

        input_path = os.path.join(UPLOAD_FOLDER, clean_name)
        file.save(input_path)

        base_name, _ = os.path.splitext(clean_name)
        pdf_name = f"{base_name}.pdf"
        output_path = os.path.join(OUTPUT_FOLDER, pdf_name)

        # Cek info gambar jika docx
        image_info = None
        if clean_name.lower().endswith(".docx"):
            try:
                image_info = converter.inspect_docx(input_path)
            except Exception:
                pass

        try:
            res_pdf = converter.convert(input_path, output_path)
            size_bytes = os.path.getsize(res_pdf)
            size_kb = round(size_bytes / 1024, 1)

            converted_files.append({
                "original_name": raw_name,
                "pdf_name": pdf_name,
                "download_url": f"/api/download/{pdf_name}",
                "size_kb": size_kb,
                "image_info": image_info
            })
        except WordConverterError as e:
            errors.append({"file": raw_name, "error": str(e)})
        except Exception as e:
            errors.append({"file": raw_name, "error": f"Error tak terduga: {str(e)}"})
        finally:
            if os.path.exists(input_path):
                try:
                    os.remove(input_path)
                except Exception:
                    pass

    # Jika banyak file berhasil dikonversi, buat file ZIP batch
    zip_url = None
    if len(converted_files) > 1:
        zip_filename = "converted_pdfs_hd.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)
        try:
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for item in converted_files:
                    fpath = os.path.join(OUTPUT_FOLDER, item["pdf_name"])
                    if os.path.exists(fpath):
                        zf.write(fpath, arcname=item["pdf_name"])
            zip_url = f"/api/download/{zip_filename}"
        except Exception:
            zip_url = None

    return jsonify({
        "success": len(converted_files) > 0,
        "converted": converted_files,
        "errors": errors,
        "zip_url": zip_url,
        "total_converted": len(converted_files),
        "total_errors": len(errors)
    })


@app.route("/api/download/<filename>")
def download_file(filename: str):
    """
    Mengunduh file PDF atau ZIP hasil konversi.
    """
    clean_name = secure_filename(filename)
    target_path = os.path.join(OUTPUT_FOLDER, clean_name)
    if not os.path.isfile(target_path):
        return jsonify({"error": "File tidak ditemukan"}), 404

    return send_file(target_path, as_attachment=True, download_name=clean_name)


@app.route("/api/open-folder", methods=["POST"])
def open_folder():
    """
    Membuka folder output langsung di Windows File Explorer.
    """
    try:
        folder_to_open = os.path.abspath(OUTPUT_FOLDER)
        subprocess.Popen(f'explorer "{folder_to_open}"')
        return jsonify({"success": True, "path": folder_to_open})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("=" * 60)
    print("   HD WORD TO PDF CONVERTER (LOSSLESS PRINT QUALITY)")
    print("=" * 60)
    print(f"Server berjalan di: http://127.0.0.1:{port}")
    print("Tekan Ctrl+C untuk berhenti.")
    print("=" * 60)
    app.run(host="127.0.0.1", port=port, debug=False)
