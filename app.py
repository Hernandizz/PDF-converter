"""
HD Converter Pro - Web Application Server
==========================================
Flask web server untuk:
  1. Konversi Word (.docx/.doc) ke PDF kualitas Ultra-HD Lossless.
  2. Konversi PDF ke Word (.docx) via Microsoft Word COM Automation.
"""

import os
import sys
import shutil
import zipfile
import subprocess
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from converter import HDWordToPdfConverter, WordConverterError
from pdf_to_word_converter import PdfToWordConverter, PdfToWordError
from pdf_tools import (
    merge_pdfs,
    split_pdf,
    get_pdf_info,
    images_to_pdf,
    PdfToolsError
)

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {".docx", ".doc"}
ALLOWED_PDF_EXTENSIONS = {".pdf"}
ALLOWED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff"}


def is_allowed_file(filename: str) -> bool:
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS


def is_allowed_pdf(filename: str) -> bool:
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_PDF_EXTENSIONS


def is_allowed_image(filename: str) -> bool:
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_IMAGE_EXTENSIONS


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

        # Fast count info gambar jika docx (tanpa membuka image Pillow)
        image_info = None
        if clean_name.lower().endswith(".docx"):
            try:
                with zipfile.ZipFile(input_path, 'r') as zf:
                    media_files = [f for f in zf.namelist() if f.startswith("word/media/")]
                    image_info = {"image_count": len(media_files), "estimated_quality": "Ultra HD Lossless"}
            except Exception:
                pass

        try:
            res_pdf = converter.convert(input_path, output_path, keep_word_open=True)
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


@app.route("/api/pdf-to-word", methods=["POST"])
def pdf_to_word():
    """
    Menerima satu atau banyak file PDF, mengonversinya ke DOCX via
    Microsoft Word COM Automation, dan mengembalikan URL unduh.
    """
    files = request.files.getlist("files")
    if not files or len(files) == 0 or (len(files) == 1 and files[0].filename == ""):
        return jsonify({"error": "Harap pilih minimal satu file PDF"}), 400

    output_format = request.form.get("output_format", "docx").lower().strip().lstrip('.')
    if output_format not in ("docx", "doc"):
        output_format = "docx"

    converter = PdfToWordConverter()
    converted_files = []
    errors = []

    for file in files:
        raw_name = file.filename
        if not raw_name or not is_allowed_pdf(raw_name):
            errors.append({"file": raw_name, "error": "Format file tidak didukung (harus .pdf)"})
            continue

        clean_name = secure_filename(raw_name)
        if not clean_name:
            clean_name = "document.pdf"

        input_path = os.path.join(UPLOAD_FOLDER, clean_name)
        file.save(input_path)

        base_name, _ = os.path.splitext(clean_name)
        word_name = f"{base_name}.{output_format}"
        output_path = os.path.join(OUTPUT_FOLDER, word_name)

        try:
            res_word = converter.convert(input_path, output_path, output_format=output_format)
            size_bytes = os.path.getsize(res_word)
            size_kb = round(size_bytes / 1024, 1)

            converted_files.append({
                "original_name": raw_name,
                "word_name": word_name,
                "docx_name": word_name,
                "format": output_format,
                "download_url": f"/api/download/{word_name}",
                "size_kb": size_kb,
            })
        except PdfToWordError as e:
            errors.append({"file": raw_name, "error": str(e)})
        except Exception as e:
            errors.append({"file": raw_name, "error": f"Error tak terduga: {str(e)}"})
        finally:
            if os.path.exists(input_path):
                try:
                    os.remove(input_path)
                except Exception:
                    pass

    # Buat ZIP jika ada lebih dari satu file
    zip_url = None
    if len(converted_files) > 1:
        zip_filename = f"converted_words.zip"
        zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)
        try:
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for item in converted_files:
                    fpath = os.path.join(OUTPUT_FOLDER, item["word_name"])
                    if os.path.exists(fpath):
                        zf.write(fpath, arcname=item["word_name"])
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


@app.route("/api/pdf/info", methods=["POST"])
def pdf_info():
    """
    Mengambil informasi jumlah halaman dan metadata dari sebuah file PDF.
    """
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file PDF yang diunggah"}), 400

    file = request.files["file"]
    if not file or not file.filename or not is_allowed_pdf(file.filename):
        return jsonify({"error": "File harus berformat PDF (.pdf)"}), 400

    clean_name = secure_filename(file.filename) or "temp_doc.pdf"
    temp_path = os.path.join(UPLOAD_FOLDER, f"info_{clean_name}")
    try:
        file.save(temp_path)
        info = get_pdf_info(temp_path)
        return jsonify({"success": True, "filename": clean_name, "info": info})
    except PdfToolsError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Gagal membaca PDF: {str(e)}"}), 500
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


@app.route("/api/pdf/merge", methods=["POST"])
def pdf_merge():
    """
    Menggabungkan beberapa file PDF menjadi satu file PDF utuh.
    """
    files = request.files.getlist("files")
    if not files or len(files) < 2:
        return jsonify({"error": "Pilih minimal 2 file PDF untuk digabungkan"}), 400

    saved_paths = []
    custom_name = request.form.get("output_name", "").strip()
    if custom_name:
        clean_custom = secure_filename(custom_name)
        if not clean_custom.lower().endswith(".pdf"):
            clean_custom += ".pdf"
    else:
        clean_custom = "merged_document.pdf"

    output_path = os.path.join(OUTPUT_FOLDER, clean_custom)

    try:
        for idx, file in enumerate(files):
            raw_name = file.filename
            if not raw_name or not is_allowed_pdf(raw_name):
                return jsonify({"error": f"File '{raw_name}' bukan file PDF yang valid"}), 400

            clean_name = secure_filename(raw_name) or f"doc_{idx+1}.pdf"
            tmp_path = os.path.join(UPLOAD_FOLDER, f"merge_{idx}_{clean_name}")
            file.save(tmp_path)
            saved_paths.append(tmp_path)

        res_path = merge_pdfs(saved_paths, output_path)
        size_bytes = os.path.getsize(res_path)

        return jsonify({
            "success": True,
            "filename": clean_custom,
            "download_url": f"/api/download/{clean_custom}",
            "size_kb": round(size_bytes / 1024, 1),
            "total_files_merged": len(saved_paths)
        })
    except PdfToolsError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Error saat menggabungkan PDF: {str(e)}"}), 500
    finally:
        for p in saved_paths:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass


@app.route("/api/pdf/split", methods=["POST"])
def pdf_split():
    """
    Memisahkan halaman PDF berdasarkan mode 'all' (semua halaman) atau 'range' (rentang halaman).
    """
    if "file" not in request.files:
        return jsonify({"error": "Tidak ada file PDF yang diunggah"}), 400

    file = request.files["file"]
    if not file or not file.filename or not is_allowed_pdf(file.filename):
        return jsonify({"error": "File harus berformat PDF (.pdf)"}), 400

    mode = request.form.get("mode", "all").strip().lower()
    page_range = request.form.get("page_range", "").strip()

    clean_name = secure_filename(file.filename) or "doc.pdf"
    input_path = os.path.join(UPLOAD_FOLDER, f"split_{clean_name}")

    try:
        file.save(input_path)
        result = split_pdf(
            input_path=input_path,
            output_dir=OUTPUT_FOLDER,
            mode=mode,
            range_str=page_range if mode == "range" else None
        )

        primary_file = result["primary_file"]
        file_path = os.path.join(OUTPUT_FOLDER, primary_file)
        size_bytes = os.path.getsize(file_path) if os.path.exists(file_path) else 0

        return jsonify({
            "success": True,
            "filename": primary_file,
            "download_url": f"/api/download/{primary_file}",
            "size_kb": round(size_bytes / 1024, 1),
            "is_zip": result["is_zip"],
            "total_pages": result["total_pages"],
            "extracted_pages_count": result["extracted_pages_count"]
        })
    except PdfToolsError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Error saat memisahkan PDF: {str(e)}"}), 500
    finally:
        if os.path.exists(input_path):
            try:
                os.remove(input_path)
            except Exception:
                pass


@app.route("/api/image-to-pdf", methods=["POST"])
def image_to_pdf_endpoint():
    """
    Mengonversi satu atau banyak file gambar (JPG, PNG, WEBP, BMP) menjadi PDF berkualitas tinggi.
    """
    files = request.files.getlist("files")
    if not files or len(files) == 0 or (len(files) == 1 and files[0].filename == ""):
        return jsonify({"error": "Pilih minimal satu file gambar (JPG/PNG/WEBP/BMP)"}), 400

    page_size = request.form.get("page_size", "fit").strip().lower()
    custom_name = request.form.get("output_name", "").strip()
    if custom_name:
        clean_custom = secure_filename(custom_name)
        if not clean_custom.lower().endswith(".pdf"):
            clean_custom += ".pdf"
    else:
        clean_custom = "images_document.pdf"

    output_path = os.path.join(OUTPUT_FOLDER, clean_custom)
    saved_images = []

    try:
        for idx, file in enumerate(files):
            raw_name = file.filename
            if not raw_name or not is_allowed_image(raw_name):
                return jsonify({"error": f"File '{raw_name}' bukan format gambar yang didukung (JPG, PNG, WEBP, BMP)"}), 400

            clean_name = secure_filename(raw_name) or f"img_{idx+1}.jpg"
            tmp_path = os.path.join(UPLOAD_FOLDER, f"img_{idx}_{clean_name}")
            file.save(tmp_path)
            saved_images.append(tmp_path)

        res_path = images_to_pdf(saved_images, output_path, page_size=page_size)
        size_bytes = os.path.getsize(res_path)

        return jsonify({
            "success": True,
            "filename": clean_custom,
            "download_url": f"/api/download/{clean_custom}",
            "size_kb": round(size_bytes / 1024, 1),
            "total_images": len(saved_images)
        })
    except PdfToolsError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Error konversi gambar ke PDF: {str(e)}"}), 500
    finally:
        for p in saved_images:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass


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
    print("   HD CONVERTER PRO — Word<->PDF (LOSSLESS / NATIVE COM)")
    print("=" * 60)
    print(f"Server berjalan di: http://127.0.0.1:{port}")
    print("Tekan Ctrl+C untuk berhenti.")
    print("=" * 60)
    app.run(host="127.0.0.1", port=port, debug=False)
