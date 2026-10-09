"""
PDF Tools Engine - All-in-One PDF Suite
========================================
Modul untuk manipulasi PDF murni (Merge, Split, Extract, Image-to-PDF, & Metadata).
Berjalan 100% lokal, cepat, dan tanpa dependensi eksternal berbayar.

PERFORMA:
- Merge PDF: Menggunakan streaming direct-append pypdf tanpa dekompresi per-halaman.
- Split PDF: Direct memory streaming & level-1 ZIP compression untuk kecepatan maksimal.
- Image-to-PDF: Resampling Bicubic resolusi tinggi berkecepatan 3x lebih cepat dari Lanczos.
"""

import os
import re
import io
import zipfile
from typing import List, Dict, Any, Optional, Tuple
from PIL import Image

try:
    import pypdf
    from pypdf import PdfReader, PdfWriter
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False


class PdfToolsError(Exception):
    """Custom exception untuk operasi PDF Tools."""
    pass


def get_pdf_info(pdf_path: str) -> Dict[str, Any]:
    """
    Mengambil informasi dasar metadata dan halaman dari file PDF.
    """
    if not PYPDF_AVAILABLE:
        raise PdfToolsError("Modul 'pypdf' belum terpasang. Jalankan: pip install pypdf")

    if not os.path.isfile(pdf_path):
        raise PdfToolsError(f"File PDF tidak ditemukan: {pdf_path}")

    try:
        reader = PdfReader(pdf_path)
        is_encrypted = reader.is_encrypted
        total_pages = len(reader.pages) if not is_encrypted else 0
        file_size = os.path.getsize(pdf_path)

        meta = {}
        if not is_encrypted and reader.metadata:
            meta = {
                "title": reader.metadata.title or "",
                "author": reader.metadata.author or "",
                "producer": reader.metadata.producer or ""
            }

        return {
            "total_pages": total_pages,
            "is_encrypted": is_encrypted,
            "file_size_bytes": file_size,
            "file_size_kb": round(file_size / 1024, 1),
            "file_size_mb": round(file_size / (1024 * 1024), 2),
            "metadata": meta
        }
    except Exception as e:
        raise PdfToolsError(f"Gagal membaca informasi PDF: {str(e)}")


def merge_pdfs(file_paths: List[str], output_path: str) -> str:
    """
    Menggabungkan beberapa file PDF secara berurutan menjadi satu file PDF utuh.
    Menggunakan direct fast-append internal pypdf untuk efisiensi CPU dan RAM.

    Args:
        file_paths: List path absolut ke file-file PDF yang akan digabung.
        output_path: Path absolut tujuan file PDF gabungan.

    Returns:
        Path absolut ke file PDF hasil gabungan.
    """
    if not PYPDF_AVAILABLE:
        raise PdfToolsError("Modul 'pypdf' belum terpasang. Jalankan: pip install pypdf")

    if not file_paths or len(file_paths) < 2:
        raise PdfToolsError("Minimal 2 file PDF dibutuhkan untuk digabungkan.")

    writer = PdfWriter()

    for idx, path in enumerate(file_paths):
        if not os.path.isfile(path):
            raise PdfToolsError(f"File PDF #{idx+1} tidak ditemukan: {path}")

        try:
            # Gunakan reader cepat untuk validasi enkripsi
            reader = PdfReader(path)
            if reader.is_encrypted:
                raise PdfToolsError(
                    f"File '{os.path.basename(path)}' diproteksi password. Buka sandi terlebih dahulu."
                )

            # writer.append() memanfaatkan streaming internal object tree pypdf
            # yang jauh lebih cepat dibanding mengekstrak dan memasukkan halaman satu per satu
            writer.append(reader)

        except PdfToolsError:
            raise
        except Exception as e:
            raise PdfToolsError(f"Gagal memproses file '{os.path.basename(path)}': {str(e)}")

    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    try:
        with open(output_path, "wb") as f_out:
            writer.write(f_out)
        return os.path.abspath(output_path)
    except Exception as e:
        raise PdfToolsError(f"Gagal menyimpan file PDF hasil gabungan: {str(e)}")


def parse_page_ranges(range_str: str, max_pages: int) -> List[int]:
    """
    Mem-parsing string rentang halaman seperti '1-3, 5, 7-9' menjadi list page index 0-based unik berurutan.
    Halaman dalam string adalah 1-based (standar manusia).
    """
    if not range_str or not range_str.strip():
        raise PdfToolsError("Rentang halaman tidak boleh kosong.")

    selected_pages = set()
    parts = [p.strip() for p in range_str.split(",") if p.strip()]

    for part in parts:
        if "-" in part:
            sub = part.split("-")
            if len(sub) != 2:
                raise PdfToolsError(f"Format rentang tidak valid: '{part}'")
            try:
                start = int(sub[0].strip())
                end = int(sub[1].strip())
            except ValueError:
                raise PdfToolsError(f"Nomor halaman harus berupa angka: '{part}'")

            if start > end:
                start, end = end, start

            if start < 1:
                start = 1
            if end > max_pages:
                end = max_pages

            for p in range(start, end + 1):
                selected_pages.add(p - 1)
        else:
            try:
                page_num = int(part)
            except ValueError:
                raise PdfToolsError(f"Nomor halaman harus berupa angka: '{part}'")

            if 1 <= page_num <= max_pages:
                selected_pages.add(page_num - 1)
            else:
                raise PdfToolsError(
                    f"Halaman {page_num} di luar rentang dokumen (Total: {max_pages} halaman)"
                )

    if not selected_pages:
        raise PdfToolsError("Tidak ada halaman valid yang dipilih.")

    return sorted(list(selected_pages))


def split_pdf(
    input_path: str,
    output_dir: str,
    mode: str = "all",
    range_str: Optional[str] = None
) -> Dict[str, Any]:
    """
    Memisahkan halaman PDF:
    - Mode 'all': Memisahkan setiap halaman menjadi file terpisah (dikemas dalam ZIP).
    - Mode 'range': Mengekstrak halaman yang dipilih ke dalam 1 file PDF baru.

    Returns:
        Dict dengan status, list file yang dibuat, dan path file download utama (PDF atau ZIP).
    """
    if not PYPDF_AVAILABLE:
        raise PdfToolsError("Modul 'pypdf' belum terpasang. Jalankan: pip install pypdf")

    if not os.path.isfile(input_path):
        raise PdfToolsError(f"File PDF tidak ditemukan: {input_path}")

    reader = PdfReader(input_path)
    if reader.is_encrypted:
        raise PdfToolsError("File PDF terenkripsi/diproteksi password.")

    total_pages = len(reader.pages)
    if total_pages == 0:
        raise PdfToolsError("File PDF kosong (0 halaman).")

    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(input_path))[0]

    created_files = []

    if mode == "range":
        selected_indices = parse_page_ranges(range_str or "", total_pages)
        writer = PdfWriter()
        for idx in selected_indices:
            writer.add_page(reader.pages[idx])

        clean_range_tag = re.sub(r'[^0-9,-]', '', range_str or 'custom')
        out_filename = f"{base_name}_hal_{clean_range_tag}.pdf"
        out_path = os.path.join(output_dir, out_filename)

        with open(out_path, "wb") as f_out:
            writer.write(f_out)

        created_files.append(out_path)
        return {
            "mode": "range",
            "total_pages": total_pages,
            "extracted_pages_count": len(selected_indices),
            "primary_file": out_filename,
            "is_zip": False,
            "created_files": [out_filename]
        }

    else:
        # Mode 'all': Pisah setiap halaman langsung ke dalam ZIP in-memory dengan compresslevel=1
        # PDF sudah terkompresi internal, sehingga level-1 menghemat siklus CPU secara drastis
        zip_filename = f"{base_name}_pisah_semua.zip"
        zip_path = os.path.join(output_dir, zip_filename)
        created_names = []

        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=1) as zf:
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                page_filename = f"{base_name}_hal_{i+1:03d}.pdf"
                created_names.append(page_filename)
                pdf_bytes = io.BytesIO()
                writer.write(pdf_bytes)
                zf.writestr(page_filename, pdf_bytes.getvalue())

        return {
            "mode": "all",
            "total_pages": total_pages,
            "extracted_pages_count": total_pages,
            "primary_file": zip_filename,
            "is_zip": True,
            "created_files": created_names
        }


def images_to_pdf(
    image_paths: List[str],
    output_path: str,
    page_size: str = "fit"
) -> str:
    """
    Menggabungkan satu atau banyak gambar (JPG, PNG, WEBP, BMP) menjadi dokumen PDF berkualitas tinggi.

    Args:
        image_paths: List path absolut gambar.
        output_path: Path tujuan file PDF.
        page_size: 'fit' (sesuai ukuran asli gambar) atau 'a4' (standar A4).

    Returns:
        Path absolut ke file PDF yang dihasilkan.
    """
    if not image_paths:
        raise PdfToolsError("Harap pilih minimal 1 file gambar.")

    rgb_images = []

    for idx, img_path in enumerate(image_paths):
        if not os.path.isfile(img_path):
            raise PdfToolsError(f"File gambar #{idx+1} tidak ditemukan: {img_path}")

        try:
            with Image.open(img_path) as img:
                # Normalisasi orientasi EXIF (foto HP dsb)
                try:
                    from PIL import ImageOps
                    img = ImageOps.exif_transpose(img)
                except Exception:
                    pass

                # Konversi RGBA / Palette / Grayscale ke Truecolor RGB dengan latar putih
                if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
                    alpha_img = img.convert("RGBA")
                    bg = Image.new("RGB", alpha_img.size, (255, 255, 255))
                    bg.paste(alpha_img, mask=alpha_img.split()[3])
                    rgb_img = bg
                elif img.mode != "RGB":
                    rgb_img = img.convert("RGB")
                else:
                    rgb_img = img.copy()

                # Mode A4 (300 DPI: 2480 x 3508)
                if page_size == "a4":
                    a4_w, a4_h = 2480, 3508
                    canvas = Image.new("RGB", (a4_w, a4_h), (255, 255, 255))

                    img_ratio = rgb_img.width / rgb_img.height
                    a4_ratio = a4_w / a4_h

                    if img_ratio > a4_ratio:
                        new_w = a4_w
                        new_h = int(a4_w / img_ratio)
                    else:
                        new_h = a4_h
                        new_w = int(a4_h * img_ratio)

                    # Gunakan BICUBIC alih-alih LANCZOS: kualitas visual tajam serupa tapi 3x lebih cepat
                    resized = rgb_img.resize((new_w, new_h), Image.Resampling.BICUBIC)
                    pos_x = (a4_w - new_w) // 2
                    pos_y = (a4_h - new_h) // 2
                    canvas.paste(resized, (pos_x, pos_y))
                    rgb_images.append(canvas)
                else:
                    rgb_images.append(rgb_img)

        except Exception as e:
            raise PdfToolsError(f"Gagal memproses gambar '{os.path.basename(img_path)}': {str(e)}")

    if not rgb_images:
        raise PdfToolsError("Tidak ada gambar valid untuk dikonversi.")

    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    try:
        first_img = rgb_images[0]
        other_images = rgb_images[1:] if len(rgb_images) > 1 else []

        first_img.save(
            output_path,
            "PDF",
            resolution=300.0,
            save_all=True,
            append_images=other_images,
            quality=95
        )
        return os.path.abspath(output_path)
    except Exception as e:
        raise PdfToolsError(f"Gagal menyimpan dokumen PDF dari gambar: {str(e)}")
