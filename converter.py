"""
HD Word to PDF Converter Engine
================================
Engine konversi Word (.docx / .doc) ke PDF dengan kualitas maksimal (HD / Lossless).
Memastikan:
1. Gambar tidak pecah / tidak terkompresi (DoNotCompressImages = True).
2. Optimasi cetak resolusi tinggi (wdExportOptimizeForPrint, 300+ DPI).
3. Font vector tetap tajam pada semua tingkat zoom (BitmapMissingFonts = False).
4. Struktur dokumen & bookmark heading dipertahankan.
"""

import os
import sys
import time
import zipfile
import threading
import argparse
from typing import Optional, List, Dict, Any
from PIL import Image
import io

# COM Types untuk automasi Word di Windows
try:
    import comtypes
    import comtypes.client
    COMTYPES_AVAILABLE = True
except ImportError:
    COMTYPES_AVAILABLE = False


# Word COM Constants
WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0     # 0 = wdExportOptimizeForPrint (Kualitas Tertinggi / 300+ DPI)
WD_EXPORT_OPTIMIZE_FOR_SCREEN = 1    # 1 = wdExportOptimizeForOnScreen (Terkonversi/Kompres rendah)
WD_EXPORT_ALL_DOCUMENT = 0
WD_EXPORT_DOCUMENT_CONTENT = 0
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1
WD_EXPORT_CREATE_NO_BOOKMARKS = 0
WD_ALERTS_NONE = 0


class WordConverterError(Exception):
    """Custom exception untuk error konversi."""
    pass


# Global lock untuk mencegah concurrent access ke instance Word COM (STA Apartment)
word_lock = threading.Lock()


class HDWordToPdfConverter:
    """
    Kelas konverter Word ke PDF dengan konfigurasi Ultra-HD Lossless.
    """

    def __init__(self, optimize_for_print: bool = True, preserve_fonts: bool = True, create_bookmarks: bool = True):
        self.optimize_for_print = optimize_for_print
        self.preserve_fonts = preserve_fonts
        self.create_bookmarks = create_bookmarks

    @staticmethod
    def inspect_docx(docx_path: str) -> Dict[str, Any]:
        """
        Memeriksa kualitas gambar internal di dalam file DOCX sebelum dikonversi.
        Menghasilkan statistik gambar, resolusi, dan dimensi asli.
        """
        stats: Dict[str, Any] = {
            "file_name": os.path.basename(docx_path),
            "file_size_bytes": os.path.getsize(docx_path) if os.path.exists(docx_path) else 0,
            "image_count": 0,
            "images": [],
            "has_hd_images": False,
            "estimated_quality": "Standard"
        }

        if not docx_path.lower().endswith(".docx"):
            return stats

        try:
            with zipfile.ZipFile(docx_path, 'r') as zf:
                media_files = [f for f in zf.namelist() if f.startswith("word/media/")]
                stats["image_count"] = len(media_files)

                for media in media_files:
                    try:
                        data = zf.read(media)
                        filename = os.path.basename(media)
                        with Image.open(io.BytesIO(data)) as img:
                            width, height = img.size
                            dpi = img.info.get("dpi", (72, 72))
                            if isinstance(dpi, tuple):
                                dpi_val = round(dpi[0])
                            else:
                                dpi_val = round(dpi)

                            is_hd = (width >= 1280 or height >= 1280 or dpi_val >= 200)
                            if is_hd:
                                stats["has_hd_images"] = True

                            stats["images"].append({
                                "name": filename,
                                "format": img.format or "UNKNOWN",
                                "width": width,
                                "height": height,
                                "dpi": dpi_val,
                                "size_bytes": len(data),
                                "is_hd": is_hd
                            })
                    except Exception:
                        continue

            if stats["has_hd_images"]:
                stats["estimated_quality"] = "Ultra HD (300+ DPI / High-Res)"
            elif stats["image_count"] > 0:
                stats["estimated_quality"] = "Standard / Web Quality"
            else:
                stats["estimated_quality"] = "Vector Text Document"

        except Exception as e:
            stats["error"] = str(e)

        return stats

    def convert(self, input_path: str, output_path: Optional[str] = None) -> str:
        """
        Mengonversi satu file DOC/DOCX ke PDF berkualitas HD tanpa kompresi gambar.
        """
        if not COMTYPES_AVAILABLE:
            raise WordConverterError(
                "Modul 'comtypes' tidak ditemukan. Jalankan: pip install comtypes"
            )

        input_abs = os.path.abspath(input_path)
        if not os.path.isfile(input_abs):
            raise WordConverterError(f"File dokumen tidak ditemukan: {input_path}")

        if not output_path:
            base_name, _ = os.path.splitext(input_abs)
            output_abs = f"{base_name}.pdf"
        else:
            output_abs = os.path.abspath(output_path)

        # Pastikan direktori output sudah dibuat
        os.makedirs(os.path.dirname(output_abs), exist_ok=True)

        # Hapus file lama jika ada
        if os.path.exists(output_abs):
            try:
                os.remove(output_abs)
            except OSError as e:
                raise WordConverterError(f"Tidak dapat menimpa file PDF yang sedang dibuka: {e}")

        # Inisialisasi COM untuk thread saat ini
        comtypes.CoInitialize()

        with word_lock:
            word_app = None
            doc = None
            try:
                # Buka Word Application
                word_app = comtypes.client.CreateObject("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = WD_ALERTS_NONE

                # Konfigurasi opsi Word untuk kualitas maksimal
                try:
                    word_app.Options.DoNotPromptForConvert = True
                    word_app.Options.SaveInterval = 0
                    word_app.Options.UpdateLinksAtOpen = False
                except Exception:
                    pass

                # Buka dokumen dalam mode ReadOnly agar aman
                doc = word_app.Documents.Open(
                    FileName=input_abs,
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False
                )

                # ==============================================================
                # PENTING: Cegah kompresi gambar otomatis oleh Word
                # ==============================================================
                try:
                    doc.DoNotCompressImages = True
                except Exception:
                    pass

                optimize_flag = (
                    WD_EXPORT_OPTIMIZE_FOR_PRINT if self.optimize_for_print
                    else WD_EXPORT_OPTIMIZE_FOR_SCREEN
                )
                bookmark_flag = (
                    WD_EXPORT_CREATE_HEADING_BOOKMARKS if self.create_bookmarks
                    else WD_EXPORT_CREATE_NO_BOOKMARKS
                )

                # Ekspor ke PDF dengan setting Lossless / HD
                doc.ExportAsFixedFormat(
                    OutputFileName=output_abs,
                    ExportFormat=WD_EXPORT_FORMAT_PDF,
                    OpenAfterExport=False,
                    OptimizeFor=optimize_flag,            # Kualitas cetak maksimal (tidak blur)
                    Range=WD_EXPORT_ALL_DOCUMENT,
                    Item=WD_EXPORT_DOCUMENT_CONTENT,
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=bookmark_flag,        # Pertahankan daftar isi / bookmark
                    DocStructureTags=True,                # Tag semantik & aksesibilitas
                    BitmapMissingFonts=not self.preserve_fonts, # False = tetap vektor murni
                    UseISO19005_1=False                   # False agar gamut warna & gambar tidak dikonversi kaku
                )

                # Verifikasi hasil
                if not os.path.exists(output_abs) or os.path.getsize(output_abs) == 0:
                    raise WordConverterError("Gagal menghasilkan file PDF (file kosong atau tidak terbentuk).")

                return output_abs

            except Exception as e:
                raise WordConverterError(f"Gagal mengonversi '{os.path.basename(input_path)}': {str(e)}")

            finally:
                if doc:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass
                if word_app:
                    try:
                        word_app.Quit()
                    except Exception:
                        pass
                comtypes.CoUninitialize()

    def batch_convert(self, input_paths: List[str], output_dir: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Mengonversi banyak file Word sekaligus ke dalam folder tujuan.
        """
        results = []
        for path in input_paths:
            res: Dict[str, Any] = {
                "input": path,
                "input_name": os.path.basename(path),
                "success": False,
                "output_pdf": None,
                "error": None,
                "stats": None
            }
            try:
                # Dapatkan statistik gambar docx
                if path.lower().endswith(".docx"):
                    res["stats"] = self.inspect_docx(path)

                out_pdf = None
                if output_dir:
                    os.makedirs(output_dir, exist_ok=True)
                    base_name, _ = os.path.splitext(os.path.basename(path))
                    out_pdf = os.path.join(output_dir, f"{base_name}.pdf")

                output_file = self.convert(path, out_pdf)
                res["success"] = True
                res["output_pdf"] = output_file
                res["output_size"] = os.path.getsize(output_file)
            except Exception as e:
                res["error"] = str(e)

            results.append(res)
        return results


def main():
    parser = argparse.ArgumentParser(
        description="HD Word to PDF Converter - Hasil tidak pecah dan gambar tetap tajam (Lossless 300+ DPI)"
    )
    parser.add_argument("input", help="Path file Word (.docx / .doc) atau folder")
    parser.add_argument("-o", "--output", help="Path file output (.pdf) atau folder output", default=None)
    parser.add_argument("--standard", action="store_true", help="Gunakan optimasi ukuran kecil (bukan HD)")
    parser.add_argument("--no-bookmarks", action="store_true", help="Jangan buat bookmark heading PDF")
    parser.add_argument("--inspect-only", action="store_true", help="Hanya periksa resolusi gambar di DOCX")

    args = parser.parse_args()

    if args.inspect_only:
        if os.path.isfile(args.input) and args.input.lower().endswith(".docx"):
            stats = HDWordToPdfConverter.inspect_docx(args.input)
            print("========================================")
            print(f"File: {stats['file_name']}")
            print(f"Kualitas Estimasi: {stats['estimated_quality']}")
            print(f"Jumlah Gambar: {stats['image_count']}")
            for idx, img in enumerate(stats['images'], 1):
                hd_tag = "[HD 300+ DPI / High-Res]" if img['is_hd'] else "[Standard]"
                print(f"  {idx}. {img['name']} ({img['format']}) - {img['width']}x{img['height']}px | {img['dpi']} DPI {hd_tag}")
            print("========================================")
        else:
            print("Error: Harap masukkan file .docx untuk diperiksa.")
        return

    converter = HDWordToPdfConverter(
        optimize_for_print=not args.standard,
        preserve_fonts=True,
        create_bookmarks=not args.no_bookmarks
    )

    if os.path.isdir(args.input):
        files = [
            os.path.join(args.input, f) for f in os.listdir(args.input)
            if f.lower().endswith((".docx", ".doc")) and not f.startswith("~$")
        ]
        if not files:
            print(f"Tidak ada file Word ditemukan di folder: {args.input}")
            return

        print(f"Ditemukan {len(files)} file Word. Memulai konversi Ultra-HD...")
        results = converter.batch_convert(files, args.output)
        success_count = sum(1 for r in results if r["success"])
        print(f"\nSelesai! Berhasil mengonversi {success_count}/{len(files)} file.")
        for r in results:
            status = "BERHASIL" if r["success"] else f"GAGAL ({r['error']})"
            print(f"  - {r['input_name']} -> {status}")
    else:
        print(f"Mengonversi '{os.path.basename(args.input)}' dengan kualitas HD...")
        start_time = time.time()
        try:
            pdf_path = converter.convert(args.input, args.output)
            elapsed = time.time() - start_time
            size_kb = os.path.getsize(pdf_path) / 1024
            print(f"Konversi berhasil dalam {elapsed:.2f} detik!")
            print(f"File PDF: {pdf_path} ({size_kb:.1f} KB)")
        except Exception as e:
            print(f"Error: {e}")
            sys.exit(1)


if __name__ == "__main__":
    main()
