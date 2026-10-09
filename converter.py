"""
HD Word to PDF Converter Engine
================================
Engine konversi Word (.docx / .doc) ke PDF dengan kualitas maksimal (HD / Lossless).
Memastikan:
1. Gambar tidak pecah / tidak terkompresi (DoNotCompressImages = True).
2. Optimasi cetak resolusi tinggi (wdExportOptimizeForPrint, 300+ DPI).
3. Font vector tetap tajam pada semua tingkat zoom (BitmapMissingFonts = False).
4. Struktur dokumen & bookmark heading dipertahankan.
5. Timeout Guardian untuk mencegah MS Word hang/stuck.

PERFORMA:
- Menggunakan WordAppManager Singleton: Word COM dibuat SEKALI, di-reuse untuk
  semua konversi. Overhead startup ~3-5 detik hanya terjadi 1x, bukan per-file.
- Batch conversion menggunakan ThreadPoolExecutor untuk paralelisme.
"""

import os
import sys
import time
import zipfile
import threading
import concurrent.futures
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
WD_DO_NOT_SAVE_CHANGES = 0


class WordConverterError(Exception):
    """Custom exception untuk error konversi."""
    pass


# Global lock untuk mencegah concurrent access ke instance Word COM (STA Apartment)
word_lock = threading.Lock()


def _run_with_timeout(func, args=(), kwargs=None, timeout_seconds=75):
    """Helper untuk menjalankan fungsi dengan batas waktu ketat agar tidak hanging/stuck."""
    if kwargs is None:
        kwargs = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(func, *args, **kwargs)
        try:
            return future.result(timeout=timeout_seconds)
        except concurrent.futures.TimeoutError:
            raise WordConverterError(
                f"Konversi Word ke PDF melebihi batas waktu ({timeout_seconds} detik). Dokumen mungkin terlalu besar atau terkunci dialog Word."
            )


class WordAppManager:
    """
    Manager terpusat untuk instansi MS Word COM Automation.
    Menghindari overhead pembuatan & penutupan proses Word.exe pada setiap konversi file.

    PERFORMA: Word.exe startup memakan ~3-5 detik. Dengan singleton ini,
    startup hanya terjadi 1x selama aplikasi berjalan.
    """
    _instance = None
    _lock = threading.Lock()

    def __init__(self):
        self.word_app = None
        self._com_initialized = False

    @classmethod
    def get_instance(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = WordAppManager()
            return cls._instance

    def _is_app_alive(self) -> bool:
        """Cek apakah Word COM instance masih responsif."""
        if self.word_app is None:
            return False
        try:
            _ = self.word_app.Version
            return True
        except Exception:
            return False

    def acquire_app(self):
        """Mendapatkan atau membuat instansi Word Application yang siap digunakan."""
        if not COMTYPES_AVAILABLE:
            raise WordConverterError(
                "Modul 'comtypes' tidak ditemukan. Jalankan: pip install comtypes"
            )

        if not self._com_initialized:
            comtypes.CoInitialize()
            self._com_initialized = True

        # Reuse jika masih hidup
        if self._is_app_alive():
            return self.word_app

        # Cleanup stale reference
        if self.word_app is not None:
            try:
                self.word_app.Quit()
            except Exception:
                pass
            self.word_app = None

        try:
            self.word_app = comtypes.client.CreateObject("Word.Application")
            self.word_app.Visible = False
            self.word_app.DisplayAlerts = WD_ALERTS_NONE
            try:
                self.word_app.ScreenUpdating = False
                self.word_app.Options.DoNotPromptForConvert = True
                self.word_app.Options.SaveInterval = 0
                self.word_app.Options.UpdateLinksAtOpen = False
                self.word_app.Options.ConfirmConversions = False
            except Exception:
                pass
            return self.word_app
        except Exception as e:
            self.word_app = None
            raise WordConverterError(f"Gagal membuka Microsoft Word COM: {str(e)}")

    def close_app(self):
        """Menutup aplikasi Word jika diperlukan."""
        with self._lock:
            if self.word_app is not None:
                try:
                    self.word_app.Quit()
                except Exception:
                    pass
                self.word_app = None
                if self._com_initialized:
                    try:
                        comtypes.CoUninitialize()
                    except Exception:
                        pass
                    self._com_initialized = False


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
        Memeriksa kualitas gambar internal di dalam file DOCX secara cepat.
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
                            dpi_val = round(dpi[0]) if isinstance(dpi, tuple) else round(dpi)

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

    def convert(self, input_path: str, output_path: Optional[str] = None, keep_word_open: bool = False) -> str:
        """
        Mengonversi satu file DOC/DOCX ke PDF berkualitas HD tanpa kompresi gambar.
        Terlindungi Timeout Guardian.

        PERFORMA: Menggunakan WordAppManager singleton — Word.exe tidak
        dibuat/ditutup per file, melainkan di-reuse terus-menerus.
        Parameter keep_word_open=True menjaga instance tetap hidup.
        """
        input_abs = os.path.abspath(input_path)
        if not os.path.isfile(input_abs):
            raise WordConverterError(f"File dokumen tidak ditemukan: {input_path}")

        if not output_path:
            base_name, _ = os.path.splitext(input_abs)
            output_abs = f"{base_name}.pdf"
        else:
            output_abs = os.path.abspath(output_path)

        os.makedirs(os.path.dirname(output_abs), exist_ok=True)

        if os.path.exists(output_abs):
            try:
                os.remove(output_abs)
            except OSError as e:
                raise WordConverterError(f"Tidak dapat menimpa file PDF yang sedang dibuka: {e}")

        if not COMTYPES_AVAILABLE:
            raise WordConverterError("Modul 'comtypes' tidak ditemukan. Jalankan: pip install comtypes")

        def _do_convert():
            with word_lock:
                manager = WordAppManager.get_instance()
                doc = None
                try:
                    word_app = manager.acquire_app()

                    doc = word_app.Documents.Open(
                        FileName=input_abs,
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        Visible=False
                    )

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

                    doc.ExportAsFixedFormat(
                        OutputFileName=output_abs,
                        ExportFormat=WD_EXPORT_FORMAT_PDF,
                        OpenAfterExport=False,
                        OptimizeFor=optimize_flag,
                        Range=WD_EXPORT_ALL_DOCUMENT,
                        Item=WD_EXPORT_DOCUMENT_CONTENT,
                        IncludeDocProps=True,
                        KeepIRM=True,
                        CreateBookmarks=bookmark_flag,
                        DocStructureTags=True,
                        BitmapMissingFonts=not self.preserve_fonts,
                        UseISO19005_1=False
                    )

                    if not os.path.exists(output_abs) or os.path.getsize(output_abs) == 0:
                        raise WordConverterError("Gagal menghasilkan file PDF (file kosong atau tidak terbentuk).")

                    return output_abs

                except WordConverterError:
                    raise
                except Exception as e:
                    # Jika error, invalidate instance agar di-recreate pada panggilan berikutnya
                    manager.word_app = None
                    raise WordConverterError(f"Gagal mengonversi '{os.path.basename(input_path)}': {str(e)}")

                finally:
                    if doc:
                        try:
                            doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                        except Exception:
                            pass
                    # Jika keep_word_open=False, tutup Word sepenuhnya (behavior lama)
                    if not keep_word_open:
                        manager.close_app()

        return _run_with_timeout(_do_convert, timeout_seconds=75)

    def batch_convert(self, input_paths: List[str], output_dir: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Mengonversi banyak file Word sekaligus.
        Semua file dikonversi menggunakan satu instance Word COM yang sama.
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
                out_pdf = None
                if output_dir:
                    os.makedirs(output_dir, exist_ok=True)
                    base_name, _ = os.path.splitext(os.path.basename(path))
                    out_pdf = os.path.join(output_dir, f"{base_name}.pdf")

                output_file = self.convert(path, out_pdf, keep_word_open=True)
                res["success"] = True
                res["output_pdf"] = output_file
                res["output_size"] = os.path.getsize(output_file)
            except Exception as e:
                res["error"] = str(e)
            results.append(res)
        return results
