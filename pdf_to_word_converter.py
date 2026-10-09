"""
PDF to Word Converter Engine
==============================
Engine konversi PDF ke DOCX menggunakan pdf2docx (High Performance Multi-Core)
dengan Fallback ke Microsoft Word COM Automation (Windows).

PERFORMA:
- pdf2docx: Menggunakan multi-processing otomatis untuk dokumen multi-halaman
  sehingga utilisasi CPU optimal.
- Fallback MS Word: Menggunakan singleton WordAppManager sehingga proses Word.exe
  tidak dibuka-tutup berulang kali (menghemat 3-5 detik per file).
"""

import os
import sys
import threading
import concurrent.futures
from typing import Optional, Dict, Any

try:
    from pdf2docx import Converter as Pdf2DocxConverter
    PDF2DOCX_AVAILABLE = True
except ImportError:
    PDF2DOCX_AVAILABLE = False

try:
    import comtypes
    import comtypes.client
    from converter import WordAppManager
    COMTYPES_AVAILABLE = True
except ImportError:
    COMTYPES_AVAILABLE = False


# Word COM Constants
WD_FORMAT_DOCUMENT_DEFAULT = 16   # .docx (Open XML)
WD_FORMAT_DOCUMENT_97     = 0    # .doc  (Word 97-2003)
WD_ALERTS_NONE = 0
WD_DO_NOT_SAVE_CHANGES = 0


class PdfToWordError(Exception):
    """Custom exception untuk error konversi PDF ke Word."""
    pass


# Lock global agar Word COM tidak diakses concurrent (STA thread model)
pdf_word_lock = threading.Lock()


def _run_with_timeout(func, args=(), kwargs=None, timeout_seconds=90):
    """Helper untuk menjalankan fungsi dengan batas waktu ketat agar tidak hanging/stuck."""
    if kwargs is None:
        kwargs = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(func, *args, **kwargs)
        try:
            return future.result(timeout=timeout_seconds)
        except concurrent.futures.TimeoutError:
            raise PdfToWordError(
                f"Proses konversi melebihi batas waktu ({timeout_seconds} detik). File PDF mungkin terlalu besar atau memiliki struktur tabel kompleks."
            )


class PdfToWordConverter:
    """
    Konverter PDF ke DOCX via pdf2docx (Parallel Multi-Core Engine)
    atau Microsoft Word COM Automation sebagai Fallback.
    """

    def __init__(self):
        pass

    def convert(
        self,
        input_pdf_path: str,
        output_docx_path: Optional[str] = None,
        output_format: str = 'docx',
        keep_word_open: bool = True
    ) -> str:
        """
        Mengonversi satu file PDF ke DOCX atau DOC.

        Args:
            input_pdf_path: Path ke file PDF sumber.
            output_docx_path: Path output (opsional).
            output_format: 'docx' (default) atau 'doc'.
            keep_word_open: Jika menggunakan COM fallback, pertahankan Word tetap hidup.

        Returns:
            Path absolut ke file Word hasil konversi.
        """
        input_abs = os.path.abspath(input_pdf_path)
        if not os.path.isfile(input_abs):
            raise PdfToWordError(f"File PDF tidak ditemukan: {input_pdf_path}")

        if not input_abs.lower().endswith(".pdf"):
            raise PdfToWordError(f"File bukan berformat PDF: {input_pdf_path}")

        fmt = (output_format or 'docx').lower().strip().lstrip('.')
        if fmt not in ('docx', 'doc'):
            fmt = 'docx'
        target_ext = f".{fmt}"

        if not output_docx_path:
            base, _ = os.path.splitext(input_abs)
            output_abs = f"{base}{target_ext}"
        else:
            base, _ = os.path.splitext(os.path.abspath(output_docx_path))
            output_abs = f"{base}{target_ext}"

        out_dir = os.path.dirname(output_abs)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        if os.path.exists(output_abs):
            try:
                os.remove(output_abs)
            except OSError as e:
                raise PdfToWordError(f"Tidak dapat menimpa file Word yang sedang dibuka: {e}")

        # OPSI 1: Gunakan pdf2docx jika output adalah DOCX
        if PDF2DOCX_AVAILABLE and fmt == 'docx':
            def _convert_pdf2docx():
                cv = Pdf2DocxConverter(input_abs)
                try:
                    # Deteksi multi-core untuk mempercepat konversi dokumen panjang
                    cpu_cnt = os.cpu_count() or 1
                    use_mp = cpu_cnt > 2
                    cv.convert(output_abs, start=0, end=None, multi_processing=use_mp, cpu_count=min(cpu_cnt, 4))
                finally:
                    cv.close()

            try:
                _run_with_timeout(_convert_pdf2docx, timeout_seconds=90)
                if os.path.exists(output_abs) and os.path.getsize(output_abs) > 0:
                    return output_abs
            except Exception:
                # Jika pdf2docx gagal atau timeout, lanjut mencoba MS Word COM Fallback...
                pass

        # OPSI 2: Fallback ke MS Word COM Automation (Singleton WordAppManager)
        if not COMTYPES_AVAILABLE:
            raise PdfToWordError(
                "Tidak ada engine konversi PDF ke Word yang tersedia (install pdf2docx atau comtypes)."
            )

        word_format = WD_FORMAT_DOCUMENT_DEFAULT if fmt == 'docx' else WD_FORMAT_DOCUMENT_97

        def _convert_word_com():
            with pdf_word_lock:
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

                    doc.SaveAs2(
                        FileName=output_abs,
                        FileFormat=word_format,
                        AddToRecentFiles=False,
                    )

                    if not os.path.exists(output_abs) or os.path.getsize(output_abs) == 0:
                        raise PdfToWordError("Gagal menghasilkan file DOCX (file kosong atau tidak terbentuk).")

                    return output_abs

                except PdfToWordError:
                    raise
                except Exception as e:
                    manager.word_app = None
                    raise PdfToWordError(f"Gagal mengonversi '{os.path.basename(input_pdf_path)}': {str(e)}")
                finally:
                    if doc:
                        try:
                            doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                        except Exception:
                            pass
                    if not keep_word_open:
                        manager.close_app()

        return _run_with_timeout(_convert_word_com, timeout_seconds=75)

    def get_pdf_info(self, pdf_path: str) -> Dict[str, Any]:
        """
        Mengambil informasi dasar file PDF (ukuran, nama).
        """
        abs_path = os.path.abspath(pdf_path)
        info: Dict[str, Any] = {
            "file_name": os.path.basename(abs_path),
            "file_size_bytes": 0,
            "file_size_kb": 0,
            "exists": os.path.isfile(abs_path),
        }
        if info["exists"]:
            size = os.path.getsize(abs_path)
            info["file_size_bytes"] = size
            info["file_size_kb"] = round(size / 1024, 1)
        return info
