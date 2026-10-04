"""
PDF to Word Converter Engine
==============================
Engine konversi PDF ke DOCX menggunakan Word COM Automation (Windows).
Strategi: Buka PDF langsung via Microsoft Word (built-in PDF reader),
kemudian Save As ke format DOCX — mempertahankan layout, teks, dan gambar
semaksimal mungkin sesuai kemampuan Word.

Catatan:
- Membutuhkan Microsoft Word yang terinstal (Word 2013+, direkomendasikan Word 2019/365).
- Word 2013+ mendukung membuka dan mengedit PDF secara native (PDF Reflow).
- Tidak bergantung pada library pihak ketiga berbayar.
"""

import os
import threading
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


class PdfToWordConverter:
    """
    Konverter PDF ke DOCX via pdf2docx (High Performance Native Python)
    atau Microsoft Word COM Automation sebagai Fallback.
    """

    def __init__(self):
        pass

    def convert(self, input_pdf_path: str, output_docx_path: Optional[str] = None, output_format: str = 'docx') -> str:
        """
        Mengonversi satu file PDF ke DOCX atau DOC.

        Args:
            input_pdf_path: Path ke file PDF sumber.
            output_docx_path: Path output (opsional).
            output_format: 'docx' (default) atau 'doc'.

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

        # OPSI 1: Gunakan pdf2docx jika tersedia dan output adalah DOCX (Sangat Cepat & Tanpa Word)
        if PDF2DOCX_AVAILABLE and fmt == 'docx':
            try:
                cv = Pdf2DocxConverter(input_abs)
                cv.convert(output_abs, start=0, end=None)
                cv.close()

                if os.path.exists(output_abs) and os.path.getsize(output_abs) > 0:
                    return output_abs
            except Exception as e:
                # Jika pdf2docx gagal (misal font khusus), coba fallback ke MS Word COM
                pass

        # OPSI 2: Fallback ke MS Word COM Automation via WordAppManager
        if not COMTYPES_AVAILABLE:
            raise PdfToWordError(
                "Tidak ada engine konversi PDF ke Word yang tersedia (install pdf2docx atau comtypes)."
            )

        word_format = WD_FORMAT_DOCUMENT_DEFAULT if fmt == 'docx' else WD_FORMAT_DOCUMENT_97
        manager = WordAppManager.get_instance()

        with pdf_word_lock:
            word_app = manager.acquire_app()
            doc = None
            try:
                doc = word_app.Documents.Open(
                    FileName=input_abs,
                    ConfirmConversions=False,
                    ReadOnly=False,
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
                raise PdfToWordError(f"Gagal mengonversi '{os.path.basename(input_pdf_path)}': {str(e)}")
            finally:
                if doc:
                    try:
                        doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                    except Exception:
                        pass

    def get_pdf_info(self, pdf_path: str) -> Dict[str, Any]:
        """
        Mengambil informasi dasar file PDF (ukuran, nama).
        Tidak membutuhkan library tambahan.
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
