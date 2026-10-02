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
    import comtypes
    import comtypes.client
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
    Konverter PDF ke DOCX via Microsoft Word COM Automation.
    Word membuka PDF menggunakan engine PDF Reflow built-in,
    lalu menyimpan hasilnya sebagai DOCX.
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
        if not COMTYPES_AVAILABLE:
            raise PdfToWordError(
                "Modul 'comtypes' tidak ditemukan. Jalankan: pip install comtypes"
            )

        input_abs = os.path.abspath(input_pdf_path)
        if not os.path.isfile(input_abs):
            raise PdfToWordError(f"File PDF tidak ditemukan: {input_pdf_path}")

        if not input_abs.lower().endswith(".pdf"):
            raise PdfToWordError(f"File bukan berformat PDF: {input_pdf_path}")

        # Tentukan format dan COM constant
        fmt = (output_format or 'docx').lower().strip().lstrip('.')
        if fmt not in ('docx', 'doc'):
            fmt = 'docx'
        target_ext = f".{fmt}"
        word_format = WD_FORMAT_DOCUMENT_DEFAULT if fmt == 'docx' else WD_FORMAT_DOCUMENT_97

        # Tentukan path output
        if not output_docx_path:
            base, _ = os.path.splitext(input_abs)
            output_abs = f"{base}{target_ext}"
        else:
            base, _ = os.path.splitext(os.path.abspath(output_docx_path))
            output_abs = f"{base}{target_ext}"

        # Buat direktori output jika belum ada
        out_dir = os.path.dirname(output_abs)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Hapus file lama agar tidak terjadi conflict
        if os.path.exists(output_abs):
            try:
                os.remove(output_abs)
            except OSError as e:
                raise PdfToWordError(
                    f"Tidak dapat menimpa file Word yang sedang dibuka: {e}"
                )

        # Inisialisasi COM untuk thread saat ini
        comtypes.CoInitialize()

        with pdf_word_lock:
            word_app = None
            doc = None
            try:
                # Buka Microsoft Word
                word_app = comtypes.client.CreateObject("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = WD_ALERTS_NONE

                # Nonaktifkan dialog dan auto-update
                try:
                    word_app.Options.DoNotPromptForConvert = True
                    word_app.Options.UpdateLinksAtOpen = False
                    word_app.Options.WarnBeforeSavingAll = False
                except Exception:
                    pass

                # Buka PDF — Word akan otomatis menggunakan PDF Reflow
                doc = word_app.Documents.Open(
                    FileName=input_abs,
                    ConfirmConversions=False,
                    ReadOnly=False,          # Harus False agar bisa disimpan ulang
                    AddToRecentFiles=False,
                    Visible=False
                )

                # Simpan ke format yang dipilih
                doc.SaveAs2(
                    FileName=output_abs,
                    FileFormat=word_format,
                    AddToRecentFiles=False,
                )

                # Verifikasi hasil
                if not os.path.exists(output_abs) or os.path.getsize(output_abs) == 0:
                    raise PdfToWordError(
                        "Gagal menghasilkan file DOCX (file kosong atau tidak terbentuk)."
                    )

                return output_abs

            except PdfToWordError:
                raise
            except Exception as e:
                raise PdfToWordError(
                    f"Gagal mengonversi '{os.path.basename(input_pdf_path)}': {str(e)}"
                )
            finally:
                if doc:
                    try:
                        doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                    except Exception:
                        pass
                if word_app:
                    try:
                        word_app.Quit()
                    except Exception:
                        pass
                try:
                    comtypes.CoUninitialize()
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
