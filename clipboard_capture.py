import hashlib
import shutil
import tempfile
import time
from pathlib import Path

import win32com.client
from PySide6.QtCore import QByteArray, QBuffer, QIODevice, QObject, QTimer, Signal
from PySide6.QtGui import QGuiApplication, QImage
from PySide6.QtWidgets import QMessageBox


SUPPORTED_IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".gif",
    ".webp",
    ".tif",
    ".tiff",
}


def _document_key(document):
    try:
        return document.FullName
    except Exception:
        return f"UNSAVED::{document.Name}"


class ImageCaptureManager(QObject):
    count_changed = Signal(int)
    state_changed = Signal(str)
    capture_finished = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)

        self.clipboard = QGuiApplication.clipboard()
        self.is_recording = False
        self.captured_files = []
        self.capture_directory = None

        self._last_signature = ""
        self._last_signature_time = 0.0
        self._capture_number = 0

    @property
    def count(self):
        return len(self.captured_files)

    def start(self):
        if self.is_recording:
            return

        self.clear()
        self.capture_directory = Path(
            tempfile.mkdtemp(prefix="ToolPy_Captured_Images_")
        )

        self.is_recording = True
        self._last_signature = ""
        self._last_signature_time = 0.0

        self.clipboard.dataChanged.connect(self._clipboard_changed)

        self.state_changed.emit("Recording — copy images now")
        self.count_changed.emit(0)

    def stop(self):
        if not self.is_recording:
            return

        try:
            self.clipboard.dataChanged.disconnect(self._clipboard_changed)
        except (RuntimeError, TypeError):
            pass

        self.is_recording = False
        self.state_changed.emit("Idle")
        self.capture_finished.emit()

    def cancel(self):
        self.stop()
        self.clear()
        self.state_changed.emit("Idle")

    def clear(self):
        self.captured_files.clear()
        self._capture_number = 0
        self.count_changed.emit(0)

        if self.capture_directory and self.capture_directory.exists():
            shutil.rmtree(self.capture_directory, ignore_errors=True)

        self.capture_directory = None

    def _clipboard_changed(self):
        if not self.is_recording:
            return

        # Some applications update the clipboard in several steps.
        QTimer.singleShot(120, self._capture_current_clipboard)

    def _capture_current_clipboard(self):
        if not self.is_recording or self.capture_directory is None:
            return

        mime_data = self.clipboard.mimeData()

        if mime_data is None:
            return

        images = []
        signature_parts = []

        if mime_data.hasUrls():
            for url in mime_data.urls():
                if not url.isLocalFile():
                    continue

                file_path = Path(url.toLocalFile())

                if file_path.suffix.lower() not in SUPPORTED_IMAGE_EXTENSIONS:
                    continue

                image = QImage(str(file_path))

                if image.isNull():
                    continue

                images.append(image)
                signature_parts.append(
                    f"file:{file_path.resolve()}:{file_path.stat().st_mtime_ns}"
                )

        # Browsers, Paint, screenshots, Discord, and similar apps usually
        # provide actual image data rather than file paths.
        if not images and mime_data.hasImage():
            image = self.clipboard.image()

            if not image.isNull():
                images.append(image)
                signature_parts.append(
                    "image:" + self._image_hash(image)
                )

        if not images:
            return

        signature = "|".join(signature_parts)
        now = time.monotonic()

        # Clipboard signals can fire more than once for one copy operation.
        if (
            signature == self._last_signature
            and now - self._last_signature_time < 1.0
        ):
            return

        self._last_signature = signature
        self._last_signature_time = now

        added = 0

        for image in images:
            self._capture_number += 1
            output_path = (
                self.capture_directory
                / f"capture_{self._capture_number:04d}.png"
            )

            if image.save(str(output_path), "PNG"):
                self.captured_files.append(output_path)
                added += 1

        if added:
            self.count_changed.emit(self.count)
            self.state_changed.emit("Recording — copy more or paste")

    @staticmethod
    def _image_hash(image):
        byte_array = QByteArray()
        buffer = QBuffer(byte_array)
        buffer.open(QIODevice.OpenModeFlag.WriteOnly)
        image.save(buffer, "PNG")
        buffer.close()

        return hashlib.sha256(bytes(byte_array)).hexdigest()

    def paste_into_word(self, document_dropdown, parent):
        if not self.captured_files:
            QMessageBox.warning(
                parent,
                "No Captured Images",
                "Start Capture and copy at least one image first.",
            )
            return False

        selected_key = document_dropdown.currentData() or ""

        if not selected_key:
            QMessageBox.warning(
                parent,
                "No Document Selected",
                "Choose an open Word document first.",
            )
            return False

        try:
            word = win32com.client.GetActiveObject("Word.Application")
            document = None

            for index in range(1, word.Documents.Count + 1):
                open_document = word.Documents(index)

                if _document_key(open_document) == selected_key:
                    document = open_document
                    break

            if document is None:
                QMessageBox.warning(
                    parent,
                    "Document Not Available",
                    "That Word document is no longer open.",
                )
                return False

            document.Activate()
            pasted_count = 0

            for image_path in self.captured_files:
                insert_position = max(document.Content.End - 1, 0)
                insert_range = document.Range(
                    Start=insert_position,
                    End=insert_position,
                )

                document.InlineShapes.AddPicture(
                    FileName=str(image_path),
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Range=insert_range,
                )

                paragraph_position = max(document.Content.End - 1, 0)
                paragraph_range = document.Range(
                    Start=paragraph_position,
                    End=paragraph_position,
                )
                paragraph_range.InsertParagraphAfter()

                pasted_count += 1

            QMessageBox.information(
                parent,
                "Images Pasted",
                f"Successfully pasted {pasted_count} image(s).\n\n"
                "The Word document was not automatically saved.",
            )

            self.stop()
            self.clear()
            self.state_changed.emit("Idle")
            return True

        except Exception as error:
            QMessageBox.critical(
                parent,
                "Paste Error",
                "ToolPy could not paste the captured images into Word.\n\n"
                f"Error:\n{error}",
            )
            return False
