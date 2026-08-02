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


class ReplacementQueueManager(QObject):
    state_changed = Signal(str)
    count_changed = Signal(int)
    current_changed = Signal(int, int)

    def __init__(self, parent=None):
        super().__init__(parent)

        self.clipboard = QGuiApplication.clipboard()
        self.is_recording = False
        self.queue_locked = False
        self.captured_files = []
        self.current_index = 0
        self.capture_directory = None

        self._last_signature = ""
        self._last_signature_time = 0.0
        self._capture_number = 0

    @property
    def count(self):
        return len(self.captured_files)

    @property
    def has_remaining(self):
        return self.current_index < self.count

    @property
    def remaining(self):
        return max(self.count - self.current_index, 0)

    def start_capture(self):
        if self.is_recording:
            return

        self.clear()

        self.capture_directory = Path(
            tempfile.mkdtemp(prefix="ToolPy_Replacement_Queue_")
        )
        self.is_recording = True
        self.queue_locked = False
        self.clipboard.dataChanged.connect(self._clipboard_changed)

        self.state_changed.emit("Recording — copy images now")

    def finish_capture(self, parent):
        if not self.is_recording:
            return False

        if self.count == 0:
            QMessageBox.warning(
                parent,
                "No Images Captured",
                "Copy at least one image first.",
            )
            return False

        self._disconnect_clipboard()
        self.is_recording = False
        self.queue_locked = True
        self.current_index = 0

        self.state_changed.emit("Replacement mode")
        self.current_changed.emit(1, self.count)
        return True

    def clear(self):
        self._disconnect_clipboard()

        self.is_recording = False
        self.queue_locked = False
        self.current_index = 0
        self.captured_files.clear()
        self._capture_number = 0
        self._last_signature = ""
        self._last_signature_time = 0.0

        if self.capture_directory and self.capture_directory.exists():
            shutil.rmtree(
                self.capture_directory,
                ignore_errors=True,
            )

        self.capture_directory = None
        self.state_changed.emit("Idle")
        self.count_changed.emit(0)
        self.current_changed.emit(0, 0)

    def replace_selected(self, parent):
        if not self._queue_is_ready(parent):
            return False

        replacement_file = self.captured_files[self.current_index]

        try:
            word = win32com.client.GetActiveObject(
                "Word.Application"
            )
            selection = word.Selection

            if selection.InlineShapes.Count > 0:
                self._replace_inline_shape(
                    selection.InlineShapes(1),
                    replacement_file,
                )
            else:
                try:
                    shape_range = selection.ShapeRange
                except Exception:
                    shape_range = None

                if shape_range is None or shape_range.Count == 0:
                    QMessageBox.warning(
                        parent,
                        "No Image Selected",
                        "Select one image inside Microsoft Word first.",
                    )
                    return False

                self._replace_floating_shape(
                    shape_range(1),
                    replacement_file,
                )

            self._advance_queue()
            return True

        except Exception as error:
            QMessageBox.critical(
                parent,
                "Replace Error",
                "ToolPy could not replace the selected image.\n\n"
                f"Error:\n{error}",
            )
            return False

    def paste_next(self, parent):
        if not self._queue_is_ready(parent):
            return False

        image_file = self.captured_files[self.current_index]

        try:
            word = win32com.client.GetActiveObject(
                "Word.Application"
            )
            selection = word.Selection
            insert_range = selection.Range.Duplicate
            insert_range.Collapse(0)

            selection.Document.InlineShapes.AddPicture(
                FileName=str(image_file),
                LinkToFile=False,
                SaveWithDocument=True,
                Range=insert_range,
            )

            self._advance_queue()
            return True

        except Exception as error:
            QMessageBox.critical(
                parent,
                "Paste Error",
                "ToolPy could not paste the next queued image.\n\n"
                f"Error:\n{error}",
            )
            return False

    def _queue_is_ready(self, parent):
        if not self.queue_locked:
            QMessageBox.warning(
                parent,
                "Queue Not Ready",
                "Start Capture, copy images, then finish capture.",
            )
            return False

        if not self.has_remaining:
            QMessageBox.warning(
                parent,
                "Queue Complete",
                "Every captured image has already been used.",
            )
            return False

        return True

    def _advance_queue(self):
        self.current_index += 1

        if self.current_index >= self.count:
            self.state_changed.emit("Queue complete")
            self.current_changed.emit(self.count, self.count)
        else:
            self.state_changed.emit(
                "Replacement mode — choose the next action"
            )
            self.current_changed.emit(
                self.current_index + 1,
                self.count,
            )

    def _clipboard_changed(self):
        if self.is_recording:
            QTimer.singleShot(
                120,
                self._capture_current_clipboard,
            )

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
                    f"file:{file_path.resolve()}:"
                    f"{file_path.stat().st_mtime_ns}"
                )

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
                / f"replacement_{self._capture_number:04d}.png"
            )

            if image.save(str(output_path), "PNG"):
                self.captured_files.append(output_path)
                added += 1

        if added:
            self.count_changed.emit(self.count)
            self.state_changed.emit(
                "Recording — F finishes, Esc cancels"
            )

    @staticmethod
    def _image_hash(image):
        byte_array = QByteArray()
        buffer = QBuffer(byte_array)
        buffer.open(QIODevice.OpenModeFlag.WriteOnly)
        image.save(buffer, "PNG")
        buffer.close()

        return hashlib.sha256(
            bytes(byte_array)
        ).hexdigest()

    def _disconnect_clipboard(self):
        try:
            self.clipboard.dataChanged.disconnect(
                self._clipboard_changed
            )
        except (RuntimeError, TypeError):
            pass

    @staticmethod
    def _replace_inline_shape(old_image, replacement_file):
        width = old_image.Width
        height = old_image.Height
        lock_aspect_ratio = old_image.LockAspectRatio

        old_range = old_image.Range.Duplicate
        document = old_range.Document

        old_image.Delete()
        old_range.Collapse(1)

        new_image = document.InlineShapes.AddPicture(
            FileName=str(replacement_file),
            LinkToFile=False,
            SaveWithDocument=True,
            Range=old_range,
        )

        new_image.LockAspectRatio = False
        new_image.Width = width
        new_image.Height = height
        new_image.LockAspectRatio = lock_aspect_ratio

    @staticmethod
    def _replace_floating_shape(old_image, replacement_file):
        document = old_image.Anchor.Document
        anchor = old_image.Anchor.Duplicate

        width = old_image.Width
        height = old_image.Height
        left = old_image.Left
        top = old_image.Top
        rotation = old_image.Rotation
        lock_aspect_ratio = old_image.LockAspectRatio
        wrap_type = old_image.WrapFormat.Type
        relative_horizontal = old_image.RelativeHorizontalPosition
        relative_vertical = old_image.RelativeVerticalPosition

        try:
            layout_in_cell = old_image.LayoutInCell
        except Exception:
            layout_in_cell = None

        try:
            lock_anchor = old_image.LockAnchor
        except Exception:
            lock_anchor = None

        old_image.Delete()

        new_image = document.Shapes.AddPicture(
            FileName=str(replacement_file),
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left,
            Top=top,
            Width=width,
            Height=height,
            Anchor=anchor,
        )

        new_image.LockAspectRatio = False
        new_image.Width = width
        new_image.Height = height
        new_image.Left = left
        new_image.Top = top
        new_image.Rotation = rotation
        new_image.RelativeHorizontalPosition = relative_horizontal
        new_image.RelativeVerticalPosition = relative_vertical
        new_image.WrapFormat.Type = wrap_type

        if layout_in_cell is not None:
            new_image.LayoutInCell = layout_in_cell

        if lock_anchor is not None:
            new_image.LockAnchor = lock_anchor

        new_image.LockAspectRatio = lock_aspect_ratio
