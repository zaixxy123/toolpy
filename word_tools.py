import math

import win32com.client
from PySide6.QtWidgets import QMessageBox


POINTS_PER_CM = 28.3464567
WD_COLLAPSE_END = 0
WD_PAGE_BREAK = 7
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_CELL_ALIGN_VERTICAL_CENTER = 1


def _document_key(document):
    try:
        return document.FullName
    except Exception:
        return f"UNSAVED::{document.Name}"


def _find_selected_document(dropdown, parent):
    selected_key = dropdown.currentData() or ""

    if not selected_key:
        QMessageBox.warning(
            parent,
            "No Document Selected",
            "Choose an open Word document first.",
        )
        return None, None

    try:
        word = win32com.client.GetActiveObject("Word.Application")

        for index in range(1, word.Documents.Count + 1):
            document = word.Documents(index)

            if _document_key(document) == selected_key:
                return word, document

        QMessageBox.warning(
            parent,
            "Document Not Available",
            "That document is no longer open.\n\n"
            "Click Refresh and choose another document.",
        )
        return None, None

    except Exception as error:
        QMessageBox.critical(
            parent,
            "Word Error",
            "Microsoft Word could not be accessed.\n\n"
            f"Error:\n{error}",
        )
        return None, None


def refresh_open_documents(
    dropdown,
    parent,
    show_warning=True,
):
    previous_key = dropdown.currentData() or ""

    dropdown.blockSignals(True)
    dropdown.clear()
    dropdown.addItem("Choose an open Word document", "")

    try:
        word = win32com.client.GetActiveObject("Word.Application")

        if word.Documents.Count == 0:
            if show_warning:
                QMessageBox.warning(
                    parent,
                    "No Documents",
                    "Word is open, but no document is currently open.",
                )
            return

        selected_index = 1

        for index in range(1, word.Documents.Count + 1):
            document = word.Documents(index)
            key = _document_key(document)

            dropdown.addItem(document.Name, key)

            if key == previous_key:
                selected_index = dropdown.count() - 1

        dropdown.setCurrentIndex(selected_index)

    except Exception:
        if show_warning:
            QMessageBox.warning(
                parent,
                "Word Not Open",
                "Open Microsoft Word and at least one document first.",
            )

    finally:
        dropdown.blockSignals(False)


def set_all_images_behind_text(dropdown, parent):
    word, document = _find_selected_document(dropdown, parent)

    if document is None:
        return

    try:
        document.Activate()
        converted_count = 0

        for index in range(document.InlineShapes.Count, 0, -1):
            document.InlineShapes(index).ConvertToShape()
            converted_count += 1

        shape_count = document.Shapes.Count

        for index in range(1, shape_count + 1):
            document.Shapes(index).WrapFormat.Type = 3

        QMessageBox.information(
            parent,
            "Finished",
            f"Done!\n\n"
            f"Document: {document.Name}\n"
            f"Inline pictures converted: {converted_count}\n"
            f"Total shapes processed: {shape_count}\n\n"
            "The document was not automatically saved.",
        )

    except Exception as error:
        QMessageBox.critical(
            parent,
            "Word Error",
            "The selected Word document could not be processed.\n\n"
            f"Error:\n{error}",
        )


def resize_images(
    dropdown,
    parent,
    width_cm,
    height_cm,
    resize_all,
):
    word, document = _find_selected_document(dropdown, parent)

    if document is None:
        return

    width_points = width_cm * POINTS_PER_CM
    height_points = height_cm * POINTS_PER_CM

    try:
        document.Activate()
        resized = 0

        if resize_all:
            for index in range(1, document.InlineShapes.Count + 1):
                picture = document.InlineShapes(index)
                picture.LockAspectRatio = False
                picture.Width = width_points
                picture.Height = height_points
                resized += 1

            for index in range(1, document.Shapes.Count + 1):
                picture = document.Shapes(index)
                picture.LockAspectRatio = False
                picture.Width = width_points
                picture.Height = height_points
                resized += 1

        else:
            selection = word.Selection

            if selection.InlineShapes.Count > 0:
                for index in range(1, selection.InlineShapes.Count + 1):
                    picture = selection.InlineShapes(index)
                    picture.LockAspectRatio = False
                    picture.Width = width_points
                    picture.Height = height_points
                    resized += 1
            else:
                try:
                    selected_shapes = selection.ShapeRange

                    for index in range(1, selected_shapes.Count + 1):
                        picture = selected_shapes(index)
                        picture.LockAspectRatio = False
                        picture.Width = width_points
                        picture.Height = height_points
                        resized += 1
                except Exception:
                    pass

        if resized == 0:
            QMessageBox.warning(
                parent,
                "No Image Selected",
                "Select an image inside Word first.",
            )
            return

        QMessageBox.information(
            parent,
            "Finished",
            f"Resized {resized} image(s).\n\n"
            f"Width: {width_cm:.2f} cm\n"
            f"Height: {height_cm:.2f} cm\n\n"
            "The document was not automatically saved.",
        )

    except Exception as error:
        QMessageBox.critical(
            parent,
            "Resize Error",
            "Could not resize the image(s).\n\n"
            f"Error:\n{error}",
        )
