import math

import win32com.client
from PySide6.QtWidgets import QMessageBox


POINTS_PER_CM = 28.3464567
WD_COLLAPSE_END = 0
WD_GOTO_PAGE = 1
WD_GOTO_ABSOLUTE = 1
WD_PAGE_BREAK = 7
WD_SECTION_BREAK_CONTINUOUS = 3
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_CELL_ALIGN_VERTICAL_CENTER = 1
WD_STATISTIC_PAGES = 2
WD_FIELD_PAGE = 33
WD_HEADER_FOOTER_PRIMARY = 1
WD_HEADER_FOOTER_FIRST_PAGE = 2
WD_HEADER_FOOTER_EVEN_PAGES = 3
WD_PAGE_NUMBER_STYLE_ARABIC = 0


def _document_key(document):
    try:
        return document.FullName
    except Exception:
        return f"UNSAVED::{document.Name}"


def _find_selected_document(
    dropdown,
    parent,
    show_warning=True,
):
    selected_key = dropdown.currentData() or ""

    if not selected_key:
        if show_warning:
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

        if show_warning:
            QMessageBox.warning(
                parent,
                "Document Not Available",
                "That document is no longer open.\n\n"
                "Click Refresh and choose another document.",
            )
        return None, None

    except Exception as error:
        if show_warning:
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

    except Exception as error:
        QMessageBox.critical(
            parent,
            "Resize Error",
            "Could not resize the image(s).\n\n"
            f"Error:\n{error}",
        )


def _remove_page_fields(footer):
    for index in range(footer.Range.Fields.Count, 0, -1):
        field = footer.Range.Fields(index)

        if field.Type == WD_FIELD_PAGE:
            field.Delete()


def _position_page_fields(footer, horizontal_offset_points):
    for index in range(1, footer.Range.Fields.Count + 1):
        field = footer.Range.Fields(index)

        if field.Type != WD_FIELD_PAGE:
            continue

        number_format = field.Result.ParagraphFormat
        number_format.Alignment = WD_ALIGN_PARAGRAPH_CENTER
        number_format.LeftIndent = horizontal_offset_points
        number_format.RightIndent = -horizontal_offset_points


def _add_centered_page_number(footer, horizontal_offset_points):
    footer.PageNumbers.Add(
        PageNumberAlignment=WD_ALIGN_PARAGRAPH_CENTER,
        FirstPage=True,
    )
    _position_page_fields(footer, horizontal_offset_points)


def add_page_numbers(
    dropdown,
    parent,
    start_page,
    horizontal_offset_cm=0.0,
    footer_distance_cm=1.27,
):
    word, document = _find_selected_document(dropdown, parent)

    if document is None:
        return

    try:
        document.Activate()
        page_count = document.ComputeStatistics(WD_STATISTIC_PAGES)
        horizontal_offset_points = (
            horizontal_offset_cm * POINTS_PER_CM
        )
        footer_distance_points = footer_distance_cm * POINTS_PER_CM

        if start_page > page_count:
            QMessageBox.warning(
                parent,
                "Page Not Found",
                f"This document has {page_count} page(s).\n\n"
                f"Choose a starting page from 1 to {page_count}.",
            )
            return

        page_range = document.GoTo(
            What=WD_GOTO_PAGE,
            Which=WD_GOTO_ABSOLUTE,
            Count=start_page,
        )
        page_start = page_range.Start
        current_section = page_range.Sections(1)

        if start_page > 1 and current_section.Range.Start != page_start:
            break_range = document.Range(page_start, page_start)
            break_range.InsertBreak(WD_SECTION_BREAK_CONTINUOUS)
            page_range = document.GoTo(
                What=WD_GOTO_PAGE,
                Which=WD_GOTO_ABSOLUTE,
                Count=start_page,
            )
            current_section = page_range.Sections(1)

        start_section_index = None

        for index in range(1, document.Sections.Count + 1):
            section = document.Sections(index)

            if section.Range.Start == current_section.Range.Start:
                start_section_index = index
                break

        if start_section_index is None:
            raise RuntimeError("Could not locate the starting section.")

        footer_types = (
            WD_HEADER_FOOTER_PRIMARY,
            WD_HEADER_FOOTER_FIRST_PAGE,
            WD_HEADER_FOOTER_EVEN_PAGES,
        )

        # Remove only PAGE fields; other footer text remains untouched.
        for section_index in range(1, document.Sections.Count + 1):
            section = document.Sections(section_index)

            for footer_type in footer_types:
                _remove_page_fields(section.Footers(footer_type))

        for section_index in range(
            start_section_index,
            document.Sections.Count + 1,
        ):
            section = document.Sections(section_index)
            is_start_section = section_index == start_section_index
            section.PageSetup.FooterDistance = footer_distance_points

            for footer_type in footer_types:
                footer = section.Footers(footer_type)

                if is_start_section:
                    footer.LinkToPrevious = False

                if is_start_section or not footer.LinkToPrevious:
                    _add_centered_page_number(
                        footer,
                        horizontal_offset_points,
                    )

                footer.PageNumbers.RestartNumberingAtSection = (
                    is_start_section
                )
                footer.PageNumbers.NumberStyle = (
                    WD_PAGE_NUMBER_STYLE_ARABIC
                )

                if is_start_section:
                    footer.PageNumbers.StartingNumber = 1

        document.Fields.Update()

        QMessageBox.information(
            parent,
            "Page Numbers Added",
            f"Page {start_page} now starts at 1. "
            "The following pages continue as 2, 3, and so on.",
        )

    except Exception as error:
        QMessageBox.critical(
            parent,
            "Page Number Error",
            "Could not add page numbers to the document.\n\n"
            f"Error:\n{error}",
        )


def update_page_number_position(
    dropdown,
    parent,
    start_page,
    horizontal_offset_cm,
    footer_distance_cm,
):
    word, document = _find_selected_document(
        dropdown,
        parent,
        show_warning=False,
    )

    if document is None:
        return

    try:
        page_count = document.ComputeStatistics(WD_STATISTIC_PAGES)

        if start_page > page_count:
            return

        page_range = document.GoTo(
            What=WD_GOTO_PAGE,
            Which=WD_GOTO_ABSOLUTE,
            Count=start_page,
        )
        current_section = page_range.Sections(1)
        start_section_index = None

        for index in range(1, document.Sections.Count + 1):
            section = document.Sections(index)

            if section.Range.Start == current_section.Range.Start:
                start_section_index = index
                break

        if start_section_index is None:
            return

        horizontal_offset_points = (
            horizontal_offset_cm * POINTS_PER_CM
        )
        footer_distance_points = footer_distance_cm * POINTS_PER_CM
        footer_types = (
            WD_HEADER_FOOTER_PRIMARY,
            WD_HEADER_FOOTER_FIRST_PAGE,
            WD_HEADER_FOOTER_EVEN_PAGES,
        )

        for section_index in range(
            start_section_index,
            document.Sections.Count + 1,
        ):
            section = document.Sections(section_index)
            section.PageSetup.FooterDistance = footer_distance_points

            for footer_type in footer_types:
                _position_page_fields(
                    section.Footers(footer_type),
                    horizontal_offset_points,
                )

    except Exception:
        # Live position previews should never interrupt typing with dialogs.
        return
