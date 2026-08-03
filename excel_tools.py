import re
from datetime import date, datetime, timedelta

import pywintypes
import win32com.client
from PySide6.QtWidgets import QMessageBox


MONTH_FORMATS = [
    "%B %d, %Y",
    "%B %d %Y",
    "%b %d, %Y",
    "%b %d %Y",
    "%d %B %Y",
    "%d %b %Y",
    "%B %d, %y",
    "%B %d %y",
    "%b %d, %y",
    "%b %d %y",
    "%d %B %y",
    "%d %b %y",
]

OUTPUT_FORMATS = {
    "MM/DD/YYYY": "mm/dd/yyyy",
    "DD/MM/YYYY": "dd/mm/yyyy",
    "Month DD, YYYY": "mmmm dd, yyyy",
    "MMM DD, YYYY": "mmm dd, yyyy",
    "YYYY-MM-DD": "yyyy-mm-dd",
}

ALIGNMENTS = {
    "Left": -4131,
    "Center": -4108,
    "Right": -4152,
}


def _get_excel():
    return win32com.client.GetActiveObject(
        "Excel.Application"
    )


def _find_workbook(excel, workbook_name):
    for workbook in excel.Workbooks:
        if workbook.Name == workbook_name:
            return workbook
    return None


def _find_worksheet(workbook, worksheet_name):
    for worksheet in workbook.Worksheets:
        if worksheet.Name == worksheet_name:
            return worksheet
    return None


def _validate_target(
    parent,
    workbook_dropdown,
    worksheet_dropdown,
    start_cell_text,
):
    workbook_name = workbook_dropdown.currentData()
    worksheet_name = worksheet_dropdown.currentData()
    start_cell_text = start_cell_text.strip().upper()

    if not workbook_name:
        QMessageBox.warning(
            parent,
            "No Workbook Selected",
            "Choose an open Excel workbook first.",
        )
        return None

    if not worksheet_name:
        QMessageBox.warning(
            parent,
            "No Worksheet Selected",
            "Choose a worksheet first.",
        )
        return None

    if not re.fullmatch(
        r"\$?[A-Z]{1,3}\$?[1-9]\d*",
        start_cell_text,
    ):
        QMessageBox.warning(
            parent,
            "Invalid Start Cell",
            "Enter a valid cell address such as A2, E2, or J12.",
        )
        return None

    try:
        excel = _get_excel()
        workbook = _find_workbook(excel, workbook_name)

        if workbook is None:
            QMessageBox.warning(
                parent,
                "Workbook Not Found",
                "The selected workbook is no longer open. Click Refresh.",
            )
            return None

        worksheet = _find_worksheet(
            workbook,
            worksheet_name,
        )

        if worksheet is None:
            QMessageBox.warning(
                parent,
                "Worksheet Not Found",
                "The selected worksheet no longer exists. Click Refresh.",
            )
            return None

        workbook.Activate()
        worksheet.Activate()

        start_cell = worksheet.Range(start_cell_text)

        return (
            workbook,
            worksheet,
            start_cell,
            start_cell_text,
        )

    except Exception as error:
        QMessageBox.critical(
            parent,
            "Excel Error",
            "ToolPy could not access the selected Excel location.\n\n"
            f"Error:\n{error}",
        )
        return None


def _collect_downward_cells(worksheet, start_cell):
    cells = []
    row = start_cell.Row
    column = start_cell.Column

    while True:
        cell = worksheet.Cells(row, column)
        value = cell.Value

        if value is None or str(value).strip() == "":
            break

        cells.append(cell)
        row += 1

    return cells


def _apply_alignment(cell, alignment):
    if alignment == "No Change":
        return

    cell.HorizontalAlignment = ALIGNMENTS[alignment]




def _column_number_to_letters(column_number):
    letters = ""

    while column_number > 0:
        column_number, remainder = divmod(
            column_number - 1,
            26,
        )
        letters = chr(65 + remainder) + letters

    return letters


def _cell_address(row, column):
    return (
        f"{_column_number_to_letters(column)}"
        f"{row}"
    )


def activate_excel_selection_target(
    parent,
    workbook_dropdown,
    worksheet_dropdown,
):
    workbook_name = workbook_dropdown.currentData()
    worksheet_name = worksheet_dropdown.currentData()

    if not workbook_name:
        QMessageBox.warning(
            parent,
            "No Workbook Selected",
            "Choose an open Excel workbook first.",
        )
        return None

    if not worksheet_name:
        QMessageBox.warning(
            parent,
            "No Worksheet Selected",
            "Choose a worksheet first.",
        )
        return None

    try:
        excel = _get_excel()

        if excel.Workbooks.Count == 0:
            QMessageBox.warning(
                parent,
                "No Active Excel Workbook",
                "No active Excel workbook found.\n\n"
                "Please open an Excel workbook first.",
            )
            return None

        workbook = _find_workbook(
            excel,
            workbook_name,
        )

        if workbook is None:
            QMessageBox.warning(
                parent,
                "Workbook Not Found",
                "The workbook selected in ToolPy "
                "is no longer open.",
            )
            return None

        worksheet = _find_worksheet(
            workbook,
            worksheet_name,
        )

        if worksheet is None:
            QMessageBox.warning(
                parent,
                "Worksheet Not Found",
                "The worksheet selected in ToolPy "
                "could not be found.",
            )
            return None

        workbook.Activate()
        worksheet.Activate()
        excel.Visible = True

        active_cell = excel.ActiveCell

        initial_address = None

        if active_cell is not None:
            initial_address = _cell_address(
                int(active_cell.Row),
                int(active_cell.Column),
            )

        return {
            "excel": excel,
            "workbook_name": workbook_name,
            "worksheet_name": worksheet_name,
            "initial_address": initial_address,
        }

    except Exception as error:
        QMessageBox.critical(
            parent,
            "Cell Selection Error",
            "ToolPy could not start Excel cell selection.\n\n"
            f"Error:\n{error}",
        )
        return None


def read_active_excel_cell(
    selection_context,
):
    try:
        excel = selection_context["excel"]

        active_workbook = excel.ActiveWorkbook
        active_sheet = excel.ActiveSheet
        active_cell = excel.ActiveCell

        if (
            active_workbook is None
            or active_sheet is None
            or active_cell is None
        ):
            return None

        if (
            active_workbook.Name
            != selection_context["workbook_name"]
        ):
            return None

        if (
            active_sheet.Name
            != selection_context["worksheet_name"]
        ):
            return None

        return _cell_address(
            int(active_cell.Row),
            int(active_cell.Column),
        )

    except Exception:
        return None


def refresh_open_workbooks(
    workbook_dropdown,
    worksheet_dropdown,
    parent,
    show_warning=True,
):
    current_workbook = workbook_dropdown.currentData()

    workbook_dropdown.blockSignals(True)
    workbook_dropdown.clear()
    workbook_dropdown.addItem(
        "Choose an open Excel workbook",
        "",
    )

    try:
        excel = _get_excel()
        names = [
            workbook.Name
            for workbook in excel.Workbooks
        ]

        for name in names:
            workbook_dropdown.addItem(name, name)

        if current_workbook:
            index = workbook_dropdown.findData(
                current_workbook
            )
            if index >= 0:
                workbook_dropdown.setCurrentIndex(index)

        if show_warning and not names:
            QMessageBox.warning(
                parent,
                "No Excel Workbooks",
                "Open at least one workbook in Microsoft Excel.",
            )

    except Exception:
        if show_warning:
            QMessageBox.warning(
                parent,
                "Excel Not Found",
                "Open Microsoft Excel and at least one workbook.",
            )

    finally:
        workbook_dropdown.blockSignals(False)

    refresh_worksheets(
        workbook_dropdown,
        worksheet_dropdown,
        parent,
        show_warning=False,
    )


def refresh_worksheets(
    workbook_dropdown,
    worksheet_dropdown,
    parent,
    show_warning=True,
):
    current_sheet = worksheet_dropdown.currentData()
    workbook_name = workbook_dropdown.currentData()

    worksheet_dropdown.blockSignals(True)
    worksheet_dropdown.clear()
    worksheet_dropdown.addItem(
        "Choose a worksheet",
        "",
    )

    try:
        if not workbook_name:
            return

        excel = _get_excel()
        workbook = _find_workbook(
            excel,
            workbook_name,
        )

        if workbook is None:
            if show_warning:
                QMessageBox.warning(
                    parent,
                    "Workbook Not Found",
                    "The selected workbook is no longer open.",
                )
            return

        for worksheet in workbook.Worksheets:
            worksheet_dropdown.addItem(
                worksheet.Name,
                worksheet.Name,
            )

        if current_sheet:
            index = worksheet_dropdown.findData(
                current_sheet
            )
            if index >= 0:
                worksheet_dropdown.setCurrentIndex(index)

    except Exception as error:
        if show_warning:
            QMessageBox.warning(
                parent,
                "Worksheet Error",
                "ToolPy could not list the worksheets.\n\n"
                f"Error:\n{error}",
            )

    finally:
        worksheet_dropdown.blockSignals(False)


def apply_date_format(
    parent,
    workbook_dropdown,
    worksheet_dropdown,
    start_cell_text,
    date_order,
    output_format,
    alignment,
):
    target = _validate_target(
        parent,
        workbook_dropdown,
        worksheet_dropdown,
        start_cell_text,
    )

    if target is None:
        return

    _workbook, worksheet, start_cell, address = target
    cells = _collect_downward_cells(
        worksheet,
        start_cell,
    )

    if not cells:
        QMessageBox.warning(
            parent,
            "No Dates Found",
            f"{address} is empty.",
        )
        return

    try:
        detected_order = date_order

        if date_order == "Auto Detect":
            detected_order = _detect_numeric_order(cells)

        number_format = OUTPUT_FORMATS[output_format]

        for cell in cells:
            parsed, _was_real_date = _parse_value(
                cell.Value,
                detected_order,
            )

            if parsed is None:
                continue

            cell.Value = parsed
            cell.NumberFormat = number_format
            _apply_alignment(cell, alignment)

    except Exception as error:
        QMessageBox.critical(
            parent,
            "Date Format Error",
            "ToolPy could not format the dates.\n\n"
            f"Error:\n{error}",
        )


def apply_text_cleanup(
    parent,
    workbook_dropdown,
    worksheet_dropdown,
    start_cell_text,
    action,
    alignment,
):
    target = _validate_target(
        parent,
        workbook_dropdown,
        worksheet_dropdown,
        start_cell_text,
    )

    if target is None:
        return

    _workbook, worksheet, start_cell, address = target
    cells = _collect_downward_cells(
        worksheet,
        start_cell,
    )

    if not cells:
        QMessageBox.warning(
            parent,
            "No Text Found",
            f"{address} is empty.",
        )
        return

    try:
        for cell in cells:
            value = cell.Value

            if value is None:
                continue

            text = str(value)

            if action == "Trim Spaces":
                cleaned = " ".join(text.split())
            elif action == "Proper Case":
                cleaned = " ".join(text.split()).title()
            elif action == "UPPERCASE":
                cleaned = " ".join(text.split()).upper()
            elif action == "lowercase":
                cleaned = " ".join(text.split()).lower()
            elif action == "Sentence case":
                normalized = " ".join(text.split())
                cleaned = (
                    normalized[:1].upper()
                    + normalized[1:].lower()
                    if normalized
                    else normalized
                )
            else:
                cleaned = text

            cell.Value = cleaned
            _apply_alignment(cell, alignment)

    except Exception as error:
        QMessageBox.critical(
            parent,
            "Text Cleanup Error",
            "ToolPy could not clean the text.\n\n"
            f"Error:\n{error}",
        )



def apply_quick_calculate(
    parent,
    workbook_dropdown,
    worksheet_dropdown,
    mode,
    direction,
    start_cell_text,
    end_cell_text,
    operation,
    alignment,
):
    target = _validate_target(
        parent,
        workbook_dropdown,
        worksheet_dropdown,
        start_cell_text,
    )

    if target is None:
        return

    _workbook, worksheet, start_cell, start_address = target

    try:
        if mode == "Auto":
            cells, result_cell = _collect_quick_calculate_auto(
                worksheet,
                start_cell,
                direction,
            )

            if not cells:
                QMessageBox.warning(
                    parent,
                    "No Values Found",
                    f"{start_address} is empty.",
                )
                return

        else:
            end_cell_text = end_cell_text.strip().upper()

            if not re.fullmatch(
                r"\$?[A-Z]{1,3}\$?[1-9]\d*",
                end_cell_text,
            ):
                QMessageBox.warning(
                    parent,
                    "Invalid End Cell",
                    "Enter a valid end cell such as A5 or E2.",
                )
                return

            end_cell = worksheet.Range(end_cell_text)

            cells, result_cell = _collect_quick_calculate_range(
                parent,
                worksheet,
                start_cell,
                end_cell,
                direction,
            )

            if cells is None:
                return

        numeric_values = []

        for cell in cells:
            numeric_value = _to_number(cell.Value)

            if numeric_value is not None:
                numeric_values.append(numeric_value)

        if operation == "Count":
            result = len(numeric_values)

        else:
            if not numeric_values:
                QMessageBox.warning(
                    parent,
                    "No Numbers Found",
                    "The selected cells do not contain numeric values.",
                )
                return

            if operation == "Sum":
                result = sum(numeric_values)

            elif operation == "Average":
                result = sum(numeric_values) / len(numeric_values)

            else:
                QMessageBox.warning(
                    parent,
                    "Invalid Operation",
                    "Choose Sum, Average, or Count.",
                )
                return

        result_cell.Value = result
        _apply_alignment(result_cell, alignment)

    except Exception as error:
        QMessageBox.critical(
            parent,
            "Quick Calculate Error",
            "ToolPy could not calculate the selected cells.\n\n"
            f"Error:\n{error}",
        )


def _collect_quick_calculate_auto(
    worksheet,
    start_cell,
    direction,
):
    cells = []

    row = start_cell.Row
    column = start_cell.Column

    while True:
        cell = worksheet.Cells(row, column)
        value = cell.Value

        if value is None or str(value).strip() == "":
            return cells, cell

        cells.append(cell)

        if direction == "Right":
            column += 1
        else:
            row += 1


def _collect_quick_calculate_range(
    parent,
    worksheet,
    start_cell,
    end_cell,
    direction,
):
    start_row = start_cell.Row
    start_column = start_cell.Column
    end_row = end_cell.Row
    end_column = end_cell.Column

    cells = []

    if direction == "Down":
        if start_column != end_column:
            QMessageBox.warning(
                parent,
                "Invalid Down Range",
                "For Down direction, the start and end cells "
                "must be in the same column.",
            )
            return None, None

        if end_row < start_row:
            QMessageBox.warning(
                parent,
                "Invalid Down Range",
                "The end cell must be below the start cell.",
            )
            return None, None

        for row in range(start_row, end_row + 1):
            cells.append(
                worksheet.Cells(row, start_column)
            )

        result_cell = worksheet.Cells(
            end_row + 1,
            start_column,
        )

    else:
        if start_row != end_row:
            QMessageBox.warning(
                parent,
                "Invalid Right Range",
                "For Right direction, the start and end cells "
                "must be in the same row.",
            )
            return None, None

        if end_column < start_column:
            QMessageBox.warning(
                parent,
                "Invalid Right Range",
                "The end cell must be to the right of the start cell.",
            )
            return None, None

        for column in range(
            start_column,
            end_column + 1,
        ):
            cells.append(
                worksheet.Cells(start_row, column)
            )

        result_cell = worksheet.Cells(
            start_row,
            end_column + 1,
        )

    return cells, result_cell


def _to_number(value):
    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(",", "")

    if not text:
        return None

    try:
        return float(text)
    except ValueError:
        return None



def _detect_numeric_order(cells):
    mm_dd_evidence = 0
    dd_mm_evidence = 0

    for cell in cells:
        value = cell.Value

        if _is_real_excel_date(value):
            continue

        parts = _numeric_parts(str(value))

        if parts is None:
            continue

        first, second, _year = parts

        if first > 12 and second <= 12:
            dd_mm_evidence += 1
        elif second > 12 and first <= 12:
            mm_dd_evidence += 1

    if dd_mm_evidence > mm_dd_evidence:
        return "DD/MM/YYYY"

    return "MM/DD/YYYY"


def _is_real_excel_date(value):
    return isinstance(
        value,
        (
            datetime,
            date,
            pywintypes.TimeType,
        ),
    )


def _parse_value(value, numeric_order):
    if isinstance(value, pywintypes.TimeType):
        return datetime(
            value.year,
            value.month,
            value.day,
            value.hour,
            value.minute,
            value.second,
        ), True

    if isinstance(value, datetime):
        return value, True

    if isinstance(value, date):
        return datetime(
            value.year,
            value.month,
            value.day,
        ), True

    if isinstance(value, (int, float)) and 1 <= float(value) <= 2958465:
        excel_epoch = datetime(1899, 12, 30)

        try:
            return excel_epoch + timedelta(
                days=float(value)
            ), True
        except (OverflowError, ValueError):
            return None, False

    text = str(value).strip()

    if not text:
        return None, False

    month_name_date = _parse_month_name(text)

    if month_name_date is not None:
        return month_name_date, False

    iso_date = _parse_iso_date(text)

    if iso_date is not None:
        return iso_date, False

    parts = _numeric_parts(text)

    if parts is None:
        return None, False

    first, second, year = parts

    if year < 100:
        year += 2000 if year < 70 else 1900

    try:
        if first > 12 and second <= 12:
            day = first
            month = second
        elif second > 12 and first <= 12:
            month = first
            day = second
        elif numeric_order == "DD/MM/YYYY":
            day = first
            month = second
        else:
            month = first
            day = second

        return datetime(year, month, day), False

    except ValueError:
        return None, False


def _parse_month_name(text):
    normalized = re.sub(
        r"\s+",
        " ",
        text.replace(".", "").strip(),
    )

    for pattern in MONTH_FORMATS:
        try:
            return datetime.strptime(
                normalized,
                pattern,
            )
        except ValueError:
            continue

    return None


def _parse_iso_date(text):
    try:
        return datetime.strptime(
            text,
            "%Y-%m-%d",
        )
    except ValueError:
        return None


def _numeric_parts(text):
    match = re.fullmatch(
        r"\s*(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\s*",
        text,
    )

    if match is None:
        return None

    return tuple(
        int(part)
        for part in match.groups()
    )
