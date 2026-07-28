#!/usr/bin/env python3
"""
Lập các bảng độ nhạy theo công suất lắp đặt và hệ số công suất:
Project NPV, Equity NPV, Project IRR và Equity IRR.

Mặc định, chương trình mở PTKTTC_Ver1.xlsx nằm cùng thư mục với file Python.
Có thể truyền đường dẫn workbook làm tham số:

    python NPV-P-CF.py "D:\\DuAn\\PTKTTC_Ver1.xlsx"

Yêu cầu:
    - Windows có cài Microsoft Excel.
    - Python đã cài pywin32:
          py -m pip install pywin32

Kết quả:
    - In bốn bảng kết quả trên terminal.
    - Tạo một workbook Excel mới, đưa bốn bảng vào cùng một sheet.
    - Chỉ tính lại mô hình một lần cho mỗi phương án rồi đọc cả bốn chỉ tiêu.
    - Không tự động lưu workbook kết quả; người dùng tự chọn Save/Save As.
    - Không lưu thay đổi vào PTKTTC_Ver1.xlsx.
"""

from __future__ import annotations

import numbers
import sys
import time
from pathlib import Path
from typing import Sequence


WORKBOOK_NAME = "PTKTTC_Ver1.xlsx"

INPUT_SHEET = "Input"
CAPACITY_CELL = "D12"
CF_CELL = "D41"

OUTPUT_SHEET = "Summary"

RESULT_METRICS = (
    {
        "key": "project_npv",
        "title": "PROJECT NPV",
        "cell": "E11",
        "unit": "tỷ VND",
        "number_format": '#,##0.00;[Red](#,##0.00);-',
        "terminal_format": ",.2f",
    },
    {
        "key": "equity_npv",
        "title": "EQUITY NPV",
        "cell": "E12",
        "unit": "tỷ VND",
        "number_format": '#,##0.00;[Red](#,##0.00);-',
        "terminal_format": ",.2f",
    },
    {
        "key": "project_irr",
        "title": "PROJECT IRR",
        "cell": "B11",
        "unit": "%",
        "number_format": '0.00%;[Red](0.00%);-',
        "terminal_format": ".2%",
    },
    {
        "key": "equity_irr",
        "title": "EQUITY IRR",
        "cell": "B12",
        "unit": "%",
        "number_format": '0.00%;[Red](0.00%);-',
        "terminal_format": ".2%",
    },
)

# Kiểm tra nhãn để tránh đọc/ghi nhầm ô khi cấu trúc mô hình thay đổi.
MODEL_CELL_LABELS = (
    (INPUT_SHEET, "B12", "Công suất lắp đặt"),
    (INPUT_SHEET, "B41", "Hệ số công suất ròng"),
    (OUTPUT_SHEET, "D11", "Project NPV"),
    (OUTPUT_SHEET, "D12", "Equity NPV"),
    (OUTPUT_SHEET, "A11", "Project IRR"),
    (OUTPUT_SHEET, "A12", "Equity IRR"),
)

CAPACITIES_MW = tuple(range(60, 181, 20))
CAPACITY_FACTORS_PERCENT = tuple(range(25, 37))
TABLE_GAP_ROWS = 2

# Excel constants, khai báo trực tiếp để không phụ thuộc
# win32com.client.constants.
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
XL_CALCULATION_DONE = 0
XL_CENTER = -4108
XL_THIN = 2
XL_CONTINUOUS = 1


def resolve_workbook_path(arguments: Sequence[str]) -> Path:
    """Lấy đường dẫn workbook từ tham số hoặc từ thư mục chứa script."""
    if len(arguments) > 1:
        raise ValueError(
            'Chỉ truyền tối đa một tham số: đường dẫn đến "PTKTTC_Ver1.xlsx".'
        )

    if arguments:
        workbook_path = Path(arguments[0]).expanduser()
    else:
        workbook_path = Path(__file__).resolve().with_name(WORKBOOK_NAME)

    workbook_path = workbook_path.resolve()
    if not workbook_path.is_file():
        raise FileNotFoundError(
            f'Không tìm thấy workbook: "{workbook_path}".\n'
            f'Hãy đặt "{WORKBOOK_NAME}" cùng thư mục với script hoặc truyền '
            "đường dẫn workbook khi chạy."
        )
    return workbook_path


def wait_for_excel(excel, timeout_seconds: float = 300.0) -> None:
    """Chờ Excel tính xong hoặc báo lỗi nếu vượt quá thời gian cho phép."""
    deadline = time.monotonic() + timeout_seconds
    while excel.CalculationState != XL_CALCULATION_DONE:
        if time.monotonic() >= deadline:
            raise TimeoutError(
                f"Excel chưa tính xong sau {timeout_seconds:.0f} giây."
            )
        time.sleep(0.05)


def excel_color(red: int, green: int, blue: int) -> int:
    """Đổi RGB sang giá trị màu mà Excel COM sử dụng."""
    return red + green * 256 + blue * 65536


def validate_model_layout(workbook) -> None:
    """Xác nhận các ô quan trọng vẫn đúng với cấu trúc mô hình hiện tại."""
    mismatches = []
    for sheet_name, label_cell, expected_label in MODEL_CELL_LABELS:
        actual_label = workbook.Worksheets(sheet_name).Range(label_cell).Value2
        if str(actual_label).strip() != expected_label:
            mismatches.append(
                f'{sheet_name}!{label_cell}: cần "{expected_label}", '
                f"nhưng đang là {actual_label!r}"
            )

    if mismatches:
        raise ValueError(
            "Cấu trúc workbook không khớp với phiên bản script:\n  - "
            + "\n  - ".join(mismatches)
        )


def read_numeric_result(cell, metric, capacity_mw: int, cf_percent: int) -> float:
    """Đọc và kiểm tra một chỉ tiêu số sau khi Excel tính lại."""
    value = cell.Value2
    if isinstance(value, bool) or not isinstance(value, numbers.Real):
        raise ValueError(
            f"{OUTPUT_SHEET}!{metric['cell']} ({metric['title']}) "
            f"không trả về số tại MW={capacity_mw}, CF={cf_percent}%: "
            f"{value!r}"
        )
    return float(value)


def write_metric_table(result_sheet, metric, results, start_row: int) -> int:
    """Ghi và định dạng một ma trận chỉ tiêu; trả về hàng cuối của ma trận."""
    last_column = len(CAPACITIES_MW) + 1
    header_row = start_row + 2
    data_start_row = start_row + 3
    last_row = data_start_row + len(CAPACITY_FACTORS_PERCENT) - 1

    title_range = result_sheet.Range(
        result_sheet.Cells(start_row, 1),
        result_sheet.Cells(start_row, last_column),
    )
    title_range.Merge()
    title_range.Value2 = (
        f"{metric['title']} THEO CÔNG SUẤT VÀ HỆ SỐ CÔNG SUẤT"
    )

    unit_range = result_sheet.Range(
        result_sheet.Cells(start_row + 1, 1),
        result_sheet.Cells(start_row + 1, last_column),
    )
    unit_range.Merge()
    unit_range.Value2 = f"Đơn vị: {metric['unit']}"

    headers = ("CF \\ MW", *CAPACITIES_MW)
    table_rows = [
        (cf_percent / 100.0, *row)
        for cf_percent, row in zip(CAPACITY_FACTORS_PERCENT, results)
    ]

    header_range = result_sheet.Range(
        result_sheet.Cells(header_row, 1),
        result_sheet.Cells(header_row, last_column),
    )
    header_range.Value2 = (headers,)

    data_range = result_sheet.Range(
        result_sheet.Cells(data_start_row, 1),
        result_sheet.Cells(last_row, last_column),
    )
    data_range.Value2 = tuple(table_rows)

    dark_blue = excel_color(31, 78, 121)
    light_blue = excel_color(221, 235, 247)
    light_border = excel_color(191, 191, 191)

    title_range.Interior.Color = dark_blue
    title_range.Font.Color = excel_color(255, 255, 255)
    title_range.Font.Bold = True
    title_range.Font.Size = 14
    title_range.HorizontalAlignment = XL_CENTER
    title_range.RowHeight = 24

    unit_range.Font.Italic = True
    unit_range.HorizontalAlignment = XL_CENTER

    header_range.Interior.Color = light_blue
    header_range.Font.Bold = True
    header_range.HorizontalAlignment = XL_CENTER

    row_header_range = result_sheet.Range(
        result_sheet.Cells(data_start_row, 1),
        result_sheet.Cells(last_row, 1),
    )
    row_header_range.Interior.Color = light_blue
    row_header_range.Font.Bold = True
    row_header_range.HorizontalAlignment = XL_CENTER
    row_header_range.NumberFormat = "0%"

    capacity_header_range = result_sheet.Range(
        result_sheet.Cells(header_row, 2),
        result_sheet.Cells(header_row, last_column),
    )
    capacity_header_range.NumberFormat = '0 "MW"'

    result_value_range = result_sheet.Range(
        result_sheet.Cells(data_start_row, 2),
        result_sheet.Cells(last_row, last_column),
    )
    result_value_range.NumberFormat = metric["number_format"]

    table_range = result_sheet.Range(
        result_sheet.Cells(header_row, 1),
        result_sheet.Cells(last_row, last_column),
    )
    for border_index in range(7, 13):
        border = table_range.Borders(border_index)
        border.LineStyle = XL_CONTINUOUS
        border.Weight = XL_THIN
        border.Color = light_border

    return last_row


def create_results_workbook(excel, results):
    """Tạo workbook mới chứa bốn ma trận trên cùng một sheet, chưa lưu."""
    result_workbook = excel.Workbooks.Add()
    result_sheet = result_workbook.Worksheets(1)
    result_sheet.Name = "Kết quả"

    # Chỉ giữ lại một worksheet trong workbook kết quả.
    while result_workbook.Worksheets.Count > 1:
        result_workbook.Worksheets(
            result_workbook.Worksheets.Count
        ).Delete()

    start_row = 1
    for metric in RESULT_METRICS:
        last_row = write_metric_table(
            result_sheet,
            metric,
            results[metric["key"]],
            start_row,
        )
        start_row = last_row + TABLE_GAP_ROWS + 1

    result_sheet.Columns(1).ColumnWidth = 12
    result_sheet.Range(
        result_sheet.Columns(2),
        result_sheet.Columns(len(CAPACITIES_MW) + 1),
    ).ColumnWidth = 14
    result_sheet.Activate()
    result_sheet.Range("A1").Select()
    excel.ActiveWindow.DisplayGridlines = False
    return result_workbook


def calculate_sensitivity(workbook_path: Path):
    """Tính bốn chỉ tiêu và mở workbook kết quả chưa lưu trong Excel."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "Chưa có thư viện pywin32. Hãy cài bằng lệnh:\n"
            "    py -m pip install pywin32"
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    result_workbook = None
    keep_excel_open = False

    try:
        # DispatchEx tạo một phiên Excel riêng, không can thiệp workbook người
        # dùng đang mở trong các cửa sổ Excel khác.
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False

        workbook = excel.Workbooks.Open(
            str(workbook_path),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
        )

        # Một số phiên bản Excel chỉ cho đổi chế độ tính toán sau khi đã mở
        # workbook. Nếu Excel vẫn từ chối, CalculateFullRebuild() bên dưới
        # vẫn buộc mô hình tính lại đầy đủ cho từng phương án.
        try:
            excel.Calculation = XL_CALCULATION_MANUAL
        except Exception:
            pass

        input_sheet = workbook.Worksheets(INPUT_SHEET)
        output_sheet = workbook.Worksheets(OUTPUT_SHEET)
        validate_model_layout(workbook)

        capacity_cell = input_sheet.Range(CAPACITY_CELL)
        cf_cell = input_sheet.Range(CF_CELL)
        output_cells = {
            metric["key"]: output_sheet.Range(metric["cell"])
            for metric in RESULT_METRICS
        }

        original_capacity = capacity_cell.Value2
        original_cf = cf_cell.Value2
        results: dict[str, list[list[float]]] = {
            metric["key"]: [] for metric in RESULT_METRICS
        }

        total_cases = len(CAPACITIES_MW) * len(CAPACITY_FACTORS_PERCENT)
        completed_cases = 0

        print(f'Workbook: "{workbook_path}"')
        print(
            f"Đang tính {total_cases} phương án "
            f"({len(CAPACITY_FACTORS_PERCENT)} mức CF × "
            f"{len(CAPACITIES_MW)} mức công suất)..."
        )

        for cf_percent in CAPACITY_FACTORS_PERCENT:
            metric_rows: dict[str, list[float]] = {
                metric["key"]: [] for metric in RESULT_METRICS
            }
            cf_cell.Value2 = cf_percent / 100.0

            for capacity_mw in CAPACITIES_MW:
                capacity_cell.Value2 = capacity_mw

                # FullRebuild buộc Excel dựng lại toàn bộ chuỗi phụ thuộc,
                # tránh đọc nhầm giá trị NPV lưu trong bộ nhớ đệm của workbook.
                excel.CalculateFullRebuild()
                wait_for_excel(excel)

                for metric in RESULT_METRICS:
                    metric_rows[metric["key"]].append(
                        read_numeric_result(
                            output_cells[metric["key"]],
                            metric,
                            capacity_mw,
                            cf_percent,
                        )
                    )

                completed_cases += 1

            for metric in RESULT_METRICS:
                results[metric["key"]].append(metric_rows[metric["key"]])
            print(
                f"  Đã xong CF {cf_percent}% "
                f"({completed_cases}/{total_cases} phương án)"
            )

        # Hoàn nguyên giá trị trong bộ nhớ trước khi đóng. Workbook vẫn được
        # đóng với SaveChanges=False nên file gốc không bị sửa.
        capacity_cell.Value2 = original_capacity
        cf_cell.Value2 = original_cf

        result_workbook = create_results_workbook(excel, results)

        # Đóng file nguồn mà không lưu, chỉ giữ workbook kết quả mới trong
        # Excel. Workbook kết quả chưa có đường dẫn cho đến khi người dùng tự
        # chọn Save hoặc Save As.
        workbook.Close(SaveChanges=False)
        workbook = None

        try:
            excel.Calculation = XL_CALCULATION_AUTOMATIC
        except Exception:
            pass

        excel.DisplayAlerts = True
        excel.ScreenUpdating = True
        excel.EnableEvents = True
        excel.Visible = True
        result_workbook.Activate()
        keep_excel_open = True
        return results

    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            if keep_excel_open:
                # Không gọi Quit(): chuyển workbook kết quả chưa lưu cho người
                # dùng tiếp tục thao tác trong Excel.
                excel.Visible = True
            else:
                # Nếu có lỗi trước khi hoàn tất, đóng phiên Excel ẩn để không
                # để lại tiến trình nền.
                excel.Quit()
        pythoncom.CoUninitialize()


def format_terminal_table(metric, results: Sequence[Sequence[float]]) -> str:
    """Tạo một bảng chỉ tiêu dễ đọc để in trên terminal."""
    headers = ["CF \\ MW", *(f"{mw} MW" for mw in CAPACITIES_MW)]
    body = [
        [
            f"{cf_percent}%",
            *(format(value, metric["terminal_format"]) for value in row),
        ]
        for cf_percent, row in zip(CAPACITY_FACTORS_PERCENT, results)
    ]

    widths = [
        max(len(headers[col]), *(len(row[col]) for row in body))
        for col in range(len(headers))
    ]

    def format_row(row: Sequence[str]) -> str:
        return " | ".join(
            value.ljust(widths[col]) if col == 0 else value.rjust(widths[col])
            for col, value in enumerate(row)
        )

    separator = "-+-".join("-" * width for width in widths)
    return "\n".join(
        [
            f"{metric['title']} ({metric['unit'].upper()})",
            format_row(headers),
            separator,
            *(format_row(row) for row in body),
        ]
    )


def main() -> int:
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8")

        workbook_path = resolve_workbook_path(sys.argv[1:])
        results = calculate_sensitivity(workbook_path)

        print()
        print(
            "\n\n".join(
                format_terminal_table(metric, results[metric["key"]])
                for metric in RESULT_METRICS
            )
        )
        print()
        print(
            "Đã tạo workbook kết quả mới với 4 ma trận trên cùng một sheet "
            "và mở trong Excel."
        )
        print("Workbook kết quả chưa được lưu; hãy dùng Save hoặc Save As.")
        print(f'File nguồn "{WORKBOOK_NAME}" không bị thay đổi.')
        return 0

    except Exception as exc:
        print(f"LỖI: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
