#!/usr/bin/env python3
"""
Lập bảng độ nhạy Project NPV theo công suất lắp đặt và hệ số công suất.

Mặc định, chương trình mở PTKTTC_Ver1.xlsx nằm cùng thư mục với file Python.
Có thể truyền đường dẫn workbook làm tham số:

    python NPV-P-CF.py "D:\\DuAn\\PTKTTC_Ver1.xlsx"

Yêu cầu:
    - Windows có cài Microsoft Excel.
    - Python đã cài pywin32:
          py -m pip install pywin32

Kết quả:
    - In bảng Project NPV (tỷ VND) trên terminal.
    - Tạo một file Excel mới chứa bảng kết quả trong cùng thư mục với workbook
      nguồn. Tên mặc định là NPV-P-CF.xlsx; nếu đã tồn tại, chương trình tự
      thêm hậu tố _1, _2, ... để không ghi đè.
    - Không lưu thay đổi vào PTKTTC_Ver1.xlsx.
"""

from __future__ import annotations

import numbers
import sys
import time
from pathlib import Path
from typing import Sequence


WORKBOOK_NAME = "PTKTTC_Ver1.xlsx"
OUTPUT_WORKBOOK_NAME = "NPV-P-CF.xlsx"

INPUT_SHEET = "Input"
CAPACITY_CELL = "D12"
CF_CELL = "D44"

OUTPUT_SHEET = "Summary"
PROJECT_NPV_CELL = "E11"

CAPACITIES_MW = tuple(range(60, 181, 20))
CAPACITY_FACTORS_PERCENT = tuple(range(25, 37))

# Excel constants, khai báo trực tiếp để không phụ thuộc
# win32com.client.constants.
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_DONE = 0
XL_OPEN_XML_WORKBOOK = 51
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


def resolve_output_path(workbook_path: Path) -> Path:
    """Tạo đường dẫn output mới, không ghi đè file kết quả đã có."""
    output_path = workbook_path.with_name(OUTPUT_WORKBOOK_NAME)
    if not output_path.exists():
        return output_path

    stem = output_path.stem
    suffix = output_path.suffix
    index = 1
    while True:
        candidate = output_path.with_name(f"{stem}_{index}{suffix}")
        if not candidate.exists():
            return candidate
        index += 1


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


def save_results_workbook(excel, results, output_path: Path) -> None:
    """Tạo workbook Excel mới và ghi bảng Project NPV vào đó."""
    result_workbook = None
    try:
        result_workbook = excel.Workbooks.Add()
        result_sheet = result_workbook.Worksheets(1)
        result_sheet.Name = "Project NPV"

        # Chỉ giữ lại một worksheet trong file kết quả.
        while result_workbook.Worksheets.Count > 1:
            result_workbook.Worksheets(
                result_workbook.Worksheets.Count
            ).Delete()

        last_column = len(CAPACITIES_MW) + 1
        last_row = len(CAPACITY_FACTORS_PERCENT) + 3

        title_range = result_sheet.Range(
            result_sheet.Cells(1, 1),
            result_sheet.Cells(1, last_column),
        )
        title_range.Merge()
        title_range.Value2 = "PROJECT NPV THEO CÔNG SUẤT VÀ HỆ SỐ CÔNG SUẤT"

        unit_range = result_sheet.Range(
            result_sheet.Cells(2, 1),
            result_sheet.Cells(2, last_column),
        )
        unit_range.Merge()
        unit_range.Value2 = "Đơn vị: tỷ VND"

        headers = ("CF \\ MW", *CAPACITIES_MW)
        table_rows = [
            (cf_percent / 100.0, *row)
            for cf_percent, row in zip(CAPACITY_FACTORS_PERCENT, results)
        ]

        header_range = result_sheet.Range(
            result_sheet.Cells(3, 1),
            result_sheet.Cells(3, last_column),
        )
        header_range.Value2 = (headers,)

        data_range = result_sheet.Range(
            result_sheet.Cells(4, 1),
            result_sheet.Cells(last_row, last_column),
        )
        data_range.Value2 = tuple(table_rows)

        # Định dạng tối giản, rõ ràng để có thể sử dụng ngay.
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
            result_sheet.Cells(4, 1),
            result_sheet.Cells(last_row, 1),
        )
        row_header_range.Interior.Color = light_blue
        row_header_range.Font.Bold = True
        row_header_range.HorizontalAlignment = XL_CENTER
        row_header_range.NumberFormat = "0%"

        capacity_header_range = result_sheet.Range(
            result_sheet.Cells(3, 2),
            result_sheet.Cells(3, last_column),
        )
        capacity_header_range.NumberFormat = '0 "MW"'

        npv_range = result_sheet.Range(
            result_sheet.Cells(4, 2),
            result_sheet.Cells(last_row, last_column),
        )
        npv_range.NumberFormat = '#,##0.00;[Red](#,##0.00);-'

        table_range = result_sheet.Range(
            result_sheet.Cells(3, 1),
            result_sheet.Cells(last_row, last_column),
        )
        for border_index in range(7, 13):
            border = table_range.Borders(border_index)
            border.LineStyle = XL_CONTINUOUS
            border.Weight = XL_THIN
            border.Color = light_border

        result_sheet.Columns(1).ColumnWidth = 12
        result_sheet.Range(
            result_sheet.Columns(2),
            result_sheet.Columns(last_column),
        ).ColumnWidth = 14
        result_sheet.Range("A1").Select()
        excel.ActiveWindow.DisplayGridlines = False

        result_workbook.SaveAs(
            str(output_path),
            FileFormat=XL_OPEN_XML_WORKBOOK,
        )
        result_workbook.Close(SaveChanges=False)
        result_workbook = None

    finally:
        if result_workbook is not None:
            result_workbook.Close(SaveChanges=False)


def calculate_sensitivity(workbook_path: Path):
    """Tính Project NPV và lưu toàn bộ kết quả bằng Microsoft Excel."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "Chưa có thư viện pywin32. Hãy cài bằng lệnh:\n"
            "    py -m pip install pywin32"
        ) from exc

    output_path = resolve_output_path(workbook_path)
    pythoncom.CoInitialize()
    excel = None
    workbook = None

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

        capacity_cell = input_sheet.Range(CAPACITY_CELL)
        cf_cell = input_sheet.Range(CF_CELL)
        project_npv_cell = output_sheet.Range(PROJECT_NPV_CELL)

        original_capacity = capacity_cell.Value2
        original_cf = cf_cell.Value2
        results: list[list[float]] = []

        total_cases = len(CAPACITIES_MW) * len(CAPACITY_FACTORS_PERCENT)
        completed_cases = 0

        print(f'Workbook: "{workbook_path}"')
        print(
            f"Đang tính {total_cases} phương án "
            f"({len(CAPACITY_FACTORS_PERCENT)} mức CF × "
            f"{len(CAPACITIES_MW)} mức công suất)..."
        )

        for cf_percent in CAPACITY_FACTORS_PERCENT:
            row: list[float] = []
            cf_cell.Value2 = cf_percent / 100.0

            for capacity_mw in CAPACITIES_MW:
                capacity_cell.Value2 = capacity_mw

                # FullRebuild buộc Excel dựng lại toàn bộ chuỗi phụ thuộc,
                # tránh đọc nhầm giá trị NPV lưu trong bộ nhớ đệm của workbook.
                excel.CalculateFullRebuild()
                wait_for_excel(excel)

                npv_value = project_npv_cell.Value2
                if (
                    isinstance(npv_value, bool)
                    or not isinstance(npv_value, numbers.Real)
                ):
                    raise ValueError(
                        f"{OUTPUT_SHEET}!{PROJECT_NPV_CELL} không trả về số "
                        f"tại MW={capacity_mw}, CF={cf_percent}%: "
                        f"{npv_value!r}"
                    )

                row.append(float(npv_value))
                completed_cases += 1

            results.append(row)
            print(
                f"  Đã xong CF {cf_percent}% "
                f"({completed_cases}/{total_cases} phương án)"
            )

        # Hoàn nguyên giá trị trong bộ nhớ trước khi đóng. Workbook vẫn được
        # đóng với SaveChanges=False nên file gốc không bị sửa.
        capacity_cell.Value2 = original_capacity
        cf_cell.Value2 = original_cf

        save_results_workbook(excel, results, output_path)
        return results, output_path

    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def format_terminal_table(results: Sequence[Sequence[float]]) -> str:
    """Tạo bảng dễ đọc để in trên terminal."""
    headers = ["CF \\ MW", *(f"{mw} MW" for mw in CAPACITIES_MW)]
    body = [
        [f"{cf_percent}%", *(f"{value:,.2f}" for value in row)]
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
            "PROJECT NPV (TỶ VND)",
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
        results, output_path = calculate_sensitivity(workbook_path)

        print()
        print(format_terminal_table(results))
        print()
        print(f'Đã tạo file kết quả: "{output_path}"')
        print(f'File nguồn "{WORKBOOK_NAME}" không bị thay đổi.')
        return 0

    except Exception as exc:
        print(f"LỖI: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
