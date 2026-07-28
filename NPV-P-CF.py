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
    - Copy cùng bảng ở dạng TSV vào clipboard để dán trực tiếp vào Excel.
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
CF_CELL = "D44"

OUTPUT_SHEET = "Summary"
PROJECT_NPV_CELL = "E11"

CAPACITIES_MW = tuple(range(60, 181, 20))
CAPACITY_FACTORS_PERCENT = tuple(range(25, 37))

# Excel constants, khai báo trực tiếp để không phụ thuộc win32com.client.constants.
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_DONE = 0
XL_DECIMAL_SEPARATOR = 3


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


def calculate_sensitivity(workbook_path: Path):
    """Tính Project NPV cho toàn bộ tổ hợp MW và CF bằng Microsoft Excel."""
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

    try:
        # DispatchEx tạo một phiên Excel riêng, không can thiệp workbook người dùng
        # đang mở trong các cửa sổ Excel khác.
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

        decimal_separator = str(excel.International(XL_DECIMAL_SEPARATOR))
        return results, decimal_separator

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


def format_clipboard_table(
    results: Sequence[Sequence[float]], decimal_separator: str
) -> str:
    """Tạo TSV dùng dấu thập phân theo thiết lập Excel trên máy."""

    def format_number(value: float) -> str:
        text = f"{value:.2f}"
        if decimal_separator != ".":
            text = text.replace(".", decimal_separator)
        return text

    rows = [["CF \\ MW", *(str(mw) for mw in CAPACITIES_MW)]]
    rows.extend(
        [
            [f"{cf_percent}%", *(format_number(value) for value in row)]
            for cf_percent, row in zip(CAPACITY_FACTORS_PERCENT, results)
        ]
    )
    return "\r\n".join("\t".join(row) for row in rows)


def copy_to_clipboard(text: str) -> None:
    """Copy Unicode text vào clipboard bằng thư viện chuẩn tkinter."""
    try:
        import tkinter

        root = tkinter.Tk()
        root.withdraw()
        root.clipboard_clear()
        root.clipboard_append(text)
        # update() chuyển dữ liệu sang Windows clipboard trước khi đóng cửa sổ.
        root.update()
        root.destroy()
    except Exception as exc:
        raise RuntimeError(
            "Đã tính xong nhưng không copy được bảng vào clipboard."
        ) from exc


def main() -> int:
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8")

        workbook_path = resolve_workbook_path(sys.argv[1:])
        results, decimal_separator = calculate_sensitivity(workbook_path)

        terminal_table = format_terminal_table(results)
        clipboard_table = format_clipboard_table(results, decimal_separator)
        copy_to_clipboard(clipboard_table)

        print()
        print(terminal_table)
        print()
        print("Đã copy bảng vào clipboard. Có thể dán trực tiếp vào Excel.")
        print(f'File "{WORKBOOK_NAME}" không bị thay đổi.')
        return 0

    except Exception as exc:
        print(f"LỖI: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
