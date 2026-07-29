#!/usr/bin/env python3
"""
Lập các bảng độ nhạy theo công suất lắp đặt và hệ số công suất:
Project NPV, Equity NPV, Project IRR và Equity IRR.

Trước khi chạy, chương trình hiển thị giao diện chọn Min, Max và Bước cho
công suất (MW) và CF (%). Các giá trị được lưu trong NPV-P-CF.settings.json
cạnh script và tự động nạp lại ở lần chạy sau.

Mặc định, chương trình mở PTKTTC_Ver1.xlsx nằm cùng thư mục với file Python.
Có thể truyền đường dẫn workbook làm tham số:

    python NPV-P-CF.py "D:\\DuAn\\PTKTTC_Ver1.xlsx"

Yêu cầu:
    - Windows có cài Microsoft Excel.
    - Python đã cài pywin32:
          py -m pip install pywin32

Kết quả:
    - In bốn bảng kết quả trên terminal.
    - Tạo một workbook Excel mới với bốn bảng trên sheet "Kết quả".
    - Tạo sheet "Input" mới và sao chép giá trị, định dạng, kích thước
      hàng/cột từ mô hình; không sao chép công thức hoặc liên kết.
    - Chỉ tính lại mô hình một lần cho mỗi phương án rồi đọc cả bốn chỉ tiêu.
    - Không tự động lưu workbook kết quả; người dùng tự chọn Save/Save As.
    - Không lưu thay đổi vào PTKTTC_Ver1.xlsx.
"""

from __future__ import annotations

import json
import numbers
import sys
import time
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Sequence


WORKBOOK_NAME = "PTKTTC_Ver1.xlsx"

INPUT_SHEET = "Input"
CAPACITY_CELL = "D12"
CF_LABEL = "Hệ số công suất ròng"
INPUT_VALUE_COLUMN = 4  # Cột D chứa giá trị input cùng hàng với nhãn.

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
    (OUTPUT_SHEET, "D11", "Project NPV"),
    (OUTPUT_SHEET, "D12", "Equity NPV"),
    (OUTPUT_SHEET, "A11", "Project IRR"),
    (OUTPUT_SHEET, "A12", "Equity IRR"),
)

SETTINGS_NAME = "NPV-P-CF.settings.json"
DEFAULT_SETTINGS = {
    "capacity_min": "60",
    "capacity_max": "180",
    "capacity_step": "20",
    "cf_min": "25",
    "cf_max": "36",
    "cf_step": "1",
}

# Hai dãy này được cập nhật từ hộp thoại trước khi chạy. Giữ giá trị mặc
# định để các hàm vẫn có thể được nhập và kiểm thử độc lập.
CAPACITIES_MW = tuple(range(60, 181, 20))
CAPACITY_FACTORS_PERCENT = tuple(range(25, 37))
TABLE_GAP_ROWS = 2
MAX_MATRIX_CASES = 5000

# Excel constants, khai báo trực tiếp để không phụ thuộc
# win32com.client.constants.
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
XL_CALCULATION_DONE = 0
XL_CENTER = -4108
XL_THIN = 2
XL_CONTINUOUS = 1
XL_PASTE_FORMATS = -4122


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


def settings_path() -> Path:
    """Trả về đường dẫn file cấu hình nằm cạnh script."""
    return Path(__file__).resolve().with_name(SETTINGS_NAME)


def load_settings() -> dict[str, str]:
    """Đọc cấu hình lần chạy trước; dùng mặc định nếu file chưa có hoặc lỗi."""
    settings = DEFAULT_SETTINGS.copy()
    path = settings_path()
    if not path.is_file():
        return settings
    try:
        saved = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(saved, dict):
            for key in settings:
                if key in saved:
                    settings[key] = str(saved[key])
    except (OSError, ValueError, TypeError):
        pass
    return settings


def save_settings(settings: dict[str, str]) -> None:
    """Lưu lựa chọn hợp lệ để tự động nạp lại ở lần chạy sau."""
    settings_path().write_text(
        json.dumps(settings, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def decimal_range(minimum: Decimal, maximum: Decimal, step: Decimal) -> tuple:
    """Tạo dãy số có cả giá trị max khi max nằm đúng trên bước."""
    values = []
    current = minimum
    while current <= maximum:
        value = float(current)
        values.append(int(value) if value.is_integer() else value)
        current += step
    return tuple(values)


def parse_range(settings: dict[str, str], prefix: str, label: str) -> tuple:
    """Kiểm tra Min/Max/Bước và trả về dãy giá trị."""
    try:
        minimum = Decimal(settings[f"{prefix}_min"].strip().replace(",", "."))
        maximum = Decimal(settings[f"{prefix}_max"].strip().replace(",", "."))
        step = Decimal(settings[f"{prefix}_step"].strip().replace(",", "."))
    except (InvalidOperation, AttributeError) as exc:
        raise ValueError(f"{label}: Min, Max và Bước phải là số.") from exc

    if not all(value.is_finite() for value in (minimum, maximum, step)):
        raise ValueError(f"{label}: giá trị phải là số hữu hạn.")
    if minimum < 0:
        raise ValueError(f"{label}: Min không được âm.")
    if maximum < minimum:
        raise ValueError(f"{label}: Max phải lớn hơn hoặc bằng Min.")
    if step <= 0:
        raise ValueError(f"{label}: Bước phải lớn hơn 0.")

    values = decimal_range(minimum, maximum, step)
    if not values:
        raise ValueError(f"{label}: không tạo được giá trị nào.")
    return values


def choose_matrix_ranges():
    """Hiện hộp thoại chọn dải công suất/CF và lưu lại lựa chọn."""
    try:
        import tkinter as tk
        from tkinter import messagebox, ttk
    except ImportError as exc:
        raise RuntimeError(
            "Python hiện tại không có Tkinter để hiển thị giao diện."
        ) from exc

    initial = load_settings()
    selected = None
    root = tk.Tk()
    root.title("Thiết lập ma trận độ nhạy")
    root.resizable(False, False)

    frame = ttk.Frame(root, padding=14)
    frame.grid(row=0, column=0, sticky="nsew")
    ttk.Label(
        frame,
        text="Chọn phạm vi tính toán trước khi chạy",
        font=("", 11, "bold"),
    ).grid(row=0, column=0, columnspan=4, pady=(0, 10))

    ttk.Label(frame, text="Thông số").grid(row=1, column=0, padx=5)
    for column, heading in enumerate(("Min", "Max", "Bước"), start=1):
        ttk.Label(frame, text=heading).grid(row=1, column=column, padx=5)

    variables = {key: tk.StringVar(value=value) for key, value in initial.items()}
    entries = []
    for label, prefix, row in (
        ("Công suất (MW)", "capacity", 2),
        ("CF (%)", "cf", 3),
    ):
        ttk.Label(frame, text=label).grid(
            row=row, column=0, sticky="w", padx=5, pady=4
        )
        for column, suffix in enumerate(("min", "max", "step"), start=1):
            entry = ttk.Entry(
                frame,
                textvariable=variables[f"{prefix}_{suffix}"],
                width=11,
                justify="right",
            )
            entry.grid(row=row, column=column, padx=5, pady=4)
            entries.append(entry)

    def accept():
        nonlocal selected
        raw = {key: variable.get().strip() for key, variable in variables.items()}
        try:
            capacities = parse_range(raw, "capacity", "Công suất")
            capacity_factors = parse_range(raw, "cf", "CF")
            if max(capacity_factors) > 100:
                raise ValueError("CF: Max không được lớn hơn 100%.")
            total_cases = len(capacities) * len(capacity_factors)
            if total_cases > MAX_MATRIX_CASES:
                raise ValueError(
                    f"Ma trận có {total_cases:,} phương án, vượt giới hạn "
                    f"{MAX_MATRIX_CASES:,}. Hãy tăng bước hoặc thu hẹp phạm vi."
                )
            save_settings(raw)
        except (ValueError, OSError) as exc:
            messagebox.showerror("Giá trị không hợp lệ", str(exc), parent=root)
            return
        selected = capacities, capacity_factors
        root.destroy()

    def cancel():
        root.destroy()

    buttons = ttk.Frame(frame)
    buttons.grid(row=4, column=0, columnspan=4, sticky="e", pady=(12, 0))
    ttk.Button(buttons, text="Hủy", command=cancel).grid(
        row=0, column=0, padx=(0, 8)
    )
    ttk.Button(buttons, text="Chạy", command=accept).grid(row=0, column=1)

    root.protocol("WM_DELETE_WINDOW", cancel)
    root.bind("<Return>", lambda _event: accept())
    root.bind("<Escape>", lambda _event: cancel())
    entries[0].focus_set()
    root.update_idletasks()
    root.geometry(
        f"+{(root.winfo_screenwidth() - root.winfo_width()) // 2}"
        f"+{(root.winfo_screenheight() - root.winfo_height()) // 2}"
    )
    root.mainloop()
    return selected


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



def normalize_label(value) -> str:
    """Chuẩn hóa nhãn Excel để tìm ổn định khi có thừa khoảng trắng."""
    if value is None:
        return ""
    return " ".join(str(value).split()).casefold()


def find_input_value_cell(input_sheet, expected_label: str):
    """Tìm nhãn trong UsedRange và trả về ô giá trị ở cột D cùng hàng."""
    used_range = input_sheet.UsedRange
    values = used_range.Value2

    if used_range.Rows.Count == 1 and used_range.Columns.Count == 1:
        value_rows = ((values,),)
    elif used_range.Rows.Count == 1:
        value_rows = (values,)
    else:
        value_rows = values

    target = normalize_label(expected_label)
    matches = []
    first_row = int(used_range.Row)
    first_column = int(used_range.Column)

    for row_offset, row_values in enumerate(value_rows):
        if not isinstance(row_values, tuple):
            row_values = (row_values,)
        for column_offset, value in enumerate(row_values):
            if normalize_label(value) == target:
                matches.append(
                    (first_row + row_offset, first_column + column_offset)
                )

    if not matches:
        raise ValueError(
            f'Không tìm thấy nhãn "{expected_label}" trong sheet '
            f'"{input_sheet.Name}".'
        )
    if len(matches) > 1:
        locations = ", ".join(
            input_sheet.Cells(row, column).Address(False, False)
            for row, column in matches
        )
        raise ValueError(
            f'Có nhiều ô mang nhãn "{expected_label}" trong sheet '
            f'"{input_sheet.Name}": {locations}.'
        )

    label_row, _label_column = matches[0]
    return input_sheet.Cells(label_row, INPUT_VALUE_COLUMN)


def read_numeric_result(cell, metric, capacity_mw, cf_percent) -> float:
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


def copy_input_snapshot(
    excel,
    source_input_sheet,
    result_workbook,
    after_sheet,
):
    """Tạo sheet Input mới và sao chép giá trị, định dạng, kích thước."""
    target_input_sheet = result_workbook.Worksheets.Add(After=after_sheet)
    target_input_sheet.Name = INPUT_SHEET

    source_range = source_input_sheet.UsedRange
    first_row = int(source_range.Row)
    first_column = int(source_range.Column)
    last_row = first_row + int(source_range.Rows.Count) - 1
    last_column = first_column + int(source_range.Columns.Count) - 1

    target_range = target_input_sheet.Range(
        target_input_sheet.Cells(first_row, first_column),
        target_input_sheet.Cells(last_row, last_column),
    )

    # Value2 chỉ lấy giá trị hiện tại của ô, không mang công thức hoặc liên
    # kết về workbook nguồn sang workbook kết quả.
    target_range.Value2 = source_range.Value2

    # Chỉ dán định dạng ô. Không dùng Worksheet.Copy nên không sao chép công
    # thức, named range, validation hoặc liên kết ngoài của worksheet nguồn.
    try:
        source_range.Copy()
        target_range.PasteSpecial(Paste=XL_PASTE_FORMATS)
    finally:
        excel.CutCopyMode = False

    # Sao chép chính xác độ rộng, trạng thái ẩn của từng cột trong UsedRange.
    for column_number in range(first_column, last_column + 1):
        source_column = source_input_sheet.Columns(column_number)
        target_column = target_input_sheet.Columns(column_number)
        target_column.ColumnWidth = source_column.ColumnWidth
        target_column.Hidden = source_column.Hidden

    # Sao chép chính xác chiều cao, trạng thái ẩn của từng hàng trong UsedRange.
    for row_number in range(first_row, last_row + 1):
        source_row = source_input_sheet.Rows(row_number)
        target_row = target_input_sheet.Rows(row_number)
        target_row.RowHeight = source_row.RowHeight
        target_row.Hidden = source_row.Hidden

    # PasteSpecial(xlPasteFormats) có thể xử lý ô gộp khác nhau giữa các phiên
    # bản Excel; duyệt và tái tạo rõ ràng để bố cục luôn giống sheet nguồn.
    merged_addresses = set()
    for row_number in range(first_row, last_row + 1):
        for column_number in range(first_column, last_column + 1):
            source_cell = source_input_sheet.Cells(row_number, column_number)
            if not bool(source_cell.MergeCells):
                continue
            merge_address = source_cell.MergeArea.Address(False, False)
            if merge_address in merged_addresses:
                continue
            target_merge_area = target_input_sheet.Range(merge_address)
            if not bool(target_merge_area.MergeCells):
                target_merge_area.Merge()
            merged_addresses.add(merge_address)

    return target_input_sheet


def create_results_workbook(excel, results, source_input_sheet):
    """Tạo workbook kết quả gồm bốn ma trận và ảnh chụp sheet Input."""
    result_workbook = excel.Workbooks.Add()
    result_sheet = result_workbook.Worksheets(1)
    result_sheet.Name = "Kết quả"

    # Chỉ giữ lại một worksheet trong workbook kết quả.
    while result_workbook.Worksheets.Count > 1:
        result_workbook.Worksheets(
            result_workbook.Worksheets.Count
        ).Delete()

    # Tạo sheet Input mới rồi chỉ sao chép giá trị, định dạng và kích thước.
    # Không dùng Worksheet.Copy để tránh công thức/liên kết và lỗi COM Bad index.
    copy_input_snapshot(
        excel,
        source_input_sheet,
        result_workbook,
        result_sheet,
    )

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
        cf_cell = find_input_value_cell(input_sheet, CF_LABEL)
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

        # Tính lại sau khi hoàn nguyên để sheet Input được sao chép đúng với
        # bộ giả định ban đầu, không giữ trạng thái của phương án cuối cùng.
        excel.CalculateFullRebuild()
        wait_for_excel(excel)

        result_workbook = create_results_workbook(
            excel,
            results,
            input_sheet,
        )

        # Đóng file nguồn mà không lưu, chỉ giữ workbook kết quả mới (gồm
        # sheet Kết quả và bản sao sheet Input) trong Excel. Workbook kết quả
        # chưa có đường dẫn cho đến khi người dùng tự
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
    headers = ["CF \\ MW", *(f"{mw:g} MW" for mw in CAPACITIES_MW)]
    body = [
        [
            f"{cf_percent:g}%",
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

        ranges = choose_matrix_ranges()
        if ranges is None:
            print("Đã hủy; chưa mở Excel và chưa thực hiện tính toán.")
            return 0

        global CAPACITIES_MW, CAPACITY_FACTORS_PERCENT
        CAPACITIES_MW, CAPACITY_FACTORS_PERCENT = ranges
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
            "Đã tạo workbook kết quả mới với 4 ma trận trên sheet Kết quả, "
            "tạo sheet Input dạng bản chụp giá trị/định dạng và mở trong Excel."
        )
        print("Workbook kết quả chưa được lưu; hãy dùng Save hoặc Save As.")
        print(f'File nguồn "{WORKBOOK_NAME}" không bị thay đổi.')
        return 0

    except Exception as exc:
        print(f"LỖI: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
