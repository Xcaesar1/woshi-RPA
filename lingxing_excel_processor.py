from __future__ import annotations

import argparse
import json
import re
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.rich_text import CellRichText, InlineFont, TextBlock
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


SOURCE_REQUIRED_HEADERS = [
    "序号",
    "MSKU",
    "FNSKU",
    "品名",
    "SKU",
    "发货量",
    "单箱数量",
    "箱数",
    "箱号",
]

TEMPLATE_REQUIRED_HEADERS = [
    "NO.",
    "SKU",
    "工厂型号",
    "品名",
    "内盒标签",
    "箱号",
    "套",
    "箱数",
    "套/箱",
    "票数",
    "FBA号",
    "备注/品线",
    "说明书/包装上-品牌",
]

MSKU_MAP_REQUIRED_HEADERS = [
    "店铺",
    "MSKU",
]

STORE_DETAIL_REQUIRED_HEADERS = [
    "店铺+站点",
    "店铺",
    "店铺简称",
]

FIELD_MAPPING = {
    "NO.": "序号",
    "SKU": "MSKU",
    "工厂型号": "SKU",
    "品名": "品名",
    "内盒标签": "FNSKU",
    "箱号": "箱号清洗后加前缀“编号 ”",
    "套": "发货量",
    "箱数": "箱数",
    "套/箱": "单箱数量",
    "票数": "店铺简称+海运第x票",
    "FBA号": "货件单号",
    "备注/品线": "品名关键词归类(浴缸/厨房/淋浴/面盆)",
    "说明书/包装上-品牌": "MSKU -> MSKU对应品线表[店铺] -> 店铺明细表[店铺]",
}

PRODUCT_LINE_KEYWORDS = {
    "浴缸": "浴缸",
    "厨房": "厨房",
    "淋浴": "淋浴",
    "面盆": "面盆",
}

BOX_RANGE_RE = re.compile(r"(\d+)\s*[～~\-－—–至]+\s*(\d+)\s*$")
BOX_SINGLE_RE = re.compile(r"(\d+)\s*$")
DATE_IN_FILENAME_RE = re.compile(r"_(\d{8})(?:[-_]|$)")
THIN_SIDE = Side(style="thin", color="000000")
FULL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
TITLE_ALIGNMENT = Alignment(horizontal="center", vertical="center")
HEADER_FBA_FONT = InlineFont(rFont="宋体", b=True, sz=14, color="FF0000")
DATA_FONT = Font(name="宋体", size=12)
TITLE_FONT = Font(name="宋体", size=20)
DATA_ROW_HEIGHT = 30
HEADER_ROW_HEIGHT = 20
OUTPUT_COLUMN_WIDTHS = {
    1: 8,
    2: 18,
    3: 18,
    4: 32,
    5: 20,
    6: 18,
    7: 10,
    8: 10,
    9: 12,
    10: 18,
    11: 20,
    12: 16,
    13: 24,
}


@dataclass
class WorkbookSelection:
    path: Path
    sheet_name: str
    header_row: int
    headers: dict[str, int]


@dataclass
class StoreLookupResult:
    row_number: int
    msku: str | None
    store_site: str | None
    store_name: str | None
    store_short: str | None
    mapped_line: str | None
    brand_name: str | None
    ticket_value: str | None
    lookup_ok: bool


def iter_xlsx_files(base_dir: Path) -> list[Path]:
    return sorted(path for path in base_dir.glob("*.xlsx") if not path.name.startswith("~$"))


def iter_source_files(base_dir: Path) -> list[Path]:
    candidates = [
        path
        for path in iter_xlsx_files(base_dir)
        if path.name.startswith("FBA")
    ]
    return sorted(candidates, key=lambda path: (path.stat().st_ctime_ns, path.name))


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).strip()


def normalize_lookup_key(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).upper()


def is_blank(value: Any) -> bool:
    return value is None or str(value).strip() == ""


def convert_numeric(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return int(value) if float(value).is_integer() else float(value)

    text = str(value).strip()
    if not text:
        return None

    try:
        number = float(text)
    except ValueError:
        return text

    return int(number) if number.is_integer() else number


def dedupe_preserve_order(items: list[str]) -> list[str]:
    seen: set[str] = set()
    output: list[str] = []
    for item in items:
        if item not in seen:
            seen.add(item)
            output.append(item)
    return output


def find_matching_sheet(workbook_path: Path, required_headers: list[str]) -> WorkbookSelection:
    workbook = load_workbook(workbook_path)
    try:
        for sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            for row_idx in range(1, min(worksheet.max_row, 20) + 1):
                row_headers: dict[str, int] = {}
                for col_idx in range(1, worksheet.max_column + 1):
                    normalized = normalize_header(worksheet.cell(row=row_idx, column=col_idx).value)
                    if normalized:
                        row_headers[normalized] = col_idx
                if all(header in row_headers for header in required_headers):
                    return WorkbookSelection(
                        path=workbook_path,
                        sheet_name=sheet_name,
                        header_row=row_idx,
                        headers=row_headers,
                    )
    finally:
        workbook.close()

    raise ValueError(f"未在 {workbook_path.name} 中找到包含以下表头的工作表：{', '.join(required_headers)}")


def locate_template_file(base_dir: Path) -> Path:
    exact_template = base_dir / "模板.xlsx"
    if exact_template.exists():
        return exact_template

    candidates = [path for path in iter_xlsx_files(base_dir) if "模板" in path.stem]
    if candidates:
        return sorted(candidates, key=lambda path: (len(path.name), path.name))[0]

    raise FileNotFoundError("未找到模板 Excel 文件。")


def locate_source_file(base_dir: Path) -> Path:
    candidates = [path for path in iter_source_files(base_dir) if "_NO_PIC" in path.name]
    if candidates:
        return candidates[0]

    candidates = iter_source_files(base_dir)
    if candidates:
        return candidates[0]

    raise FileNotFoundError("未找到领星下载的源数据 Excel 文件。")


def locate_msku_mapping_file(base_dir: Path) -> Path:
    exact_path = base_dir / "MSKU对应品线表.xlsx"
    if exact_path.exists():
        return exact_path

    candidates = [path for path in iter_xlsx_files(base_dir) if "MSKU" in path.name]
    if candidates:
        return sorted(candidates, key=lambda path: (len(path.name), path.name))[0]

    raise FileNotFoundError("未找到 MSKU 对应品线表 Excel 文件。")


def locate_store_detail_file(base_dir: Path) -> Path:
    exact_path = base_dir / "店铺明细表.xlsx"
    if exact_path.exists():
        return exact_path

    candidates = [path for path in iter_xlsx_files(base_dir) if "店铺明细表" in path.name]
    if candidates:
        return sorted(candidates, key=lambda path: (len(path.name), path.name))[0]

    raise FileNotFoundError("未找到店铺明细表 Excel 文件。")


def extract_metadata(worksheet: Worksheet, header_row: int) -> dict[str, Any]:
    metadata: dict[str, Any] = {}
    for row_idx in range(1, header_row):
        for col_idx in range(1, worksheet.max_column):
            label = normalize_header(worksheet.cell(row=row_idx, column=col_idx).value)
            if not label:
                continue
            value = worksheet.cell(row=row_idx, column=col_idx + 1).value
            if value is not None and label not in metadata:
                metadata[label] = value
    return metadata


def extract_detail_rows(worksheet: Worksheet, selection: WorkbookSelection) -> list[dict[str, Any]]:
    detail_rows: list[dict[str, Any]] = []
    relevant_headers = [header for header in SOURCE_REQUIRED_HEADERS if header in selection.headers]

    for row_idx in range(selection.header_row + 1, worksheet.max_row + 1):
        row_data = {
            header: worksheet.cell(row=row_idx, column=selection.headers[header]).value
            for header in relevant_headers
        }
        if all(is_blank(row_data[header]) for header in relevant_headers):
            continue
        detail_rows.append(row_data)

    return detail_rows


def build_lookup_index(workbook_path: Path, selection: WorkbookSelection, key_header: str) -> dict[str, list[dict[str, Any]]]:
    workbook = load_workbook(workbook_path, data_only=True)
    try:
        worksheet = workbook[selection.sheet_name]
        index: dict[str, list[dict[str, Any]]] = {}
        headers = list(selection.headers.keys())
        key_col = selection.headers[key_header]

        for row_idx in range(selection.header_row + 1, worksheet.max_row + 1):
            key_value = worksheet.cell(row=row_idx, column=key_col).value
            normalized_key = normalize_lookup_key(key_value)
            if not normalized_key:
                continue

            row_data = {
                header: worksheet.cell(row=row_idx, column=selection.headers[header]).value
                for header in headers
            }
            row_data["_row_number"] = row_idx
            index.setdefault(normalized_key, []).append(row_data)

        return index
    finally:
        workbook.close()


def clean_box_number(raw_value: Any) -> tuple[Any, str | None]:
    if is_blank(raw_value):
        return None, "箱号为空，已保留为空值"

    text = str(raw_value).strip().rstrip("；;，,")
    range_match = BOX_RANGE_RE.search(text)
    if range_match:
        start = str(int(range_match.group(1)))
        end = str(int(range_match.group(2)))
        return f"{start}-{end}", None

    single_match = BOX_SINGLE_RE.search(text)
    if single_match:
        return str(int(single_match.group(1))), None

    return text, f"箱号无法识别，保留原值：{text}"


def format_box_display(box_value: Any) -> Any:
    if is_blank(box_value):
        return None
    return f"编号 {box_value}"


def extract_fba_number(raw_value: Any) -> tuple[Any, str | None]:
    if is_blank(raw_value):
        return None, "未找到货件单号，FBA号留空"

    text = str(raw_value).strip()
    if re.fullmatch(r"[A-Za-z0-9-]+", text):
        return text, None

    tokens = re.findall(r"[A-Za-z0-9-]+", text)
    if len(tokens) == 1:
        return tokens[0], None

    return text, f"FBA号无法判断，已保留原始货件单号：{text}"


def classify_product_line(product_name: Any) -> tuple[Any, str | None]:
    if is_blank(product_name):
        return None, "品名为空，备注/品线留空"

    text = str(product_name).strip()
    for keyword, line_name in PRODUCT_LINE_KEYWORDS.items():
        if keyword in text:
            return line_name, None

    return None, f"备注/品线无法根据品名判断：{text}"


def format_store_brand(store_name: Any) -> str | None:
    if is_blank(store_name):
        return None

    text = str(store_name).strip()

    def repl(match: re.Match[str]) -> str:
        word = match.group(0)
        return word[:1].upper() + word[1:].lower()

    return re.sub(r"[A-Za-z]+", repl, text)


def select_source_files(base_dir: Path) -> list[Path]:
    preferred = [path for path in iter_source_files(base_dir) if "_NO_PIC" in path.name]
    return preferred or iter_source_files(base_dir)


def extract_ticket_date(metadata: dict[str, Any], source_path: Path) -> tuple[str, str | None]:
    cargo_name = "" if is_blank(metadata.get("货件名称")) else str(metadata.get("货件名称")).strip()
    cargo_match = re.search(r"(20\d{6})(?!\d)", cargo_name)
    if cargo_match:
        full_date = cargo_match.group(1)
        return full_date[2:], None

    filename_match = DATE_IN_FILENAME_RE.search(source_path.name)
    if filename_match:
        full_date = filename_match.group(1)
        return full_date[2:], "货件名称中未找到日期，已回退使用源文件名日期"

    return datetime.now().strftime("%y%m%d"), "货件名称和源文件名中都未找到日期，已回退使用系统日期"


def extract_station_code(store_site: Any) -> str | None:
    if is_blank(store_site):
        return None
    text = str(store_site).strip()
    parts = [part for part in re.split(r"[-_]", text) if part]
    return parts[-1] if parts else text


def build_title_store_label(store_short: Any, store_site: Any) -> str | None:
    if is_blank(store_short):
        return None
    station_code = extract_station_code(store_site)
    if is_blank(station_code):
        return None
    return f"{str(store_short).strip()}-{station_code}"


def ensure_merge_range(worksheet: Worksheet, cell_range: str) -> None:
    merge_ranges = {str(existing) for existing in worksheet.merged_cells.ranges}
    if cell_range not in merge_ranges:
        worksheet.merge_cells(cell_range)


def sync_merged_range_borders(worksheet: Worksheet, cell_range: str) -> None:
    for merged_range in worksheet.merged_cells.ranges:
        if str(merged_range) == cell_range:
            merged_range.format()
            break


def border_side_is_thin(side: Side) -> bool:
    return side.style == "thin"


def reinforce_header_block_borders(worksheet: Worksheet, header_row: int, header_blank_row: int, max_col: int) -> None:
    for col_idx in range(1, max_col + 1):
        merged_range = (
            f"{get_column_letter(col_idx)}{header_row}:{get_column_letter(col_idx)}{header_blank_row}"
        )
        sync_merged_range_borders(worksheet, merged_range)


def header_block_borders_are_complete(
    worksheet: Worksheet,
    header_row: int,
    header_blank_row: int,
    max_col: int,
) -> bool:
    for col_idx in range(1, max_col + 1):
        top_cell = worksheet.cell(row=header_row, column=col_idx)
        bottom_cell = worksheet.cell(row=header_blank_row, column=col_idx)
        if not (
            border_side_is_thin(top_cell.border.left)
            and border_side_is_thin(top_cell.border.right)
            and border_side_is_thin(top_cell.border.top)
            and border_side_is_thin(top_cell.border.bottom)
        ):
            return False
        if not (
            border_side_is_thin(bottom_cell.border.left)
            and border_side_is_thin(bottom_cell.border.right)
            and border_side_is_thin(bottom_cell.border.bottom)
        ):
            return False
    return True


def prepare_output_sheet(workbook, keep_sheet_name: str, output_sheet_name: str) -> Worksheet:
    for sheet_name in list(workbook.sheetnames):
        if sheet_name != keep_sheet_name:
            workbook.remove(workbook[sheet_name])
    worksheet = workbook[keep_sheet_name]
    worksheet.title = output_sheet_name
    return worksheet


def create_output_workbook(output_sheet_name: str) -> tuple[Any, Worksheet, WorkbookSelection]:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = output_sheet_name
    for col_idx, width in OUTPUT_COLUMN_WIDTHS.items():
        worksheet.column_dimensions[get_column_letter(col_idx)].width = width

    header_map = {header: idx for idx, header in enumerate(TEMPLATE_REQUIRED_HEADERS, start=1)}
    selection = WorkbookSelection(
        path=Path("<generated>"),
        sheet_name=output_sheet_name,
        header_row=2,
        headers=header_map,
    )
    return workbook, worksheet, selection


def build_block_positions(start_row: int, detail_row_count: int) -> dict[str, int]:
    return {
        "title_row": start_row,
        "header_row": start_row + 1,
        "header_blank_row": start_row + 2,
        "detail_start_row": start_row + 3,
        "summary_row": start_row + 3 + detail_row_count,
        "next_start_row": start_row + 4 + detail_row_count,
    }


def apply_header_block(
    worksheet: Worksheet,
    block_start_row: int,
    header_values: dict[int, Any],
    fba_number: Any,
    max_col: int,
) -> None:
    header_row = block_start_row + 1
    header_blank_row = block_start_row + 2

    if block_start_row > 1:
        worksheet.row_dimensions[block_start_row].height = worksheet.row_dimensions[1].height
        clone_row_format(worksheet, 2, header_row, max_col)
        clone_row_format(worksheet, 3, header_blank_row, max_col)
    else:
        worksheet.row_dimensions[header_row].height = HEADER_ROW_HEIGHT
        worksheet.row_dimensions[header_blank_row].height = HEADER_ROW_HEIGHT

    ensure_merge_range(worksheet, build_title_merge_range(block_start_row))
    for col_idx in range(1, max_col + 1):
        column_letter = get_column_letter(col_idx)
        merged_range = f"{column_letter}{header_row}:{column_letter}{header_blank_row}"
        ensure_merge_range(worksheet, merged_range)
        header_cell = worksheet.cell(row=header_row, column=col_idx)
        if col_idx != 6:
            header_cell.value = header_values.get(col_idx)
        header_cell.alignment = copy(HEADER_ALIGNMENT)
        header_cell.border = FULL_BORDER
        sync_merged_range_borders(worksheet, merged_range)

    box_header_ref = f"{worksheet.cell(row=header_row, column=6).column_letter}{header_row}"
    apply_box_header_style(worksheet, box_header_ref, fba_number)
    reinforce_header_block_borders(worksheet, header_row, header_blank_row, max_col)


def detect_data_start_row(worksheet: Worksheet, header_row: int) -> int:
    max_header_row = header_row
    for merged_range in worksheet.merged_cells.ranges:
        if merged_range.min_row <= header_row <= merged_range.max_row:
            max_header_row = max(max_header_row, merged_range.max_row)
    return max_header_row + 1


def clone_row_format(worksheet: Worksheet, source_row: int, target_row: int, max_col: int) -> None:
    source_dimension = worksheet.row_dimensions[source_row]
    target_dimension = worksheet.row_dimensions[target_row]
    target_dimension.height = source_dimension.height
    target_dimension.hidden = source_dimension.hidden
    target_dimension.outlineLevel = source_dimension.outlineLevel

    for col_idx in range(1, max_col + 1):
        source_cell = worksheet.cell(row=source_row, column=col_idx)
        target_cell = worksheet.cell(row=target_row, column=col_idx)
        if isinstance(source_cell, MergedCell):
            continue
        if source_cell.has_style:
            target_cell._style = copy(source_cell._style)
        if source_cell.font:
            target_cell.font = copy(source_cell.font)
        if source_cell.fill:
            target_cell.fill = copy(source_cell.fill)
        if source_cell.border:
            target_cell.border = copy(source_cell.border)
        if source_cell.alignment:
            target_cell.alignment = copy(source_cell.alignment)
        if source_cell.protection:
            target_cell.protection = copy(source_cell.protection)
        if source_cell.number_format:
            target_cell.number_format = source_cell.number_format


def clear_target_range(worksheet: Worksheet, start_row: int, end_row: int, max_col: int) -> None:
    if end_row < start_row:
        return

    for row_idx in range(start_row, end_row + 1):
        for col_idx in range(1, max_col + 1):
            worksheet.cell(row=row_idx, column=col_idx).value = None


def ensure_header_room(worksheet: Worksheet, header_row: int) -> None:
    next_row = header_row + 1
    current_header_height = worksheet.row_dimensions[header_row].height or 15
    current_next_height = worksheet.row_dimensions[next_row].height or 15
    worksheet.row_dimensions[header_row].height = max(current_header_height, 20)
    worksheet.row_dimensions[next_row].height = max(current_next_height, 20)


def build_title_merge_range(title_row: int) -> str:
    return f"A{title_row}:M{title_row}"


def apply_title_style(worksheet: Worksheet, title_row: int, title_text: str) -> None:
    ensure_merge_range(worksheet, build_title_merge_range(title_row))
    title_cell = worksheet.cell(row=title_row, column=1)
    title_cell.value = title_text
    title_cell.font = copy(TITLE_FONT)
    title_cell.alignment = copy(TITLE_ALIGNMENT)
    worksheet.row_dimensions[title_row].height = 30


def apply_box_header_style(worksheet: Worksheet, header_cell_ref: str, fba_number: Any) -> None:
    header_cell = worksheet[header_cell_ref]
    fba_text = "" if is_blank(fba_number) else str(fba_number).strip()
    if fba_text:
        header_cell.value = CellRichText("箱号\n", TextBlock(HEADER_FBA_FONT, fba_text))
    else:
        header_cell.value = "箱号"
    header_cell.alignment = copy(HEADER_ALIGNMENT)
    header_cell.border = FULL_BORDER
    sync_merged_range_borders(
        worksheet,
        f"{header_cell.column_letter}{header_cell.row}:{header_cell.column_letter}{header_cell.row + 1}",
    )
    ensure_header_room(worksheet, header_cell.row)


def apply_data_cell_style(cell) -> None:
    wrap_text = bool(cell.alignment.wrapText) if cell.alignment else False
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap_text)
    cell.border = FULL_BORDER
    cell.font = copy(DATA_FONT)


def apply_data_row_style(worksheet: Worksheet, row_idx: int, max_col: int) -> None:
    worksheet.row_dimensions[row_idx].height = DATA_ROW_HEIGHT
    for col_idx in range(1, max_col + 1):
        apply_data_cell_style(worksheet.cell(row=row_idx, column=col_idx))


def build_output_name(store_short: str | None, mapped_line: str | None, ticket_date: str | None) -> str:
    short_part = store_short or "UNKNOWN"
    line_part = mapped_line or "未知品线"
    date_part = ticket_date or datetime.now().strftime("%y%m%d")
    return f"{short_part}-{line_part}-货代信息表（排序表）-{date_part}.xlsx"


def save_workbook_with_fallback(workbook, preferred_path: Path) -> Path:
    if not preferred_path.exists():
        try:
            workbook.save(preferred_path)
            return preferred_path
        except PermissionError:
            pass

    for index in range(1, 100):
        fallback_path = preferred_path.with_name(f"{preferred_path.stem}_样式调整{index}{preferred_path.suffix}")
        try:
            workbook.save(fallback_path)
            return fallback_path
        except PermissionError:
            continue

    raise PermissionError(f"无法保存输出文件：{preferred_path.name}")


def resolve_store_lookup(
    row_number: int,
    msku_value: Any,
    ticket_index: int,
    msku_index: dict[str, list[dict[str, Any]]],
    store_index: dict[str, list[dict[str, Any]]],
    anomalies: list[str],
) -> StoreLookupResult:
    msku_text = None if is_blank(msku_value) else str(msku_value).strip()
    if is_blank(msku_text):
        anomalies.append(f"第 {row_number} 行：MSKU 为空，无法查店铺")
        return StoreLookupResult(row_number, msku_text, None, None, None, None, None, None, False)

    msku_matches = msku_index.get(normalize_lookup_key(msku_text), [])
    if not msku_matches:
        anomalies.append(f"第 {row_number} 行：MSKU 在 MSKU对应品线表.xlsx 中找不到：{msku_text}")
        return StoreLookupResult(row_number, msku_text, None, None, None, None, None, None, False)

    if len(msku_matches) > 1:
        anomalies.append(f"第 {row_number} 行：MSKU 在 MSKU对应品线表.xlsx 中匹配到多条记录：{msku_text}")
        return StoreLookupResult(row_number, msku_text, None, None, None, None, None, None, False)

    mapped_line_raw = msku_matches[0].get("品线")
    mapped_line = None if is_blank(mapped_line_raw) else str(mapped_line_raw).strip()
    store_site_value = msku_matches[0].get("店铺")
    store_site_text = None if is_blank(store_site_value) else str(store_site_value).strip()
    if is_blank(store_site_text):
        anomalies.append(f"第 {row_number} 行：MSKU 对应的店铺字段为空：{msku_text}")
        return StoreLookupResult(row_number, msku_text, None, None, None, mapped_line, None, None, False)

    store_matches = store_index.get(normalize_lookup_key(store_site_text), [])
    if not store_matches:
        anomalies.append(f"第 {row_number} 行：店铺明细表.xlsx 中找不到对应店铺+站点：{store_site_text}")
        return StoreLookupResult(row_number, msku_text, store_site_text, None, None, mapped_line, None, None, False)

    if len(store_matches) > 1:
        anomalies.append(f"第 {row_number} 行：店铺明细表.xlsx 中店铺+站点匹配到多条记录：{store_site_text}")
        return StoreLookupResult(row_number, msku_text, store_site_text, None, None, mapped_line, None, None, False)

    store_name_raw = store_matches[0].get("店铺")
    store_short_raw = store_matches[0].get("店铺简称")
    store_name = None if is_blank(store_name_raw) else str(store_name_raw).strip()
    store_short = None if is_blank(store_short_raw) else str(store_short_raw).strip()
    brand_name = format_store_brand(store_name)
    ticket_value = None if is_blank(store_short) else f"{store_short}海运第{ticket_index}票"

    if is_blank(store_short):
        anomalies.append(f"第 {row_number} 行：店铺简称为空，票数无法填写：{store_site_text}")
    if is_blank(store_name):
        anomalies.append(f"第 {row_number} 行：店铺名称为空，说明书/包装上-品牌无法填写：{store_site_text}")
    if is_blank(mapped_line):
        anomalies.append(f"第 {row_number} 行：MSKU 对应的品线为空：{msku_text}")

    lookup_ok = not is_blank(store_short) and not is_blank(store_name) and not is_blank(mapped_line)
    return StoreLookupResult(
        row_number=row_number,
        msku=msku_text,
        store_site=store_site_text,
        store_name=store_name,
        store_short=store_short,
        mapped_line=mapped_line,
        brand_name=brand_name,
        ticket_value=ticket_value,
        lookup_ok=lookup_ok,
    )


def process_workbooks(resource_dir: Path, source_dir: Path, output_dir: Path) -> dict[str, Any]:
    source_files = select_source_files(source_dir)
    if not source_files:
        raise FileNotFoundError("未找到可处理的领星源数据 Excel 文件。")

    msku_map_path = locate_msku_mapping_file(resource_dir)
    store_detail_path = locate_store_detail_file(resource_dir)

    msku_selection = find_matching_sheet(msku_map_path, MSKU_MAP_REQUIRED_HEADERS)
    store_selection = find_matching_sheet(store_detail_path, STORE_DETAIL_REQUIRED_HEADERS)

    msku_index = build_lookup_index(msku_map_path, msku_selection, "MSKU")
    store_index = build_lookup_index(store_detail_path, store_selection, "店铺+站点")

    output_sheet_name = f"{len(source_files)}票"
    template_workbook, template_sheet, template_selection = create_output_workbook(output_sheet_name)

    anomalies: list[str] = []
    try:
        target_max_col = len(TEMPLATE_REQUIRED_HEADERS)
        output_sheet_name = template_selection.sheet_name
        header_values = {
            col_idx: TEMPLATE_REQUIRED_HEADERS[col_idx - 1]
            for col_idx in range(1, target_max_col + 1)
        }

        block_reports: list[dict[str, Any]] = []
        lookup_rows_for_report: list[dict[str, Any]] = []
        all_store_sites: list[str] = []
        all_store_shorts: list[str] = []
        all_store_names: list[str] = []
        all_mapped_lines: list[str] = []
        ticket_date_for_filename: str | None = None
        filename_store_short: str | None = None
        filename_mapped_line: str | None = None
        current_start_row = 1

        for ticket_index, source_path in enumerate(source_files, start=1):
            source_selection = find_matching_sheet(source_path, SOURCE_REQUIRED_HEADERS)
            source_workbook = load_workbook(source_path, data_only=False)
            try:
                source_sheet = source_workbook[source_selection.sheet_name]
                metadata = extract_metadata(source_sheet, source_selection.header_row)
                shipment_no = metadata.get("货件单号")
                fba_number, fba_anomaly = extract_fba_number(shipment_no)
                if fba_anomaly:
                    anomalies.append(f"第{ticket_index}票：{fba_anomaly}")

                ticket_date, ticket_date_anomaly = extract_ticket_date(metadata, source_path)
                if ticket_index == 1:
                    ticket_date_for_filename = ticket_date
                if ticket_date_anomaly:
                    anomalies.append(f"第{ticket_index}票：{ticket_date_anomaly}")

                source_rows = extract_detail_rows(source_sheet, source_selection)
                positions = build_block_positions(current_start_row, len(source_rows))
                clear_target_range(
                    template_sheet,
                    positions["detail_start_row"],
                    max(template_sheet.max_row, positions["summary_row"]),
                    target_max_col,
                )
                apply_header_block(template_sheet, current_start_row, header_values, fba_number, target_max_col)

                lookup_results: list[StoreLookupResult] = []
                expected_set_total = sum(convert_numeric(row.get("发货量")) or 0 for row in source_rows)
                expected_carton_total = sum(convert_numeric(row.get("箱数")) or 0 for row in source_rows)

                for offset, source_row in enumerate(source_rows):
                    target_row = positions["detail_start_row"] + offset
                    if offset > 0:
                        clone_row_format(template_sheet, positions["detail_start_row"], target_row, target_max_col)

                    box_number, box_anomaly = clean_box_number(source_row.get("箱号"))
                    if box_anomaly:
                        anomalies.append(f"第{ticket_index}票第 {offset + 1} 行：{box_anomaly}")

                    product_line, product_anomaly = classify_product_line(source_row.get("品名"))
                    if product_anomaly:
                        anomalies.append(f"第{ticket_index}票第 {offset + 1} 行：{product_anomaly}")

                    lookup_result = resolve_store_lookup(
                        row_number=offset + 1,
                        msku_value=source_row.get("MSKU"),
                        ticket_index=ticket_index,
                        msku_index=msku_index,
                        store_index=store_index,
                        anomalies=anomalies,
                    )
                    lookup_results.append(lookup_result)

                    row_payload = {
                        "NO.": convert_numeric(source_row.get("序号")),
                        "SKU": source_row.get("MSKU"),
                        "工厂型号": source_row.get("SKU"),
                        "品名": source_row.get("品名"),
                        "内盒标签": source_row.get("FNSKU"),
                        "箱号": format_box_display(box_number),
                        "套": convert_numeric(source_row.get("发货量")),
                        "箱数": convert_numeric(source_row.get("箱数")),
                        "套/箱": convert_numeric(source_row.get("单箱数量")),
                        "票数": lookup_result.ticket_value,
                        "FBA号": fba_number,
                        "备注/品线": product_line,
                        "说明书/包装上-品牌": lookup_result.brand_name,
                    }

                    for header, value in row_payload.items():
                        col_idx = template_selection.headers[header]
                        template_sheet.cell(row=target_row, column=col_idx).value = value

                    apply_data_row_style(template_sheet, target_row, target_max_col)

                    lookup_rows_for_report.append(
                        {
                            "ticket_index": ticket_index,
                            "source_file": source_path.name,
                            "row_number": offset + 1,
                            "msku": lookup_result.msku,
                            "store_site": lookup_result.store_site,
                            "store_short": lookup_result.store_short,
                            "store_name": lookup_result.store_name,
                            "mapped_line": lookup_result.mapped_line,
                            "ticket_value": lookup_result.ticket_value,
                            "brand_value": lookup_result.brand_name,
                        }
                    )

                if source_rows:
                    clone_row_format(template_sheet, positions["detail_start_row"], positions["summary_row"], target_max_col)
                template_sheet.cell(row=positions["summary_row"], column=template_selection.headers["套"]).value = expected_set_total
                template_sheet.cell(row=positions["summary_row"], column=template_selection.headers["箱数"]).value = expected_carton_total
                apply_data_row_style(template_sheet, positions["summary_row"], target_max_col)

                resolved_store_sites = dedupe_preserve_order(
                    [result.store_site for result in lookup_results if not is_blank(result.store_site)]
                )
                resolved_store_shorts = dedupe_preserve_order(
                    [result.store_short for result in lookup_results if not is_blank(result.store_short)]
                )
                resolved_store_names = dedupe_preserve_order(
                    [result.store_name for result in lookup_results if not is_blank(result.store_name)]
                )
                resolved_mapped_lines = dedupe_preserve_order(
                    [result.mapped_line for result in lookup_results if not is_blank(result.mapped_line)]
                )

                all_store_sites.extend(resolved_store_sites)
                all_store_shorts.extend(resolved_store_shorts)
                all_store_names.extend(resolved_store_names)
                all_mapped_lines.extend(resolved_mapped_lines)

                if len(resolved_store_shorts) > 1:
                    anomalies.append(f"第{ticket_index}票：同一仓库文件存在多个店铺简称：{', '.join(resolved_store_shorts)}")
                if len(resolved_store_sites) > 1:
                    anomalies.append(f"第{ticket_index}票：同一仓库文件存在多个店铺+站点：{', '.join(resolved_store_sites)}")
                if len(resolved_mapped_lines) > 1:
                    anomalies.append(f"第{ticket_index}票：同一仓库文件存在多个品线：{', '.join(resolved_mapped_lines)}")

                title_label = None
                if lookup_results:
                    title_label = build_title_store_label(lookup_results[0].store_short, lookup_results[0].store_site)
                if is_blank(title_label) and resolved_store_shorts and resolved_store_sites:
                    title_label = build_title_store_label(resolved_store_shorts[0], resolved_store_sites[0])
                    anomalies.append(f"第{ticket_index}票：首行SKU未能成功解析标题店铺信息，标题改用首个可用店铺简称和站点")
                if is_blank(title_label):
                    title_label = "待确认"
                    anomalies.append(f"第{ticket_index}票：无法确定标题中的店铺简称和站点")

                title_text = f"装箱单（{title_label}）海运第{ticket_index}票"
                apply_title_style(template_sheet, positions["title_row"], title_text)

                if ticket_index == 1:
                    filename_store_short = resolved_store_shorts[0] if resolved_store_shorts else None
                    filename_mapped_line = resolved_mapped_lines[0] if resolved_mapped_lines else None

                block_reports.append(
                    {
                        "ticket_index": ticket_index,
                        "source_file": source_path.name,
                        "source_sheet": source_selection.sheet_name,
                        "source_header_row": source_selection.header_row,
                        "shipment_number": shipment_no,
                        "fba_number": fba_number,
                        "ticket_date": ticket_date,
                        "title_text": title_text,
                        "positions": positions,
                        "detail_row_count": len(source_rows),
                        "expected_set_total": expected_set_total,
                        "expected_carton_total": expected_carton_total,
                        "resolved_store_sites": resolved_store_sites,
                        "resolved_store_shorts": resolved_store_shorts,
                        "resolved_store_names": resolved_store_names,
                        "resolved_mapped_lines": resolved_mapped_lines,
                    }
                )
                current_start_row = positions["next_start_row"]
            finally:
                source_workbook.close()

        deduped_store_shorts = dedupe_preserve_order(all_store_shorts)
        deduped_mapped_lines = dedupe_preserve_order(all_mapped_lines)
        if len(deduped_store_shorts) > 1:
            anomalies.append(f"整个输出文件存在多个店铺简称：{', '.join(deduped_store_shorts)}")
        if len(deduped_mapped_lines) > 1:
            anomalies.append(f"整个输出文件存在多个品线：{', '.join(deduped_mapped_lines)}")

        preferred_output_path = output_dir / build_output_name(filename_store_short, filename_mapped_line, ticket_date_for_filename)
        output_path = save_workbook_with_fallback(template_workbook, preferred_output_path)

        written_workbook = load_workbook(output_path, data_only=False, rich_text=True)
        try:
            written_sheet = written_workbook[written_workbook.sheetnames[0]]
            actual_sheet_title = written_sheet.title
            actual_sheet_count = len(written_workbook.sheetnames)
            total_expected_rows = sum(block["detail_row_count"] for block in block_reports)
            total_actual_rows = 0
            total_expected_set = sum(block["expected_set_total"] for block in block_reports)
            total_actual_set = 0
            total_expected_carton = sum(block["expected_carton_total"] for block in block_reports)
            total_actual_carton = 0
            block_title_ok = True
            block_merge_ok = True
            block_title_style_ok = True
            block_box_header_ok = True
            block_summary_ok = True
            block_header_border_ok = True
            ticket_values_filled = True
            brand_values_filled = True
            box_numbers_prefixed = True
            added_cells_centered = True
            added_cells_bordered = True
            added_cells_font_ok = True
            added_rows_height_ok = True

            for block in block_reports:
                positions = block["positions"]
                detail_rows = list(range(positions["detail_start_row"], positions["detail_start_row"] + block["detail_row_count"]))
                styled_rows = detail_rows + [positions["summary_row"]]
                total_actual_rows += sum(
                    1
                    for row_idx in detail_rows
                    if any(not is_blank(written_sheet.cell(row=row_idx, column=col_idx).value) for col_idx in range(1, target_max_col + 1))
                )
                block_set_total = sum(
                    convert_numeric(written_sheet.cell(row=row_idx, column=template_selection.headers["套"]).value) or 0
                    for row_idx in detail_rows
                )
                block_carton_total = sum(
                    convert_numeric(written_sheet.cell(row=row_idx, column=template_selection.headers["箱数"]).value) or 0
                    for row_idx in detail_rows
                )
                total_actual_set += block_set_total
                total_actual_carton += block_carton_total
                summary_set_total = convert_numeric(
                    written_sheet.cell(row=positions["summary_row"], column=template_selection.headers["套"]).value
                )
                summary_carton_total = convert_numeric(
                    written_sheet.cell(row=positions["summary_row"], column=template_selection.headers["箱数"]).value
                )
                block_summary_ok = (
                    block_summary_ok
                    and block["expected_set_total"] == summary_set_total
                    and block["expected_carton_total"] == summary_carton_total
                )

                actual_title = written_sheet.cell(row=positions["title_row"], column=1).value
                block_title_ok = block_title_ok and str(actual_title) == block["title_text"]
                block_merge_ok = block_merge_ok and build_title_merge_range(positions["title_row"]) in {
                    str(cell_range) for cell_range in written_sheet.merged_cells.ranges
                }
                block_title_style_ok = (
                    block_title_style_ok
                    and written_sheet.cell(row=positions["title_row"], column=1).font.name == "宋体"
                    and int(round(written_sheet.cell(row=positions["title_row"], column=1).font.sz or 0)) == 20
                    and written_sheet.cell(row=positions["title_row"], column=1).alignment.horizontal == "center"
                    and written_sheet.cell(row=positions["title_row"], column=1).alignment.vertical == "center"
                )

                box_header_value = written_sheet.cell(row=positions["header_row"], column=template_selection.headers["箱号"]).value
                block_box_header_ok = block_box_header_ok and (
                    is_blank(block["fba_number"]) or str(block["fba_number"]) in str(box_header_value)
                )
                block_header_border_ok = block_header_border_ok and header_block_borders_are_complete(
                    written_sheet,
                    positions["header_row"],
                    positions["header_blank_row"],
                    target_max_col,
                )

                ticket_values_filled = ticket_values_filled and all(
                    re.fullmatch(r".+海运第\d+票", str(written_sheet.cell(row=row_idx, column=template_selection.headers["票数"]).value or ""))
                    for row_idx in detail_rows
                )
                brand_values_filled = brand_values_filled and all(
                    not is_blank(written_sheet.cell(row=row_idx, column=template_selection.headers["说明书/包装上-品牌"]).value)
                    for row_idx in detail_rows
                )
                box_numbers_prefixed = box_numbers_prefixed and all(
                    re.fullmatch(r"编号 \d+(?:-\d+)?", str(written_sheet.cell(row=row_idx, column=template_selection.headers["箱号"]).value or ""))
                    for row_idx in detail_rows
                )
                added_cells_centered = added_cells_centered and all(
                    written_sheet.cell(row=row_idx, column=col_idx).alignment.horizontal == "center"
                    and written_sheet.cell(row=row_idx, column=col_idx).alignment.vertical == "center"
                    for row_idx in styled_rows
                    for col_idx in range(1, target_max_col + 1)
                )
                added_cells_bordered = added_cells_bordered and all(
                    written_sheet.cell(row=row_idx, column=col_idx).border.left.style == "thin"
                    and written_sheet.cell(row=row_idx, column=col_idx).border.right.style == "thin"
                    and written_sheet.cell(row=row_idx, column=col_idx).border.top.style == "thin"
                    and written_sheet.cell(row=row_idx, column=col_idx).border.bottom.style == "thin"
                    for row_idx in styled_rows
                    for col_idx in range(1, target_max_col + 1)
                )
                added_cells_font_ok = added_cells_font_ok and all(
                    written_sheet.cell(row=row_idx, column=col_idx).font.name == "宋体"
                    and int(round((written_sheet.cell(row=row_idx, column=col_idx).font.sz or 0))) == 12
                    for row_idx in styled_rows
                    for col_idx in range(1, target_max_col + 1)
                )
                added_rows_height_ok = added_rows_height_ok and all(
                    int(round(written_sheet.row_dimensions[row_idx].height or 0)) == DATA_ROW_HEIGHT
                    for row_idx in styled_rows
                )
        finally:
            written_workbook.close()

        validations = {
            "sheet_name": {
                "expected": output_sheet_name,
                "actual": actual_sheet_title,
                "match": actual_sheet_title == output_sheet_name,
            },
            "sheet_count_is_one": actual_sheet_count == 1,
            "data_row_count": {
                "expected": total_expected_rows,
                "actual": total_actual_rows,
                "match": total_expected_rows == total_actual_rows,
            },
            "set_total": {
                "expected": total_expected_set,
                "actual": total_actual_set,
                "match": total_expected_set == total_actual_set,
            },
            "carton_total": {
                "expected": total_expected_carton,
                "actual": total_actual_carton,
                "match": total_expected_carton == total_actual_carton,
            },
            "summary_row_totals": block_summary_ok,
            "box_numbers_prefixed": box_numbers_prefixed,
            "ticket_values_filled": ticket_values_filled,
            "brand_values_filled": brand_values_filled,
            "all_titles_match": block_title_ok,
            "all_title_merges_ok": block_merge_ok,
            "all_title_styles_ok": block_title_style_ok,
            "all_box_headers_contain_fba": block_box_header_ok,
            "all_header_borders_complete": block_header_border_ok,
            "added_cells_centered": added_cells_centered,
            "added_cells_bordered": added_cells_bordered,
            "added_cells_songti_12": added_cells_font_ok,
            "added_rows_height_30": added_rows_height_ok,
        }

        if not validations["sheet_name"]["match"]:
            anomalies.append("工作表名称未按总票数命名")
        if not validations["sheet_count_is_one"]:
            anomalies.append("输出工作簿不是单工作表")
        if not validations["data_row_count"]["match"]:
            anomalies.append("数据行数校验不一致")
        if not validations["set_total"]["match"]:
            anomalies.append("套合计校验不一致")
        if not validations["carton_total"]["match"]:
            anomalies.append("箱数合计校验不一致")
        if not validations["summary_row_totals"]:
            anomalies.append("汇总行的套或箱数与明细合计不一致")
        if not validations["box_numbers_prefixed"]:
            anomalies.append("存在未按要求添加“编号”前缀的箱号")
        if not validations["ticket_values_filled"]:
            anomalies.append("票数未全部按“店铺简称+海运第x票”填写")
        if not validations["brand_values_filled"]:
            anomalies.append("说明书/包装上-品牌未全部根据SKU查表填写")
        if not validations["all_titles_match"]:
            anomalies.append("存在票块标题内容未按规则生成")
        if not validations["all_title_merges_ok"]:
            anomalies.append("存在票块标题未保持 A:M 合并")
        if not validations["all_title_styles_ok"]:
            anomalies.append("存在票块标题样式不是宋体20居中")
        if not validations["all_box_headers_contain_fba"]:
            anomalies.append("存在票块箱号表头未写入对应FBA号")
        if not validations["all_header_borders_complete"]:
            anomalies.append("存在票块表头边框不完整的情况")
        if not validations["added_cells_centered"]:
            anomalies.append("存在未居中的新增数据单元格")
        if not validations["added_cells_bordered"]:
            anomalies.append("存在未添加边框的新增数据单元格")
        if not validations["added_cells_songti_12"]:
            anomalies.append("存在未设置为宋体 12 号的新增数据单元格")
        if not validations["added_rows_height_30"]:
            anomalies.append("存在行高不是 30 磅的新增数据行")

        anomalies = dedupe_preserve_order(anomalies)

        report = {
            "files_read": {
                "rpa_docx": "rpa.docx",
                "template_workbook": "generated_in_code",
                "template_sheet": template_selection.sheet_name,
                "resource_dir": str(resource_dir),
                "source_dir": str(source_dir),
                "output_dir": str(output_dir),
                "source_workbooks": [path.name for path in source_files],
                "msku_mapping_workbook": msku_map_path.name,
                "msku_mapping_sheet": msku_selection.sheet_name,
                "store_detail_workbook": store_detail_path.name,
                "store_detail_sheet": store_selection.sheet_name,
            },
            "header_rows": {
                "template_header_row": template_selection.header_row,
                "template_data_start_row": detect_data_start_row(template_sheet, template_selection.header_row),
                "msku_mapping_header_row": msku_selection.header_row,
                "store_detail_header_row": store_selection.header_row,
            },
            "source_file_order": {
                "ticket_count": len(source_files),
                "all_source_files": [path.name for path in source_files],
            },
            "field_mapping": FIELD_MAPPING,
            "store_lookup_summary": {
                "resolved_store_sites": dedupe_preserve_order(all_store_sites),
                "resolved_store_shorts": dedupe_preserve_order(all_store_shorts),
                "resolved_store_names": dedupe_preserve_order(all_store_names),
                "resolved_mapped_lines": dedupe_preserve_order(all_mapped_lines),
                "rows": lookup_rows_for_report,
            },
            "block_reports": block_reports,
            "styling_updates": {
                "sheet_name": output_sheet_name,
                "box_value_prefix": "编号",
                "title_format": "装箱单（店铺简称-站点）海运第x票",
                "box_header_extra_line": "每票块显示对应FBA号",
                "added_data_font": "宋体 12",
                "added_row_height": DATA_ROW_HEIGHT,
                "summary_row_added": True,
            },
            "validations": validations,
            "anomalies": anomalies,
            "output_workbook": output_path.name,
        }

        report_path = output_path.with_name(f"{output_path.stem}_report.json")
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        report["report_file"] = report_path.name
        return report
    finally:
        template_workbook.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="领星 Excel 批量整理工具")
    parser.add_argument("--resource-dir", default=str(Path(__file__).resolve().parent))
    parser.add_argument("--source-dir", default=str(Path(__file__).resolve().parent))
    parser.add_argument("--output-dir", default=None)
    args = parser.parse_args()

    resource_dir = Path(args.resource_dir).resolve()
    source_dir = Path(args.source_dir).resolve()
    output_dir = Path(args.output_dir).resolve() if args.output_dir else source_dir
    report = process_workbooks(resource_dir, source_dir, output_dir)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
