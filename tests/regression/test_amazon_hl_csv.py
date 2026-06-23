import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from lingxing_excel_processor import (
    classify_source_workbook,
    extract_detail_rows,
    locate_msku_mapping_file,
    process_workbooks,
    read_source_metadata,
)

from app.services.amazon_hl_csv_service import (
    convert_amazon_hl_csv_to_source_workbook,
    parse_amazon_hl_csv,
)
from app.services.workflow_service import validate_hl_upload_filename


AMAZON_HL_SAMPLE = '''"工作流程名称","wffec879c4-0f2d-42f5-adba-17db47eeaa3e"
"货件编号","FBA19GPGFQ5Q"
"货件名称","FBA STA (06/23/2026 08:10)-POC1"
"配送地址","FONTANA, CA"
"箱子数量","30"
"SKU 数量","1"
"商品数量","150"

"原厂包装发货（30 箱）"
"SKU","商品名称","ASIN","FNSKU","原厂包装模板名称","状况","预处理类型","包装箱重量（磅）","箱子长度（英寸）","箱子宽度（英寸）","箱子高度（英寸）","每箱件数","箱子总数","商品总数","箱号"
"B2207-BN","Roman Bathtub Faucet Deck Mount Tub Faucet for Bathroom Widespread Elegant Classic Spout with 3 Holes 2 Handle Valve Cartridge Included, Brush Nickel","B09VGDD6DF","X0036T55NF","B2207","新品","无需进行预处理","44.22","19.69","17.91","13.58","5","30","150","FBA19GPGFQ5QU000030,FBA19GPGFQ5QU000011,FBA19GPGFQ5QU000021,FBA19GPGFQ5QU000010,FBA19GPGFQ5QU000014,FBA19GPGFQ5QU000029,FBA19GPGFQ5QU000007,FBA19GPGFQ5QU000026,FBA19GPGFQ5QU000020,FBA19GPGFQ5QU000027,FBA19GPGFQ5QU000013,FBA19GPGFQ5QU000024,FBA19GPGFQ5QU000017,FBA19GPGFQ5QU000012,FBA19GPGFQ5QU000005,FBA19GPGFQ5QU000002,FBA19GPGFQ5QU000001,FBA19GPGFQ5QU000003,FBA19GPGFQ5QU000009,FBA19GPGFQ5QU000015,FBA19GPGFQ5QU000004,FBA19GPGFQ5QU000018,FBA19GPGFQ5QU000006,FBA19GPGFQ5QU000022,FBA19GPGFQ5QU000008,FBA19GPGFQ5QU000019,FBA19GPGFQ5QU000016,FBA19GPGFQ5QU000025,FBA19GPGFQ5QU000028,FBA19GPGFQ5QU000023"
'''


class AmazonHlCsvTests(unittest.TestCase):
    def write_sample(self, directory: Path, name: str = "FBA19GPGFQ5Q.csv") -> Path:
        path = directory / name
        path.write_text(AMAZON_HL_SAMPLE, encoding="utf-8-sig")
        return path

    def test_parse_amazon_hl_csv_extracts_metadata_and_detail_rows(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            csv_path = self.write_sample(Path(tmp))

            shipment = parse_amazon_hl_csv(csv_path)

            self.assertEqual(shipment.fba_code, "FBA19GPGFQ5Q")
            self.assertEqual(shipment.cargo_name, "FBA STA (06/23/2026 08:10)-POC1")
            self.assertEqual(shipment.destination, "FONTANA, CA")
            self.assertEqual(shipment.carton_count, 30)
            self.assertEqual(shipment.sku_count, 1)
            self.assertEqual(shipment.total_quantity, 150)
            self.assertEqual(len(shipment.items), 1)
            item = shipment.items[0]
            self.assertEqual(item.msku, "B2207-BN")
            self.assertEqual(item.factory_sku, "B2207")
            self.assertEqual(item.fnsku, "X0036T55NF")
            self.assertEqual(item.quantity_per_box, 5)
            self.assertEqual(item.carton_count, 30)
            self.assertEqual(item.total_quantity, 150)
            self.assertEqual(item.box_range, "1-30")

    def test_convert_amazon_hl_csv_creates_existing_one_sku_source_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            csv_path = self.write_sample(root)

            workbook_path, shipment = convert_amazon_hl_csv_to_source_workbook(
                csv_path,
                root / "source",
                resource_dir=Path.cwd(),
            )

            self.assertEqual(shipment.fba_code, "FBA19GPGFQ5Q")
            self.assertTrue(workbook_path.name.startswith("FBA19GPGFQ5Q"))
            source_info = classify_source_workbook(workbook_path)
            self.assertIsNotNone(source_info)
            self.assertEqual(source_info.format_type, "ONE_SKU")

            metadata = read_source_metadata(source_info)
            self.assertEqual(metadata["货件单号"], "FBA19GPGFQ5Q")
            self.assertEqual(metadata["货件名称"], "FBA STA (06/23/2026 08:10)-POC1")
            self.assertEqual(metadata["配送地址"], "FONTANA, CA")

            workbook = load_workbook(workbook_path, data_only=True)
            try:
                rows = extract_detail_rows(workbook[source_info.selection.sheet_name], source_info.selection)
                self.assertEqual(len(rows), 1)
                self.assertEqual(rows[0]["序号"], 1)
                self.assertEqual(rows[0]["MSKU"], "B2207-BN")
                self.assertEqual(rows[0]["SKU"], "B2207")
                self.assertEqual(rows[0]["FNSKU"], "X0036T55NF")
                self.assertEqual(rows[0]["品名"], "22系列三件套缸边浴缸龙头拉丝封油")
                self.assertEqual(rows[0]["品线"], "缸边浴缸")
                self.assertEqual(rows[0]["发货量"], 150)
                self.assertEqual(rows[0]["单箱数量"], 5)
                self.assertEqual(rows[0]["箱数"], 30)
                self.assertEqual(rows[0]["箱号"], "1-30")
            finally:
                workbook.close()

    def test_process_amazon_hl_csv_uses_mapping_table_product_name_and_line(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            csv_path = self.write_sample(root)
            source_dir = root / "source"
            output_dir = root / "output"
            convert_amazon_hl_csv_to_source_workbook(csv_path, source_dir, resource_dir=Path.cwd())

            report = process_workbooks(Path.cwd(), source_dir, output_dir)

            output_path = output_dir / report["output_workbook"]
            workbook = load_workbook(output_path, data_only=True)
            try:
                worksheet = workbook[workbook.sheetnames[0]]
                header_row = None
                headers = {}
                for row_idx in range(1, worksheet.max_row + 1):
                    row_headers = {
                        str(worksheet.cell(row=row_idx, column=col_idx).value).strip(): col_idx
                        for col_idx in range(1, worksheet.max_column + 1)
                        if worksheet.cell(row=row_idx, column=col_idx).value is not None
                    }
                    if "品名" in row_headers and "备注/品线" in row_headers:
                        header_row = row_idx
                        headers = row_headers
                        break

                self.assertIsNotNone(header_row)
                data_row = header_row + 1
                self.assertEqual(worksheet.cell(row=data_row, column=headers["品名"]).value, "22系列三件套缸边浴缸龙头拉丝封油")
                self.assertEqual(worksheet.cell(row=data_row, column=headers["备注/品线"]).value, "缸边浴缸")
            finally:
                workbook.close()

    def test_hl_upload_filename_accepts_only_csv(self) -> None:
        validate_hl_upload_filename("FBA19GPGFQ5Q.csv")

        for name in ["fba_manifest.txt", "fba_manifest.xlsx", "FBA19GPGFQ5Q"]:
            with self.subTest(name=name):
                with self.assertRaises(ValueError):
                    validate_hl_upload_filename(name)

    def test_msku_mapping_file_prefers_updated_product_code_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            old_mapping = root / "MSKU对应品线表.xlsx"
            updated_mapping = root / "MSKU对应品线表_已更新产品编码_20260508_193206.xlsx"
            old_mapping.touch()
            updated_mapping.touch()

            self.assertEqual(locate_msku_mapping_file(root), updated_mapping)


if __name__ == "__main__":
    unittest.main()
