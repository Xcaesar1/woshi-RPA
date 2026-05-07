import unittest

from lingxing_excel_processor import format_mul_warehouse_address


class UpsWarehouseAddressTests(unittest.TestCase):
    def test_address_keeps_commas_inside_parentheses_and_avoids_duplicate_fc(self) -> None:
        raw_address = "XPB2 - AFTX 3P FC (Fort Pierce, FL, US),4661 Apopka Logistics Pkwy,,Apopka,FL,32712,US,"

        formatted = format_mul_warehouse_address("XPB2", raw_address)

        self.assertEqual(
            formatted,
            "仓库地址：XPB2 - AFTX 3P FC (Fort Pierce, FL, US) - 4661 Apopka Logistics Pkwy 32712 - Apopka, FL - United States",
        )


if __name__ == "__main__":
    unittest.main()
