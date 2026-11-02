import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

import excel_to_iphone_vcf as app


class ContactValidationTests(unittest.TestCase):
    def test_required_headers_are_accepted(self):
        app.validate_headers(" 姓名 ", "电话")

    def test_wrong_headers_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "表头不正确"):
            app.validate_headers("客户", "邮箱")

    def test_non_phone_text_is_not_exportable(self):
        contacts = []
        app.add_contact_row(contacts, 2, "张三", "abc@example.com")
        self.assertEqual(contacts[0]["status"], app.INVALID_PHONE_STATUS)

    def test_common_phone_separators_are_normalized(self):
        contacts = []
        app.add_contact_row(contacts, 2, "张三", "+52 (55) 1234-5678")
        self.assertEqual(contacts[0]["phone"], "+525512345678")
        self.assertEqual(contacts[0]["status"], app.VALID_STATUS)

    def test_numeric_phone_with_zero_mask_recovers_leading_zero(self):
        phone, issue = app.excel_phone_value(123456789, "0000000000")
        self.assertEqual(phone, "0123456789")
        self.assertIsNone(issue)

    def test_unformatted_numeric_phone_is_blocked(self):
        phone, issue = app.excel_phone_value(123456789, "General")
        self.assertEqual(phone, "123456789")
        self.assertEqual(issue, app.NUMERIC_PHONE_STATUS)

    def test_scientific_notation_keeps_integer_zeroes(self):
        self.assertEqual(app.normalize_number_text("1E+10"), "10000000000")

    def test_xlsx_cell_types_are_handled_safely(self):
        from openpyxl import Workbook

        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["姓名", "电话"])
        sheet.append(["文本电话", "+52 (55) 1234-5678"])
        sheet.append(["错误电话", "abc@example.com"])
        sheet.append(["补零电话", 123456789])
        sheet["B4"].number_format = "0000000000"
        sheet.append(["普通数字", 123456789])

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "phones.xlsx"
            workbook.save(path)
            contacts = app.read_contacts(str(path))

        self.assertEqual(contacts[0]["phone"], "+525512345678")
        self.assertEqual(contacts[0]["status"], app.VALID_STATUS)
        self.assertEqual(contacts[1]["status"], app.INVALID_PHONE_STATUS)
        self.assertEqual(contacts[2]["phone"], "0123456789")
        self.assertEqual(contacts[2]["status"], app.VALID_STATUS)
        self.assertEqual(contacts[3]["status"], app.NUMERIC_PHONE_STATUS)


if __name__ == "__main__":
    unittest.main()
