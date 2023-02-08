import unittest
from xlwings import Book
from locate import prepend_sys_path
with prepend_sys_path():
    from util import functions_book


class TestDictsToTable(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("Test_ExcelBook"):
                func = book.macro("Test__Helper_MiscExcel.Test_ExcelBook")
                self.assertTrue(func())

            with self.subTest("Test_ExcelBook_tempFile"):
                func = book.macro("Test__Helper_MiscExcel.Test_ExcelBook_tempFile")
                self.assertTrue(func())

            with self.subTest("Test_ExcelBook_tempFile_2"):
                func = book.macro("Test__Helper_MiscExcel.Test_ExcelBook_tempFile_2")
                self.assertTrue(func())

            with self.subTest("Test_ExcelBook_tempFile_fail"):
                func = book.macro("Test__Helper_MiscExcel.Test_ExcelBook_tempFile_fail")
                self.assertTrue(func())

            with self.subTest("Test_ExcelBook_tempFile_fail_2"):
                func = book.macro(
                    "Test__Helper_MiscExcel.Test_ExcelBook_tempFile_fail_2"
                )
                self.assertTrue(func())

            with self.subTest("Test_fail_ExcelBook"):
                func = book.macro("Test__Helper_MiscExcel.Test_fail_ExcelBook")
                self.assertTrue(func())

            with self.subTest("Test_fail_ExcelBook_2"):
                func = book.macro("Test__Helper_MiscExcel.Test_fail_ExcelBook_2")
                self.assertTrue(func())

            with self.subTest("Test_OpenWorkbook"):
                func = book.macro("Test__Helper_MiscExcel.Test_OpenWorkbook")
                self.assertTrue(func())

            with self.subTest("LastRow"):
                func = book.macro("Test__Helper_MiscExcel.Test_LastRow")
                self.assertTrue(func())

            with self.subTest("Test_LastColumn"):
                func = book.macro("Test__Helper_MiscExcel.Test_LastColumn")
                self.assertTrue(func())

            with self.subTest("Test_LastCell_1"):
                func = book.macro("Test__Helper_MiscExcel.Test_LastCell_1")
                self.assertTrue(func())

            with self.subTest("Test_LastCell_2"):
                func = book.macro("Test__Helper_MiscExcel.Test_LastCell_2")
                self.assertTrue(func())

            with self.subTest("Test_RelevantRange"):
                func = book.macro("Test__Helper_MiscExcel.Test_RelevantRange")
                self.assertTrue(func())

            with self.subTest("Test_RelevantRange2"):
                func = book.macro("Test__Helper_MiscExcel.Test_RelevantRange2")
                self.assertTrue(func())

            with self.subTest("SanitiseExcelName"):
                func = book.macro("MiscExcel.SanitiseExcelName")

                self.assertEqual("_1", func("1"))
                self.assertEqual("a_b", func("a b"))
                self.assertEqual(
                    "_____________________________",
                    func("- /*+=^!@#$%&?`~:;[](){}" "'|,<>"),
                )


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
