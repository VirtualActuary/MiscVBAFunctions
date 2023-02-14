import unittest
from xlwings import Book
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscExcel(unittest.TestCase):
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

            with self.subTest("RenameSheet"):

                func_RenameSheet = book.macro("MiscExcel.RenameSheet")
                func_RenameSheet(book.sheets[0], "foo")
                self.assertEqual("foo", book.sheets[0].name)
                func_RenameSheet(book.sheets[0], "Sheet1")

            with self.subTest("renameSheet_2"):
                func_RenameSheet = book.macro("MiscExcel.RenameSheet")
                book.sheets.add()
                func_RenameSheet(book.sheets[0], "foo")
                func_RenameSheet(book.sheets[1], "foo")
                self.assertEqual("foo", book.sheets[0].name)
                self.assertEqual("foo (1)", book.sheets[1].name)
                func_RenameSheet(book.sheets[0], "Sheet1")
                func_RenameSheet(book.sheets[1], "Sheet2")

            with self.subTest("renameSheet_string_1"):
                func_RenameSheet(book.sheets[1], "temp")
                func_RenameSheet("temp", "Bar")
                self.assertEqual("Bar", book.sheets[1].name)

            with self.subTest("renameSheet_string_2"):
                func_RenameSheet(book.sheets[1], "temp")
                func_RenameSheet("temp", "Bar")
                func_RenameSheet("Bar", "Baz")
                self.assertEqual("Baz", book.sheets[1].name)

            with self.subTest("renameSheet_string_fail"):
                func = book.macro("Test__Helper_MiscExcel.Test_RenameSheet_fail")
                self.assertTrue(func())

            with self.subTest("Test_AddWS"):
                func = book.macro("Test__Helper_MiscExcel.Test_AddWS")
                self.assertTrue(func())

            with self.subTest("Test_DeleteSheet"):
                func_ContainsSheet = book.macro("MiscExcel.ContainsSheet")
                func_DeleteSheet = book.macro("MiscExcel.DeleteSheet")
                book.sheets.add("NewSheet")
                self.assertTrue((func_ContainsSheet("NewSheet")))
                func_DeleteSheet("NewSheet")
                self.assertFalse((func_ContainsSheet("NewSheet")))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
