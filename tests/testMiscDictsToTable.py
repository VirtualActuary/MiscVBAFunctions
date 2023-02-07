import unittest
from aa_py_xl import Table
from xlwings import Book
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book, vba_dict


class TestDictsToTable(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("DictsToTable_1"):
                func_dicts_to_table = book.macro("MiscDictsToTable.DictsToTable")
                func_col = book.macro("MiscCollectionCreate.Col")

                d1 = vba_dict({"a": 1, "b": 2})
                func_dicts_to_table(
                    func_col(d1, d1),
                    book.sheets["Sheet1"].range("A1"),
                    "Table1",
                    False,
                )

                table = Table.get_from_book(book, "Table1")

                self.assertEqual("$A$1:$B$3", table.range.address)
                self.assertEqual("Table1", table.name)
                self.assertEqual("Sheet1", table.sheet.name)
                self.assertEqual(
                    [["a", "b"], [1, 2], [1, 2]],
                    table.range.value,
                )

            with self.subTest("DictsToTable_2"):
                func_dicts_to_table = book.macro("MiscDictsToTable.DictsToTable")
                func_col = book.macro("MiscCollectionCreate.Col")

                func_dicts_to_table(
                    func_col(
                        vba_dict({"a": 1, "b": 2, "c": 3}),
                        vba_dict({"a": 10, "c": 20, "b": 30}),
                    ),
                    book.sheets["Sheet1"].range("D1"),
                    "Table2",
                    False,
                )

                table = Table.get_from_book(book, "Table2")

                self.assertEqual("$D$1:$F$3", table.range.address)
                self.assertEqual("Table2", table.name)
                self.assertEqual("Sheet1", table.sheet.name)
                self.assertEqual(
                    [["a", "b", "c"], [1, 2, 3], [10, 30, 20]],
                    table.range.value,
                )

            with self.subTest("Test_DictsToTable_fail_1"):
                func = book.macro(
                    "Test__Helper_MiscDictsToTable.Test_DictsToTable_fail_1"
                )
                func_col = book.macro("MiscCollectionCreate.Col")
                dict1 = vba_dict({"col1": 1, "col2": 2, "col3": 3, "col4": 30})
                dict2 = vba_dict({"col1": 10, "col2": 20, "col3": 30})

                table_dict = func_col(dict1, dict2)
                range_obj = book.sheets.active.range("A10")
                self.assertTrue((func(table_dict, range_obj)))

            with self.subTest("Test_DictsToTable_fail_2"):
                func = book.macro(
                    "Test__Helper_MiscDictsToTable.Test_DictsToTable_fail_2"
                )
                func_col = book.macro("MiscCollectionCreate.Col")
                dict1 = vba_dict({"col1": 1, "col2": 2, "col3": 3})
                dict2 = vba_dict({"col1": 10, "col2": 20, "col4": 30})

                table_dict = func_col(dict1, dict2)
                range_obj = book.sheets.active.range("A10")
                self.assertTrue((func(table_dict, range_obj)))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
