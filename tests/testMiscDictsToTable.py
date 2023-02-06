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


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
