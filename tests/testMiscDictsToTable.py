import unittest

from aa_py_xl import Table
from xlwings import Book

from .util import functions_book, vba_dict


class TestDictsToTable(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            func_dicts_to_table = book.macro("MiscDictsToTable.DictsToTable")
            func_col = book.macro("MiscCollectionCreate.Col")

            func_dicts_to_table(
                func_col(
                    vba_dict({"a": 1, "b": 2}),
                    vba_dict({"a": 3, "b": 4}),
                    vba_dict({"a": 5, "b": 6}),
                ),
                book.sheets["Sheet1"].range("A1"),
                "Table1",
                False,
            )

            table = Table.get_from_book(book, "Table1")

            self.assertEqual("$A$1:$B$4", table.range.address)
            self.assertEqual("Table1", table.name)
            self.assertEqual("Sheet1", table.sheet.name)
            self.assertEqual(
                [["a", "b"], [1, 2], [3, 4], [5, 6]],
                table.range.value,
            )


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
