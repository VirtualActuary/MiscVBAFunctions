import unittest
from pathlib import Path
from aa_py_xl import Table
from locate import prepend_sys_path, this_dir
from xlwings import Book

with prepend_sys_path():
    from util import functions_book

base_dir = this_dir().joinpath("..")


class MiscCsv(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("Test_CsvToLO"):
                func_CreateTextFile = book.macro("MiscCSV.CsvToLO")

                func_CreateTextFile(
                    book.sheets[0].cells(10, 10),
                    str(Path(base_dir, r".\test_data\Csv\MiscCsv.csv").resolve()),
                    "MyTable",
                )
                table = Table.get_from_book(book, "MyTable")
                self.assertEqual("MyTable", table.name)
                self.assertEqual("1", table.range(1, 1).value)
                self.assertEqual(11, table.range(3, 1).value)


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
