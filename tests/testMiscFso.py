import unittest
from pathlib import Path
from xlwings import Book
from locate import prepend_sys_path, this_dir

with prepend_sys_path():
    from util import functions_book


class TestDictsToTable(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            func = book.macro("Test__Helper_MiscFso.Test_GetAllFilesRecursive")

            self.assertTrue(
                func(str(Path(this_dir().parent, r"test_data\GetAllFiles").absolute()))
            )


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
