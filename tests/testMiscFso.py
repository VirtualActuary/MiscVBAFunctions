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
            look_in_path = Path(this_dir().parent, r"test_data\GetAllFiles")

            self.assertEqual(
                [
                    "empty file.txt",
                    "folder1/empty file.txt",
                    "folder1/folder1/empty file.xlsx",
                    "folder2/empty file.docx",
                    "folder2/folder1/folder1/empty file.txt",
                ],
                [
                    Path(p).relative_to(look_in_path).as_posix()
                    for p in book.macro(
                        "Test__Helper_MiscFso.Test_GetAllFilesRecursive"
                    )(str(look_in_path.absolute()))
                ],
            )


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
