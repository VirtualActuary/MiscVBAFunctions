import shutil
import unittest
from pathlib import Path

from xlwings import Book
from .util import functions_book
import os


class TestDictsToTable(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("ExpandEnvironmentalVariables"):
                func_ExpandEnvironmentalVariables = book.macro(
                    "MiscOs.ExpandEnvironmentalVariables"
                )

                self.assertEqual(
                    os.environ["windir"], func_ExpandEnvironmentalVariables("%windir%")
                )
                self.assertEqual(
                    os.environ["username"],
                    func_ExpandEnvironmentalVariables("%username%"),
                )
                self.assertEqual(
                    os.environ["windir"]
                    + "\\%foo\\bar%\\%username\\"
                    + os.environ["username"],
                    func_ExpandEnvironmentalVariables(
                        r"%windir%\%foo\bar%\%username\%username%"
                    ),
                )

            with self.subTest("MakeDirs"):
                func_MakeDirs = book.macro("MiscOs.MakeDirs")
                func_ExpandEnvironmentalVariables = book.macro(
                    "MiscOs.ExpandEnvironmentalVariables"
                )
                Dir = (
                    Path(
                        func_ExpandEnvironmentalVariables("%temp%"),
                        "MakeDirs_folder1",
                        "folder2",
                        "folder3",
                    )
                    .resolve()
                    .__str__()
                )
                try:
                    func_MakeDirs(Dir)
                    self.assertTrue(Path(Dir).is_dir())
                finally:
                    Dir = (
                        Path(
                            func_ExpandEnvironmentalVariables("%temp%"),
                            "MakeDirs_folder1",
                        )
                        .resolve()
                        .__str__()
                    )
                    if Path(Dir).is_dir():
                        shutil.rmtree(Dir)

            with self.subTest("RunShell"):
                func_RunShell = book.macro("MiscOs.RunShell")
                self.assertEqual(0, func_RunShell("cmd /c echo hello", True))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )