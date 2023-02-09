import shutil
import unittest
from pathlib import Path

from xlwings import Book
import os
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscOs(unittest.TestCase):
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

            with self.subTest("CreateFolders"):
                func_CreateFolders = book.macro("MiscOs.CreateFolders")
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
                    func_CreateFolders(Dir)
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

            with self.subTest("Test_is64BitXl"):
                func = book.macro("Test__Helper_MiscOs.Test_is64BitXl")
                self.assertTrue(func())


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
