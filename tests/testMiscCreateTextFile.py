import unittest
from pathlib import Path
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscCreateTextFile(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            with self.subTest("CreateTextFile"):
                func_CreateTextFile = book.macro("MiscCreateTextFile.CreateTextFile")
                func_EvalPath = book.macro("MiscPath.EvalPath")
                inputText = "my test text."
                FilePath = func_EvalPath(r".\test_data\MiscCreateTextFile\test.txt")
                func_CreateTextFile(inputText, FilePath)

                with open(Path(FilePath).resolve()) as file:
                    lines = [line.rstrip() for line in file]

                self.assertEqual("my test text.", lines[0])


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
