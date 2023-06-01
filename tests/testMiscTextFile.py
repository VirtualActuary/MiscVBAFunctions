import unittest
from pathlib import Path
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscTextFile(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            with self.subTest("CreateTextFile"):
                func_CreateTextFile = book.macro("MiscTextFile.CreateTextFile")
                input_text = "my test text."
                file_path = str(
                    Path(book.api.Path, r".\test_data\MiscCreateTextFile\test.txt")
                )
                func_CreateTextFile(input_text, file_path)

                with open(Path(file_path).resolve()) as file:
                    lines = [line.rstrip() for line in file]

                self.assertEqual("my test text.", lines[0])

            with self.subTest("ReadTextFile"):
                func_CreateTextFile = book.macro("MiscTextFile.CreateTextFile")
                func_ReadTextFile = book.macro("MiscTextFile.ReadTextFile")
                input_text = "my test text."
                file_path = str(
                    Path(book.api.Path, r".\test_data\MiscCreateTextFile\test.txt")
                )
                func_CreateTextFile(input_text, file_path)
                self.assertEqual("my test text.", func_ReadTextFile(file_path).strip())


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
