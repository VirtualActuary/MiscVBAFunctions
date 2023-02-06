import os
import unittest

import locate
from xlwings import Book
from pathlib import Path
from locate import prepend_sys_path
base_dir = locate.this_dir().joinpath("..")

with prepend_sys_path():
    from util import functions_book


class TestDictsToTable(unittest.TestCase):
    def test_glob(self) -> None:
        book: Book
        with functions_book() as book:
            func_Glob = book.macro("MiscGlob.Glob")
            func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")
            base_path = str(Path(base_dir, r".\test_data\GetAllFiles").resolve())
            test_path = str(Path(base_dir, r".\test_data\GetAllFiles").resolve())
            Path(os.environ["temp"], "output.txt").write_text(base_path)

            with self.subTest("Glob_simple_1"):
                arr = func_Col_to_arr(func_Glob(test_path, "folder1"))
                self.assertEqual(1, len(arr))
                self.assertEqual(base_path + r"\folder1", arr[0])

            with self.subTest("Glob_simple_2"):
                arr = func_Col_to_arr(func_Glob(test_path, "folder[1-9]"))
                self.assertEqual(2, len(arr))
                self.assertEqual(base_path + r"\folder1", arr[0])
                self.assertEqual(base_path + r"\folder2", arr[1])

            with self.subTest("Glob_simple_3"):
                arr = func_Col_to_arr(func_Glob(test_path, "folder?"))
                self.assertEqual(2, len(arr))
                self.assertEqual(base_path + r"\folder1", arr[0])
                self.assertEqual(base_path + r"\folder2", arr[1])

            with self.subTest("Glob_asterisk_1"):
                arr = func_Col_to_arr(func_Glob(test_path, "*"))
                self.assertEqual(3, len(arr))
                self.assertEqual(base_path + r"\empty file.txt", arr[0])
                self.assertEqual(base_path + r"\folder1", arr[1])
                self.assertEqual(base_path + r"\folder2", arr[2])

            with self.subTest("Glob_asterisk_2"):
                arr = func_Col_to_arr(func_Glob(test_path + r"\folder1\folder1", "*"))
                self.assertEqual(1, len(arr))
                self.assertEqual(
                    base_path + r"\folder1\folder1\empty file.xlsx", arr[0]
                )

            with self.subTest("Glob_asterisk_3"):
                arr = func_Col_to_arr(func_Glob(test_path, r"*\*er[1-9]\*"))
                self.assertEqual(2, len(arr))
                self.assertEqual(
                    base_path + r"\folder1\folder1\empty file.xlsx", arr[0]
                )
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[1])

            with self.subTest("Glob_recursion_x1_1"):
                arr = func_Col_to_arr(func_Glob(test_path, r"folder1\**"))
                self.assertEqual(2, len(arr))
                self.assertEqual(base_path + r"\folder1", arr[0])
                self.assertEqual(base_path + r"\folder1\folder1", arr[1])

            with self.subTest("Glob_recursion_x1_2"):
                arr = func_Col_to_arr(func_Glob(test_path, r"**\*er[1-9]\*.xlsx"))
                self.assertEqual(1, len(arr))
                self.assertEqual(
                    base_path + r"\folder1\folder1\empty file.xlsx", arr[0]
                )

            with self.subTest("Glob_recursion_x1_3"):
                arr = func_Col_to_arr(func_Glob(test_path, r"*\*er[1-9]\**"))
                self.assertEqual(3, len(arr))
                self.assertEqual(base_path + r"\folder1\folder1", arr[0])
                self.assertEqual(base_path + r"\folder2\folder1", arr[1])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[2])

            with self.subTest("Glob_recursion_x1_4"):
                arr = func_Col_to_arr(func_Glob(test_path, r"**\*"))
                self.assertEqual(10, len(arr))
                self.assertEqual(base_path + r"\empty file.txt", arr[0])
                self.assertEqual(base_path + r"\folder1", arr[1])
                self.assertEqual(base_path + r"\folder1\empty file.txt", arr[2])
                self.assertEqual(base_path + r"\folder1\folder1", arr[3])
                self.assertEqual(
                    base_path + r"\folder1\folder1\empty file.xlsx", arr[4]
                )
                self.assertEqual(base_path + r"\folder2", arr[5])
                self.assertEqual(base_path + r"\folder2\empty file.docx", arr[6])
                self.assertEqual(base_path + r"\folder2\folder1", arr[7])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[8])
                self.assertEqual(
                    base_path + r"\folder2\folder1\folder1\empty file.txt", arr[9]
                )

            with self.subTest("Glob_recursion_x1_5"):
                arr = func_Col_to_arr(func_Glob(test_path, r"**"))
                self.assertEqual(6, len(arr))
                self.assertEqual(base_path + r"", arr[0])
                self.assertEqual(base_path + r"\folder1", arr[1])
                self.assertEqual(base_path + r"\folder1\folder1", arr[2])
                self.assertEqual(base_path + r"\folder2", arr[3])
                self.assertEqual(base_path + r"\folder2\folder1", arr[4])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[5])

            with self.subTest("Glob_recursion_x1_6"):
                arr = func_Col_to_arr(func_Glob(test_path, r"*\folder1\**\*.xlsx"))
                self.assertEqual(1, len(arr))
                self.assertEqual(
                    base_path + r"\folder1\folder1\empty file.xlsx", arr[0]
                )

            with self.subTest("Glob_recursion_multiple_1"):
                arr = func_Col_to_arr(func_Glob(test_path, r"**\**"))
                self.assertEqual(6, len(arr))
                self.assertEqual(base_path + r"", arr[0])
                self.assertEqual(base_path + r"\folder1", arr[1])
                self.assertEqual(base_path + r"\folder1\folder1", arr[2])
                self.assertEqual(base_path + r"\folder2", arr[3])
                self.assertEqual(base_path + r"\folder2\folder1", arr[4])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[5])

            with self.subTest("Glob_recursion_multiple_2"):
                arr = func_Col_to_arr(func_Glob(test_path, r"**\folder1\**"))
                self.assertEqual(4, len(arr))
                self.assertEqual(base_path + r"\folder1", arr[0])
                self.assertEqual(base_path + r"\folder1\folder1", arr[1])
                self.assertEqual(base_path + r"\folder2\folder1", arr[2])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[3])

            with self.subTest("Glob_recursion_multiple_3"):
                arr = func_Col_to_arr(func_Glob(test_path, r"**\folder1\**\*"))
                self.assertEqual(5, len(arr))
                self.assertEqual(base_path + r"\folder1\empty file.txt", arr[0])
                self.assertEqual(base_path + r"\folder1\folder1", arr[1])
                self.assertEqual(
                    base_path + r"\folder1\folder1\empty file.xlsx", arr[2]
                )
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[3])
                self.assertEqual(
                    base_path + r"\folder2\folder1\folder1\empty file.txt", arr[4]
                )

    def test_rglob(self) -> None:
        book: Book
        with functions_book() as book:
            func_RGlob = book.macro("MiscGlob.RGlob")
            func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")
            base_path = str(Path(base_dir, r".\test_data\GetAllFiles").resolve())
            test_path = str(Path(base_dir, r".\test_data\GetAllFiles").resolve())

            with self.subTest("RGlob_simple_1"):
                arr = func_Col_to_arr(func_RGlob(test_path, r"folder1"))
                self.assertEqual(4, len(arr))
                self.assertEqual(base_path + r"\folder1", arr[0])
                self.assertEqual(base_path + r"\folder1\folder1", arr[1])
                self.assertEqual(base_path + r"\folder2\folder1", arr[2])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[3])

            with self.subTest("RGlob_simple_2"):
                arr = func_Col_to_arr(func_RGlob(test_path, r"*.xlsx"))
                self.assertEqual(1, len(arr))
                self.assertEqual(
                    base_path + r"\folder1\folder1\empty file.xlsx", arr[0]
                )

            with self.subTest("RGlob_simple_3"):
                arr = func_Col_to_arr(func_RGlob(test_path, r"*.txt"))
                self.assertEqual(3, len(arr))
                self.assertEqual(base_path + r"\empty file.txt", arr[0])
                self.assertEqual(base_path + r"\folder1\empty file.txt", arr[1])
                self.assertEqual(
                    base_path + r"\folder2\folder1\folder1\empty file.txt", arr[2]
                )

            with self.subTest("RGlob_simple_4"):
                arr = func_Col_to_arr(func_RGlob(test_path, r"*1*"))
                self.assertEqual(4, len(arr))
                self.assertEqual(base_path + r"\folder1", arr[0])
                self.assertEqual(base_path + r"\folder1\folder1", arr[1])
                self.assertEqual(base_path + r"\folder2\folder1", arr[2])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[3])

            with self.subTest("RGlob_simple_5"):
                arr = func_Col_to_arr(func_RGlob(test_path, r"folder[1-9]"))
                self.assertEqual(5, len(arr))
                self.assertEqual(base_path + r"\folder1", arr[0])
                self.assertEqual(base_path + r"\folder1\folder1", arr[1])
                self.assertEqual(base_path + r"\folder2", arr[2])
                self.assertEqual(base_path + r"\folder2\folder1", arr[3])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[4])

            with self.subTest("RGlob_simple_6"):
                arr = func_Col_to_arr(func_RGlob(test_path, r"folder?"))
                self.assertEqual(5, len(arr))
                self.assertEqual(base_path + r"\folder1", arr[0])
                self.assertEqual(base_path + r"\folder1\folder1", arr[1])
                self.assertEqual(base_path + r"\folder2", arr[2])
                self.assertEqual(base_path + r"\folder2\folder1", arr[3])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[4])

            with self.subTest("RGlob_asterisk_1"):
                arr = func_Col_to_arr(func_RGlob(test_path + r"\folder2", r"*"))
                self.assertEqual(4, len(arr))
                self.assertEqual(base_path + r"\folder2\empty file.docx", arr[0])
                self.assertEqual(base_path + r"\folder2\folder1", arr[1])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[2])
                self.assertEqual(
                    base_path + r"\folder2\folder1\folder1\empty file.txt", arr[3]
                )

            with self.subTest("RGlob_asterisk_2"):
                arr = func_Col_to_arr(func_RGlob(test_path, r""))
                self.assertEqual(6, len(arr))
                self.assertEqual(base_path + r"", arr[0])
                self.assertEqual(base_path + r"\folder1", arr[1])
                self.assertEqual(base_path + r"\folder1\folder1", arr[2])
                self.assertEqual(base_path + r"\folder2", arr[3])
                self.assertEqual(base_path + r"\folder2\folder1", arr[4])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[5])

            with self.subTest("RGlob_asterisk_3"):
                arr = func_Col_to_arr(func_RGlob(test_path, r"**"))
                self.assertEqual(6, len(arr))
                self.assertEqual(base_path + r"", arr[0])
                self.assertEqual(base_path + r"\folder1", arr[1])
                self.assertEqual(base_path + r"\folder1\folder1", arr[2])
                self.assertEqual(base_path + r"\folder2", arr[3])
                self.assertEqual(base_path + r"\folder2\folder1", arr[4])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[5])

            with self.subTest("RGlob_asterisk_4"):
                arr = func_Col_to_arr(func_RGlob(test_path, r"*"))
                self.assertEqual(10, len(arr))
                self.assertEqual(base_path + r"\empty file.txt", arr[0])
                self.assertEqual(base_path + r"\folder1", arr[1])
                self.assertEqual(base_path + r"\folder1\empty file.txt", arr[2])
                self.assertEqual(base_path + r"\folder1\folder1", arr[3])
                self.assertEqual(
                    base_path + r"\folder1\folder1\empty file.xlsx", arr[4]
                )
                self.assertEqual(base_path + r"\folder2", arr[5])
                self.assertEqual(base_path + r"\folder2\empty file.docx", arr[6])
                self.assertEqual(base_path + r"\folder2\folder1", arr[7])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[8])
                self.assertEqual(
                    base_path + r"\folder2\folder1\folder1\empty file.txt", arr[9]
                )

            with self.subTest("RGlob_asterisk_5"):
                arr = func_Col_to_arr(func_RGlob(test_path, r"*\*er[1-9]\*"))
                self.assertEqual(3, len(arr))
                self.assertEqual(
                    base_path + r"\folder1\folder1\empty file.xlsx", arr[0]
                )
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[1])
                self.assertEqual(
                    base_path + r"\folder2\folder1\folder1\empty file.txt", arr[2]
                )

            with self.subTest("RGlob_recursive"):
                arr = func_Col_to_arr(func_RGlob(test_path, r"folder2\**"))
                self.assertEqual(3, len(arr))
                self.assertEqual(base_path + r"\folder2", arr[0])
                self.assertEqual(base_path + r"\folder2\folder1", arr[1])
                self.assertEqual(base_path + r"\folder2\folder1\folder1", arr[2])


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
