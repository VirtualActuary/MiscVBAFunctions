import unittest
from xlwings import Book
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book, vba_dict


class MiscHasKey(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("Test_HasKey_Collection"):
                func = book.macro("Test__Helper_MiscHasKey.Test_HasKey_Collection")
                self.assertTrue(func())

            with self.subTest("Test_HasKey_Workbook"):
                func = book.macro("Test__Helper_MiscHasKey.Test_HasKey_Workbook")
                self.assertTrue(func())

            with self.subTest("Test_HasKey_Dictionary"):
                func_col = book.macro("MiscCollectionCreate.Col")
                func_hasKey = book.macro("MiscHasKey.hasKey")

                d = vba_dict({"a": "foo", "b": func_col("x", "y", "z")})
                self.assertTrue(func_hasKey(d, "a"))
                self.assertTrue(func_hasKey(d, "b"))
                self.assertFalse(func_hasKey(d, "A"))

            with self.subTest("Test_HasKey_Dictionary_object"):
                func = book.macro("Test__Helper_MiscHasKey.Test_HasKey_Dictionary_object")
                self.assertTrue(func())

            with self.subTest("Test_HasKey_Dictionary_fail"):
                func = book.macro("Test__Helper_MiscHasKey.Test_HasKey_Dictionary_fail")
                self.assertTrue(func())


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
