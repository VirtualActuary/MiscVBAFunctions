import unittest
from xlwings import Book
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscHasKey(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("Test_GetNewKey1"):
                func_GetNewKey = book.macro("Test__Helper_MiscNewKeys.Test_GetNewKey1")
                self.assertTrue(func_GetNewKey())

            with self.subTest("Test_GetNewKey2"):
                func_GetNewKey = book.macro("Test__Helper_MiscNewKeys.Test_GetNewKey2")
                self.assertTrue(func_GetNewKey())


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
