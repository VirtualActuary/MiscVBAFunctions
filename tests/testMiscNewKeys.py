import unittest
from xlwings import Book
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscHasKey(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("Test_GetNewKey"):
                func = book.macro(
                    "Test__Helper_MiscNewKeys.Test_GetNewKey"
                )
                self.assertTrue(func())


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
