import unittest
from xlwings import Book
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscAssign(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("MiscAssign_variant"):
                func_assign = book.macro("MiscAssign.assign")
                x = 1
                self.assertEqual(5, func_assign(x, 5))
                x = func_assign(x, 5)
                self.assertEqual(5, x)

                self.assertEqual(1.4, func_assign(x, 1.4))
                x = func_assign(x, 1.4)
                self.assertEqual(1.4, x)

            with self.subTest("MiscAssign_object"):
                func_col = book.macro("MiscCollectionCreate.Col")
                func = book.macro(
                    "Test__Helper_MiscAssign.Test_MiscAssign_object"
                )
                self.assertTrue(func(func_col(4, 5, 6)))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
