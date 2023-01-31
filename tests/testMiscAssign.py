import unittest
from xlwings import Book
from .util import functions_book


class TestMin(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
