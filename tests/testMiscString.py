import unittest
from xlwings import Book
from .util import functions_book


class TestMin(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("randomString"):
                func_randomString = book.macro("MiscString.randomString")
                self.assertEqual(4, len(func_randomString(4)))
                self.assertNotEqual(func_randomString(5), func_randomString(5))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
