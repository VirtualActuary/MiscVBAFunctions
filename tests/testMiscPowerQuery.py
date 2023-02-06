import unittest
from xlwings import Book
from .util import functions_book


class TestDictsToTable(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("doesQueryExist"):
                func_doesQueryExist = book.macro("MiscPowerQuery.doesQueryExist")
                self.assertFalse(func_doesQueryExist("foo"))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
