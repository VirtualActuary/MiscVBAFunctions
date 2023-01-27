import unittest

from .util import functions_book


class TestMin(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            func_min = book.macro("MiscCollection.Min")
            func_col = book.macro("MiscCollectionCreate.Col")

            self.assertEqual(1, func_min(func_col(1, 2, 3)))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
