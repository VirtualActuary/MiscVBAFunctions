import unittest

from .util import functions_book


class TestPath(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            func = book.macro("MiscPath.Path")

            with self.subTest("Toets een"):
                self.assertEqual(r"a\b\c", func("a", "b", "c"))

            with self.subTest("Toets twee"):
                self.assertEqual(r"a\b\c", func("a", "b", "c"))
                self.assertEqual(r"a\b\c", func("a", "b", "c"))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
