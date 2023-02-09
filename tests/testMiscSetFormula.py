import unittest
from xlwings import Book
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscSetFormula(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("SetFormula"):
                func = book.macro("MiscSetFormula.SetFormula")
                func(book.sheets[0].range("B2"), "=1+2.1")
                self.assertEqual(3.1, book.sheets[0].range("B2").value)


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
