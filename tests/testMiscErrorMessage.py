import unittest
from xlwings import Book
from locate import prepend_sys_path
with prepend_sys_path():
    from util import functions_book

class TestDictsToTable(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("ErrorMessage"):
                func_ErrorMessage = book.macro("MiscErrorMessage.ErrorMessage")

                self.assertEqual(
                    "This array is fixed or temporarily locked", func_ErrorMessage(10)
                )
                self.assertEqual(
                    "Out of memory: a fix is required before continuing",
                    func_ErrorMessage(7, "a fix is required before continuing"),
                )
                self.assertEqual("Unknown error", func_ErrorMessage(77))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
