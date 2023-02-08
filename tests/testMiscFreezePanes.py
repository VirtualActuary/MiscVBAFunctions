import unittest
from xlwings import Book
from locate import prepend_sys_path, this_dir
from aa_py_xl import excel_app

with prepend_sys_path():
    from util import functions_book, extra_book


class TestDictsToTable(unittest.TestCase):
    def test_1(self) -> None:
        with excel_app(True, True) as app:
            book: Book
            with functions_book(app) as book:
                repo_path = this_dir().parent
                repo_path = repo_path.joinpath(
                    r"test_data\MiscFreezePanes\MiscFreezePanes.xlsx"
                )
                book_extra: Book
                with extra_book(app, repo_path) as book_extra:

                    with self.subTest("Test_FreezePanes"):
                        func = book.macro(
                            "Test__Helper_MiscFreezePanes.Test_FreezePanes"
                        )
                        self.assertTrue(func(book_extra.sheets[0].range("D6")))

                    with self.subTest("Test_UnFreezePanes"):
                        func = book.macro(
                            "Test__Helper_MiscFreezePanes.Test_UnFreezePanes"
                        )
                        self.assertTrue(func(book_extra.sheets[0].range("D6")))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
