import unittest
from xlwings import Book
from locate import prepend_sys_path, this_dir
from aa_py_xl import excel_app

with prepend_sys_path():
    from util import functions_book, extra_book


class MiscGroupOnIndentations(unittest.TestCase):
    def test_1(self) -> None:
        with excel_app(True, True) as app:
            book: Book
            with functions_book(app=app) as book:
                repo_path = this_dir().parent
                repo_path = repo_path.joinpath(
                    r"test_data\MiscGroupOnIndentations\MiscGroupOnIndentations.xlsx"
                )
                book_extra: Book
                with extra_book(app, repo_path) as book_extra:
                    with self.subTest("TestGroupOnIndentationsRows"):
                        func = book.macro(
                            "Test__Helper_MiscGroupOnIndent.TestGroupOnIndentationsRows"
                        )
                        self.assertTrue(func(book_extra))

                    with self.subTest("TestGroupOnIndentationsColumns"):
                        func = book.macro(
                            "Test__Helper_MiscGroupOnIndent.TestGroupOnIndentationsColumns"
                        )
                        self.assertTrue(func(book_extra))

                    with self.subTest("TestUnGroupOnIndentationsRow"):
                        func = book.macro(
                            "Test__Helper_MiscGroupOnIndent.TestUnGroupOnIndentationsRow"
                        )
                        self.assertTrue(func(book_extra))

                    with self.subTest("TestUnGroupOnIndentationsCol"):
                        func = book.macro(
                            "Test__Helper_MiscGroupOnIndent.TestUnGroupOnIndentationsCol"
                        )
                        self.assertTrue(func(book_extra))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
