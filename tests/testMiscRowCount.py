import unittest
from xlwings import Book
from locate import prepend_sys_path
from locate import this_dir
from aa_py_xl import excel

with prepend_sys_path():
    from util import functions_book


class MiscRowCount(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            func_ActiveRowsDown = book.macro("MiscRowCount.ActiveRowsDown")

            repo_path = this_dir().parent
            repo_path = repo_path.joinpath(
                r"test_data\MiscRowCount\ActiveRowsDown.xlsx"
            )
            book_temp: Book
            with excel(
                path=repo_path,
                save=False,
                quiet=True,
                close_book=True,
                close_excel=True,
                must_exist=True,
                read_only=True,
                events=False,
            ) as book_temp:

                with self.subTest("Test_NoFilter"):
                    range_obj = book_temp.sheets["Sheet1"].range("B2")
                    num_rows = func_ActiveRowsDown(range_obj)
                    self.assertEqual(5, num_rows)

                with self.subTest("Test_UnusedFilter"):
                    range_obj = book_temp.sheets["Sheet2"].range("B2")
                    num_rows = func_ActiveRowsDown(range_obj)
                    self.assertEqual(5, num_rows)

                with self.subTest("Test_FilterInMiddle"):
                    range_obj = book_temp.sheets["Sheet3"].range("C2")
                    num_rows = func_ActiveRowsDown(range_obj)
                    self.assertEqual(4, num_rows)

                with self.subTest("Test_FilterAtEnd"):
                    range_obj = book_temp.sheets["Sheet4"].range("D2")
                    num_rows = func_ActiveRowsDown(range_obj)
                    self.assertEqual(6, num_rows)

                with self.subTest("Test_FilterAtEnd_2"):
                    range_obj = book_temp.sheets["Sheet4"].range("C2:D3")
                    num_rows = func_ActiveRowsDown(range_obj)
                    self.assertEqual(4, num_rows)


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
