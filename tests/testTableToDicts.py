import unittest
from xlwings import Book
from locate import this_dir, prepend_sys_path
from aa_py_xl import excel

with prepend_sys_path():
    from util import functions_book, vba_dict


class TestDictsToTable(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            repo_path = this_dir().parent
            repo_path = repo_path.joinpath(
                r"test_data\MiscTableToDicts\MiscTableToDicts.xlsx"
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
            ) as book_temp:
                with self.subTest("TableLookupValue"):
                    func_TableLookupValue = book.macro(
                        "MiscTableToDicts.TableLookupValue"
                    )
                    func_col = book.macro("MiscCollectionCreate.Col")
                    func_dicti = book.macro("MiscDictionaryCreate.dicti")

                    Table = func_col(
                        func_dicti("a", 1, "b", 2, "c", 5),
                        func_dicti("a", 3, "b", 4, "c", 6),
                        func_dicti("a", "foo", "b", "bar"),
                    )

                    self.assertEqual(
                        6,
                        func_TableLookupValue(
                            Table, func_col("a", "b"), func_col(3, 4), "c"
                        ),
                    )
                    self.assertEqual(
                        "foo",
                        func_TableLookupValue(
                            Table, func_col("a", "b"), func_col(3, 400), "c", "foo"
                        ),
                    )

                with self.subTest("TableDictToArray"):
                    func_TableDictToArray = book.macro(
                        "MiscTableToDicts.TableDictToArray"
                    )
                    func_col = book.macro("MiscCollectionCreate.Col")
                    func_dict = book.macro("MiscDictionaryCreate.dict")

                    col1 = func_col(
                        func_dict("a", 1, "b", 2), func_dict("b", 11, "a", 12)
                    )
                    arr = func_TableDictToArray(col1)

                    self.assertEqual("a", arr[0][0])
                    self.assertEqual("b", arr[0][1])
                    self.assertEqual(1, arr[1][0])
                    self.assertEqual(2, arr[1][1])
                    self.assertEqual(12, arr[2][0])
                    self.assertEqual(11, arr[2][1])


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
