import unittest
from xlwings import Book
from locate import this_dir, prepend_sys_path
from aa_py_xl import excel_app

with prepend_sys_path():
    from util import functions_book, extra_book


class MiscTableToDicts(unittest.TestCase):
    def test_1(self) -> None:
        with excel_app(True, True) as app:
            book: Book
            with functions_book(app) as book:
                repo_path = this_dir().parent
                repo_path = repo_path.joinpath(
                    r"test_data\MiscTableToDicts\MiscTableToDicts.xlsx"
                )
                book_extra: Book
                with extra_book(app, repo_path) as book_extra:

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

                    with self.subTest("TestListObjectsToDicts1"):
                        func_TableToDicts = book.macro("MiscTableToDicts.TableToDicts")
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestListObjectsToDicts1"
                        )

                        table_dicts = func_TableToDicts("ListObject1", book_extra)
                        self.assertTrue(func(table_dicts))

                    with self.subTest("TestListObjectsToDicts2"):
                        func_TableToDicts = book.macro("MiscTableToDicts.TableToDicts")
                        func_col = book.macro("MiscCollectionCreate.Col")
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestListObjectsToDicts2"
                        )

                        table_dicts = func_TableToDicts(
                            "ListObject1", book_extra, func_col("a", "C")
                        )
                        self.assertTrue(func(table_dicts))

                    with self.subTest("TestNamedRangeToDicts1"):
                        func_TableToDicts = book.macro("MiscTableToDicts.TableToDicts")
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestNamedRangeToDicts1"
                        )

                        table_dicts = func_TableToDicts("NamedRange1", book_extra)
                        self.assertTrue(func(table_dicts))

                    with self.subTest("TestNamedRangeToDicts2"):
                        func_TableToDicts = book.macro("MiscTableToDicts.TableToDicts")
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestNamedRangeToDicts2"
                        )
                        func_col = book.macro("MiscCollectionCreate.Col")

                        table_dicts = func_TableToDicts(
                            "NamedRange1", book_extra, func_col("a", "C")
                        )
                        self.assertTrue(func(table_dicts))

                    with self.subTest("TestEmptyTablesToDicts1"):
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestEmptyTablesToDicts1"
                        )
                        self.assertTrue(func(book_extra))

                    with self.subTest("TestEmptyTablesToDicts2"):
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestEmptyTablesToDicts2"
                        )
                        self.assertTrue(func(book_extra))

                    with self.subTest("TestEmpty1ColumnTablesToDicts1"):
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestEmpty1ColumnTablesToDicts1"
                        )
                        self.assertTrue(func(book_extra))

                    with self.subTest("TestEmpty1ColumnTablesToDicts2"):
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestEmpty1ColumnTablesToDicts2"
                        )
                        self.assertTrue(func(book_extra))

                    with self.subTest("TestGetTableRowIndex1"):
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestGetTableRowIndex1"
                        )
                        self.assertTrue(func())

                    with self.subTest("TestGetTableRowIndex2"):
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestGetTableRowIndex2"
                        )
                        self.assertTrue(func())

                    with self.subTest("TestTableToDictsLogSource"):
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestTableToDictsLogSource"
                        )
                        self.assertTrue(func(book_extra))

                    with self.subTest("TestGetTableRowRange1"):
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestGetTableRowRange1"
                        )
                        self.assertTrue(func(book_extra))

                    with self.subTest("TestGetTableRowRange2"):
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.TestGetTableRowRange2"
                        )
                        self.assertTrue(func(book_extra))

                    with self.subTest("Test_TableDictToArray_fail_1"):
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.Test_TableDictToArray_fail_1"
                        )
                        self.assertTrue(func())

                    with self.subTest("Test_TableDictToArray_fail_2"):
                        func = book.macro(
                            "Test__Helper_MiscTableToDicts.Test_TableDictToArray_fail_2"
                        )
                        self.assertTrue(func())


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
