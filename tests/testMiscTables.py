import unittest
from xlwings import Book
from locate import this_dir, prepend_sys_path
from aa_py_xl import excel_app

with prepend_sys_path():
    from util import functions_book, vba_dict, extra_book


class MiscTables(unittest.TestCase):
    def test_1(self) -> None:
        with excel_app(quiet=True, close=True, events=False, logger=None) as app:
            book: Book
            with functions_book(app=app) as book:
                func_HasLO = book.macro("MiscTables.HasLO")
                func_GetLO = book.macro("MiscTables.GetLO")

                repo_path = this_dir().parent
                repo_path = repo_path.joinpath(
                    r"test_data\MiscTables\MiscTablesTests.xlsx"
                )

                book_extra: Book
                with extra_book(app, repo_path) as book_extra:
                    with self.subTest("HasLO"):
                        self.assertTrue(
                            func_HasLO("Table1", book_extra)
                        )  # Correct case
                        self.assertTrue(
                            func_HasLO("taBLe1", book_extra)
                        )  # Should work without correct case

                    with self.subTest("TestListAllTables"):
                        func_GetAllTables = book.macro("MiscTables.GetAllTables")
                        func_IsValueInCollection = book.macro(
                            "MiscCollection.IsValueInCollection"
                        )

                        Tables = func_GetAllTables(book_extra)

                        self.assertTrue(func_IsValueInCollection(Tables, "Table1"))
                        self.assertTrue(func_IsValueInCollection(Tables, "NamedRange1"))
                        self.assertTrue(
                            func_IsValueInCollection(Tables, "SheetScopedNamedRange1")
                        )
                    with self.subTest("Test_TableColumnToArray"):
                        func_TableColumnToArray = book.macro(
                            "MiscTables.TableColumnToArray"
                        )
                        func_col = book.macro("MiscCollectionCreate.Col")

                        col = func_col(
                            vba_dict({"a": 1, "b": 2}), vba_dict({"a": 10, "b": 20})
                        )
                        arr = func_TableColumnToArray(col, "b")
                        self.assertEqual(2, arr[0])
                        self.assertEqual(20, arr[1])

                    with self.subTest("TableColumnToCollection"):
                        func_TableColumnToCollection = book.macro(
                            "MiscTables.TableColumnToCollection"
                        )
                        func_col = book.macro("MiscCollectionCreate.Col")
                        func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")

                        col1 = func_col(
                            vba_dict({"a": 1, "b": 2}), vba_dict({"a": 10, "b": 20})
                        )

                        col2 = func_TableColumnToCollection(col1, "b")
                        arr = func_Col_to_arr(col2)

                        self.assertEqual(2, arr[0])
                        self.assertEqual(20, arr[1])

                    with self.subTest("GetTableColumnDataRange"):
                        func_GetTableColumnDataRange = book.macro(
                            "MiscTables.GetTableColumnDataRange"
                        )
                        func_RangeToArray = book.macro("MiscRangeToArray.RangeToArray")
                        LO = func_GetLO("table2", book_extra)
                        arr = func_RangeToArray(
                            func_GetTableColumnDataRange(LO, "Column2")
                        )
                        self.assertEqual(12, arr[0])
                        self.assertEqual(22, arr[1])
                        self.assertEqual(32, arr[2])

                    with self.subTest("GetTableRowNumberDataRange"):
                        func_GetTableRowNumberDataRange = book.macro(
                            "MiscTables.GetTableRowNumberDataRange"
                        )
                        func_RangeToArray = book.macro("MiscRangeToArray.RangeToArray")
                        LO = func_GetLO("table2", book_extra)
                        arr = func_RangeToArray(func_GetTableRowNumberDataRange(LO, 2))
                        self.assertEqual(21, arr[0])
                        self.assertEqual(22, arr[1])
                        self.assertEqual(23, arr[2])

                    with self.subTest("GetTableColumnDataRange_2"):
                        func_GetTableColumnDataRange = book.macro(
                            "MiscTables.GetTableColumnDataRange"
                        )
                        func_ResizeLO = book.macro("MiscTables.ResizeLO")
                        func_RangeToArray = book.macro("MiscRangeToArray.RangeToArray")
                        SelectedTable = func_GetLO("table4", book_extra)
                        range_obj = func_GetTableColumnDataRange(
                            SelectedTable, "Column2"
                        )
                        func_ResizeLO(SelectedTable, 0)
                        arr = func_RangeToArray(range_obj)
                        self.assertEqual(None, arr[0])

                    with self.subTest("TestHasLO_and_GetLO"):
                        func = book.macro("Test__Helper_MiscTables.TestHasLO_and_GetLO")
                        self.assertTrue(func(book_extra))

                    with self.subTest("TestTableRange"):
                        func = book.macro("Test__Helper_MiscTables.TestTableRange")
                        self.assertTrue(func(book_extra))

                    with self.subTest("Test_CopyTable"):
                        func = book.macro("Test__Helper_MiscTables.Test_CopyTable")
                        self.assertTrue(func(book_extra, book))

                    with self.subTest("Test_CopyTable"):
                        func = book.macro(
                            "Test__Helper_MiscTables.Test_TableColumnToCollection"
                        )
                        self.assertTrue(func())

                    with self.subTest("Test_ResizeLO_1"):
                        func_ResizeLO = book.macro("MiscTables.ResizeLO")
                        LO = func_GetLO("Table1", book_extra)
                        func_ResizeLO(LO, 3)
                        func = book.macro("Test__Helper_MiscTables.Test_ResizeLO_1")
                        self.assertTrue(func(LO))

                    with self.subTest("Test_ResizeLO_2"):
                        func_ResizeLO = book.macro("MiscTables.ResizeLO")
                        LO = func_GetLO("Table1", book_extra)
                        func_ResizeLO(LO, 0)
                        func = book.macro("Test__Helper_MiscTables.Test_ResizeLO_2")
                        self.assertTrue(func(LO))

                    with self.subTest("Test_ResizeLO_3"):
                        func_ResizeLO = book.macro("MiscTables.ResizeLO")
                        LO = func_GetLO("Table1", book_extra)
                        func_ResizeLO(LO, 5)
                        func_ResizeLO(LO, 2)
                        func = book.macro("Test__Helper_MiscTables.Test_ResizeLO_3")
                        self.assertTrue(func(LO))

                    with self.subTest("Test_ResizeLO_4"):
                        func_ResizeLO = book.macro("MiscTables.ResizeLO")
                        LO = func_GetLO("Table1", book_extra)
                        func_ResizeLO(LO, 0)
                        func_ResizeLO(LO, 1)
                        func = book.macro("Test__Helper_MiscTables.Test_ResizeLO_4")
                        self.assertTrue(func(LO))

                    with self.subTest("Test_GetTableColumnDataRange_fail"):
                        func = book.macro(
                            "Test__Helper_MiscTables.Test_GetTableColumnDataRange_fail"
                        )
                        LO = func_GetLO("table2", book_extra)
                        self.assertTrue(func(LO))

                    with self.subTest("Test_GetTableRowNumberDataRange_fail"):
                        func = book.macro(
                            "Test__Helper_MiscTables.Test_GetTableRowNumberDataRange_fail"
                        )
                        LO = func_GetLO("table2", book_extra)
                        self.assertTrue(func(LO))

                    repo_path = this_dir().parent
                    repo_path = repo_path.joinpath(
                        r"test_data\MiscTableToDicts\MiscTableToDicts.xlsx"
                    )

                    book_temp: Book
                    with extra_book(app, repo_path) as book_temp:
                        with self.subTest("Test_GetTableRowRange1"):
                            func_col = book.macro("MiscCollectionCreate.Col")

                            func = book.macro(
                                "Test__Helper_MiscTables.Test_GetTableRowRange1"
                            )

                            func_GetTableRowRange = book.macro(
                                "MiscTables.GetTableRowRange"
                            )
                            R = func_GetTableRowRange(
                                "ListObject1",
                                func_col("a", "b"),
                                func_col(4, 5),
                                book_temp,
                            )

                            self.assertTrue(func(R))

                        with self.subTest("Test_GetTableRowRange2"):
                            func_col = book.macro("MiscCollectionCreate.Col")

                            func = book.macro(
                                "Test__Helper_MiscTables.Test_GetTableRowRange2"
                            )

                            func_GetTableRowRange = book.macro(
                                "MiscTables.GetTableRowRange"
                            )
                            R = func_GetTableRowRange(
                                "NamedRange1",
                                func_col("a", "b"),
                                func_col(4, 5),
                                book_temp,
                            )

                            self.assertTrue(func(R))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
