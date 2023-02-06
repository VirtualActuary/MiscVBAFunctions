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
            func_HasLO = book.macro("MiscTables.HasLO")
            func_GetLO = book.macro("MiscTables.GetLO")

            repo_path = this_dir().parent
            repo_path = repo_path.joinpath(r"test_data\MiscTables\MiscTablesTests.xlsx")

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
                with self.subTest("HasLO_and_GetLO"):
                    pass
                    self.assertTrue(func_HasLO("Table1", book_temp))  # Correct case
                    self.assertTrue(
                        func_HasLO("taBLe1", book_temp)
                    )  # Should work without correct case

                with self.subTest("HasLO_and_GetLO"):
                    func_GetAllTables = book.macro("MiscTables.GetAllTables")
                    func_IsValueInCollection = book.macro(
                        "MiscCollection.IsValueInCollection"
                    )

                    Tables = func_GetAllTables(book_temp)

                    self.assertTrue(func_IsValueInCollection(Tables, "Table1"))
                    self.assertTrue(func_IsValueInCollection(Tables, "NamedRange1"))
                    self.assertTrue(
                        func_IsValueInCollection(Tables, "SheetScopedNamedRange1")
                    )

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
                    LO = func_GetLO("table2", book_temp)
                    arr = func_RangeToArray(func_GetTableColumnDataRange(LO, "Column2"))
                    self.assertEqual(12, arr[0])
                    self.assertEqual(22, arr[1])
                    self.assertEqual(32, arr[2])

                with self.subTest("GetTableRowNumberDataRange"):
                    func_GetTableRowNumberDataRange = book.macro(
                        "MiscTables.GetTableRowNumberDataRange"
                    )
                    func_RangeToArray = book.macro("MiscRangeToArray.RangeToArray")
                    LO = func_GetLO("table2", book_temp)
                    arr = func_RangeToArray(func_GetTableRowNumberDataRange(LO, 2))
                    self.assertEqual(21, arr[0])
                    self.assertEqual(22, arr[1])
                    self.assertEqual(23, arr[2])

                with self.subTest("GetTableRowRange"):

                    func_GetTableColumnDataRange = book.macro(
                        "MiscTables.GetTableColumnDataRange"
                    )
                    func_ResizeLO = book.macro("MiscTables.ResizeLO")
                    func_RangeToArray = book.macro("MiscRangeToArray.RangeToArray")
                    SelectedTable = func_GetLO("table4", book_temp)
                    range_obj = func_GetTableColumnDataRange(SelectedTable, "Column2")
                    func_ResizeLO(SelectedTable, 0)
                    arr = func_RangeToArray(range_obj)
                    self.assertEqual(None, arr[0])


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
