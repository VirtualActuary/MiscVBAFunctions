from datetime import datetime

from aa_py_xl import Table

from ..util import TestCaseWithFunctionBook


class TestArrayToNewTable(TestCaseWithFunctionBook):
    def test_1(self) -> None:
        func_ArrayToNewTable = self.book.macro("MiscArray.ArrayToNewTable")

        arr = [["col1", "col2", "col3"], ["=[d]", "=d", 1]]

        range_obj = self.book.sheets["Sheet1"].range("M1")
        func_ArrayToNewTable("TestTable", arr, range_obj, True)

        table = Table.get_from_book(self.book, "TestTable")

        self.assertEqual("$M$1:$O$2", table.range.address)
        self.assertEqual("TestTable", table.name)
        self.assertEqual("Sheet1", table.sheet.name)
        self.assertEqual(
            [["col1", "col2", "col3"], ["=[d]", "=d", 1]],
            table.range.value,
        )

    def test_string_dates(self) -> None:
        func = self.book.macro(
            "Test__Helper_MiscArray.Test_ArrayToNewTable_StringDates"
        )

        range_obj = self.book.sheets["Sheet1"].range("B1000")

        self.assertTrue(func(range_obj))

    def test_funky_headers(self) -> None:
        func_ArrayToNewTable = self.book.macro("MiscArray.ArrayToNewTable")

        arr = [["asdf", 1234, "2022/11/02", False], ["a", "b", "c", "d"]]

        range_obj = self.book.sheets["Sheet1"].range("Q1")
        func_ArrayToNewTable("TestTable2", arr, range_obj, True)

        table = Table.get_from_book(self.book, "TestTable2")

        self.assertEqual("$Q$1:$T$2", table.range.address)
        self.assertEqual("TestTable2", table.name)
        self.assertEqual("Sheet1", table.sheet.name)
        self.assertEqual(
            [["asdf", "1234", "2022/11/02", "FALSE"], ["a", "b", "c", "d"]],
            table.range.value,
        )
        self.assertEqual(
            8,
            table.range.count,
        )
        self.assertEqual(
            4,
            table.range.columns.count,
        )
        self.assertEqual(
            2,
            table.range.rows.count,
        )

    def test_1d_array(self) -> None:
        func_ArrayToNewTable = self.book.macro("MiscArray.ArrayToNewTable")
        func_Ensure2dArray = self.book.macro("MiscArray.Ensure2dArray")

        range_obj = self.book.sheets["Sheet1"].range("V1")
        func_ArrayToNewTable(
            "TestTable3",
            func_Ensure2dArray(["col1", "col2", "col3"]),
            range_obj,
            True,
        )
        table = Table.get_from_book(self.book, "TestTable3")

        self.assertEqual("TestTable3", table.name)
        self.assertEqual("col2", table.range(1, 2).value)

    def test_fail(self) -> None:
        func = self.book.macro("Test__Helper_MiscArray.Test_ArrayToNewTable_fail")
        arr = [["col1", "col2", "col3"], ["-[d]", "=d", 1]]
        range_obj = self.book.sheets.active.range("B4")
        self.assertTrue(func(arr, range_obj))

    def test_header_without_rows(self) -> None:
        func_ArrayToNewTable = self.book.macro("MiscArray.ArrayToNewTable")

        arr = [["a", "b", "c"]]

        table_name = "Table1"
        func_ArrayToNewTable(
            table_name,
            arr,
            self.book.sheets["Sheet1"].range("A1"),
            False,
            self.book.macro("col")("0.00", "General", "yyyy/mm/dd"),
        )

        table = Table.get_from_book(self.book, table_name)
        self.assertEqual(
            [
                ["a", "b", "c"],
                [None, None, None],
            ],
            table.range.value,
        )
        self.assertEqual(
            [
                # Header is always text.
                ["@", "@", "@"],
                # Number formats are ignored if there are no data rows.
                # This empty row is added because a ListObject requires at least one row.
                ["General", "General", "General"],
            ],
            [[cell.number_format for cell in row] for row in table.range.rows],
        )

    def test_single_data_row(self) -> None:
        func_ArrayToNewTable = self.book.macro("MiscArray.ArrayToNewTable")

        arr = [["a", "b", "c"], [1, 2, 3]]

        table_name = "Table1"
        func_ArrayToNewTable(
            table_name,
            arr,
            self.book.sheets["Sheet1"].range("A1"),
            False,
            self.book.macro("col")("0.00", "General", "yyyy/mm/dd"),
        )

        table = Table.get_from_book(self.book, table_name)
        self.assertEqual(
            [
                ["a", "b", "c"],
                [1, 2, datetime(1900, 1, 2)],
            ],
            table.range.value,
        )
        self.assertEqual(
            [
                # Header is always text.
                ["@", "@", "@"],
                # The given number formats are applied.
                ["0.00", "General", "yyyy/mm/dd"],
            ],
            [[cell.number_format for cell in row] for row in table.range.rows],
        )
