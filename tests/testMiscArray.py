import unittest
from xlwings import Book
from .util import functions_book
from aa_py_xl import Table


class TestMin(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("EnsureDotSeparatorTransformation"):
                func_EnsureDotSeparatorTransformation = book.macro(
                    "MiscArray.EnsureDotSeparatorTransformation"
                )

                self.assertEqual(
                    (("100.2", "1.9"), ("2.1", "2.2")),
                    func_EnsureDotSeparatorTransformation([[100.2, 1.9], [2.1, 2.2]]),
                )

                self.assertEqual(
                    ("1.2", "2.1", "3.8"),
                    func_EnsureDotSeparatorTransformation([1.2, 2.1, 3.8]),
                )

            with self.subTest("DateToStringTransformation"):
                import datetime as dt

                def time_helper(date):
                    return date + dt.timedelta(hours=2)

                func_DateToStringTransformation = book.macro(
                    "MiscArray.DateToStringTransformation"
                )

                self.assertEqual(
                    ((100.2, "2021-01-02"), (2.1, "2021-01-28")),
                    func_DateToStringTransformation(
                        [
                            [100.2, (dt.datetime(2021, 1, 2, 2, 0, 0))],
                            [2.1, (dt.datetime(2021, 1, 28, 2, 0, 0))],
                        ]
                    ),
                )

                self.assertEqual(
                    (1.2, 2.1, "2021-03-28"),
                    func_DateToStringTransformation(
                        [1.2, 2.1, dt.datetime(2021, 3, 28, 10, 2, 10)]
                    ),
                )

                self.assertEqual(
                    "2021-01",
                    func_DateToStringTransformation(
                        [dt.datetime(2021, 1, 28, 10, 2, 10)], "yyyy-mm"
                    )[0],
                )

                self.assertEqual(
                    "2021/01/28",
                    func_DateToStringTransformation(
                        [dt.datetime(2021, 1, 28, 10, 2, 10)], "yyyy/mm/dd"
                    )[0],
                )

                self.assertEqual(
                    "2021-01-28 10:02:10",
                    func_DateToStringTransformation(
                        [time_helper(dt.datetime(2021, 1, 28, 10, 2, 10))],
                        "yyyy-mm-dd hh:mm:ss",
                    )[0],
                )

            with self.subTest("ArrayToCollection"):
                func_ArrayToCollection = book.macro("MiscArray.ArrayToCollection")

                func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")

                self.assertEqual(
                    (10, 11, 12, 13),
                    func_Col_to_arr(func_ArrayToCollection([10, 11, 12, 13])),
                )

            with self.subTest("ArrayToRange"):
                func_ArrayToRange = book.macro("MiscArray.ArrayToRange")
                range_obj = book.sheets["Sheet1"].range("A1")

                func_ArrayToRange(
                    [["a", "b"], [1, 2], ["aa", "bb"]], range_obj, False, True
                )
                self.assertEqual(
                    [["a", "b"], [1, 2], ["aa", "bb"]],
                    book.sheets[0].range("A1:B3").value,
                )

            with self.subTest("ArrayToRange_2"):
                func_ArrayToRange = book.macro("MiscArray.ArrayToRange")

                range_obj = book.sheets["Sheet1"].range("F1")

                func_ArrayToRange([["a"], ["=[d]"]], range_obj, True)
                self.assertEqual(
                    ["a", "=[d]"],
                    book.sheets[0].range("F1:F2").value,
                )

            with self.subTest("ArrayToRange_3"):
                func_ArrayToRange = book.macro("MiscArray.ArrayToRange")

                range_obj = book.sheets["Sheet1"].range("H1")

                func_ArrayToRange(
                    [["asdf", 1234, "2022/11/02", False], ["a", "b", "c", "d"]],
                    range_obj,
                    False,
                    True,
                )
                self.assertEqual(
                    [["asdf", 1234, "2022/11/02", False], ["a", "b", "c", "d"]],
                    book.sheets[0].range("H1:K2").value,
                )

            with self.subTest("ArrayToNewTable"):
                func_ArrayToNewTable = book.macro("MiscArray.ArrayToNewTable")

                arr = [["col1", "col2", "col3"], ["=[d]", "=d", 1]]

                range_obj = book.sheets["Sheet1"].range("M1")
                func_ArrayToNewTable("TestTable", arr, range_obj, True)

                table = Table.get_from_book(book, "TestTable")

                self.assertEqual("$M$1:$O$2", table.range.address)
                self.assertEqual("TestTable", table.name)
                self.assertEqual("Sheet1", table.sheet.name)
                self.assertEqual(
                    [["col1", "col2", "col3"], ["=[d]", "=d", 1]],
                    table.range.value,
                )

            with self.subTest("ArrayToNewTable_FunkyHeaders"):

                func_ArrayToNewTable = book.macro("MiscArray.ArrayToNewTable")

                arr = [["asdf", 1234, "2022/11/02", False], ["a", "b", "c", "d"]]

                range_obj = book.sheets["Sheet1"].range("Q1")
                func_ArrayToNewTable("TestTable2", arr, range_obj, True)

                table = Table.get_from_book(book, "TestTable2")

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

            with self.subTest("ArrayToNewTable_1dArray"):
                func_ArrayToNewTable = book.macro("MiscArray.ArrayToNewTable")
                func_Ensure2dArray = book.macro("MiscArray.Ensure2dArray")

                range_obj = book.sheets["Sheet1"].range("V1")
                func_ArrayToNewTable(
                    "TestTable3",
                    func_Ensure2dArray(["col1", "col2", "col3"]),
                    range_obj,
                    True,
                )
                table = Table.get_from_book(book, "TestTable3")

                self.assertEqual("TestTable3", table.name)
                self.assertEqual("col2", table.range(1, 2).value)

            with self.subTest("Ensure2DArray"):
                func_Ensure2dArray = book.macro("MiscArray.Ensure2dArray")
                self.assertEqual((("a", "b", "c"),), func_Ensure2dArray(["a", "b", "c"]))
                self.assertEqual((("a", "b", "c"),), func_Ensure2dArray([["a", "b", "c"]]))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
