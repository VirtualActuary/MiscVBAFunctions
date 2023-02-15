import unittest
from xlwings import Book
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscRange(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("Test_RangeToLO"):
                func_RangeToLO = book.macro("MiscRange.RangeToLO")
                func_ArrayToRange = book.macro("MiscArray.ArrayToRange")
                func_HasLO = book.macro("MiscTables.HasLO")

                arr = [["col1", "col2", "col3"], ["=[d]", "=d", 1]]
                range_start = book.sheets[0].range("B4")
                range_test = func_ArrayToRange(arr, range_start, True)
                LO = func_RangeToLO(book.sheets[0], range_test, "myTable")
                self.assertEqual("col2", LO.Range(1, 2).Value)
                self.assertEqual(1, LO.Range(2, 3).Value)
                self.assertTrue(func_HasLO("myTable", book))

            with self.subTest("Test_RangeToLO_fail"):
                func = book.macro("Test__Helper_MiscRange.Test_RangeToLO_fail")
                self.assertTrue(func())


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
