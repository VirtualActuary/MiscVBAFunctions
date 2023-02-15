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

            with self.subTest("Test_IsInRange"):
                func_IsInRange = book.macro("MiscRange.IsInRange")
                func_ArrayToRange = book.macro("MiscArray.ArrayToRange")

                arr = [[11, 22, 33], [44, "111", "222"], ["333", "444", "555"]]

                range_start = book.sheets[0].range("B4")
                range_test = func_ArrayToRange(arr, range_start, True)

                self.assertTrue(func_IsInRange(range_test, 11))
                self.assertFalse(func_IsInRange(range_test, 123))
                self.assertTrue(func_IsInRange(range_test, "111"))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
