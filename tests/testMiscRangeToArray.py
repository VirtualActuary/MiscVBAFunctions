import unittest
from xlwings import Book
from locate import this_dir
from aa_py_xl import excel
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscRangeToArray(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            func_RangeToArray = book.macro("MiscRangeToArray.RangeToArray")

            repo_path = this_dir().parent
            repo_path = repo_path.joinpath(
                r"test_data\MiscRangeToArray\RangeToArray.xlsx"
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
                with self.subTest("RangeToArray_2D"):
                    range_obj = book_temp.sheets["Sheet1"].range("A1:C2")
                    arr = func_RangeToArray(range_obj)
                    self.assertEqual(11, arr[0][0])
                    self.assertEqual(12, arr[0][1])
                    self.assertEqual(9, arr[0][2])
                    self.assertEqual(13, arr[1][0])
                    self.assertEqual(14, arr[1][1])
                    self.assertEqual(15, arr[1][2])
                    self.assertEqual(2, len(arr))
                    self.assertEqual(3, len(arr[1]))

                with self.subTest("RangeToArray_1_value"):
                    range_obj = book_temp.sheets["Sheet2"].range("A1")
                    arr = func_RangeToArray(range_obj)
                    self.assertEqual(4, arr[0])
                    self.assertEqual(1, len(arr))

                with self.subTest("RangeToArray_1D_row"):
                    range_obj = book_temp.sheets["Sheet3"].range("A1:C1")
                    arr = func_RangeToArray(range_obj)
                    self.assertEqual(1, arr[0])
                    self.assertEqual(2, arr[1])
                    self.assertEqual(3, arr[2])
                    self.assertEqual(3, len(arr))

                with self.subTest("RangeToArray_1D_column"):
                    range_obj = book_temp.sheets["Sheet4"].range("A1:A3")
                    arr = func_RangeToArray(range_obj)
                    self.assertEqual(66, arr[0])
                    self.assertEqual(77, arr[1])
                    self.assertEqual(88, arr[2])
                    self.assertEqual(3, len(arr))

                with self.subTest("Test_RangeToFlatArray"):
                    func_RangeToFlatArray = book.macro("MiscRangeToArray.RangeToFlatArray")
                    func_ArrayToRange = book.macro("MiscArray.ArrayToRange")

                    arr = [[11, 22, 33], [44, "111", "222"], ["333", "444", "555"]]

                    range_start = book.sheets[0].range("B4")
                    range_test = func_ArrayToRange(arr, range_start, True)

                    ArrayOutput = func_RangeToFlatArray(range_test)

                    self.assertEqual((11, 22, 33, 44, "111", "222", "333", "444", "555"), ArrayOutput)


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
