from ..util import TestCaseWithFunctionBook


class TestArrayToRange(TestCaseWithFunctionBook):
    def test_1(self) -> None:
        func_ArrayToRange = self.book.macro("MiscArray.ArrayToRange")
        range_obj = self.book.sheets["Sheet1"].range("A1")

        func_ArrayToRange([["a", "b"], [1, 2], ["aa", "bb"]], range_obj, False, True)
        self.assertEqual(
            [["a", "b"], [1, 2], ["aa", "bb"]],
            self.book.sheets[0].range("A1:B3").value,
        )

    def test_2(self) -> None:
        func_ArrayToRange = self.book.macro("MiscArray.ArrayToRange")

        range_obj = self.book.sheets["Sheet1"].range("F1")

        func_ArrayToRange([["a"], ["=[d]"]], range_obj, True)
        self.assertEqual(
            ["a", "=[d]"],
            self.book.sheets[0].range("F1:F2").value,
        )

    def test_3(self) -> None:
        func_ArrayToRange = self.book.macro("MiscArray.ArrayToRange")

        range_obj = self.book.sheets["Sheet1"].range("H1")

        func_ArrayToRange(
            [["asdf", 1234, "2022/11/02", False], ["a", "b", "c", "d"]],
            range_obj,
            False,
            True,
        )
        self.assertEqual(
            [["asdf", 1234, "2022/11/02", False], ["a", "b", "c", "d"]],
            self.book.sheets[0].range("H1:K2").value,
        )

    def test_fail(self) -> None:
        func_Test_ArrayToRange_fail = self.book.macro(
            "Test__Helper_MiscArray.Test_ArrayToRange_fail"
        )

        arr = ["col1", "col2", "col3"]
        range_obj = self.book.sheets.active.range("B4")
        self.assertTrue(func_Test_ArrayToRange_fail(arr, range_obj))
