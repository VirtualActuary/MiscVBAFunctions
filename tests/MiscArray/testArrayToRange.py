from ..util import TestCaseWithFunctionBook


class TestArrayToRange(TestCaseWithFunctionBook):
    def setUp(self) -> None:
        super().setUp()
        self.array_to_range = self.book.macro("MiscArray.ArrayToRange")
        self.array_to_range_fail = self.book.macro(
            "Test__Helper_MiscArray.Test_ArrayToRange_fail"
        )

    def test_1(self) -> None:
        range_obj = self.book.sheets["Sheet1"].range("A1")

        self.array_to_range([["a", "b"], [1, 2], ["aa", "bb"]], range_obj, False, True)
        self.assertEqual(
            [["a", "b"], [1, 2], ["aa", "bb"]],
            self.book.sheets[0].range("A1:B3").value,
        )

    def test_2(self) -> None:
        range_obj = self.book.sheets["Sheet1"].range("F1")

        self.array_to_range([["a"], ["=[d]"]], range_obj, True)
        self.assertEqual(
            ["a", "=[d]"],
            self.book.sheets[0].range("F1:F2").value,
        )

    def test_3(self) -> None:
        range_obj = self.book.sheets["Sheet1"].range("H1")

        self.array_to_range(
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
        arr = ["col1", "col2", "col3"]
        range_obj = self.book.sheets.active.range("B4")
        self.assertTrue(self.array_to_range_fail(arr, range_obj))
