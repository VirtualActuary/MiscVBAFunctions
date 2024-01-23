from typing import List, Any

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

    def test_preserve_all_data_types(self) -> None:
        """
        Data types must be preserved:
        - Strings should stay strings, even if it looks like it might contain something else.
        - Numbers should stay numbers. Excel sometimes wants to convert them to dates.

        If data needs to be converted from one type to another, that belongs on a business logic level, and not
        in a low-level function like this one.

        See https://github.com/AutoActuary/autory/issues/1150
        """
        data: List[List[Any]] = [
            ["TRUE", "FALSE"],  # Strings that look like booleans.
            ["#VALUE!", "#N/A"],  # Strings that look like errors.
            ["1", "2"],  # Strings that look like integers.
            ["1.2", "2.3"],  # Strings that look like floats.
            ["1.2.3", "2.3.4"],  # Strings that look like version numbers.
            ["2022/11/02", "20231020T123456"],  # Strings that look like dates / times.
            [1.2, 2.3],  # These are floats, but Excel sometimes sees them as dates.
            [1, 2],  # These are integers, but Excel sometimes sees them as dates.
            [True, False],  # Booleans.
        ]

        self.array_to_range(
            data,
            self.book.sheets["Sheet1"].range("A1"),
            False,  # Escape formulas.
            False,  # No header.
        )
        self.assertEqual(
            data,
            self.book.sheets[0].range((1, 1), (len(data), len(data[0]))).value,
        )
