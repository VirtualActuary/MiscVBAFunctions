from typing import List, Any

from ..util import TestCaseWithFunctionBook


class TestArrayGetNumDimensions(TestCaseWithFunctionBook):
    def test_1(self) -> None:
        func_ArrayGetNumDimensions = self.book.macro("MiscArray.ArrayGetNumDimensions")

        arr: List[Any]
        arr = []
        self.assertEqual(1, func_ArrayGetNumDimensions(arr))

        arr = [[], []]
        self.assertEqual(2, func_ArrayGetNumDimensions(arr))

        arr = [[[], []], [[], []]]
        self.assertEqual(3, func_ArrayGetNumDimensions(arr))

        arr = [[[[], []], [[], []]], [[[], []], [[], []]]]
        self.assertEqual(4, func_ArrayGetNumDimensions(arr))
