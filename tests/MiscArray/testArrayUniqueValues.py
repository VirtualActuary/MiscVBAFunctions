from ..util import TestCaseWithFunctionBook


class TestArrayUniqueValues(TestCaseWithFunctionBook):
    def test_1(self) -> None:
        func_ArrayUniqueValues = self.book.macro("MiscArray.ArrayUniqueValues")
        arr = [1, "1", 1, "3.4", "asdf", 3.4, 3.4, "1"]
        arr2 = func_ArrayUniqueValues(arr)
        self.assertEqual(5, len(arr2))
        self.assertEqual(arr2, (1, "1", "3.4", "asdf", 3.4))

    def test_2D(self) -> None:
        func_ArrayUniqueValues = self.book.macro("MiscArray.ArrayUniqueValues")
        arr = [[1, "1", 1, "foo"], [3.4, 1, "asdf", 3.4]]
        arr2 = func_ArrayUniqueValues(arr)
        self.assertEqual(5, len(arr2))
        self.assertEqual(arr2, (1, "1", "foo", 3.4, "asdf"))
