from ..util import TestCaseWithFunctionBook


class TestIsInArray(TestCaseWithFunctionBook):
    def test_1(self) -> None:
        func_IsInArray = self.book.macro("MiscArray.IsInArray")
        arr = [1, "2", "k", 3.4]

        self.assertTrue(func_IsInArray(arr, 1))
        self.assertFalse(func_IsInArray(arr, "1"))
        self.assertFalse(func_IsInArray(arr, 2))
        self.assertTrue(func_IsInArray(arr, "2"))
        self.assertTrue(func_IsInArray(arr, 3.4))
