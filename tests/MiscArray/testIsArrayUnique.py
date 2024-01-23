from ..util import TestCaseWithFunctionBook


class TestIsArrayUnique(TestCaseWithFunctionBook):
    def test_1(self) -> None:
        func_IsArrayUnique = self.book.macro("MiscArray.IsArrayUnique")
        arr = [1, "1", "3.4", 3.4]
        arr2 = ["asdf", 3.4, 3.4, "1"]
        self.assertTrue(func_IsArrayUnique(arr))
        self.assertFalse(func_IsArrayUnique(arr2))
