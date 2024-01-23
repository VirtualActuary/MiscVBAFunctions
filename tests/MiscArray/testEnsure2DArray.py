from ..util import TestCaseWithFunctionBook


class TestEnsure2DArray(TestCaseWithFunctionBook):
    def test_1(self) -> None:
        func_Ensure2dArray = self.book.macro("MiscArray.Ensure2dArray")
        self.assertEqual((("a", "b", "c"),), func_Ensure2dArray(["a", "b", "c"]))
        self.assertEqual((("a", "b", "c"),), func_Ensure2dArray([["a", "b", "c"]]))
