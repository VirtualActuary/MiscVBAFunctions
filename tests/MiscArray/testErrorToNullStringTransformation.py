from ..util import TestCaseWithFunctionBook


class TestErrorToNullStringTransformation(TestCaseWithFunctionBook):
    def test_1(self) -> None:
        func_Test_ErrorToNullStringTransformation = self.book.macro(
            "Test__Helper_MiscArray.Test_ErrorToNullStringTransformation_1"
        )

        self.assertTrue(func_Test_ErrorToNullStringTransformation())

    def test_2(self) -> None:
        func_Test_ErrorToNullStringTransformation = self.book.macro(
            "Test__Helper_MiscArray.Test_ErrorToNullStringTransformation_2"
        )

        self.assertTrue(func_Test_ErrorToNullStringTransformation())
