from ..util import TestCaseWithFunctionBook


class TestEnsureDotSeparatorTransformation(TestCaseWithFunctionBook):
    def test_1(self) -> None:
        func_EnsureDotSeparatorTransformation = self.book.macro(
            "MiscArray.EnsureDotSeparatorTransformation"
        )

        self.assertEqual(
            (("100.2", "1.9"), ("2.1", "2.2")),
            func_EnsureDotSeparatorTransformation([[100.2, 1.9], [2.1, 2.2]]),
        )

        self.assertEqual(
            ("1.2", "2.1", "3.8"),
            func_EnsureDotSeparatorTransformation([1.2, 2.1, 3.8]),
        )
