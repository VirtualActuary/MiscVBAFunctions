from ..util import TestCaseWithFunctionBook


class TestArrayToCollection(TestCaseWithFunctionBook):
    def test_1(self) -> None:
        func_ArrayToCollection = self.book.macro("MiscArray.ArrayToCollection")

        func_Col_to_arr = self.book.macro("MiscCollection.CollectionToArray")

        self.assertEqual(
            (10, 11, 12, 13),
            func_Col_to_arr(func_ArrayToCollection([10, 11, 12, 13])),
        )
