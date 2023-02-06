import unittest
from locate import prepend_sys_path
with prepend_sys_path():
    from util import functions_book

class TestMin(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            with self.subTest("BubbleSort"):
                func_col = book.macro("MiscCollectionCreate.Col")
                func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")

                func_BubbleSort = book.macro("MiscCollectionSort.BubbleSort")

                Coll = func_col(
                    "variables10",
                    "variables",
                    "variables2",
                    "variables_10",
                    "variables_2",
                )
                Coll = func_BubbleSort(Coll)

                arr = func_Col_to_arr(Coll)
                self.assertEqual(
                    (
                        "variables",
                        "variables10",
                        "variables2",
                        "variables_10",
                        "variables_2",
                    ),
                    arr,
                )


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
