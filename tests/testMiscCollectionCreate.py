import unittest
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class TestMin(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            with self.subTest("Col"):
                func_col = book.macro("MiscCollectionCreate.Col")
                func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")
                col = func_col(1, 4, 5)
                arr = func_Col_to_arr(col)
                self.assertEqual(1, arr[0])
                self.assertEqual(4, arr[1])
                self.assertEqual(5, arr[2])


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
