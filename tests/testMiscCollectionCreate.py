import unittest
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscCollectionCreate(unittest.TestCase):
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

            with self.subTest("zip"):
                func_col = book.macro("MiscCollectionCreate.Col")
                func_zip = book.macro("MiscCollectionCreate.Zip")
                func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")

                c1 = func_zip(func_col(1, 2, 3), func_col(4, 5, 6, 7))
                arr = func_Col_to_arr(c1)
                self.assertEqual(((1, 4), (2, 5), (3, 6)), arr)

            with self.subTest("zip2"):
                func_col = book.macro("MiscCollectionCreate.Col")
                func_zip = book.macro("MiscCollectionCreate.Zip")
                func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")

                c1 = func_zip(func_col(1, 2, 3), func_col(4, 5, 6, 7), func_col(10, 12))
                arr = func_Col_to_arr(c1)
                self.assertEqual(((1, 4, 10), (2, 5, 12)), arr)


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
