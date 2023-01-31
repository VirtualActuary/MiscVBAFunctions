import unittest

from .util import functions_book


class TestMin(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:

            with self.subTest("min"):
                func_min = book.macro("MiscCollection.Min")
                func_col = book.macro("MiscCollectionCreate.Col")
                self.assertEqual(4, func_min(func_col(7, 4, 5, 6)))
                self.assertEqual(5, func_min(func_col(9, 5, 6)))

            with self.subTest("max"):
                func_max = book.macro("MiscCollection.Max")
                func_col = book.macro("MiscCollectionCreate.Col")
                self.assertEqual(6, func_max(func_col(4, 5, 6, 1, 2)))
                self.assertEqual(6.1, func_max(func_col(5.3, 6.1)))

            with self.subTest("mean"):
                func_mean = book.macro("MiscCollection.Mean")
                func_col = book.macro("MiscCollectionCreate.Col")

                self.assertEqual(4, func_mean(func_col(4, 5, 6, 3, 2)))
                self.assertEqual(6, func_mean(func_col(5, 7)))

            with self.subTest("IsValueInCollection"):
                func_IsValueInCollection = book.macro(
                    "MiscCollection.IsValueInCollection"
                )
                func_col = book.macro("MiscCollectionCreate.Col")

                self.assertTrue(func_IsValueInCollection(func_col("a", "b"), "b"))
                self.assertFalse(func_IsValueInCollection(func_col("a", "b"), "c"))
                self.assertFalse(
                    func_IsValueInCollection(func_col("a", "b"), "B", True)
                )

            with self.subTest("Join_Collections"):
                func_Join_Collections = book.macro("MiscCollection.JoinCollections")
                func_col = book.macro("MiscCollectionCreate.Col")
                func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")

                c1 = func_col(1, 2, 3)
                c2 = func_col(4, 5, 6)
                c3 = func_col(7, 8, 9)
                x = func_Join_Collections(c2, c3, c1)

                self.assertEqual((4, 5, 6, 7, 8, 9, 1, 2, 3), func_Col_to_arr(x))

            with self.subTest("CollectionToArray"):
                func_col = book.macro("MiscCollectionCreate.Col")
                func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")

                self.assertEqual((1, 2, 3), func_Col_to_arr(func_col(1, 2, 3)))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
