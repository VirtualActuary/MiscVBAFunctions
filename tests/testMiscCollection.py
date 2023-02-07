import unittest
from locate import prepend_sys_path
with prepend_sys_path():
    from util import functions_book, vba_dict


class TestMin(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            func_col = book.macro("MiscCollectionCreate.Col")

            with self.subTest("min"):
                func_min = book.macro("MiscCollection.Min")
                self.assertEqual(4, func_min(func_col(7, 4, 5, 6)))
                self.assertEqual(5, func_min(func_col(9, 5, 6)))

            with self.subTest("max"):
                func_max = book.macro("MiscCollection.Max")
                self.assertEqual(6, func_max(func_col(4, 5, 6, 1, 2)))
                self.assertEqual(6.1, func_max(func_col(5.3, 6.1)))

            with self.subTest("mean"):
                func_mean = book.macro("MiscCollection.Mean")

                self.assertEqual(4, func_mean(func_col(4, 5, 6, 3, 2)))
                self.assertEqual(6, func_mean(func_col(5, 7)))

            with self.subTest("IsValueInCollection"):
                func_IsValueInCollection = book.macro(
                    "MiscCollection.IsValueInCollection"
                )

                self.assertTrue(func_IsValueInCollection(func_col("a", "b"), "b"))
                self.assertFalse(func_IsValueInCollection(func_col("a", "b"), "c"))
                self.assertFalse(
                    func_IsValueInCollection(func_col("a", "b"), "B", True)
                )

            with self.subTest("Join_Collections"):
                func_Join_Collections = book.macro("MiscCollection.JoinCollections")
                func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")

                c1 = func_col(1, 2, 3)
                c2 = func_col(4, 5, 6)
                c3 = func_col(7, 8, 9)
                x = func_Join_Collections(c2, c3, c1)

                self.assertEqual((4, 5, 6, 7, 8, 9, 1, 2, 3), func_Col_to_arr(x))

            with self.subTest("CollectionToArray"):
                func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")

                self.assertEqual((1, 2, 3), func_Col_to_arr(func_col(1, 2, 3)))

            with self.subTest("Test_min_fail"):
                func = book.macro("Test__Helper_MiscCollection.Test_min_fail")
                self.assertTrue(func())

            with self.subTest("Test_max_fail"):
                func = book.macro("Test__Helper_MiscCollection.Test_max_fail")
                self.assertTrue(func())

            with self.subTest("Test_mean_fail"):
                func = book.macro("Test__Helper_MiscCollection.Test_mean_fail")
                self.assertTrue(func())

            with self.subTest("Test_Join_Collections_fail"):
                func_JoinCollections = book.macro("MiscCollection.JoinCollections")
                func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")

                col = func_JoinCollections(
                    func_col(1, 2, 3),
                    func_col(4, 5, 6),
                )
                arr = func_Col_to_arr(col)
                try:
                    _ = arr[6]
                    self.assertTrue(False)
                except IndexError:
                    self.assertTrue(True)
                else:
                    self.assertTrue(False)

            with self.subTest("Test_Join_Collections_fail_2"):
                func = book.macro(
                    "Test__Helper_MiscCollection.Test_Join_Collections_fail_2"
                )
                func_col = book.macro("MiscCollectionCreate.Col")
                self.assertTrue(func(vba_dict({"a": 1, "b": 2}), func_col(1, 2, 3)))

            with self.subTest("Test_Concat_Collections_fail"):
                func = book.macro(
                    "Test__Helper_MiscCollection.Test_Concat_Collections_fail"
                )
                func_col = book.macro("MiscCollectionCreate.Col")
                self.assertTrue(func(vba_dict({"a": 1, "b": 2}), func_col(1, 2, 3)))

            with self.subTest("Test_Concat_Collections"):
                func = book.macro("Test__Helper_MiscCollection.Test_Concat_Collections")
                self.assertTrue(func(func_col(1, 2), func_col(3, 4), func_col(5, 6)))

            with self.subTest("Test_CollectionToArray_empty"):
                func = book.macro(
                    "Test__Helper_MiscCollection.Test_CollectionToArray_empty"
                )
                func_Col_to_arr = book.macro("MiscCollection.CollectionToArray")
                arr = func_Col_to_arr(func_col())
                self.assertTrue(func(arr))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
