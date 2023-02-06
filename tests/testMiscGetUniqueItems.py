import unittest
from xlwings import Book
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class TestDictsToTable(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("GetUniqueItems"):
                func_GetUniqueItems = book.macro("MiscGetUniqueItems.GetUniqueItems")

                self.assertEqual(3, len(func_GetUniqueItems(["a", "b", "c", "b"])))
                self.assertEqual(4, len(func_GetUniqueItems(["a", "b", "c", "B"])))
                self.assertEqual(
                    3, len(func_GetUniqueItems(["a", "b", "c", "B"], False))
                )

                self.assertEqual(3, len(func_GetUniqueItems([1, 2, 3, 2])))
                self.assertEqual(2, len(func_GetUniqueItems([1, 1, "a", "a"])))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
