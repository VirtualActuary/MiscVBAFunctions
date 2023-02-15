import unittest
from xlwings import Book
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book, vba_dict


class MiscEnsureDictIUtil(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("EnsureDictI"):
                func_EnsureDictI = book.macro("MiscEnsureDictIUtil.EnsureDictI")
                d1 = vba_dict({"a": 1})
                d2 = func_EnsureDictI(d1)

                self.assertTrue(d1.exists("a"))
                self.assertFalse(d1.exists("A"))

                self.assertTrue(d2.exists("a"))
                self.assertTrue(d2.exists("A"))

            with self.subTest("Test_EnsureDictIContainer"):
                func = book.macro(
                    "Test__Helper_MiscEnsureDictUtil.Test_EnsureDictIContainer"
                )
                func_col = book.macro("MiscCollectionCreate.Col")
                col = func_col(
                    vba_dict({"A": "foo"}),
                    vba_dict({"b": "foo"}),
                    vba_dict({"C": "foo"}),
                )

                self.assertTrue(func(col))

            with self.subTest("Test_EnsureDictIContainer_I"):
                func = book.macro(
                    "Test__Helper_MiscEnsureDictUtil.Test_EnsureDictIContainer_I"
                )
                func_EnsureDictI = book.macro("MiscEnsureDictIUtil.EnsureDictI")

                func_col = book.macro("MiscCollectionCreate.Col")
                col = func_col(
                    vba_dict({"A": "foo"}),
                    vba_dict({"b": "foo"}),
                    vba_dict({"C": "foo"}),
                )
                col = func_EnsureDictI(col)
                self.assertTrue(func(col))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
