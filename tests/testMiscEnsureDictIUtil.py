import unittest
from xlwings import Book
from locate import prepend_sys_path
with prepend_sys_path():
    from util import functions_book,vba_dict

class TestDictsToTable(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
