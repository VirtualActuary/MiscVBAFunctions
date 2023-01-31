import unittest
from .util import functions_book


class TestMin(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            with self.subTest("dict"):
                func_dictget = book.macro("MiscDictionary.dictget")
                func_dict = book.macro("MiscDictionaryCreate.dict")

                d = func_dict("a", 2, "b", None)

                self.assertEqual(2, func_dictget(d, "a"))
                self.assertEqual(None, func_dictget(d, "b"))

            with self.subTest("dicti"):
                func_dictget = book.macro("MiscDictionary.dictget")
                func_dict = book.macro("MiscDictionaryCreate.dicti")

                d = func_dict("a", 2, "b", None)

                self.assertEqual(2, func_dictget(d, "a"))
                self.assertEqual(2, func_dictget(d, "A"))
                self.assertEqual(None, func_dictget(d, "b"))
                self.assertEqual(None, func_dictget(d, "B"))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
