import unittest
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book, vba_dict


class TestMin(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            with self.subTest("dictget"):
                func_dictget = book.macro("MiscDictionary.dictget")
                func_dict = book.macro("MiscDictionaryCreate.dict")

                d1 = func_dict("a", 2, "b", "foo")
                self.assertEqual(2, func_dictget(d1, "a"))
                self.assertEqual("foo", func_dictget(d1, "b"))
                self.assertEqual(
                    None, func_dictget(d1, "c", None)
                )  # vbNullString's numerical value???

            with self.subTest("Concat_Dicts"):
                func_ConcatDicts = book.macro("MiscDictionary.ConcatDicts")

                d1 = vba_dict({"a": 2, "b": "foo"})
                d2 = vba_dict({"v": "bar", "d": 3})
                d3 = vba_dict({"2": 2, "4": 4})

                func_ConcatDicts(d1, d2, d3)
                self.assertEqual(2, d1["a"])
                self.assertEqual("foo", d1["b"])
                self.assertEqual("bar", d1["v"])
                self.assertEqual(3, d1["d"])
                self.assertEqual(2, d1["2"])
                self.assertEqual(4, d1["4"])

            with self.subTest("Join_Dicts"):
                func_JoinDicts = book.macro("MiscDictionary.JoinDicts")
                func_dictget = book.macro("MiscDictionary.dictget")

                d1 = vba_dict({"a": 2, "b": "foo"})
                d2 = vba_dict({"v": "bar", "d": 3})
                d3 = vba_dict({"2": 2, "4": 4})

                d4 = func_JoinDicts(d1, d2, d3)

                self.assertEqual(2, func_dictget(d4, "a"))
                self.assertEqual("foo", func_dictget(d4, "b"))
                self.assertEqual("bar", func_dictget(d4, "v"))
                self.assertEqual(3, func_dictget(d4, "d"))
                self.assertEqual(2, func_dictget(d4, "2"))
                self.assertEqual(4, func_dictget(d4, "4"))

            with self.subTest("Test_dictget_fail"):
                func = book.macro("Test__Helper_MiscDictionary.Test_dictget_fail")
                self.assertTrue(func())


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
