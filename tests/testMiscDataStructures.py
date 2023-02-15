import unittest
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book, vba_dict


class MiscCollectionCreate(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            with self.subTest("Test_EnsureUniqueKey_Col"):
                func = book.macro(
                    "Test__Helper_MiscDataStructures.Test_EnsureUniqueKey_Col"
                )
                self.assertTrue(func())

            with self.subTest("Test_EnsureUniqueKey_Dict"):
                func_EnsureUniqueKey = book.macro("MiscDataStructures.EnsureUniqueKey")
                D = vba_dict({"a": 1, "b": 1, "c": 1})
                D2 = vba_dict({"a": 1, "b": 1, "b1": 1})

                self.assertEqual("d", func_EnsureUniqueKey(D, "d"))
                self.assertEqual("b2", func_EnsureUniqueKey(D2, "b"))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
