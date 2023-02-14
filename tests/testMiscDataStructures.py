import unittest
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscCollectionCreate(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            with self.subTest("Test_EnsureUniqueKey_Col"):
                # func_EnsureUniqueKey = book.macro("MiscDataStructures.EnsureUniqueKey")
                func = book.macro("Test__Helper_MiscDataStructures.Test_EnsureUniqueKey_Col")
                self.assertTrue(func())


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
