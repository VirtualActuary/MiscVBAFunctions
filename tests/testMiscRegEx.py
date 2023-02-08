import unittest
from xlwings import Book
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscRegEx(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("RenameVariableInFormula"):
                func_RenameVariableInFormula = book.macro(
                    "MiscRegEx.RenameVariableInFormula"
                )
                self.assertEqual(
                    "xyz + a ^ xyz + foo(xyz) + abc + abc1 /xyz",
                    func_RenameVariableInFormula(
                        "Ab + a ^ ab + foo(AB) + abc + abc1 /aB", "ab", "xyz"
                    ),
                )


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
