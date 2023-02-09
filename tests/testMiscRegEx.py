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

            with self.subTest("RenameVariableInFormula_2"):
                func = book.macro("MiscRegEx.RenameVariableInFormula")

                self.assertEqual(
                    'FOO + bar - foobar "foo bar foobar"',
                    func('foo + bar - foobar "foo bar foobar"', "foo", "FOO"),
                )

                self.assertEqual(
                    'foo + bar - foobar "foo bar foobar"',
                    func('foo + bar - foobar "foo bar foobar"', "fo", "FOO"),
                )

                self.assertEqual(
                    'foo + bar - FOOBAR "foo bar foobar"',
                    func('foo + bar - foobar "foo bar foobar"', "foobar", "FOOBAR"),
                )

                self.assertEqual(
                    'FOO + bar - foobar "foo bar foobar"',
                    func('foo + bar - foobar "foo bar foobar"', "FOO", "FOO"),
                )

                self.assertEqual(
                    'foo + BAR - foobar "foo bar foobar"',
                    func('foo + bar - foobar "foo bar foobar"', "bar", "BAR"),
                )

                self.assertEqual(
                    'foo+BAR-foobar "foo bar foobar"',
                    func('foo+bar-foobar "foo bar foobar"', "bar", "BAR"),
                )

                self.assertEqual(
                    'FOO + bar - foobar "FOO bar foobar"',
                    func('foo + bar - foobar "foo bar foobar"', "foo", "FOO", False),
                )

                self.assertEqual(
                    'foo + bar - foobar "foo bar foobar"',
                    func('foo + bar - foobar "foo bar foobar"', "fo", "FOO", False),
                )

                self.assertEqual(
                    'foo + bar - FOOBAR "foo bar FOOBAR"',
                    func(
                        'foo + bar - foobar "foo bar foobar"', "foobar", "FOOBAR", False
                    ),
                )

                self.assertEqual(
                    'FOO + bar - foobar "FOO bar foobar"',
                    func('foo + bar - foobar "foo bar foobar"', "FOO", "FOO", False),
                )

                self.assertEqual(
                    'foo + BAR - foobar "foo BAR foobar"',
                    func('foo + bar - foobar "foo bar foobar"', "bar", "BAR", False),
                )

                self.assertEqual(
                    'foo+BAR-foobar "foo BAR foobar"',
                    func('foo+bar-foobar "foo bar foobar"', "bar", "BAR", False),
                )


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
