import unittest

from aa_py_xl import automatically_click_vba_error

from tests.util import functions_book


class TestCompile(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            with automatically_click_vba_error():
                book.macro("compile")()
                self.assertTrue(book.macro("isCompiled")())
