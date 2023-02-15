import unittest
from xlwings import Book
from locate import prepend_sys_path

with prepend_sys_path():
    from util import functions_book


class MiscString(unittest.TestCase):
    def test_1(self) -> None:
        book: Book
        with functions_book() as book:
            with self.subTest("randomString"):
                func_randomString = book.macro("MiscString.randomString")
                self.assertEqual(4, len(func_randomString(4)))
                self.assertNotEqual(func_randomString(5), func_randomString(5))

            with self.subTest("Test_EndsWith"):
                func_EndsWith = book.macro("MiscString.EndsWith")

                self.assertTrue(func_EndsWith("foo bar baz", " baz"))
                self.assertTrue(func_EndsWith("foo bar baz", "az"))
                self.assertFalse(func_EndsWith("foo bar baz", " baz "))
                self.assertFalse(func_EndsWith("foo bar baz", "bar"))

            with self.subTest("Test_StartsWith"):
                func_StartsWith = book.macro("MiscString.StartsWith")

                self.assertTrue(func_StartsWith("foo bar baz", "foo "))
                self.assertTrue(func_StartsWith("foo bar baz", "foo bar baz"))
                self.assertFalse(func_StartsWith("foo bar baz", "bar"))
                self.assertFalse(func_StartsWith("foo bar baz", " Foo"))


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
