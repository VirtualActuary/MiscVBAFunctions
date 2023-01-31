import unittest
from .util import functions_book, vba_dict


class TestMin(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            with self.subTest("DictsToArray"):
                func_col = book.macro("MiscCollectionCreate.Col")
                func_DictsToArray = book.macro("MiscDictsToArray.DictsToArray")

                d1 = vba_dict({"a": 1, "b": 2, "c": 3})
                d2 = vba_dict({"a": 11, "b": 22, "c": 33})
                c1 = func_col(d1, d2)

                arr = func_DictsToArray(c1)

                self.assertEqual(
                    (
                        ("a", "b", "c"),
                        (1, 2, 3),
                        (11, 22, 33),
                    ),
                    arr,
                )


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
