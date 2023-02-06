import unittest
from pathlib import Path
import os
from locate import prepend_sys_path
with prepend_sys_path():
    from util import functions_book

class TestPath(unittest.TestCase):
    def test_1(self) -> None:
        with functions_book() as book:
            func_Path = book.macro("MiscPath.Path")
            func_IsAbsolutePath = book.macro("MiscPath.IsAbsolutePath")
            func_col = book.macro("MiscCollectionCreate.Col")
            func_PathGetDrive = book.macro("MiscPath.PathGetDrive")
            func_PathHasDrive = book.macro("MiscPath.PathHasDrive")
            func_PathGetServer = book.macro("MiscPath.PathGetServer")
            func_PathHasServer = book.macro("MiscPath.PathHasServer")
            func_AbsolutePath = book.macro("MiscPath.AbsolutePath")
            func_EvalPath = book.macro("MiscPath.EvalPath")

            with self.subTest("Test_path_1"):
                # Empty path should return empty path
                self.assertEqual("", func_Path(""))
                self.assertEqual("", func_Path([""]))
                self.assertEqual("", func_Path(func_col("")))

            with self.subTest("Test_path_2"):
                # Normalize slashes to '\'.
                self.assertEqual(r"C:\foo\bar", func_Path("C:", "foo", "bar"))
                self.assertEqual(r"C:\foo\bar", func_Path("C:/", "foo", "bar"))
                self.assertEqual(r"C:\foo\bar", func_Path("C:\\", "foo", "bar"))
                self.assertEqual(r"C:\foo\bar", func_Path("C:", "foo/bar/"))
                self.assertEqual(r"C:\foo\bar", func_Path("C:", "foo\\bar\\"))
                self.assertEqual(
                    r"foo\bar\asdf\zxcv", func_Path("foo/bar", r"asdf\zxcv")
                )
                self.assertEqual(
                    r"foo\bar\asdf\zxcv", func_Path("foo//bar", r"asdf\\zxcv")
                )

            with self.subTest("Test_path_3"):
                # Path segments starting with drive letters throw away what comes before them.
                self.assertEqual(r"D:\bar", func_Path("C:", "foo", "D:", "bar"))
                self.assertEqual("D:", func_Path("C:", "foo", "D:"))
                self.assertEqual("E:", func_Path("C:", "D:", "E:"))

            with self.subTest("Test_path_4"):
                # Path segments starting with slashes throw away what comes before them, but preserve the drive letter.
                self.assertEqual(r"C:\bar", func_Path("C:", "foo", "/bar"))
                self.assertEqual("C:", func_Path("C:", "foo", "/"))

            with self.subTest("Test_path_5"):
                # File extensions are preserved
                self.assertEqual(
                    r"folder\file.extension", func_Path("folder", "file.extension")
                )

            with self.subTest("Test_path_6"):
                # Single path segment is allowed
                self.assertEqual(
                    r"C:\folder1\folder2\folder3",
                    func_Path("C:/folder1/folder2/folder3"),
                )

            with self.subTest("Test_path_7"):
                # Collection is allowed instead of multiple arguments
                self.assertEqual(
                    r"foo\bar\baz", func_Path(func_col("foo", "bar", "baz"))
                )

            with self.subTest("Test_path_8"):
                # Array is allowed instead of multiple arguments
                self.assertEqual(r"foo\bar\baz", func_Path(["foo", "bar", "baz"]))

            with self.subTest("Test_path_9"):
                # Network paths
                self.assertEqual(r"\\server1\foo", func_Path(r"\\server1", "foo"))

            with self.subTest("Test_path_10"):
                # Root paths without drive letters
                self.assertEqual(r"\foo\bar", func_Path("/foo", "bar"))

            with self.subTest("Test_IsAbsolutePath"):
                self.assertEqual(r"\foo\bar", func_Path("/foo", "bar"))
                self.assertTrue(func_IsAbsolutePath("C:"))
                self.assertTrue(func_IsAbsolutePath("C:\\"))
                self.assertTrue(func_IsAbsolutePath("C:/"))
                self.assertTrue(func_IsAbsolutePath(r"\asdf"))
                self.assertTrue(func_IsAbsolutePath("/asdf"))
                self.assertTrue(func_IsAbsolutePath("\\"))
                self.assertTrue(func_IsAbsolutePath("/"))

                self.assertFalse(func_IsAbsolutePath("asdf"))
                self.assertFalse(func_IsAbsolutePath("asdf/"))
                self.assertFalse(func_IsAbsolutePath("asdf\\"))
                self.assertFalse(func_IsAbsolutePath("asdf/foo"))
                self.assertFalse(func_IsAbsolutePath(r"asdf\foo"))

            with self.subTest("Test_PathGetDrive"):
                self.assertEqual("C:", func_PathGetDrive("C:"))
                self.assertEqual("C:", func_PathGetDrive("C:\\"))
                self.assertEqual("C:", func_PathGetDrive("C:/"))
                self.assertEqual("C:", func_PathGetDrive(r"C:\asdf"))
                self.assertEqual("C:", func_PathGetDrive("C:/asdf"))
                self.assertEqual("", func_PathGetDrive(r"\asdf"))
                self.assertEqual("", func_PathGetDrive("/asdf"))
                self.assertEqual("", func_PathGetDrive("\\"))
                self.assertEqual("", func_PathGetDrive("/"))
                self.assertEqual("", func_PathGetDrive("asdf"))
                self.assertEqual("", func_PathGetDrive("asdf/"))
                self.assertEqual("", func_PathGetDrive("asdf\\"))
                self.assertEqual("", func_PathGetDrive("asdf/foo"))
                self.assertEqual("", func_PathGetDrive(r"asdf\foo"))
                self.assertEqual("", func_PathGetDrive(r"\\server1\foo"))
                self.assertEqual("", func_PathGetDrive("//server1/foo"))

            with self.subTest("Test_PathGetDrive"):
                self.assertTrue(func_PathHasDrive("C:"))
                self.assertTrue(func_PathHasDrive("C:\\"))
                self.assertTrue(func_PathHasDrive("C:/"))
                self.assertTrue(func_PathHasDrive(r"C:\asdf"))
                self.assertTrue(func_PathHasDrive("C:/asdf"))
                self.assertFalse(func_PathHasDrive(r"\asdf"))
                self.assertFalse(func_PathHasDrive("/asdf"))
                self.assertFalse(func_PathHasDrive("\\"))
                self.assertFalse(func_PathHasDrive("/"))
                self.assertFalse(func_PathHasDrive("asdf"))
                self.assertFalse(func_PathHasDrive("asdf/"))
                self.assertFalse(func_PathHasDrive("asdf\\"))
                self.assertFalse(func_PathHasDrive("asdf/foo"))
                self.assertFalse(func_PathHasDrive(r"asdf\foo"))
                self.assertFalse(func_PathHasDrive(r"\\server1\foo"))
                self.assertFalse(func_PathHasDrive("//server1/foo"))

            with self.subTest("Test_PathGetServer"):
                self.assertEqual("", func_PathGetServer("C:"))
                self.assertEqual("", func_PathGetServer("C:\\"))
                self.assertEqual("", func_PathGetServer("C:/"))
                self.assertEqual("", func_PathGetServer(r"C:\asdf"))
                self.assertEqual("", func_PathGetServer("C:/asdf"))
                self.assertEqual("", func_PathGetServer(r"\asdf"))
                self.assertEqual("", func_PathGetServer("/asdf"))
                self.assertEqual("", func_PathGetServer("\\"))
                self.assertEqual("", func_PathGetServer("/"))
                self.assertEqual("", func_PathGetServer("asdf"))
                self.assertEqual("", func_PathGetServer("asdf/"))
                self.assertEqual("", func_PathGetServer("asdf\\"))
                self.assertEqual("", func_PathGetServer("asdf/foo"))
                self.assertEqual("", func_PathGetServer(r"asdf\foo"))
                self.assertEqual(r"\\server1", func_PathGetServer(r"\\server1\foo"))
                self.assertEqual(r"//server1", func_PathGetServer(r"//server1/foo"))

            with self.subTest("Test_PathHasServer"):
                self.assertFalse(func_PathHasServer("C:"))
                self.assertFalse(func_PathHasServer("C:\\"))
                self.assertFalse(func_PathHasServer("C:/"))
                self.assertFalse(func_PathHasServer("C:\\asdf"))
                self.assertFalse(func_PathHasServer("C:/asdf"))
                self.assertFalse(func_PathHasServer("\\asdf"))
                self.assertFalse(func_PathHasServer("/asdf"))
                self.assertFalse(func_PathHasServer("\\"))
                self.assertFalse(func_PathHasServer("/"))
                self.assertFalse(func_PathHasServer("asdf"))
                self.assertFalse(func_PathHasServer("asdf/"))
                self.assertFalse(func_PathHasServer("asdf\\"))
                self.assertFalse(func_PathHasServer("asdf/foo"))
                self.assertFalse(func_PathHasServer("asdf\\foo"))
                self.assertTrue(func_PathHasServer(r"\\server1\foo"))
                self.assertTrue(func_PathHasServer("//server1/foo"))

            with self.subTest("Test_AbsolutePath"):
                # Absolute paths
                self.assertEqual("C:\\", func_AbsolutePath("C:\\"))
                self.assertEqual(r"C:\test", func_AbsolutePath(r"C:\test"))
                self.assertEqual(r"C:\foo\bar", func_AbsolutePath(r"C:\foo\bar"))
                self.assertEqual(r"C:\bar", func_AbsolutePath(r"C:\foo\..\bar"))
                self.assertEqual(r"C:\foo", func_AbsolutePath(r"C:\foo\\bar\\.."))

                # Relative paths
                self.assertEqual(Path.cwd().drive + r"\foo", func_AbsolutePath(r"\foo"))
                self.assertEqual(
                    str(Path.cwd().parent) + r"\foo\bar", func_AbsolutePath("foo/bar", book)
                )
                self.assertEqual(
                    str(Path.cwd().parent) + r"\foo", func_AbsolutePath(r".\foo", book)
                )
                self.assertEqual(
                    str(Path.cwd().parent.parent) + r"\foo", func_AbsolutePath(r"..\foo", book)
                )
                self.assertEqual(
                    str(Path(str(Path.cwd()) + r"foo\..\..").resolve()), func_AbsolutePath(r"foo\..\..", book)
                )
                self.assertEqual(
                    str(Path(str(Path.cwd()) + r"foo//..//..").resolve()), func_AbsolutePath("foo//..//..", book)
                )

                # Network Paths
                self.assertEqual(r"\\foo\bar", func_AbsolutePath(r"\\foo/bar"))
                self.assertEqual(r"\\foo\bar", func_AbsolutePath(r"\\foo\\bar"))
                self.assertEqual(
                    r"\\foo\bar", func_AbsolutePath(r"\\foo\\.\bar\\..\\bar")
                )
                self.assertEqual(r"\\foo", func_AbsolutePath(r"//foo"))
                self.assertEqual(r"\\foo\bar", func_AbsolutePath(r"//foo\\bar"))
                self.assertEqual(
                    r"\\hello\2", func_AbsolutePath(r"\\hello\world\\..\2")
                )

            with self.subTest("Test_AbsolutePath"):
                self.assertEqual(r"C:\foo", func_EvalPath(r"C:\foo"))
                self.assertEqual(r"C:\foo", func_EvalPath(r"C:/foo"))
                self.assertEqual(r"C:\c", func_EvalPath(r"C:\a\..\b\..\c"))
                self.assertEqual(
                    os.environ["HOMEDRIVE"] + "\\Users\\" + os.environ["username"],
                    func_EvalPath(r"%HOMEDRIVE%\Users\%username%"),
                )
                self.assertEqual(
                    str(Path(str(Path.cwd().parent), r"foo\bar")), func_EvalPath(r"foo/bar")
                )
                self.assertEqual(
                    str(Path(str(Path.cwd().parent), "foo\\" + os.environ["username"])),
                    func_EvalPath(r"foo/%UserName%"),
                )


if __name__ == "__main__":
    unittest.main(
        failfast=True,
    )
