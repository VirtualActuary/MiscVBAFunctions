import sys
from pathlib import Path
from tempfile import TemporaryDirectory

from aa_py_xl import vba_object_to_python_object

from .util import TestCaseWithFunctionBook


class TestSubprocess(TestCaseWithFunctionBook):
    def test_RunAndCapture(self) -> None:
        func = self.book.macro("Subprocess.RunAndCapture")

        # Write a script that outputs to stdout and to stderr and has a non-zero exit code.
        with TemporaryDirectory() as tmp_dir_str:
            script_path = Path(tmp_dir_str, "tmp.py")
            script_path.write_text(
                r"""\
import sys

print("Hello World")
print("Nice to meet you!")

print("Goodbye World", file=sys.stderr)
print("It's been fun!", file=sys.stderr)

sys.exit(3)
"""
            )

            result_vba = func(f"{sys.executable} {script_path}")

        result_py = vba_object_to_python_object(result_vba)

        self.assertEqual(
            {
                "code": 3,
                "stderr": ("Goodbye World", "It's been fun!"),
                "stdout": ("Hello World", "Nice to meet you!"),
            },
            result_py,
        )

    def test_WhereIsExe(self) -> None:
        func = self.book.macro("Subprocess.WhereIsExe")
        self.assertEqual(r"C:\Windows\System32\cmd.exe", func("cmd.exe"))
        self.assertEqual("", func("KdQM8kP0f1lDH1k8KpjY"))

    def test_EscapeAndWrapCmdArg(self) -> None:
        func = self.book.macro("Subprocess.EscapeAndWrapCmdArg")

        # Empty string
        self.assertEqual('""', func(""))

        # Simple string
        self.assertEqual('"foo"', func("foo"))

        # Spaces
        self.assertEqual('"foo bar"', func("foo bar"))

        # Quotes
        self.assertEqual(r'"foo\"bar"', func('foo"bar'))

        # Backslashes
        self.assertEqual(r'"foo\bar"', func(r"foo\bar"))

        # Backslash before quote
        self.assertEqual(r'"foo\\\"bar"', func(r"foo\"bar"))

        # Multiple backslashes before quote
        self.assertEqual(r'"foo\\\\\"bar"', func(r'foo\\"bar'))

        # Trailing backslash
        # https://docs.python.org/3/faq/design.html#why-can-t-raw-strings-r-strings-end-with-a-backslash
        self.assertEqual(r'"foo\"', func("foo\\"))

        # Only quote
        self.assertEqual(r'"\""', func('"'))

        # Only backslash
        # https://docs.python.org/3/faq/design.html#why-can-t-raw-strings-r-strings-end-with-a-backslash
        self.assertEqual(r'"\"', func("\\"))

        # Quote at start and end
        self.assertEqual(r'"\"foo\""', func('"foo"'))

        # Complex combination
        self.assertEqual(r'"foo \ \"bar\" baz"', func(r'foo \ "bar" baz'))

    def test_FindBestTerminal(self) -> None:
        func = self.book.macro("Subprocess.FindBestTerminal")
        result = func()
        self.assertIsInstance(result, str)
        if len(result):
            self.assertTrue(Path(result).exists())
