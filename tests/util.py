import shutil
from contextlib import contextmanager
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Mapping, Any, Generator, Union

from aa_py_xl import excel, excel_book
from locate import this_dir
from xlwings import Book, App

repo_path = this_dir().parent
functions_book_path = repo_path.joinpath("MiscVBAFunctions.xlsb")
template_book_path = repo_path.joinpath("MiscVBATemplate.xlsb")


@contextmanager
def functions_book(app: App = None) -> Generator[Book, None, None]:
    if not app:
        with excel(
            path=functions_book_path,
            save=False,
            quiet=True,
            close_book=True,
            close_excel=True,
            must_exist=True,
            read_only=True,
        ) as book:
            try:
                yield book
            finally:
                pass
    else:
        with excel_book(
            app,
            path=functions_book_path,
            save=False,
            close=True,
            must_exist=True,
            read_only=True,
        ) as book:
            try:
                yield book
            finally:
                pass


@contextmanager
def template_book() -> Generator[Book, None, None]:
    with excel(
        path=template_book_path,
        save=False,
        quiet=True,
        close_book=True,
        close_excel=True,
        must_exist=True,
        read_only=True,
    ) as book:
        try:
            yield book
        finally:
            pass


@contextmanager
def extra_book(app: App, path: Union[str, Path]) -> Generator[Book, None, None]:
    with excel_book(
        app,
        path=path,
        save=False,
        close=True,
        must_exist=True,
        read_only=True,
    ) as book:
        try:
            yield book
        finally:
            pass


@contextmanager
def tmp_book() -> Generator[Book, None, None]:
    with TemporaryDirectory() as tmp_dir_str:
        tmp_dir = Path(tmp_dir_str)
        with excel(
            path=tmp_dir.joinpath("tmp.xlsx"),
            save=False,
            quiet=True,
            close_book=True,
            close_excel=True,
            must_exist=False,
            read_only=False,
        ) as book:
            try:
                yield book
            finally:
                pass


def vba_dict(d: Mapping[str, Any]) -> Any:
    """
    Create a VBA dictionary.

    See https://stackoverflow.com/questions/67397267/pass-dictionary-to-excel-macro-using-win32com-and-comtypes
    """
    import win32com.client

    result = win32com.client.Dispatch("Scripting.Dictionary")

    for key, value in d.items():
        result[key] = value

    return result


@contextmanager
def temporary_copy_of_workbook(template_path: Path) -> Generator[Path, None, None]:
    with TemporaryDirectory() as tmp_dir:
        wb_path = Path(tmp_dir, "out.xlsx")
        shutil.copyfile(src=template_path, dst=wb_path)
        try:
            yield wb_path
        finally:
            pass
