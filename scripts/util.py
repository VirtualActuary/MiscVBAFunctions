import contextlib
import itertools
import os
import shutil
import tempfile
from pathlib import Path
import locate
import uuid

app_name = "MiscVBAFunctions"
app_short_name = "MiscF"

app_dir = locate.this_dir().parent.joinpath(app_name)
app_lib_dir = locate.this_dir().parent.joinpath(f"{app_name}Lib")


def backup_last_50_files(fname):
    # Backup last 50 sheets
    import time

    backup = Path(tempfile.gettempdir()).joinpath(f"{app_name}-compile-backups")
    os.makedirs(backup, exist_ok=True)

    keep = sorted(backup.glob("*"))[-50:]
    for i in backup.glob("*"):
        if i not in keep:
            os.remove(i)

    timestr = time.strftime("%Y-%m-%d--%H-%M-%S")
    try:
        shutil.copy2(fname, backup.joinpath(f"{timestr}--{Path(fname).name}"))
    except:
        pass


@contextlib.contextmanager
def get_app_clean(xl_file=None):
    """
    Function to expose a "clean" version of the app directory
    """

    with tempfile.TemporaryDirectory() as outdir:
        flattened = Path(outdir).joinpath("flattened")
        output = Path(outdir).joinpath("output")
        os.makedirs(output, exist_ok=True)

        # Flatten the bas files into a single bas file
        shutil.copytree(app_dir, flattened)
        bas_txt = ""
        for i in flattened.rglob("*"):
            if os.path.splitext(i)[-1].lower() == ".bas" and i[:6].lower() != "test__":
                with open(i) as f:
                    bas_txt = bas_txt + f.read()
                os.remove(i)
        bas_txt = (
                      f'Attribute VB_Name = "{app_short_name}"\n'
                      'Option Explicit\n'
                  ) + bas_txt.replace(
            'Attribute VB_Name = "', "\n'************\""
        ).replace(
            'Option Explicit', ''
        )
        with open(flattened.joinpath(f"{app_short_name}.bas"), "w") as fw:
            fw.write(bas_txt)

        for i in itertools.chain(flattened.rglob("*.bas"), flattened.rglob("*.txt")):
            shutil.copy2(i, output.joinpath(i.name))

        if xl_file is not None:
            shutil.copy2(Path(xl_file), output.joinpath(Path(xl_file).name))

        yield output
