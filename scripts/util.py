import contextlib
import os
import shutil
import tempfile
from pathlib import Path
import locate
from zebra_vba_packager import Source, Config

app_name = "MiscVBAFunctions"
app_short_name = "MiscF"

app_dir = locate.this_dir().parent.joinpath(app_name)
app_lib_dir = locate.this_dir().parent.joinpath(f"{app_name}Lib")


@contextlib.contextmanager
def get_app_clean():
    with tempfile.TemporaryDirectory() as f:
        tmp = Path(f, "tmp")
        output = Path(f, "output")

        conf = Config(
            Source(
                path_source=app_dir,
                auto_bas_namespace=False,
                combine_bas_files=app_short_name,
                glob_exclude=["**/test_*"]
            )
        )
        conf.run(tmp)

        os.makedirs(output)
        for i in tmp.rglob("*"):
            if str(i.name)[-4:] in (".txt", ".bas") and not i.name.startswith("z__"):
                shutil.move(i, Path(output, i.name))

        with Path(output, "thisworkbook.txt").open() as f:
            txt = f.read()

        yield output


