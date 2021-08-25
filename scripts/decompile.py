import shutil
import tempfile
from pathlib import Path
from zebra_vba_packager import decompile_xl, is_locked, pack
import locate

locate.allow_relative_location_imports(".")
from util import backup_last_50_files, get_app_clean, app_name  # noqa

# Filenames
app_xl = locate.this_dir().joinpath(f"../{app_name}.xlsb")
app_lib_dir = locate.this_dir().joinpath(f"../{app_name}Lib")
app_dir = Path(str(app_xl)[:-5])

# Backup the directory
for i in [app_dir, app_lib_dir]:
    if is_locked(i):
        raise ValueError(f"Dir '{i}' cannot be overwritten.")

    with tempfile.TemporaryDirectory() as outdir:
        if i.is_dir():
            pack(i, zipname := Path(outdir).joinpath(i.name+".7z"))
            backup_last_50_files(zipname)

# Decompile (and remove zebra files)
decompile_xl(app_xl, app_dir)

with get_app_clean() as tmp:
    shutil.rmtree(app_lib_dir, ignore_errors=True)
    shutil.copytree(tmp, app_lib_dir)


