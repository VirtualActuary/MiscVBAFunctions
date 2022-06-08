import shutil
import tempfile
from pathlib import Path
from zebra_vba_packager import decompile_xl, is_locked, pack, backup_last_50_paths
import locate

locate.allow_relative_location_imports(".")
from util import get_app_clean, app_name  # noqa

# Filenames
app_xl = locate.this_dir().joinpath(f"../{app_name}.xlsb")
app_lib_dir = locate.this_dir().joinpath(f"../{app_name}Lib")
app_dir = Path(str(app_xl)[:-5])

# Backup the directory
for i in [app_dir, app_lib_dir]:
    if i.exists():
        backup_last_50_paths(Path(tempfile.gettempdir(), f"{app_name}-compile-backups"), i)

# Decompile (and remove zebra files)
decompile_xl(app_xl, app_dir)

with get_app_clean() as tmp:
    shutil.rmtree(app_lib_dir, ignore_errors=True)
    shutil.copytree(tmp, app_lib_dir)


