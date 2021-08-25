from pathlib import Path
from zebra_vba_packager import compile_xl, is_locked, runmacro_xl
import locate
from util import app_name, get_app_clean


# Filenames
locate.allow_relative_location_imports(".")
from util import backup_last_50_files, app_lib_dir  # noqa

app_xl = locate.this_dir().joinpath(f"../{app_name}.xlsb")
app_lib_xl = locate.this_dir().joinpath(f"../{app_name}Lib.xlsb")
app_dir = Path(str(app_xl)[:-5])

for i in [app_xl, app_lib_xl]:
    if is_locked(i):
        raise ValueError(f"File '{i}' cannot be overwritten.")
    backup_last_50_files(i)

# Compile and run early bindings
compile_xl(app_dir, app_xl)
with get_app_clean(next(app_dir.rglob("*.xlsx"))) as app_lib_dir:
    compile_xl(app_lib_dir, app_lib_xl)

runmacro_xl(app_xl)
runmacro_xl(app_lib_xl)
