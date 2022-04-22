import shutil
import tempfile
from pathlib import Path
from zebra_vba_packager import compile_xl, is_locked, runmacro_xl, backup_last_50_paths
import locate

with locate.prepend_sys_path("."):
    from util import app_name, get_app_clean

app_xl = locate.this_dir().joinpath(f"../{app_name}.xlsb")
app_lib_xl = locate.this_dir().joinpath(f"../{app_name}Lib.xlsb")
app_dir = Path(str(app_xl)[:-5])

for i in [app_xl, app_lib_xl]:
    if i.exists():
        backup_last_50_paths(Path(tempfile.gettempdir(), f"{app_name}-compile-backups"), i)

# Compile and run early bindings
compile_xl(app_dir, app_xl)
with get_app_clean() as tmp:
    for i in app_dir.rglob("*.xlsx"):
        shutil.copy2(i, Path(tmp, i.name))
        
    compile_xl(tmp, app_lib_xl)

runmacro_xl(app_xl)
runmacro_xl(app_lib_xl)
