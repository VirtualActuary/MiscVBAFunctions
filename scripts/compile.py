import shutil
import tempfile
from pathlib import Path
from zebra_vba_packager import compile_xl, is_locked, runmacro_xl, backup_last_50_paths
import locate

app_name = "MiscVBAFunctions"
app_xl = locate.this_dir().joinpath(f"../{app_name}.xlsb")
app_dir = Path(str(app_xl)[:-5])

if app_xl.exists():
    backup_last_50_paths(Path(tempfile.gettempdir(), f"{app_name}-compile-backups"), app_xl)

# Compile and run early bindings
compile_xl(app_dir, app_xl)
runmacro_xl(app_xl)
