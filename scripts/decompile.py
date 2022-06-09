import shutil
import tempfile
from pathlib import Path
from zebra_vba_packager import decompile_xl, is_locked, pack, backup_last_50_paths
import locate

app_name = "MiscVBAFunctions"
app_xl = locate.this_dir().joinpath(f"../{app_name}.xlsb")
app_dir = Path(str(app_xl)[:-5])

# Backup the directory
if app_dir.exists():
    backup_last_50_paths(Path(tempfile.gettempdir(), f"{app_name}-compile-backups"), app_dir)

# Decompile (and remove zebra files)
decompile_xl(app_xl, app_dir)
