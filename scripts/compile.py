import tempfile
from pathlib import Path
from zebra_vba_packager import (
    compile_xl,
    runmacro_xl,
    backup_last_50_paths,
    Source,
    Config,
)
import locate

app_name = "MiscVBAFunctions"
app_xl = locate.this_dir().joinpath(f"../{app_name}.xlsb")
app_xl_lib = locate.this_dir().joinpath(f"../MiscVBATemplate.xlsb")
app_dir = Path(str(app_xl)[:-5])

if app_xl.exists():
    backup_last_50_paths(
        Path(tempfile.gettempdir(), f"{app_name}-compile-backups"), app_xl
    )

# Compile into Book
compile_xl(app_dir, app_xl)

# Compile Lib into Book
with tempfile.TemporaryDirectory() as outdir:
    Config(
        Source(
            path_source=app_dir,
            glob_exclude="**/Test__*",
            combine_bas_files="Fn",
            auto_cls_rename=False,
        ),
        casing="pascal",
    ).run(outdir)

    compile_xl(outdir, app_xl_lib)

# Ensure early bindings
runmacro_xl(app_xl)
runmacro_xl(app_xl_lib)
