# Misc VBA Functions

A collection of standalone VBA functions to use in other projects.

## Getting started

1. Download and install the latest release from [Github](https://github.com/VirtualActuary/MiscVBAFunctions/releases).
2. In your start menu, open MiscVba for the documentation.
3. Download the template from the home page of the documentation.

## Add your code

1. Insert a new module in VBA
2. Use prefix `Fn` to access the MiscVba functions. For example:<br>
`Set LO = Fn.GetLO("table")`

## Developer notes

Read this is you are a developer contributing to MiscVba.

### Compiling and decompiling

The code in this repo is stored in a decompiled state, as `.bas` files. This can be combined into a file called
[MiscVBAFunctions.xlsb](MiscVBAFunctions.xlsb) using the [compile.cmd](compile.cmd) script. After editing
[MiscVBAFunctions.xlsb](MiscVBAFunctions.xlsb), it can be decompiled back into `.bas` files using the
[decompile.cmd](decompile.cmd) script.

### Coding standards

- See https://docs.google.com/spreadsheets/d/1nnPorllRq35TcZDrsksJLFC5HBvHUOkP3ihLK1sgVOI/edit#gid=2120525226
- Run `black .` from the repo root regularly.
- Run `mypy .` from the repo root regularly.

### Running tests

From the repo root:

```
python -m unittest
```

### Releasing a new version

1. Ensure all tests pass and all coding standards are met.
2. Update dependencies: `app-builder -d`
3. Make a local release and test it: `app-builder -l`
4. Make a release on GitHub: `app-builder -g`
