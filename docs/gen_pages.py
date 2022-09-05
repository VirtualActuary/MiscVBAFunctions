from pathlib import Path
import mkdocs_gen_files


def main():
    repo_dir = Path("..")
    nav = mkdocs_gen_files.Nav()

    # Copy the template so that we can link to it from the docs.
    with mkdocs_gen_files.open("MiscVbaTemplate.xlsb", "wb") as dst:
        with open("../MiscVbaTemplate.xlsb", "rb") as s:
            dst.write(s.read())

    nav["Introduction"] = "index.md"
    with mkdocs_gen_files.open("index.md", "w") as dst:
        dst.write(
            """---
title: MiscVba
---
Click [here](MiscVbaTemplate.xlsb) to download the template.
"""
        )

        # Inject the repo's README.md file into the documentation as the starting page.
        with repo_dir.joinpath("README.md").open("r") as src:
            dst.write(src.read())

    # Generate docs from source code and inject navigation for generated files.
    # https://github.com/mkdocstrings/mkdocstrings/blob/5802b1ef5ad9bf6077974f777bd55f32ce2bc219/docs/gen_doc_stubs.py
    for src_dir in [
        repo_dir.joinpath("MiscVBAFunctions"),
    ]:
        for src_path in [
            i
            for i in sorted(src_dir.rglob("*.bas"))
            if not i.name.lower().startswith("test__")
        ]:
            doc_path = src_path.relative_to(repo_dir).with_suffix(".md")

            nav[src_path.relative_to(repo_dir).parts] = doc_path.as_posix()

            with mkdocs_gen_files.open(doc_path, "w") as f:
                f.write(
                    f"""\
---
title: '{src_path.name}'
---

# `{src_path.relative_to(repo_dir)}`

::: {src_path}
    handler: vba
"""
                )

    with mkdocs_gen_files.open("nav.md", "w") as f:
        f.writelines(nav.build_literate_nav())


main()
