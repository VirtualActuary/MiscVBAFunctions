from pathlib import Path
import mkdocs_gen_files


def main():
    repo_dir = Path("..")
    nav = mkdocs_gen_files.Nav()

    # Inject the repo's README.md file into the documentation as the starting page.
    nav["README.md"] = "index.md"
    with mkdocs_gen_files.open("index.md", "w") as dst:
        with repo_dir.joinpath("README.md").open("r") as src:
            dst.write(
                """---
title: README.md
---
"""
            )
            dst.write(src.read())

    # Generate docs from source code and inject navigation for generated files.
    # https://github.com/mkdocstrings/mkdocstrings/blob/5802b1ef5ad9bf6077974f777bd55f32ce2bc219/docs/gen_doc_stubs.py
    for src_dir in [
        repo_dir.joinpath("MiscVBAFunctions"),
        repo_dir.joinpath("MiscVBAFunctionsLib"),
    ]:
        for src_path in sorted(src_dir.rglob("*.bas")):
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
