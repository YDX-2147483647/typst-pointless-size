# Prerequisites:
# - https://just.systems
# - mkdir, ln, etc.
# - https://mikefarah.gitbook.io/yq
# - https://github.com/chmln/sd

INSTALL_NAME := replace(data_directory(), '\', '/') + "/typst/packages/local/pointless-size"
VERSION := `yq .package.version typst.toml`

# Derived variables
INSTALL_DIR := INSTALL_NAME + '/' + VERSION

# List available recipes
@default:
    just --list

# Install the library to @local by creating a symlink
install: && check-install
    mkdir --parents {{ INSTALL_NAME }}
    ln --symbolic {{ replace(source_directory(), '\', '/') }} {{ INSTALL_DIR }}

# Remove the library installed to @local
uninstall:
    rm --interactive {{ INSTALL_DIR }}

[private]
target-dir:
    @mkdir -p target/

# Check the library is importable
[private]
check-install:
    echo '#import "@local/pointless-size:{{ VERSION }}"' \
    | typst compile - - --format svg > /dev/null

# Run tests
test:
    typst compile src/zihao.test.typ - --format svg > /dev/null
    typst compile docs/conversion-table.typ - --format svg --root . > /dev/null
    typst compile docs/multiples.typ - --format svg --root . > /dev/null

# Create package.7z for submission
package: target-dir
    sd --fixed-strings './' 'https://github.com/YDX-2147483647/typst-pointless-size/blob/main/' README.md
    7z a target/package.7z LICENSE README.md typst.toml src/ -x!src/*test*

# Format *.typ and check for unused files
[group("docs")]
[working-directory("ref")]
ref-check:
    #!/usr/bin/env nu
    typstyle . --inplace --column 120

    let _ = (
        typst-dev eval
        --in main.typ
        --input assets=(ls assets/ --short-names | get name | to json)
        --input files=(ls | where type == file | get name | to json)
        '
        let assets = json(bytes(sys.inputs.assets))
        let files = json(bytes(sys.inputs.files))

        let used-assets = query(image).map(i => i.source.trim("assets/", at: start)).dedup()
        let attached-files = query(pdf.attach).map(i => i.path).dedup()

        let unused-assets = assets.filter(x => x not in used-assets)
        let unused-files = files.filter(x => x not in attached-files and x != "main.pdf")

        assert.eq(unused-assets, (), message: "Found unused assets: " + repr(unused-assets))
        assert.eq(unused-files, (), message: "Found unused files: " + repr(unused-files))
        '
    )

# Export the document to PDF
[group("docs")]
ref-build: target-dir
    typst compile --input revision=$(git describe --tags --dirty) ref/main.typ target/ref.pdf
