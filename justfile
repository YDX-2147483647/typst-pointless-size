# Prerequisites:
# - https://just.systems
# - mkdir, 7z, etc.
# - https://mikefarah.gitbook.io/yq
# - https://github.com/chmln/sd

VERSION := `yq .package.version typst.toml`

# List available recipes
@default:
    just --list

[private]
target-dir:
    @mkdir -p target/

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
[env("TYPST_FEATURES", "bundle")]
[group("docs")]
[working-directory("ref")]
ref-check:
    #!/usr/bin/env nu
    typstyle . --inplace --column 120

    let _ = (
        typst eval
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

# Watch the document and recompile to one-pdf format on changes
[env("TYPST_FEATURES", "bundle")]
[group("docs")]
ref-watch: target-dir
    typst watch ref/main.typ target/ref.pdf --open

# Export the document to all supported formats
[env("TYPST_FEATURES", "bundle")]
[group("docs")]
ref-build: target-dir
    typst compile \
        --input revision=$(git describe --tags --dirty) \
        --input log="$(git log --pretty=format:'commit %H%nAuthor: %an%nDate:   %ad%n%n%w(0,2,2)%B%w(0,0,0)' --date iso ref/)" \
        ref/main.typ target/ref.pdf
    typst compile \
        --input revision=$(git describe --tags --dirty) \
        --input log="$(git log --pretty=format:'commit %H%nAuthor: %an%nDate:   %ad%n%n%w(0,2,2)%B%w(0,0,0)' --date iso ref/)" \
        --input mode=split-pdf \
        --format bundle \
        ref/main.typ target/ref/
