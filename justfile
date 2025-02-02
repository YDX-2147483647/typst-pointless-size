# Prerequisites:
# - https://just.systems
# - mkdir, ln, etc.
# - https://mikefarah.gitbook.io/yq

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

# Check the library is importable
[private]
check-install:
    echo '#import "@local/pointless-size:{{ VERSION }}"' \
    | typst compile - - --format svg > /dev/null

# Run tests
test:
    typst compile src/zihao.test.typ - --format svg > /dev/null
