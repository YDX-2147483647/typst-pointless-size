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
    @mkdir -p target/assets/

# Run tests
test: target-dir
    typst compile src/zihao.test.typ - --format svg > /dev/null
    typst compile docs/conversion-table.typ target/assets/conversion-table.svg --root .
    typst compile docs/multiples.typ target/assets/multiples.svg --root .

# Create package.7z for submission
package: target-dir
    sd --fixed-strings './' 'https://github.com/YDX-2147483647/typst-pointless-size/blob/v{{VERSION}}/' README.md
    7z a target/package.7z LICENSE README.md typst.toml src/ -x!src/*test*

# Build target/ for GitHub Pages
gh-pages: test ref-build
    typst compile docs/export-readme.typ target/index.html --root . --features html

# Format *.typ and check for unused files
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

# Check the status of web archive snapshots
[group("docs")]
check-web-archive: target-dir
    #!/usr/bin/env nu
    use std/assert

    let headers = {
      # https://archive.org/developers/bots.html#user-agent-requirements
      User-Agent: "pointless-size/2026-08-21"
    }

    let checked = (
      try { open target/web-archive.toml | get record.url } catch { [] }
      | append (try { open ref/unarchived-external-links.txt | lines } catch { [] })
    )
    print --stderr $"已加载之前获取的存档状态，共计 ($checked | length) 条。"

    for full_url in (
      typst eval --in ref/main.typ 'query(<external-links>).first().value' --format yaml | from yaml
    ) {
      let parsed = ($full_url | url parse | reject fragment)
      let url = ($parsed | reject params | url join)

      # 避免随机重定向，方便检查存档状态
      let url = ($url | str replace https://mirrors.cernet.edu.cn/ https://mirrors.tuna.tsinghua.edu.cn/)
      # 已有专门版本控制
      if ($parsed.host == "github.com" and ($parsed.path | path split | get 2 --optional) in [commit, blob]) {
        continue
      }
      # 本身就是档案性质
      if ($parsed.host == "www.unicode.org" and $parsed.path == "/cgi-bin/GetUnihanData.pl") {
        continue
      }
      if ([
        # 本身就是档案性质
        archive.org,
        openlibrary.org,
        doi.org,
        taiwanebook.ncl.edu.tw,
        worldwide.espacenet.com,

        # 已有专门版本控制
        wikimedia.org,
        wikipedia.org,
        svn.tug.org,

        # 域名不定，难以存档；且有大量备份，不必存档
        annas-archive.gl,
        z-lib.sk,

        # 不支持存档
        ss.zhizhen.com,
        opac.nlc.cn,
        instagram.com,
      ] | any {|suffix| $parsed.host | str ends-with $suffix }) {
        continue
      }

      if ($url in $checked) {
        print --stderr $"🟦 ($url) 存档状态已知。"
        continue
      }

      print --stderr $"🔎 检查 ($url)……"

      # 优先考虑 web.archive.org，不过只有 archive.today 支持微信公众号文章
      if ($parsed.host != "mp.weixin.qq.com") {
          let result = (http get $"https://archive.org/wayback/available?({ url: $url } | url build-query)" --headers $headers)

          if ($result.archived_snapshots.closest?.available | default false) {
            print --stderr "🟢 已经存档。"
            print $result.archived_snapshots.closest
            {
              record: [{
                url: $url,
                archive_snapshot: $result.archived_snapshots.closest.url,
              }]
            } | to toml | tee { save --append target/web-archive.toml }
          } else {
            print --stderr "😡 尚未存档。"
            $"($url)\n" | save --append ref/unarchived-external-links.txt
          }
      } else {
        try {
          let result = http get $"https://archive.today/timemap/($url)" --headers $headers
          let snapshot = ($result | parse --regex '\n<(?<snapshot>[^>]+)>; rel="first last memento"; datetime=".+",' | get 0.snapshot)
          print --stderr "🟢 已经存档。"
          print $result
          {
            record: [{
              url: $url,
              archive_snapshot: $snapshot,
            }]
          } | to toml | tee { save --append target/web-archive.toml }
        } catch {
            print --stderr "😡 尚未存档。"
            $"($url)\n" | save --append ref/unarchived-external-links.txt
        }
      }

      # https://archive.org/developers/bots.html#rate-limiting
      sleep 1sec
    }

    # 必须人工检查，因为有些页面的出链也要一并存档
    print --stderr "请查看 ref/unarchived-external-links.txt。"
