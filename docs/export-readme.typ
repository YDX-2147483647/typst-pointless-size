#set document(
  title: [Typst Pointless Size——字号 zìhào],
  description: [中文字号的号数制及字体度量单位 Chinese size system (hào-system) and type-related measurements units],
)
#let gh-pages = "https://ydx-2147483647.github.io/typst-pointless-size/"

#html.style(
  ```css
  main {
    max-width: 36em;
    margin-inline: auto;
    padding-inline: 1.5em;

    margin-block: 4em;

    line-height: 1.8;
    pre {
      line-height: unset;
    }

    img, pre {
      max-width: 100%;
      overflow-x: auto;
    }
  }
  ```.text,
)

#show: html.main

#import "@preview/cmarker:0.1.10"
#cmarker.render(
  read("/README.md"),
  h1-level: 0,
  set-document-title: false,
  scope: (
    image: (source, alt: none) => {
      html.img(
        src: if source.starts-with(gh-pages) {
          "./"
          source.trim(gh-pages, at: start)
        } else {
          source
        },
        ..if alt != none { (alt: alt) },
      )
    },
    link: (dest, body) => {
      let href = if dest.starts-with("./") {
        "https://github.com/YDX-2147483647/typst-pointless-size/blob/HEAD/"
        dest.trim("./", at: start)
      } else if dest.starts-with(gh-pages) {
        "./"
        dest.trim(gh-pages, at: start)
      } else {
        dest
      }
      html.a(
        href: href,
        ..if href.starts-with("https://") {
          (target: "_blank", rel: "noopener")
        },
        body,
      )
      if body.func() == html.elem and body.tag == "img" and not body.attrs.src.starts-with("https://img.shields.io") {
        parbreak()
      }
    },
  ),
)
