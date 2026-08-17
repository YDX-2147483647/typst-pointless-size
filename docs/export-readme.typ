#set document(
  title: [Typst Pointless Size——字号 zìhào],
  description: [中文字号的号数制及字体度量单位 Chinese size system (hào-system) and type-related measurements units],
)

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
      let gh-pages = "https://ydx-2147483647.github.io/typst-pointless-size/"
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
      html.a(
        href: if dest.starts-with("./") {
          "https://github.com/YDX-2147483647/typst-pointless-size/blob/HEAD/"
          dest.trim("./", at: start)
        } else {
          dest
        },
        ..if dest.starts-with("https://") {
          (target: "_blank", rel: "noopener")
        },
        body,
      )
      if body.func() == html.elem and body.tag == "img" {
        parbreak()
      }
    },
  ),
)
