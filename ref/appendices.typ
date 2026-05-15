#import "@preview/lure:0.2.0"
#show heading.where(level: 1): it => pagebreak(weak: true) + it

= 附录：相关文献 <sec:bibliography>
有些文献是数据源，前面已经引用；而另一些文献主要是分析历史、讲解如何辨别活字印刷品，前面未必引用过。
#v(1em)

#set par(justify: false)
#bibliography(title: none, full: true, "ref.bib", style: "gb-7714-2015-author-date")

= 附录：外部链接索引
单击索引中的👆可以跳转到原文位置。
#v(1em)

#context {
  let targets = query(selector(std.link).before(<sec:bibliography>))
    .filter(it => type(it.dest) == str)
    .map(it => (
      it,
      {
        let host = lure.parse(it.dest).host.trim("www.", at: start)
        (host,)

        let m = lure.parse-supplementary(it.dest).path-segments.last().match(regex("\.([^.]+)$"))
        if host != "doi.org" and m != none {
          let suffix = upper(m.captures.first())
          if suffix not in ("html", "jhtml", "php", "csp", "sty", "dtx", "cxx").map(upper) {
            (suffix,)
          }
        }
      },
    ))
    .sorted(key: ((_, (host, ..))) => host.split(".").rev())

  set list(spacing: 1em)
  set heading(outlined: false, bookmarked: true)
  let last-host = none
  for (it, tags) in targets {
    let host = tags.first().split(".").slice(-2).join(".")
    if host != last-host {
      [== #host]
      last-host = host
    }

    if tags.first() == host {
      let _ = tags.remove(0) // Omit redundancy
    }
    tags.push(std.link(it.location())[👆])

    [- #it (#tags.join(", "))]
  }
}
