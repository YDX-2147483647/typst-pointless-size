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

#set par(justify: true)
= 附录：本文件的著作权情况

生成本文件的代码采用 #link("https://choosealicense.com/licenses/mit/")[MIT 许可证]开源，可前往 #link("https://github.com/YDX-2147483647/typst-pointless-size/tree/main/ref")[GitHub 仓库 YDX-2147483647/typst-pointless-size 的`ref/`文件夹]查看。

本文件引用的数据大约属于单纯事实或通用数表，没有独创性，不存在著作权。

本文件为了说明汉字号数与点数的映射关系，摘录了很多资料原文、图像，其著作权属于原作者。由于映射关系容易出现传抄错误（例如 @刘岱伟2001 转录@source:叶重光1996），资料本身意见也容易与当代外推混合（例如#link("https://github.com/CTeX-org/ctex-kit/issues/543")[前期讨论]），讨论号数与点数映射关系不得不给出资料原貌。考虑到不少文献对映射关系的描述散落多处而且不完全一致（例如@source:小史1981），甚至还存在@source:曹洪奎1979、@source:周承民1988 这种题名相同、内容高度接近而数据却有差异的情况，本文件扩大了某些资料的摘录片段，有问题可#link("https://github.com/YDX-2147483647/typst-pointless-size/issues")[联系修改]。

#let pt = $"pt"$
此外，整理本文件曾得多人相助。有些未在正文中说明，此处再补充并感谢一些。（因不确定署名方式，仅列出 GitHub 用户页 URL 的 SHA256。）

- `f8e77f37a8a91148a394a39410233cc9899bb26d8aa3fc691004924be2bdda01`

  对数据源详情的data图提出了改进意见。

#if "log" in sys.inputs [
  = 附录：更新记录

  #show raw: set text(lang: "zh", cjk-latin-spacing: auto)
  #show raw: set strong(delta: 0) // Disable strong
  #figure(raw(sys.inputs.log, lang: "gitlog", block: true))
]
