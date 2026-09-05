#import "@preview/lure:0.2.0"
#show heading.where(level: 1): it => pagebreak(weak: true) + it

= 附录：号数名称认同规则
各个号数有很多名称，不同数据源的写法未必一致，同一数据源多次枚举时也未必一致。

为简洁，本文件的号数名称统一按「四号」「小初」形式。例如4号、四、四号、四號都认作是同一号数，而小初、小初号、新初号也认作是同一号数；行文时若不关心原始形式，则分别把它们统一记作「四号」与「小初」。

号数名称认同原则与示例如下。各原则矛盾时，优先考虑靠前的原则。

// 以下参考了 Unicode core specification §18.1.5 (Han) Unification Rules
// https://www.unicode.org/versions/Unicode17.0.0/core-spec/chapter-18/#G29313

+ *不区分字形差异。*像「号/號」「一/壹」等字，都不作区分。

  - @source:秀英1903 的一号用「壹」但二号并未用「贰」，只在notes中指出，而不作区分。

+ *保留原文差异。*若某数据源每次枚举号数都保持多种写法，则这些写法应予以区分。

  - @source:津报厂1971 在一张表格中同时出现了「大一号」与「一号」，那么二者应予以区分。
  - @source:何继曾1959 在一张表格中同时出现了「小四号」与「七号小」，后续解释时前者始终保持「小四号」，而后者始终保持「七号（小）」或「七小」，那么七号小、七号（小）、七小可认作相同，但它们不可理解为小七。

+ *接受原文认同规则。*若某数据源多次枚举号数，前后几次的不同写法明显指相同号数，则这些写法可认作相同。若数据源正文直接描述某几种写法指相同号数，则也可接受。

  - @source:曹洪奎1979 的36页写「有的也把点数铅字按号数字的大小，叫成小几号或是新几号」，那么「小□号」与「新□号」均可认作相同。
  - @source:手册1989 的40页写「……制成七种号数的铅字，即：一号字、二号字……，后又添制了大于一号的初号字……」，而41页表格枚举「……初，大，二……」，那么一号、大号可认作相同。
  - @source:王益1946 的12页正文写「頭號、二號……」，而同页插图标「特號」「大號」「二號」等，那么头号、大号也可认作相同。

  注意认同规则不可类推。一号、大号认作相同，小一、小大认作相同，并不代表特大号、大初号能当作「特一号」「一初号」理解。

+ *其余情况变通决定。*名称相似可作为认同依据，对应点数不同可作为区分依据。名称既相似，对应点数又不同时，按具体情况综合考虑。

  - @source:Ken-2-JP 是英文资料，正文写0G是初号，0G small是小初，但未明说5G、4G small等。按常理判断，把5G、4G small等理解为五号、小四等。
  - @source:沪一厂1978 表格中出现了「七行（特大号）」「五行（特号）」「四行（初号）」，但「特中号」「小特号」「小初号」未注括号。考虑到@source:沪一厂1972 写作「特大」「特」「初」与「特中」「小特」「小初」，为方便对比，@source:沪一厂1978 按特大、特号、初号而非七行、五行、四行记录。
  - @source:津报厂1971 有七倍、六倍、五倍，但并不像其它数据源的七行、六行、五行那样线度成 $7:6:5$ 比例，所以「□倍」应与「□行」区分。
  - @source:jawiki-新\只有「新□号」，并且对应点数与@source:jawiki-旧\的「□号」很接近。因此@source:jawiki-新\的「新□号」当作「□号」而非「小□号」处理，并在notes中注明。

根据以上原则，总结出以下具体规则。注意以下并不适用于@source:jawiki-新\这种notes另外注明规则的数据源；此外「老□号」与「□号」也应认作相同，不过各数据源描述号数与点数映射关系时，恰好无一写作「老□号」，因此省略。
#v(1em)
#{
  import "data.typ": g-raw
  import "util.typ": is-small, normalize-g

  // 从g的标准形式映射到原始形式
  let norm = (:)
  for raw in g-raw {
    if raw.starts-with("\\") {
      continue // 忽略天元
    }

    let g = normalize-g(raw)
    if g not in norm {
      norm.insert(g, ())
    }
    norm.at(g).push(raw)
  }

  // 人为指定号数顺序
  let cells = (
    ("初号",) + range(10).map(n => numbering("一号", n + 1)),
    ("小初",) + range(7).map(n => numbering("小一", n + 1)),
    ("大初", "大一", "七号大", "七号小"),
    ("特大", "特中", "特初", "特号", "大特", "小特"),
    ("七行", "六行", "五行", "四行"),
    ("七倍", "六倍", "五倍"),
  )
  let g-expected-list = cells.flatten()
  assert.eq(g-expected-list.filter(g => g not in norm), ())
  assert.eq(norm.keys().filter(g => g not in g-expected-list), ())

  set text(0.8em)
  let cell(n) = {
    assert(0 <= n and n < cells.len())
    context {
      set par(spacing: par.leading, hanging-indent: 3em)
      for g in cells.at(n) {
        strong[#g：]
        norm
          .at(g)
          // 调整出现顺序和折行位置
          .sorted(key: raw => ("大" in raw, raw))
          .map(raw => {
            if raw in ("头号", "小大") {
              linebreak()
            }
            raw
          })
          .join[、]
        parbreak()
      }
    }
  }
  grid(
    columns: (auto, 1fr),
    column-gutter: 3em,
    row-gutter: 2em,
    ..(0, 1).map(cell),
    grid.cell(colspan: 2, cell(2)),
    cell(3),
    grid(columns: (1fr,) * 2, ..(4, 5).map(cell)),
  )
}

= 附录：相关文献 <sec:bibliography>
有些文献是数据源，前面已经引用；而另一些文献主要是分析历史、讲解如何辨别活字印刷品，前面未必引用过。
#v(1em)

#set par(justify: false)
#bibliography(title: none, full: true, "ref.bib", style: "gb-7714-2015-author-date")

= 附录：外部链接索引
单击索引中的👆可以跳转到原文位置。
#v(1em)

#context {
  let targets = query(selector(std.link).after(<outline>).before(<sec:bibliography>))
    .filter(it => type(it.dest) == str)
    .map(it => (
      it,
      {
        let host = lure.parse(it.dest).host.trim("www.", at: start)
        (host,)

        let m = lure.parse-supplementary(it.dest).path-segments.last().match(regex("\.([^.]+)$"))
        if host != "doi.org" and host != "github.com" and m != none {
          let suffix = upper(m.captures.first())
          if suffix not in ("html", "jhtml", "php", "pl", "csp", "dtx", "def").map(upper) {
            assert(
              suffix in ("pdf", "zip").map(upper),
              message: "please add the unknown suffix (" + suffix + ") to either the ignore list or the allow list",
            )
            (suffix,)
          }
        }
      },
    ))
    .sorted(key: ((_, (host, ..))) => host.split(".").rev())

  {
    // 检查重复链接
    let expected-duplicates = (
      // 这条评论回复了多个问题，分属两个数据源，因此链接了两回
      "https://github.com/CTeX-org/ctex-kit/issues/813#issuecomment-4412583072": 2,
      // 这个PR既加了行距，又改了字号，因此被链接了两回
      "https://github.com/tailwindlabs/tailwindcss/pull/2609": 2,
    )
    let urls = targets.map(((it, ..)) => it.dest)
    let duplicate-urls = urls
      .enumerate()
      .filter(((i, a)) => i != urls.position(b => b == a))
      .map(((.., url)) => url)
      .dedup()
    assert.eq(
      duplicate-urls
        .map(url => (url, targets.filter(((it, ..)) => it.dest == url)))
        .filter(((url, targets)) => targets.len() != expected-duplicates.at(url, default: none))
        .to-dict(),
      (:),
    )
    assert.eq(expected-duplicates.keys().filter(k => k not in duplicate-urls), ())
  }

  // Export for `just check-web-archive`
  [#metadata(targets.map(((it, ..)) => it.dest).dedup())<external-links>]

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
此外，整理本文件曾得多人相助。有些未在正文中说明，此处再补充并感谢一些。（因不确定署名方式，仅列出 GitHub 用户页 URL 或社交媒体昵称的 SHA256。）

- `f8e77f37a8a91148a394a39410233cc9899bb26d8aa3fc691004924be2bdda01`

  对数据源详情的data图提出了改进意见。

- `29244d7582331f9860965319ae734ab8be517089116caed702779f3427048edb`

  提示了GB/T 12200.2—1994的蹊跷之处。

#if "log" in sys.inputs [
  = 附录：更新记录

  #show raw: set text(lang: "zh", cjk-latin-spacing: auto)
  #show raw: set strong(delta: 0) // Disable strong
  #figure(raw(sys.inputs.log, lang: "gitlog", block: true))
]
