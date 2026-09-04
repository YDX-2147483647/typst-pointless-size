#import "sources.typ": sources
#import "data.typ": data
#import "visualize.typ": build-table, draw-as-log-period
#import "template.typ": doc, title-page

#let category-meta = (
  "1-spec": [工厂规格],
  "1-code": [软件源码],
  "1-doc": [自身描述],
  "2-doc": [他人转述],
)

#doc("source-index.pdf", page[
  = 数据源索引
  本文件记录的数据源按年代与类别总结如下表，单击绿色文字可跳转到后页查看详情。注意@source:基准\未列入下表，因为它混合自多份资料，跨越多个年代且包含多个类别，无法收入表中。

  #figure({
    let simplify(edtf) = int(edtf.slice(0, 4).replace("X", "9"))
    assert.eq(("1990", "199X", "1996~").map(simplify), (1990, 1999, 1996))

    let group-meta = (
      "1-spec+1-code": ("1-spec", "1-code").map(c => category-meta.at(c)).join[与],
      "1-doc": category-meta.at("1-doc"),
      "2-doc": category-meta.at("2-doc"),
    )
    let groups = (
      "1-spec+1-code": (),
      "1-doc": (),
      "2-doc": (),
    )
    assert.eq(group-meta.keys(), groups.keys())

    // Collect into groups
    for (source, fields) in sources.pairs() {
      if "categories" not in fields {
        assert.eq(source, "基准")
        continue
      }

      let (group, edtf) = fields.categories
      if group in ("1-spec", "1-code") {
        group = "1-spec+1-code"
      }
      groups.at(group).push((source, simplify(edtf)))
    }
    for group in groups.keys() {
      // Sort by years
      groups.at(group) = groups.at(group).sorted(key: array.last)
    }

    let list-range(start, end) = {
      for items in groups.values() {
        let matched = for (source, year) in items {
          if year < start { continue }
          if year >= end { break }
          (source,)
        }
        if matched == none {
          (none,)
        } else {
          (matched.map(s => box(ref(label("source:" + s)))).join[、],)
        }
      }
    }
    set par(justify: false)
    table(
      columns: (auto, 1fr, 0.7fr, 1fr),
      align: start + top,
      table.hline(),
      table.header(
        table.cell(rowspan: 2, align: center + horizon)[*年代*],
        table.cell(colspan: groups.len(), h(1fr) + box(width: 0.7fr, inset: (x: 5%))[#h(1em)*类别*] + h(1fr)),
        table.hline(stroke: 0.5pt),
        ..group-meta.values().map(strong),
      ),
      table.hline(),

      [早期], ..list-range(1800, 1930),
      ..for start in range(1930, 2020, step: 10, inclusive: true) {
        (
          table.hline(stroke: 0.5pt),
          str(start).slice(0, 3) + "X",
          ..list-range(start, start + 10),
        )
      },
      table.hline(),
    )
  })

  #set terms(separator: h(1em, weak: true), hanging-indent: 3em)

  / 年代: 以上只是粗略划分。实际情况比较复杂：@source:zhwiki 的年代是段区间，@source:老CCT 的年代只是估算，@source:津报厂1971 的年代存在疑点，而@source:神田\的年代即使估算也存疑……详情页categories中会按 #link("https://www.loc.gov/standards/datetime/")[Extended Date Time Format (EDTF)] 给出更具体的描述。

  / 类别: *工厂规格、软件源码*是最原始的一手资料，不同时代体现为不同形式。*自身描述*是指工人、开发者、印刷厂、出版社、书店自身的描述，也是比较原始的一手资料。*他人转述*则是二手资料，通常更全面，但也更容易出现传抄错误与无依据外推。有些转述会承认号数与点数映射关系不统一，并给出多种版本。对于这种情况，本文件会在详情页notes中说明，合适时还会拆分为多个数据源。
])

#doc("source-title.pdf", title-page[
  = 数据源详情
  逐一展示每个数据源的字号定义范围与点数数值（data），介绍出处及来源、类别（brief、via、categories），并补充数据原貌、疑点等情况（notes）。

  data 部分提供了图、表两种展示方式。

  - *图*中标注了若干对号数、点数，按*传统竖排顺序*阅读（从上到下、从右到左）可*从大到小*读出。每点都位于一条*竖线*上，同一竖线上的最小点数不小于最大点数的一半，$10.5$ 所在竖线加粗。有些点严格落于一条*横线*上，但也有些点仅接近某条横线，同一横线上相邻两条竖线交点的点数是二倍关系，$10.5$ 所在加粗横线左右分别标有 $times 1$ 和 $times 2^0$。一点所在竖线下方点数*乘以*所在横线两侧*倍数*即为其点数。例如下图，三号 $16$ 位于 $10.5$ 竖线与 $times 3/2$ 横线交点上方一点点，对应 $16 gt.tilde 10.5 times 3/2 = 15.75$；而 $times 3/2$ 横线与 $times 2^(7\/12)$ 横虚线几乎重合，对应 $2^(7\/12) approx 1.49831 ≈ 1.50 = 3/2$（纯五度完全协和 🎶）。

    #figure(draw-as-log-period(data.CLReq-main, width: 20em, height: 12em))

  - *表*有号数、点数、*长度示意*三列，其中第三列由点数计算得出，方便确定点数之间的倍数关系。例如下表，初号 $42$ 两格长等于二号 $21$ 四格长，也等于一号 $28$ 三格长，所以初号、二号、一号线度之比是 $1/2 : 1/4 : 1/3 = 6:3:4$。

    #figure(build-table((初号: 42, 二号: 21, 五号: 10.5, 一号: 28, 四号: 14).pairs()))

    这些表大多提供了「*按倍数关系*」和「按点数大小」两种排列方式，前者会像上表这样将行分成几组。这种分组能方便检查对照，但*不代表数据源本身的意见*。同一组的点数也未必严格成倍数关系，未分组也不代表点数没有倍数关系。
])
#set heading(offset: 1)
#show heading.where(level: 2): it => pagebreak(weak: true) + it

#for (source, fields) in sources.pairs() [
  #show: doc.with("source-{}.pdf".replace("{}", source))

  #[= #source.replace("-", " ")] #label("source:" + source)

  == data
  #{
    let data = data.at(source, default: none)
    if data == none [
      TODO
    ] else {
      // 绘图
      figure(draw-as-log-period(data, width: 28em, height: 18em))

      // 写表
      {
        let sorted-data = data.sorted(key: ((_g, p)) => -p)
        figure(if data == sorted-data {
          build-table(data)
        } else {
          grid(
            columns: 2,
            row-gutter: 0.5em,
            column-gutter: 2em,
            [按倍数关系], [按点数大小],
            build-table(data), build-table(sorted-data),
          )
        })
      }
    }
  }

  #for (k, v) in fields.pairs() [
    == #k
    #if k == "categories" {
      let (category, edtf) = v
      [#category-meta.at(category)，#edtf]
    } else [#v]
  ]
]
