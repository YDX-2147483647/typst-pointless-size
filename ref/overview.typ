#import "@preview/lilaq:0.6.0" as lq
#import "util.typ": is-small
#import "data.typ": data
#import "template.typ": title-page

#title-page[
  = 概览
  概览展示全部数据源的情况。

  依次有三种*展示方式*：「映射表」侧重每一号数对应几点，「点数图」侧重有哪些点数被号数覆盖，「分支树」侧重数据源间的相近程度。前两者相对客观，而分支树十分主观，请酌情参考。另外注意三种方式各有特点，所以其中数据源的顺序不完全相同。

  每个*数据源*都有*代号*，例如@source:基准、@source:方正书版、@source:王选1982 等。代号只是为了方便沟通，命名方式并不统一。具体含义可单击绿色文字跳转到后页查看。

  绘制点数等图时，会将*号数*分为两个*系列*：「□号系列」包括初号、一号、二号……和特大、特中、特初、特号、小特，而「小□系列」包括小初、小一、小二……和七行、六行、五行。这种分法并没有特殊道理，只是让绘出的图方便比较。
]
#set heading(offset: 1)

// 定义输出顺序
#let source-keys = (
  "基准",
  "CTeX新",
  "CLReq-main",
  "沪新厂TTK",
  "商务厂TTK",
  "王益1946",
  "Ken-2-JP",
  "jawiki-旧",
  "神田",
  "东京1942",
  "曹洪奎1979",
  "曹洪奎1975",
  "CLReq-extra",
  "周承民1988",
  "叶重光1996",
  "京新厂1981",
  "科学1978",
  "沪一厂1972",
  "沪一厂1978",
  "沪一厂1988",
  "丹江厂1975",
  "丹江厂1980",
  "手册1989",
  "中华厂TTK",
  "小史1981",
  "王选1982",
  "方正书版",
  "朱永和1999",
  "石家庄2001",
  "方正跨媒介",
  "方正飞某",
  "新CCT",
  "老CCT",
  "华丰厂1963",
  "天元",
  "政府",
  "enwiki",
  "jawiki-新",
  "Ken-1",
  "何继曾1959",
  "申报1935",
  "秀英1903",
  "姜别利",
)
#assert.eq(source-keys.first(), data.keys().first())
#assert.eq(source-keys.sorted(), data.keys().sorted())

#page(width: auto, {
  [= 映射表]

  let palette = (
    first-seen: orange.darken(15%),
    mode: purple,
  )
  set math.cancel(stroke: gray + 0.5pt, inverted: true)

  figure({
    /// (source, G) ↦ P
    let to-p(source, g) = data.at(source).to-dict().at(g, default: none)

    /// G ↦ an array of partial cells
    let format-g(g) = {
      // 1. 输出号数
      (g,)

      let base-source = source-keys.first()
      let base-p = to-p(base-source, g)

      // pairs: An array of (source, p, first-seen)
      // count: A map from str(p) to count
      let pairs = ()
      let count = (:)
      for source in source-keys {
        let p = to-p(source, g)
        let first-seen = p != none and str(p) not in count

        if p != none {
          count.insert(str(p), count.at(str(p), default: 0) + 1)
        }
        pairs.push((source, p, first-seen))
      }

      // 2. 输出点数
      for (source, p, first-seen) in pairs {
        let result = if p == none {
          // N/A
          $dot$
        } else if source != base-source and p == base-p {
          // Same as base
          $=$
        } else {
          // Other
          if first-seen {
            // Make it special if first seen
            text(palette.first-seen, $#p$)
          } else {
            $#p$
          }
        }

        if (
          p != none
            and count.at(str(p)) == 1
            // Other values are all at least 3% away
            and count.keys().find(p-str => str(p) != p-str and calc.abs(float(p-str) - p) / p <= 0.03) == none
        ) {
          // Dim alone values
          result = math.cancel(result)
        }

        (result,)
      }

      // 3. 输出点数取值统计
      let n-max = calc.max(..count.values())
      (
        math.mat(
          ..array.zip(..count
            .pairs()
            .map(((p, n)) => ($#p$, $times #n$).map(
              x => if n == n-max {
                text(palette.mode, weight: "bold", x)
              } else {
                x
              },
            ))),
        ),
      )

      // 4. 再次输出号数
      (g,)
    }

    let row(..g-list) = {
      assert.eq(g-list.named(), (:))

      array
        .zip(..g-list.pos().map(format-g))
        .map(cells => table.cell({
          // If all partial cells are equal and nearly empty, merge into one.
          // Otherwise, join with slashes.
          if cells.first() in ($=$, $dot$) and cells.all(c => c == cells.first()) {
            cells.first()
          } else {
            cells.join(" / ")
          }
        }))
    }

    table(
      columns: source-keys.len() + 3,
      table.hline(),
      table.header(
        table.cell(rowspan: 2)[*号数*],
        table.cell(colspan: source-keys.len())[*点数*],
        table.hline(stroke: 0.5pt),
        table.cell(rowspan: 2)[*点数取值统计*],
        table.cell(rowspan: 2)[*号数*],
        ..source-keys.map(s => strong(ref(label("source:" + s)))),
      ),
      table.hline(),
      ..row("初号", "小初"),
      ..row("二号", "小二"),
      ..row("五号", "小五"),
      ..row("七号", "小七"),
      table.hline(stroke: 0.5pt),
      ..row("一号", "小一"),
      ..row("四号", "小四"),
      table.hline(stroke: 0.5pt),
      ..row("三号", "小三"),
      ..row("六号", "小六"),
      ..row("八号"),
      table.hline(stroke: 0.5pt),
      ..row("特大", "特号", "小特"),
      ..row("特中", "特初"),
      ..row("七行", "六行"),
      ..row("五行", "四行"),
      ..row("七号大", "七号小"),
      ..row("大一", "九号", "十号"),
      table.hline(),
    )
  })

  align(center, grid(
    columns: (24em,) * 3,
    gutter: 48em,
    align: start,
    [
      *点数：*数字表示点数，“$=$”表示同@source:基准，“$dot$”表示无明文直接提及（可能无定义，也可能有定义但省略了）；#text(palette.first-seen)[变色]表示这一点数在相应号数中是首次出现（指表格左方从未出现，与年份无关），#math.cancel(text(palette.first-seen)[划线])表示这一点数在相应号数中只出现了一次，并且没有与之相差 $3%$ 或更少的其它点数。
    ],
    [
      *各列顺序：*大致按相近程度排列，比较主观。比如@source:方正书版\与@source:朱永和1999 的数值有很多都不相等，可以认为区别很大；但二者定义范围相同，后者数值适当取整又完全与前者一致（例 $36.25 approx 36$），所以也可认为二者十分接近。
    ],
    [
      *点数取值统计：*第一行是点数取值，第二行是出现次数；#text(palette.mode)[*加粗变色*]对应众数（假设各数据源权重相等）。统计时忽略“$dot$”，所以取值种数始终等于#text(palette.first-seen)[变色数字]的数量，但总出现次数未必等于数据源数量。
    ],
  ))
})

/// `move-before(arr, ((x, y),))` moves `x` from its original place to the place directly before `y`
#let move-before(arr, pairs) = {
  for (value, before) in pairs {
    arr.remove(arr.position(s => s == value))
    arr.insert(arr.position(s => s == before), value)
  }
  return arr
}
// 调整输出顺序，适应展示方式
// - 表格用“=”省略了大量内容，而图中没有省略
// - 表格侧重数字是否相等，而图中侧重数字是否相近
// - 表格逐号数比较点数，而图中忽略号数只管点数
#let source-keys = move-before(source-keys, (
  ("政府", "基准"),
  ("enwiki", "基准"),
  ("石家庄2001", "小史1981"),
  ("华丰厂1963", "沪一厂1972"),
)).rev() // 反转纵轴

= 点数图
#figure({
  let large-x = ()
  let large-y = ()
  let small-x = ()
  let small-y = ()
  for (y, source) in source-keys.enumerate() {
    for (g, p) in data.at(source) {
      if is-small(g) {
        small-x.push(p)
        small-y.push(y)
      } else {
        large-x.push(p)
        large-y.push(y)
      }
    }
  }

  let xaxis-args = (
    label: [点数（对数尺度）],
    scale: "log", // Use log scale but linear ticks
    subticks: none,
  )
  show: lq.set-tick(inset: 2pt, outset: 2pt, pad: 0.25em)

  let format-ticks(..rows, gutter: 1em / 3, top: false) = {
    assert.eq(rows.named(), (:))
    for (shift, values) in rows.pos().enumerate() {
      for p in values {
        let tick = {
          if top { $#p$ }
          v(shift * gutter)
          if not top { $#p$ }
        }
        ((p, tick),)
      }
    }
  }
  // □号
  let large-ticks = format-ticks(
    (5.25, 10.5, 21, 42),
    (7, 14, 28, 56),
    (4, 8, 16, 32),
  )
  // 小□
  let small-ticks = format-ticks(
    (9, 18, 36, 45, 54, 63),
    (6, 12, 24),
    top: true,
  )

  lq.diagram(
    height: 1.5em * source-keys.len(),
    width: 100%,
    legend: (position: (100% + 0.5em, 2em)),
    yaxis: (
      label: [数据源],
      ticks: source-keys.map(s => ref(label("source:" + s))).enumerate(),
      subticks: none,
    ),

    // □号网格
    xaxis: (ticks: large-ticks) + xaxis-args,
    // 五号系列网格
    lq.vlines(5.25, 10.5, 21, 42, stroke: 1.5pt + luma(80%)),

    // 小□网格
    lq.xaxis(ticks: small-ticks, position: top, ..xaxis-args),
    lq.vlines(
      ..small-ticks.map(array.first),
      stroke: (paint: luma(80%), thickness: 0.5pt, dash: "loosely-dashed"),
    ),
    // 两网格都必须在所有数据之下

    // □号数据
    lq.scatter(label: [□号系列], large-x, large-y, mark: "o"),
    // 小□数据
    lq.scatter(label: [小□系列], small-x, small-y, mark: "s3"),
  )
})

#pagebreak()
#include "overview-tree.typ"
