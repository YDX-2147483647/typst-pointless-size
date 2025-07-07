#import "/src/lib.typ": zh

#set page(width: auto, height: auto, margin: 1em)

#set text(lang: "zh", region: "CN", font: "Source Han Serif", top-edge: "ascender", bottom-edge: "descender")

#let header = ([*号数*], [*点数*], [*意义*])

#table(
  columns: header.len() * 2,
  align: left + horizon,
  stroke: none,
  table.hline(),
  ..header,
  table.vline(stroke: 0.5pt),
  ..header,
  table.hline(stroke: 0.5pt),
  ..(
    (0, "初号"),
    ("-0", "小初"),
    ..range(1, 9).map(n => (
      (n, numbering("一号", n)),
      if n < 7 {
        (-n, numbering("小一", n))
      } else { (none, none) },
    )),
  )
    .flatten()
    .chunks(2)
    .map(((n, t)) => if n != none {
      (
        raw("zh(" + repr(n) + ")", lang: "typc"),
        [#zh(n)],
        text(zh(n), t),
      )
    } else { (none,) * 3 })
    .flatten(),
  table.hline(),
)
