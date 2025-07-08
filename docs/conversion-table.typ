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
    (("-0", 0.5), "小初"),
    ..range(1, 9)
      .map(n => (
        (n, numbering("一号", n)),
        if n < 7 {
          ((-n, n + 0.5), numbering("小一", n))
        } else { (none, none) },
      ))
      .join(),
  )
    .map(((number, name)) => if number != none {
      let numbers = if type(number) == array { number } else { (number,) }
      let n = numbers.first()
      assert.eq(numbers.map(x => zh(x)).dedup(), (zh(n),))

      (
        {
          let fmt = x => raw("zh(x)".replace("x", repr(x)), lang: "typc")
          set par(leading: 0.25em)
          numbers.map(fmt).intersperse(linebreak()).join()
        },
        [#zh(n)],
        text(zh(n), name),
      )
    } else { (none,) * 3 })
    .flatten(),
  table.hline(),
)
