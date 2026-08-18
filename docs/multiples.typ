#import "/src/lib.typ": zh

#set page(width: auto, height: auto, margin: 1em)

#set text(
  lang: "zh",
  region: "CN",
  font: "Source Han Serif",
  weight: "bold",
  top-edge: "ascender",
  bottom-edge: "descender",
)

#set align(center)

#let cell(length, body, fill: none, small: false, size: auto, height: auto) = box(
  width: length,
  height: if height == auto { length } else { height },
  align(center + horizon, box(
    // Transparent border
    width: 100% - 0.5pt,
    height: 100% - 0.5pt,
    fill: fill,
    align(center + horizon, text(
      if size == auto {
        zh(if small { "小" } + body)
      } else { size },
      body,
    )),
  )),
)
#let small = cell.with(small: true)

#let scale = 1.3

#let large-grid = grid(
  columns: 3,
  row-gutter: 0.5pt,
  grid.cell(colspan: 3, {
    let base = zh(0) * scale

    // https://wiki.evageeks.org/File:19_C338_guooo.jpg
    let cells = (
      cell(base, "初", fill: rgb("a357e9").lighten(30%)),
      cell(base / 2, "二", fill: rgb("85f458")),
      cell(base / 4, "五", fill: rgb("d55034").lighten(30%)),
      cell(base / 8, "七", fill: rgb("5c64eb").lighten(50%)),
    )

    grid(
      columns: 2,
      cells.at(0),
      grid(
        columns: 2,
        cells.at(1), cells.at(1),
        grid(
          columns: 2,
          cells.at(2), grid(columns: 2, ..(cells.at(3),) * 4),
          cells.at(2), cells.at(2),
        ),
        cells.at(1),
      ),
    )
  }),
  {
    let base = zh(1) * scale

    // https://wiki.evageeks.org/File:22_C259D_zero.jpg
    let cells = (
      cell(base, "一", fill: rgb("#7490fe")),
      cell(base / 2, "四", fill: rgb("e0cbb2")),
    )

    grid(
      columns: 1,
      cells.at(0),
      grid(
        columns: 2,
        ..(cells.at(1),) * 2
      ),
    )
  },
  {
    let base = zh(3) * scale

    // https://wiki.evageeks.org/File:19_C338_guooo.jpg
    let cells = (
      cell(base, "三", fill: rgb("e02321").lighten(20%)),
      cell(base / 2, "六", fill: rgb("fdd865")),
      cell(base / 4, "八", fill: rgb("aaaad0").lighten(30%)),
    )

    grid(
      columns: 1,
      cells.at(0),
      grid(
        columns: 2,
        grid(
          columns: 2,
          ..(cells.at(2),) * 4
        ),
        ..(cells.at(1),) * 3,
      ),
    )
  },
  cell(zh(0) * scale, [*号*], size: zh(0), height: zh(1) * scale * 3 / 2),
)

#let small-grid = grid(
  columns: 3,
  column-gutter: 0.5pt,
  grid.cell(colspan: 2, cell((zh(-3) + zh(-1)) * scale, [*小*], size: zh(0), height: zh("-0") * scale)),
  grid.cell(rowspan: 2, {
    let base = zh("-0") * scale

    // https://wiki.evageeks.org/File:19_C338_guooo.jpg
    let cells = (
      small(base, "初", fill: rgb("a357e9").lighten(30%)),
      small(base / 2, "二", fill: rgb("85f458")),
      small(base / 4, "五", fill: rgb("d55034").lighten(30%)),
    )

    grid(
      columns: 1,
      cells.at(0),
      grid(
        columns: 2,
        grid(
          columns: 2,
          ..(cells.at(2),) * 4
        ),
        ..(cells.at(1),) * 3,
      ),
    )
  }),
  {
    let base = zh(-3) * scale

    // https://wiki.evageeks.org/File:19_C338_guooo.jpg
    let cells = (
      small(base, "三", fill: rgb("e02321").lighten(20%)),
      small(base / 2, "六", fill: rgb("fdd865")),
    )

    grid(
      columns: 1,
      cells.at(0),
      grid(
        columns: 2,
        ..(cells.at(1),) * 4
      ),
    )
  },
  {
    let base = zh(-1) * scale

    // https://wiki.evageeks.org/File:22_C259D_zero.jpg
    let cells = (
      small(base, "一", fill: rgb("#7490fe")),
      small(base / 2, "四", fill: rgb("e0cbb2")),
    )

    grid(
      columns: 1,
      cells.at(0),
      grid(
        columns: 2,
        ..(cells.at(1),) * 2
      ),
    )
  },
)

#grid(
  columns: 2,
  column-gutter: 2em,
  large-grid,
  grid.cell(align: bottom, box(
    align(top, small-grid),
  )),
)
