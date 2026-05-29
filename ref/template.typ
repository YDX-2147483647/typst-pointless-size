#import "@preview/lure:0.2.0"

#let template(title: [], date: [], notes: [], body) = {
  set document(title: title)
  set page(height: auto, numbering: "1 / 1", header: counter(footnote).update(0))

  set par(justify: true)

  set text(lang: "zh", region: "CN", font: "Source Han Serif", top-edge: "ascender", bottom-edge: "descender")
  show raw: set text(font: ("DejaVu Sans Mono", "Source Han Serif"))
  show math.equation: set text(font: ("New Computer Modern Math", "Source Han Serif"))
  set list(marker: ([•], [‣], [–]).map(
    text.with(font: "Libertinus Serif"),
  ))
  show smartquote: set text(features: ("pwid",))
  show emph: set text(font: ("Libertinus Serif", "Source Han Serif"))

  show std.title: set text(1.2em)
  show heading.where(level: 1): set align(center)
  show heading.where(level: 1): set text(1.2em)
  show heading.where(level: 1): set block(spacing: 1em)

  show link: set text(blue.darken(20%))
  show cite: set text(green.darken(30%))
  show ref: it => {
    let el = it.element
    if el != none and el.func() == heading {
      link(el.location(), {
        set text(green.darken(30%))
        el.body
      })
    } else {
      it // Unchanged
    }
  }

  set footnote(numbering: "①")
  set table(stroke: none, align: center + horizon)

  set quote(block: true)
  show quote: block.with(width: 100%, stroke: (left: gray + 2pt), inset: (left: 0.5em), outset: (y: 0.5em))

  align(center, {
    std.title()
    date

    if "revision" in sys.inputs {
      linebreak()
      raw(sys.inputs.revision)
    }
  })

  outline(title: none, depth: 2)

  v(1em)
  notes

  body
}

#let link(dest, ..body) = {
  assert.eq(body.named(), (:))
  assert(body.pos().len() <= 1)

  std.link(
    lure.normalize(dest),
    body.pos().at(0, default: dest.replace(regex("^(mailto|tel)://"), "")),
  )
}

#let title-page(body) = page({
  show strong: underline.with(stroke: aqua + 0.2em, offset: 0.1em, background: true, evade: false)
  show figure: fig => {
    show underline: it => it.body // figure 内保持不变
    fig
  }

  body

  v(1em)
})
