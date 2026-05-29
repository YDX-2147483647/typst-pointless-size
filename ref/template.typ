#import "@preview/lure:0.2.0"
#import "mode.typ": mode

/// A wrapper of `document(path, body)`.
#let doc(path, body) = if mode == "split-pdf" {
  document(path, body)
} else {
  body
}

#let palette = (
  // External links
  link: blue.darken(20%),
  // Internal links
  ref: green.darken(30%),
)

#let template(title: [], date: [], notes: [], body) = {
  let revision = sys.inputs.at("revision", default: none)

  set document(title: title)
  set page(height: auto, numbering: "1 / 1")

  set page(
    // `metadata("page counter")` is a workaround for sharing the same counter across the whole bundle
    // https://github.com/typst/typst/issues/8389
    header: metadata("page counter"),
    numbering: (..nums) => {
      let current = query(metadata.where(value: "page counter").before(here())).len()
      let total = query(metadata.where(value: "page counter")).len()
      numbering("1 / 1", ..(current, total).slice(0, nums.len()))
    },
    footer: context grid(
      columns: (1fr, auto, 1fr),
      align: (left, center, right),
      {
        let prev = query(selector(document).before(here())).at(-2, default: none)
        if prev != none {
          set text(font: "KaiTi")
          link(prev.location())[前一文件]
        }
      },
      link(<outline>, counter(page).display(both: true)),
      {
        let next = query(selector(document).after(here())).first(default: none)
        if next != none {
          set text(font: "KaiTi")
          link(next.location())[后一文件]
        }
      },
    ),
  ) if mode == "split-pdf"

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

  show link: set text(palette.link)
  show link: it => {
    set text(palette.ref) if type(it.dest) != str
    it
  }
  show cite: set text(palette.ref)
  show ref: it => {
    let el = it.element
    if el != none and el.func() == heading {
      link(el.location(), el.body)
    } else {
      it // Unchanged
    }
  }

  set footnote(numbering: "①")
  set table(stroke: none, align: center + horizon)

  set quote(block: true)
  show quote: block.with(width: 100%, stroke: (left: gray + 2pt), inset: (left: 0.5em), outset: (y: 0.5em))

  doc("index.pdf", {
    align(center, {
      [#std.title()<title>]
      date

      if revision != none {
        linebreak()
        raw(revision)
      }
    })

    [#outline(title: none, depth: 2)<outline>]

    v(1em)
    notes
  })

  set page(header: {
    if mode == "split-pdf" {
      set text(0.9em)
      text(font: "KaiTi", link(<title>, title))
      if revision != none {
        h(1fr)
        raw(revision)
      }

      metadata("page counter")
    }

    counter(footnote).update(0)
  })

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
