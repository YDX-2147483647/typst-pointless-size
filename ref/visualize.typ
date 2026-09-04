#import "@preview/lilaq:0.6.0" as lq
#import "util.typ": is-small

#import calc: abs, div-euclid, exp, gcd, ln, pow, rem-euclid

/// 用表格列出一个数据源的号数与点数映射关系
///
/// - pairs: 一串 (号数, 点数)
///
/// 表格行顺序与 pairs 一致。表格还会将行分组，每组对应 pairs 中一段点数单调递减的子列。
#let build-table(pairs) = {
  let p-max = calc.max(..pairs.map(((g, p)) => p))

  table(
    columns: 3,
    align: (center, center, start).map(x => x + horizon),
    table.hline(),
    table.header[*号数*][*点数*][*长度示意*],
    table.hline(stroke: 0.5pt),
    ..pairs
      .enumerate()
      .map(((i, (g, p))) => {
        if i > 0 and p > pairs.at(i - 1).last() {
          // Add a line if increasing
          (table.hline(stroke: 0.5pt),)
        }
        (
          (
            g,
            $#p$,
            {
              let n = calc.ceil(p-max * 2 / p)
              v(-0.5em)
              grid(
                columns: (p * 1pt,) * n,
                stroke: (x: purple),
                fill: green.lighten(50%),
                box(height: 1.5em),
              )
              v(-0.5em)
            },
          ),
        )
      })
      .flatten(),
    table.hline(),
  )
}

/// Factorize p as p-ref × e^m × base^n, where start ≤ e^m < start × base and n ∈ ℤ.
/// Returns (n, m).
#let factorize(p, p-ref: 1, base: 2, start: 1) = {
  assert(p > 0)
  assert(p-ref > 0)
  assert(base > 1)
  assert(start > 0)

  let p = ln(p / p-ref)
  let base = ln(base)
  let start = ln(start)

  (
    int(div-euclid(p - start, base)),
    start + rem-euclid(p - start, base),
  )
}
#{
  let assert-eq(actual, expected) = {
    assert.eq(actual.first(), expected.first())
    assert(abs(actual.last() - expected.last()) < 1e-5)
  }

  for n in range(3) {
    assert-eq(
      factorize(10 * pow(2, n), p-ref: 10, base: 2, start: 1),
      (n, 0.0),
    )
  }
  assert-eq(
    factorize(10, p-ref: 10, base: 2, start: 2 / 3),
    (0, 0.0),
  )
  assert-eq(
    factorize(15, p-ref: 10, base: 2, start: 2 / 3),
    (1, ln(3 / 4)),
  )
  assert-eq(
    factorize(10, p-ref: 10, base: 2, start: 1 / 2),
    (1, ln(1 / 2)),
  )
}

/// 按对数周期展示一个数据源的号数与点数映射关系
///
/// - pairs: 一串 (号数, 点数)
/// - p-ref: 代表原点的参考点数，一般保留默认即可
/// - mark-scale: 数据点的放大倍数，p-ref 非默认时可一同调整
/// - start: 周期起始，详见`factorize`文档；一般保留默认即可
/// - width, height: lilaq 绘图尺寸
/// - tiny: 是否简化标注
//
// start 默认值的选取依据：
// 1. base = 2，选取 start 相当于在五号到二号之间选取一处空隙换列
// 2. 最常用的五号必须在中间，那么五号与它相邻的小五、小四必须在同一列，即排除小二到二号、五号到小四
// 3. 四号、小三、三号这段点数太密，换列最好避开，否则它们离边框太近，于是剩下只有小四到四号、三号到小二两处可选
// 4. 考虑到小四、四号毕竟都带「四」，在同一列更好，于是选择三号到小二，即要求 16/(2×10.5) < start < 18/(2×10.5)
// 5. 选取这段区间的中心，即 √(16×18) / 21 ≈ 4/5
#let draw-as-log-period(pairs, p-ref: 10.5, mark-scale: 1.8, start: 4 / 5, width: 6cm, height: 4cm, tiny: false) = {
  let if-not-tiny(v) = if not tiny { (v,) }

  let base = 2
  let factorize = factorize.with(p-ref: p-ref, base: base, start: start)

  let large-n = ()
  let large-m = ()
  let large-sizes = ()
  let small-n = ()
  let small-m = ()
  let small-sizes = ()
  for (g, p) in pairs {
    let (n, m) = factorize(p)
    if is-small(g) {
      small-n.push(n)
      small-m.push(m)
      if not tiny { small-sizes.push(mark-scale * p) }
    } else {
      large-n.push(n)
      large-m.push(m)
      if not tiny { large-sizes.push(mark-scale * p) }
    }
  }

  let large-ticks = (
    (1, 2),
    (6, 7),
    1,
    (8, 7),
    (4, 3),
    (3, 2),
    base,
  ).map(
    x => {
      let (m, label) = if type(x) != array {
        (ln(x), $times #x$)
      } else {
        let (p, q) = x
        assert.eq(gcd(p, q), 1)
        (ln(p / q), $times #p/#q$)
      }

      (m, if not tiny { label })
    },
  )

  let small-ticks = range(-12, 12).map(p-full => {
    let c = gcd(p-full, 12)
    let (p, q) = (p-full / c, 12 / c)

    let label = {
      set text(0.8em)
      if p == 0 {
        $times #base^0$
      } else if rem-euclid(p-full, 12) in (2, 4, 5, 7, 8, 10) {
        $times #base^(#p\/#q)$
      }
    }
    (p / q * ln(base), label)
  })

  let label = (
    x: lq.label(pad: 0em, grid(
      $stretch(<-, size: #{ width * 0.8 })$,
      [逐次二分],
    )),
    y: lq.label(angle: 0deg, pad: 0.25em, grid(
      columns: 2,
      rotate(-90deg, reflow: true, $stretch(<-, size: #{ height * 0.8 })$),
      {
        set par(justify: false, leading: 0.05em)
        box(width: 1em)[二倍内缩小]
      },
    )),
  )

  show: lq.set-tick(..if tiny { (inset: 0pt) })

  lq.diagram(
    width: width,
    height: height,
    xlim: (
      calc.min(-1, ..small-n, ..large-n) - 0.3,
      calc.max(2, ..small-n, ..large-n) + 0.7,
    ),
    xaxis: (
      locate-ticks: (x0, x1, ..) => {
        let start = int(calc.min(x0, x1, -1))
        let end = int(calc.max(x0, x1, 1))
        (ticks: range(start, end + 1))
      },
      format-ticks: if not tiny { (ticks, ..) => ticks.map(n => $#{ p-ref * pow(base, n) }$) },
      subticks: none,
      ..if not tiny { (label: label.x) },
    ),
    ylim: (0, ln(base)).map(m => m + ln(start)),

    // 左右 ticks
    yaxis: (ticks: large-ticks, subticks: none),
    ..if-not-tiny(lq.yaxis(
      position: right,
      ticks: small-ticks,
      subticks: 1,
      label: label.y,
    )),

    // 网格线
    lq.vlines(0, stroke: 1.5pt + luma(80%)),
    lq.hlines(0, stroke: 1.5pt + luma(80%)),
    ..if-not-tiny(lq.hlines(
      ..small-ticks.map(array.first),
      stroke: (paint: luma(80%), thickness: 0.5pt, dash: "loosely-dashed"),
    )),

    // 数据
    lq.scatter(large-n, large-m, mark: "o", ..if not tiny { (size: large-sizes) }),
    lq.scatter(small-n, small-m, mark: "s3", ..if not tiny { (size: small-sizes) }),

    // 标签
    ..if not tiny {
      for (g, p) in pairs {
        let (n, m) = factorize(p)
        let annotation = lq.place(n + 0.1, m, align: left + horizon, {
          set text(0.8em)
          place(text(stroke: stroke(thickness: 3pt, paint: white, miter-limit: 1))[#g $#p$])
          [#g $#p$]
        })
        (annotation,)
      }
    },
  )
}
