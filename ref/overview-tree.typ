#import "data.typ": data
#import "visualize.typ": draw-as-log-period

= 分支树
树叶上小图的细节可单击数据源跳转到后页查看。

#let excluded = (
  // 以下数值太小众
  "老CCT",
  "天元",
  "方正跨媒介",
  // 以下无初号
  "沪新厂TTK",
  "商务厂TTK",
  "中华厂TTK",
  "申报1935",
)
未包含奇葩数据源（#excluded.map(s => ref(label("source:" + s))).join[、]）。
#v(2em) // 避让小图

#{
  // 补全 CLReq
  let data = data + (CLReq-extra: (data.CLReq-main.to-dict() + data.CLReq-extra.to-dict()).pairs())
  // 移除奇葩
  for s in excluded {
    let _ = data.remove(s)
  }

  // 声明允许的判据类型
  let humanize(v) = {
    if type(v) == bool {
      if v [✓] else [✗]
    } else if type(v) == float {
      $#v$ // 点数
    } else if type(v) == array {
      assert.eq(v.len(), 2) // 一对点数
      let (a, b) = v
      set math.frac(style: "skewed")
      $#a / #b$
    } else if v == auto {
      [其它]
    } else if v == none {
      [无]
    } else {
      assert(false, message: repr(type(v)) + " cannot be humanized")
    }
  }

  // 声明树
  let match-tree = (
    match: (name: [数据不全], fn: pairs => pairs.len() < 6),
    cases: (
      (value: true),
      (
        value: false,
        // 没有七号也算 false
        match: (name: [七号 $= 5.5$], fn: pairs => pairs.at("七号", default: none) == 5.5),
        cases: (
          (
            value: false,
            match: (name: [初号 $= 42$], fn: pairs => pairs.初号 == 42),
            cases: (
              (
                value: false,
                match: (name: [四号], fn: pairs => pairs.四号),
                cases: (
                  (
                    value: 14.,
                    match: (name: [特号 $= 45$], fn: pairs => pairs.特号 == 45),
                    cases: (
                      (
                        value: true,
                        match: (name: [有特初], fn: pairs => "特初" in pairs),
                        cases: (
                          (value: true),
                          (value: false),
                        ),
                      ),
                      (
                        value: false,
                        match: (name: [六号], fn: pairs => pairs.六号),
                        cases: (
                          (value: 8.),
                          (value: 7.875),
                        ),
                      ),
                    ),
                  ),
                  (value: 13.75),
                ),
              ),
              (
                value: true,
                match: (name: [三号], fn: pairs => pairs.三号),
                cases: (
                  (
                    value: 16.,
                    match: (name: [一号 / 四号], fn: pairs => (pairs.一号, pairs.四号)),
                    cases: (
                      (
                        value: (28., 14.),
                        match: (name: [有小六], fn: pairs => "小六" in pairs),
                        cases: (
                          (value: true),
                          (value: false),
                        ),
                      ),
                      (
                        value: (27.5, 13.75),
                        match: (
                          name: [有小号],
                          fn: pairs => pairs.keys().position(g => g.starts-with("小")) != none,
                        ),
                        cases: (
                          (
                            value: false,
                            match: (name: [八号], fn: pairs => pairs.八号),
                            cases: (
                              (value: 4.),
                              (value: 5.25),
                            ),
                          ),
                          (value: true),
                        ),
                      ),
                    ),
                  ),
                  (
                    value: 15.75,
                    match: (name: [一号 / 四号], fn: pairs => (pairs.一号, pairs.四号)),
                    cases: (
                      (
                        value: (27.5, 13.75),
                        match: (name: [二号], fn: pairs => pairs.二号),
                        cases: (
                          (
                            value: 21.,
                            match: (name: [有六行], fn: pairs => "六行" in pairs),
                            cases: (
                              (value: true),
                              (value: false),
                            ),
                          ),
                          (value: 22.),
                        ),
                      ),
                      (
                        value: auto,
                        match: (name: [一号], fn: pairs => pairs.一号),
                        cases: (
                          (value: 28.5),
                          (value: 26.25),
                        ),
                      ),
                    ),
                  ),
                ),
              ),
            ),
          ),
          (
            value: true,
            match: (name: [六号], fn: pairs => pairs.六号),
            cases: (
              (
                value: 8.,
                match: (name: [一号], fn: pairs => pairs.一号),
                cases: (
                  (value: 28.),
                  (value: 24.),
                ),
              ),
              (value: 7.5),
            ),
          ),
        ),
      ),
    ),
  )

  // 计算

  /// Returns an array of case codes
  let calc-cases(pairs) = {
    let current = match-tree

    while true {
      if "match" not in current {
        break // Finish
      }

      // Calculate result
      let result = (current.match.fn)(pairs)

      // Match cases
      let matched = false
      for (i, arm) in current.cases.enumerate() {
        if arm.value == auto or arm.value == result {
          matched = true
          current = arm
          (i,) // Output index
          break
        }
      }
      assert(matched, message: {
        "No match found for "

        repr(result)
        " in "
        repr(current)
        ", pairs: "
        repr(pairs)
      })
    }
  }

  let at-match-tree(case-code) = {
    let current = match-tree
    for i in case-code {
      current = current.cases.at(i)
    }
    current
  }

  // An array of (source, case-code)
  let case-index = for source in data.keys() {
    let pairs = data.at(source).to-dict()
    ((source, calc-cases(pairs)),)
  }.sorted(key: array.last)

  // An array of (case-code, list of sources)
  let grouped = {
    let grouped = ()
    for (name, code) in case-index {
      let last = grouped.last(default: none)
      if last != none and code == last.first() {
        grouped.last().last().push(name)
      } else {
        grouped.push((code, (name,)))
      }
    }
    grouped
  }

  // 绘图

  let (texts, lines) = {
    let lines = ()
    let texts = ()

    let current = grouped.map(array.first).enumerate()
    let x = -1
    while true {
      x += 1

      let len = current.len()
      if len == 1 {
        break
      }

      let has-shrunk = false

      // Construct the previous level from the current level
      current = {
        let start = 0
        while start < len {
          let prefix = current.at(start).last().slice(0, -1)

          // Find the longest range [start, end) that shares a common prefix.
          let end = {
            let i = start + 1
            while i < len and current.at(i).last().slice(0, -1) == prefix {
              i += 1
            }
            i
          }

          // If there are more than one node in [start, end)
          if end - start > 1 {
            // Shrink them into one.

            let ys = current.slice(start, end).map(array.first)
            let y-mean = ys.sum() / (end - start)

            let info = at-match-tree(prefix)
            for (y, (.., last-code)) in current.slice(start, end) {
              let edge = (from: (x, y), to: (x + 1, y-mean))
              lines.push(edge)
              let text = humanize(info.cases.at(last-code).value)
              texts.push((kind: "case", ..edge, text: text))
            }
            texts.push((kind: "match", xy: (x + 1, y-mean), text: info.match.name))

            has-shrunk = true
            ((y-mean, prefix),)
          } else {
            // Otherwise, keep it unchanged
            let y = current.at(start).first()
            lines.push((from: (x, y), to: (x + 1, y)))
            (current.at(start),)
          }

          // Prepare for the next iteration.
          start = end
        }
      }

      if not has-shrunk {
        // Shrink all single-child nodes.
        let longest-code = calc.max(
          ..for (y, code) in current {
            (code.len(),)
          },
        )
        for (i, (y, code)) in current.enumerate() {
          if code.len() == longest-code {
            let _ = current.at(i).last().pop()
            lines.push((from: (x, y), to: (x + 1, y)))
          }
        }
      }
    }

    (texts, lines)
  }
  {
    let w = 300pt
    let dx = 40pt
    let dy = 5em

    let branches = curve(
      stroke: aqua.darken(10%),
      ..for (from: (x1, y1), to: (x2, y2)) in lines {
        (
          curve.move((w - x1 * dx, y1 * dy)),
          curve.line((w - x2 * dx, y2 * dy)),
        )
      },
    )

    // 分支线位于字下方
    place(branches)

    for (y, (_, source-list)) in grouped.enumerate() {
      let text = {
        set text(0.8em) if source-list.len() >= 4
        source-list.map(s => ref(label("source:" + s))).join[\ ]
      }
      place(dx: w + 0.5em, dy: y * dy - 2em, grid(
        columns: (6em, auto),
        gutter: 0.5em,
        align: horizon + center,
        text, draw-as-log-period(source-list.map(s => data.at(s)).join(), tiny: true, width: 8em, height: 4em),
      ))
    }
    for (kind, ..edge, text) in texts {
      if kind == "match" {
        let (xy: (x, y)) = edge
        place(
          dx: w - x * dx - 0.8em,
          dy: y * dy - (1em + 0.2em) / 2,
          box(fill: white, stroke: purple, inset: 0.2em, std.text(0.8em, text)),
        )
      } else {
        let (from: (x1, y1), to: (x2, y2)) = edge

        let x = w - (x1 + 2 * x2) / 3 * dx + 0.5em
        let y = (y1 + 2 * y2) / 3 * dy - 0.5em
        if calc.abs(y2 - y1) < 2 {
          x += 0.3em
          y += 0.3em * if y2 < y1 { 1 } else { -1 }
        }
        if x1 == 0 {
          x += 0.5em
        }

        place(dx: x, dy: y, box(fill: white, std.text(0.8em, text)))
      }
    }

    hide(branches)
    v(1em) // 避让页脚
  }
}
