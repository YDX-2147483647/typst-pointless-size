#import "data.typ": data
#import "visualize.typ": draw-as-log-period

= 分支树
树叶上小图的细节可单击数据源跳转到后页查看。

#let excluded = (
  // 以下数值太小众
  "天元",
  "方正跨媒介",
  // 以下无特号但似乎应该有
  "沪新厂TTK",
  "商务厂TTK",
  "中华厂TTK",
)
未包含奇葩和存疑数据源（#excluded.map(s => ref(label("source:" + s))).join[、]）。
#v(2em) // 避让小图

#{
  // 补全 CLReq
  let data = data + (CLReq-extra: (data.CLReq-main.to-dict() + data.CLReq-extra.to-dict()).pairs())
  // 移除奇葩和存疑
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
    match: (name: [数据严重不全], fn: pairs => pairs.len() < 6),
    cases: (
      (value: true),
      (
        value: false,
        match: (name: [初号 $>= 42$], fn: pairs => if "初号" in pairs { pairs.初号 >= 42 }),
        cases: (
          (
            value: true,
            match: (name: [二号], fn: pairs => pairs.二号),
            cases: (
              (
                value: 22.,
                match: (name: [五号], fn: pairs => pairs.五号),
                cases: (
                  (
                    value: 10.5,
                    match: (name: [有小三], fn: pairs => "小三" in pairs),
                    cases: (
                      (value: true),
                      (value: false),
                    ),
                  ),
                  (value: 11.),
                ),
              ),
              (
                value: 21.,
                match: (
                  name: [有小号],
                  fn: pairs => pairs.keys().position(g => g.starts-with("小")) != none,
                ),
                cases: (
                  (
                    value: true,
                    match: (name: [有小三], fn: pairs => "小三" in pairs),
                    cases: (
                      (value: true),
                      (
                        value: false,
                        match: (name: [四号], fn: pairs => pairs.四号),
                        cases: (
                          (
                            value: 14.,
                            match: (name: [三号], fn: pairs => pairs.三号),
                            cases: (
                              (
                                value: 16.,
                                match: (name: [有小初], fn: pairs => "小初" in pairs),
                                cases: (
                                  (value: false),
                                  (value: true),
                                ),
                              ),
                              (value: 15.75),
                            ),
                          ),
                          (
                            value: 13.75,
                            match: (name: [三号], fn: pairs => pairs.三号),
                            cases: (
                              (
                                value: 15.75,
                                match: (name: [大一], fn: pairs => pairs.at("大一", default: none)),
                                cases: (
                                  (value: 30.),
                                  (value: none),
                                ),
                              ),
                              (value: 16.),
                            ),
                          ),
                        ),
                      ),
                    ),
                  ),
                  (
                    value: false,
                    match: (name: [一号 / 四号], fn: pairs => (pairs.一号, pairs.四号)),
                    cases: (
                      (
                        value: (27.5, 13.75),
                        match: (name: [七号], fn: pairs => pairs.七号),
                        cases: (
                          (
                            value: 5.25,
                            match: (name: [八号], fn: pairs => pairs.at("八号", default: none)),
                            cases: (
                              (value: none),
                              (value: 4.),
                            ),
                          ),
                          (value: 6.875),
                        ),
                      ),
                      (value: (26.25, 13.125)),
                    ),
                  ),
                ),
              ),
            ),
          ),
          (
            value: false,
            match: (name: [二号], fn: pairs => pairs.二号),
            cases: (
              (
                value: 21.,
                match: (name: [四号], fn: pairs => pairs.四号),
                cases: (
                  (
                    value: 13.75,
                    match: (name: [一号], fn: pairs => pairs.一号),
                    cases: (
                      (value: 27.5),
                      (value: 27.),
                    ),
                  ),
                  (
                    value: 14.,
                    match: (name: [特号], fn: pairs => pairs.特号),
                    cases: (
                      (
                        value: 45.,
                        match: (name: [有特中], fn: pairs => "特中" in pairs),
                        cases: (
                          (
                            value: true,
                            match: (name: [小初], fn: pairs => pairs.at("小初", default: none)),
                            cases: (
                              (value: 31.5),
                              (value: 30.),
                              (value: none),
                            ),
                          ),
                          (value: false),
                        ),
                      ),
                      (
                        value: auto,
                        match: (name: [六号], fn: pairs => pairs.六号),
                        cases: (
                          (value: 8.),
                          (value: 7.875),
                        ),
                      ),
                    ),
                  ),
                ),
              ),
              (
                value: 20.,
                match: (name: [四号], fn: pairs => pairs.四号),
                cases: (
                  (value: 14.),
                  (value: 13.),
                ),
              ),
            ),
          ),
          (
            value: none,
            match: (name: [二号], fn: pairs => pairs.at("二号", default: none)),
            cases: (
              (
                value: none,
                match: (name: [四号], fn: pairs => pairs.四号),
                cases: (
                  (value: 13.5),
                  (value: 14.),
                ),
              ),
              (
                value: 21.,
                match: (name: [四号], fn: pairs => pairs.四号),
                cases: (
                  (
                    value: 14.,
                    match: (name: [一号], fn: pairs => pairs.一号),
                    cases: (
                      (value: 27.5),
                      (
                        value: 28.,
                        match: (name: [三号], fn: pairs => pairs.三号),
                        cases: (
                          (
                            value: 16.,
                            match: (name: [大一], fn: pairs => pairs.at("大一", default: none)),
                            cases: (
                              (value: 30.),
                              (value: none),
                            ),
                          ),
                          (value: 15.75),
                        ),
                      ),
                    ),
                  ),
                  (value: 13.75),
                  (value: 13.6),
                ),
              ),
              (value: 22.),
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
  let case-index = for source in data.keys().sorted() {
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

    // The current level (0 = leaves, □ + 1 = parents of □)
    let x = -1
    // An array of (y, case-code) at the current level
    let current = grouped.map(array.first).enumerate()
    while true {
      x += 1

      let len = current.len()
      if len == 1 {
        assert.eq(current.first().last(), ())
        break
      }

      // Construct the previous level from the current level
      current = {
        // Calculate the length of the longest case-code
        // Nodes of longest-code-len will be merged per common prefix, and others will be kept unchanged.
        let longest-code-len = calc.max(
          ..for (y, code) in current {
            (code.len(),)
          },
        )

        // Merge node in the range [start, end) into one
        let do-merge(start, end, prefix) = {
          // To be pushed to the outer variables
          let lines = ()
          let texts = ()

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

          let output = (y-mean, prefix)
          (output, lines, texts)
        }

        // The plan for merging, (merge-start, common-prefix)
        let merging = none
        for (i, (y, code)) in current.enumerate() {
          let action = if code.len() == longest-code-len {
            if merging == none or code.slice(0, -1) == merging.last() {
              "push-merging"
            } else {
              "finish-and-push-merging"
            }
          } else {
            if merging == none {
              "keep-unchanged"
            } else {
              "finish-merging-and-keep-unchanged"
            }
          }

          if action in ("finish-merging-and-keep-unchanged", "finish-and-push-merging") {
            // Finish merging and clear the plan.
            assert.ne(merging, none)
            let (start, prefix) = merging
            let end = i
            merging = none

            let (output, extra-lines, extra-texts) = do-merge(start, end, prefix)
            lines += extra-lines
            texts += extra-texts
            (output,)
          }
          if action in ("push-merging", "finish-and-push-merging") {
            // For finish*-merging, `merging` should already be cleared.
            if merging == none {
              // Start a plan for merging.
              merging = (i, code.slice(0, -1))
            }
          }
          if action in ("keep-unchanged", "finish-merging-and-keep-unchanged") {
            // Keep the node unchanged for the next level.
            // This node should be put after the merging plan, if it existed.
            lines.push((from: (x, y), to: (x + 1, y)))
            ((y, code),)
          }
        }
        if merging != none {
          let (start, prefix) = merging

          let (output, extra-lines, extra-texts) = do-merge(start, current.len(), prefix)
          lines += extra-lines
          texts += extra-texts
          (output,)
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
