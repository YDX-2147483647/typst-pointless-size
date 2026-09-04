#import "@local/ribbony:0.1.1": label.default-linear-label-drawer, layout, ribbon-stylizer, sankey-diagram, tinter
#import "data.typ": data

= 倍数系
此页展示号数之间的二倍关系，即各对号数在多少数据源中对应的点数成二倍关系。

#let is-two(px, py) = calc.abs(py / px - 2) <= 0.01

// 必须手工调整layers和edge-order，否则画出来太乱
#let layers = (
  // 在7之后断开的依据：
  // 1. `draw-as-log-period`选择在8到9断开，但这里无法照搬，否则三号-小初-小二会乱
  // 2. 与之最接近的断法是7到8与9到10，后者会乱小二-初号-二号，也不合适
  // 3. 7到8会乱小六-小三-六号，但这种数据源很少，可以接受
  //
  // (0, 7]
  "0": "小六、七号大、七号、七号小、小七、八号、九号、十号",
  // (7, 14]
  "1": "四号、小四、五号、小五、六号",
  // (14, 28]
  "2": "一号、小一、二号、小二、三号、小三",
  // (28, 56]
  "3": "特中、五倍、六行、特初、五行、特号、大初、小特、初号、四行、小初、大一",
  // (56, +∞)
  "4": "七倍、大特、七行、特大、六倍",
).map(v => v.split("、"))

// A list of (gx, gy, n)
#let edges = {
  let edge-order = (
    "七倍,#void",
    "大一,大特",
    "七行,#void",
    "一号,特大",
    "六倍,#void",
    "一号,特中",
    "五倍,#void",
    "六行,#void",
    "四号,一号",
    "四号,小一",
    "小一,特初",
    "五行,#void",
    "小一,特号",

    "小六,四号",
    "小六,小四",
    "七号大,四号",
    "七号,四号",
    "小四,小一",
    "七号,小四",
    "七号,五号",
    "七号小,五号",
    "小七,五号",

    "二号,特号",
    "二号,大初",
    "二号,小特",
    "二号,初号",
    "五号,二号",
    "八号,五号",

    "小五,小二",
    "小二,初号",
    "小二,四行",
    "小二,小初",
    "小初,特大",
    "小初,七倍",
    "三号,小初",

    "八号,六号",
    "六号,三号",
    "六号,小三",
    "九号,小五",
    "十号,六号",

    "小六,小三",
  ).map(row => row.split(","))

  let edges-unordered = {
    for gx in data.values().map(pairs => pairs.map(array.first)).join().dedup() {
      // 忽略天元
      if gx.starts-with("\\") { continue }

      // A map from gy to n
      let count = (:)
      for (source, pairs) in data.pairs() {
        let px = pairs.to-dict().at(gx, default: none)
        if px == none { continue }

        let hit = false
        for (gy, py) in pairs {
          // 忽略天元
          if gy.starts-with("\\") { continue }

          if is-two(px, py) {
            if gy not in count {
              count.insert(gy, 0)
            }
            count.at(gy) += 1
            hit = true
          }
        }
        if not hit {
          if "#void" not in count {
            count.insert("#void", 0)
          }
          count.at("#void") += 1
        }
      }

      for (gy, n) in count.pairs() {
        ((gx, gy, n),)
      }
    }
  }

  // Apply edge-order
  let as-key(x, y) = x + "," + y
  let edges = for (x, y) in edge-order {
    (as-key(x, y): (x: x, y: y))
  }
  for (x, y, n) in edges-unordered {
    let key = as-key(x, y)
    if "#void" in (x, y) {
      if key not in edges {
        edges.insert(key, (x: x, y: y))
      }
    } else {
      assert(key in edges, message: "please add " + key + "to the edge-order")
    }
    assert("n" not in edges.at(key))
    edges.at(key).insert("n", n)
  }
  edges.values().map(((x, y, n)) => (x, y, n))
}

#figure(sankey-diagram(
  edges,
  layout: layout.auto-linear(layers: layers),
  categories: (
    "初、二、五、七",
    "一、四",
    "三、六、八",
  )
    .map(row => row.split("、"))
    .map(row => (
      row.map(g => g + "号"),
      row.map(g => "小" + g),
    ))
    .join()
    .enumerate()
    .map(((i, v)) => (str(i), v))
    .to-dict(),
  tinter: tinter.categorical-tinter(),
  ribbon-stylizer: (edge, from-color, to-color, from-node, to-node, ..) => (
    // 采用内侧node的颜色
    fill: if from-node.layer == 0 { to-color } else { from-color }.transparentize(75%),
    stroke: none,
  ),
  draw-label: default-linear-label-drawer(
    // 标在左侧刚好对齐个位，也顺便与`draw-as-log-period`区分
    snap: left,
  ),
))
#[
  #let g = "小二"
  下面以#g;为例介绍如何阅读这种Sankey图。

  // n是数量，s是相应数据源
  #let (n, s) = {
    let n = (both: 0, only-as-x: 0, only-as-y: 0, neither: 0, uncover: 0)
    let s = (only-as-x: (), only-as-y: (), neither: ()) // both、uncover太多，不列出

    for (source, pairs) in data.pairs() {
      let p = pairs.to-dict().at(g, default: none)
      if p == none {
        n.uncover += 1
      } else {
        let as-x = false
        let as-y = false
        for (g-other, p-other) in pairs {
          // 忽略天元
          if g-other.starts-with("\\") { continue }

          if not as-x { as-x = is-two(p, p-other) }
          if not as-y { as-y = is-two(p-other, p) }
        }

        if as-x and as-y {
          n.both += 1
        } else if as-x and not as-y {
          n.only-as-x += 1
          s.only-as-x.push(source)
        } else if as-y and not as-x {
          n.only-as-y += 1
          s.only-as-y.push(source)
        } else {
          n.neither += 1
          s.neither.push(source)
        }
      }
    }
    (
      (
        ..n,
        cover: n.both + n.only-as-x + n.only-as-y + n.neither,
        any: n.both + n.only-as-x + n.only-as-y,
        as-x: n.both + n.only-as-x,
        as-y: n.both + n.only-as-y,
      ),
      s.map(sources => sources.map(s => ref(label("source:" + s))).join[、]),
    )
  }

  #let as-x = edges.filter(((x, y, n)) => x == g and y != "#void")
  #let as-y = edges.filter(((x, y, n)) => x != "#void" and y == g)
  #assert.eq(n.both + n.only-as-x, as-x.map(((x, y, n)) => n).sum())
  #assert.eq(n.both + n.only-as-y, as-y.map(((x, y, n)) => n).sum())

  + 有#n.cover;个数据源包含#g，因此图中#g;标注#n.cover，相应竖线长#n.cover;个单位。其中除#s.neither;这#n.neither;个数据源外，其余#n.any;个数据源均存在能与#g;成二倍关系的号数。

  + 在#g;能成二倍关系的#n.any;个数据源中，#n.both;个数据源既有号数是#g;的两倍，又有号数是#g;的一半；#n.only-as-x;个数据源有两倍但无一半（#s.only-as-x）；#n.only-as-y;个数据源无两倍但有一半（#s.only-as-y）。

  + 在有#g;两倍的#n.as-x;个数据源中，#for (i, (x, y, n)) in as-x.enumerate() {
      if i == 0 { ([#n;个数据源称#g;两倍为#y],) } else { ([#n;个数据源称#y],) }
    }.join[，]。因此从#g;引出条带到#for (x, y, n) in as-x { (y,) }.join[、]，宽度分别为#for (x, y, n) in as-x { ([#n],) }.join[、]个单位。

    在有#g;一半的#n.as-y;个数据源中，#{
      let ((x, y, n),) = as-y
      [全部数据源称#g;一半为#x。因此从#g;引出条带到#x，宽度为#n;个单位。]
    }

  + 竖线与条带基于@source:CTeX新\认为的二倍系列分组上色。对于@source:CTeX新\未包含的号数，小七外推算入小初、小二、小五系列，其余号数单独分为一组。这种分组比较主观，也不完全反映数据主流，请谨慎参考。
]
