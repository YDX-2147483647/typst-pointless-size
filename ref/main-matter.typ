#import "sources.typ": sources
#import "data.typ": data
#import "visualize.typ": build-table, draw-as-log-period
#import "template.typ": doc, title-page

#doc("source-title.pdf", title-page[
  = 数据源详情
  逐一展示每个数据源的字号定义范围与点数数值（data），介绍出处及来源（brief、via），并补充数据原貌、疑点等情况（notes）。

  data 部分提供了图、表两种展示方式。

  - *图*中标注了若干对号数、点数，按*传统竖排顺序*阅读（从上到下、从右到左）可*从大到小*读出。每点都位于一条*竖线*上，同一竖线上的最小点数不小于最大点数的一半，$10.5$ 所在竖线加粗。有些点严格落于一条*横线*上，但也有些点仅接近某条横线，同一横线上相邻两条竖线交点的点数是二倍关系，$10.5$ 所在加粗横线左右分别标有 $times 1$ 和 $times 2^0$。一点所在竖线下方点数*乘以*所在横线两侧*倍数*即为其点数。例如下图，三号 $16$ 位于 $10.5$ 竖线与 $times 3/2$ 横线交点上方一点点，对应 $16 gt.tilde 10.5 times 3/2 = 15.75$；而 $times 3/2$ 横线与 $times 2^(7\/12)$ 横虚线几乎重合，对应 $2^(7\/12) approx 1.49831 ≈ 1.50 = 3/2$（纯五度完全协和 🎶）。

    #figure(draw-as-log-period(data.CLReq-main, width: 20em, height: 12em))

  - *表*有号数、点数、*长度示意*三列，其中第三列由点数计算得出，方便确定点数之间的倍数关系。例如下表，初号 $42$ 两格长等于二号 $21$ 四格长，也等于一号 $28$ 三格长，所以初号、二号、一号线度之比是 $1/2 : 1/4 : 1/3 = 6:3:4$。

    #figure(build-table((初号: 42, 二号: 21, 五号: 10.5, 一号: 28, 四号: 14).pairs()))

    这些表大多提供了「*按倍数关系*」和「按点数大小」两种排列方式，前者会像上表这样将行分成几组。这种分组能方便检查对照，但*不代表数据源本身的意见*。同一组的点数也未必严格成倍数关系，未分组也不代表点数没有倍数关系。
])
#set heading(offset: 1)
#show heading.where(level: 2): it => pagebreak(weak: true) + it

#for (source, fields) in sources.pairs() [
  #show: doc.with("source-{}.pdf".replace("{}", source))

  #[= #source.replace("-", " ")] #label("source:" + source)

  == data
  #{
    let data = data.at(source, default: none)
    if data == none [
      TODO
    ] else {
      // 绘图
      figure(draw-as-log-period(data, width: 28em, height: 18em))

      // 写表
      {
        let sorted-data = data.sorted(key: ((_g, p)) => -p)
        figure(if data == sorted-data {
          build-table(data)
        } else {
          grid(
            columns: 2,
            row-gutter: 0.5em,
            column-gutter: 2em,
            [按倍数关系], [按点数大小],
            build-table(data), build-table(sorted-data),
          )
        })
      }
    }
  }

  #for (k, v) in fields.pairs() [
    == #k
    #v
  ]
]
