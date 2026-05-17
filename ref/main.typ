// Compiled with Typst v0.14.2.
#import "template.typ": link, template

#show: template.with(title: [资料汇编：汉字号数与点数的映射关系], date: [2026年5月7–17日], notes: [
  #for f in (
    "main.typ",
    "sources.typ",
    "data.typ",
    "overview.typ",
    "overview-tree.typ",
    "main-matter.typ",
    "appendices.typ",
    "ref.bib",
    "template.typ",
    "util.typ",
    "visualize.typ",
  ) {
    pdf.attach(f, relationship: "source")
  }

  注：
  - 本文件*只记录资料明文给出的映射关系*，不做任何外推（只考虑「某号是几点」说法，不考虑「某号与某号成小整数比例关系」说法）；不过这些资料本身不都是一手资料，不排除其作者自己外推过。
  - 为简洁，*号数*名称统一按「四号」「小初」形式，*点数*都不考虑有效数字位数（省略小数部分末尾的零）。`data.typ`源代码中有原始名称、数值，可从*此PDF附件*或#link("https://github.com/YDX-2147483647/typst-pointless-size/tree/main/ref/")[本项目GitHub]下载。
  - *点制*有 Fournier 点、Didot 点、美式点、TeX `pt` $1/72.27 "in"$、TeX `bp`（DTP点）$1/72 "in"$ 等。由于大部分资料都没有考虑这种差异，本文件也只好忽略，并尽量在有信息时给出说明。作为参考，大部分点都在 $350 "μm" plus.minus 5 "μm"$ 范围内，相对偏差不超过 $2%$（不过 Didot 点约 $376 "μm"$，差距较大）。另外，几乎所有资料都定义五号为 $10.5 "pt"$，所以也可理解为将五号字长度的 $1 / 10.5$ 作为一点。
  - 点是长度单位，*点数*描述字的线度，但具体描述什么线度并不完全确定。可能是金属活字的宽度或高度，也可能是汉字笔画分布范围的宽度或高度，还可能是与汉字相搭配的西文字体的点数，甚至可能是无法客观测量的汉字设计概念。由于条件有限，本文件也只好不做区分。
])

#import "data.typ": data
#import "sources.typ": sources
#metadata(data + (CLReq-extra: (data.CLReq-main.to-dict() + data.CLReq-extra.to-dict()).pairs()))<meta:data>
#metadata(sources)<meta:sources>

#include "overview.typ"

#include "main-matter.typ"
#include "appendices.typ"
