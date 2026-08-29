/// 将号数名称统一成「初号」「小初」形式
#let normalize-g(g) = {
  g
    .replace(regex("^(\d)G$"), m => {
      let x = int(m.captures.first())
      if x > 0 { numbering("一号", x) } else { "初号" }
    })
    .replace(regex("\d"), m => {
      let x = int(m.text)
      if x > 0 { numbering("一", x) } else { "初" }
    })
    .replace("小（新）", "小")
    .replace("新", "小")
    .trim(regex("[字号號]"), at: end)
    .replace(regex("[头頭壹]"), "一")
    .replace(regex("^(小)?大$"), m => m.captures.first() + "一")
    .replace(regex("^.$"), m => m.text + "号") // 统一成两个字
}
#for (input, expected) in csv(bytes(
  ```csv
  0号,初号
  小初号,小初
  5号,五号
  小5号,小五
  新五号,小五
  新5号,小五
  小（新）四号,小四
  0G,初号
  4G,四号
  七行,七行
  特大号,特大
  大特号,大特
  特号,特号
  小特号,小特
  大号,一号
  小大号,小一
  大一号,大一
  \七,\七
  头号,一号
  十号,十号
  七行字,七行
  新大號,小一
  壹號,一号
  七号大,七号大
  七号小,七号小
  ```.text,
)) {
  let actual = normalize-g(input)
  assert.eq(
    actual,
    expected,
    message: "input = `" + input + "`, expected = `" + expected + "`, actual = `" + actual + "`",
  )
  assert.eq(normalize-g(actual), actual)
}

/// 判断g是「□号」还是「小□」，只用于绘图分类
#let is-small(g) = (g.starts-with("小") and g != "小特") or g.ends-with(regex("[行倍]"))
