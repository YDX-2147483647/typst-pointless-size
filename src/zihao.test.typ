#import "./zihao.typ": *

#for rule in size_number_to_pt {
  assert.eq(rule.len(), 2)
  assert(rule.at(0) == "-0" or type(rule.at(0)) == int)
  assert.eq(type(rule.at(1)), length)
}

#{
  assert.eq(number_to_name(5), "五号")
  assert.eq(number_to_name(-5), "小五")
  assert.eq(number_to_name("-0"), "小初")
}

#{
  assert.eq(name_to_number("五号"), name_to_number("五"))

  for size in ("五号", "小五", "小初") {
    assert.eq(number_to_name(name_to_number(size)), size)
  }
}

#{
  assert.eq(type(zh(5)), length)
  assert.eq(zh("五"), zh(5))
  assert.eq(zh("五号"), zh(5))

  assert.eq(zh("小五"), zh(-5))
  assert(zh(-5) < zh(5))

  assert.eq(zh("小初"), zh("-0"))
  assert(zh(5) < zh("-0"))
}

#{
  assert.eq(zh(7), 5.5pt)
  // Override the existing definition
  let _zh = zh.with(overrides: ((7, 5.25pt),))
  assert.eq(_zh(7), 5.25pt)

}

#{
  // Panic: zh(9)
  // Add a new definition
  let _zh = zh.with(overrides: ((10, 0pt),))
  assert.eq(_zh(10), 0pt)
  assert.eq(_zh("十"), _zh(10))
  assert.eq(_zh("十号"), _zh(10))
}
