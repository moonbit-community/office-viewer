///|
test "sparse sheet axes preserve explicit dimensions and exact boundaries" {
  let axis = sheet_axis(1048576, 26.0, Json([[2, 40, -1], [1000000, 100, 5]]))
  assert_eq(axis.offset(1), 26.0)
  assert_eq(axis.offset(2), 66.0)
  assert_eq(axis.index_at(25.9), 1)
  assert_eq(axis.index_at(26.0), 2)
  assert_eq(axis.index_at(65.9), 2)
  assert_eq(axis.index_at(66.0), 3)
  assert_eq(axis.styles.get(2), None)
  assert_eq(axis.styles.get(1000000), Some(5))
  assert_eq(axis.index_at(axis.offset(1048576) - 0.1), 1048576)
}

///|
test "compressed scrollbar reaches both endpoints of the full Excel row axis" {
  let axis = sheet_axis(1048576, 26.0, Json::null())
  let view = 600.0
  let last = axis.offset(axis.count) + 26.0 - view
  assert_eq(lazy_scroll(axis, 26.0, view, 0.0), 0.0)
  assert_eq(lazy_scroll(axis, 26.0, view, 8000000.0 - view), last)
  let target = axis.offset(1000000)
  let scroll = lazy_position(axis, 26.0, view, target)
  let resolved = lazy_scroll(axis, 26.0, view, scroll)
  assert_true((resolved - target).abs() < 0.001)
}
