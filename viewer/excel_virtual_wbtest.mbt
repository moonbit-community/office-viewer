///|
test "full Excel extents stay sparse and formula values keep their source expressions" {
  let workbook = @mbtexcel.new_workbook()
  ignore(workbook.add_sheet("Sparse"))
  workbook.set_cell_int("Sparse", "A1", 1)
  workbook.set_cell_formula("Sparse", "XFD1048576", "21*2")
  let model = virtual_excel_model(workbook)
  let sheet = annotation_test_field(
    annotation_test_field(model, "sheets"),
    "Sparse",
  )
  assert_eq(annotation_test_field(sheet, "rows"), json_number(1048576))
  assert_eq(annotation_test_field(sheet, "columns"), json_number(16384))
  let cells = json_array(Some(annotation_test_field(sheet, "cells")))
  assert_eq(cells.length(), 2)
  assert_true(
    cells.any(cell => {
      cell
      is Array(
        [
          Number(1048576, ..),
          Number(16384, ..),
          String("42"),
          _,
          String("21*2"),
          False,
        ]
      )
    }),
  )
}

///|
test "a formula returning an empty string remains an empty displayed value" {
  let workbook = @mbtexcel.new_workbook()
  ignore(workbook.add_sheet("Data"))
  workbook.set_cell_formula("Data", "A1200", "\"\"")
  let model = virtual_excel_model(workbook)
  let sheet = annotation_test_field(
    annotation_test_field(model, "sheets"),
    "Data",
  )
  let cells = json_array(Some(annotation_test_field(sheet, "cells")))
  assert_true(
    cells
    is [
      Array(
        [Number(1200, ..), Number(1, ..), String(""), _, String("\"\""), False]
      ),
    ],
  )
}
