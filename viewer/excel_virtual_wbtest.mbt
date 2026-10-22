///|
test "large frozen Excel sheets mount the full virtual grid with renderer attributes" {
  let workbook = @mbtexcel.new_workbook()
  let name = "Frozen & Data"
  ignore(workbook.add_sheet(name))
  workbook.set_cell_int(name, "A1", 1)
  workbook.set_cell_int(name, "B1200", 42)
  workbook.set_panes(
    name,
    @xlsx.Panes::with_values(
      freeze=true,
      x_split=1,
      y_split=1,
      top_left_cell="B2",
    ),
  )
  match
    render_document("frozen.xlsx", "/frozen.xlsx", @mbtexcel.write(workbook)) {
    Ok(rendered) => {
      assert_true(rendered.surface_html.contains("office-virtual-table"))
      assert_false(rendered.surface_html.contains("<table>"))
      let model = fidelity_json_script(
        rendered.surface_html,
        "office-virtual-workbook",
      )
      let sheet = annotation_test_field(
        annotation_test_field(model, "sheets"),
        name,
      )
      assert_eq(annotation_test_field(sheet, "rows"), json_number(1200))
      assert_eq(annotation_test_field(sheet, "freeze_rows"), json_number(1))
      assert_eq(annotation_test_field(sheet, "freeze_columns"), json_number(1))
      let cells = json_array(Some(annotation_test_field(sheet, "cells")))
      assert_true(
        cells.any(cell => {
          cell is Array([Number(1200, ..), Number(2, ..), String("42"), ..])
        }),
      )
    }
    Err(message) => fail(message)
  }
}

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
