///|
test "viewer recognizes supported Office extensions" {
  assert_true(document_kind("notes.DOCX") is Some(Word))
  assert_true(document_kind("budget.xlsx") is Some(Excel))
  assert_true(document_kind("photo.pdf") is None)
}

///|
test "viewer keeps renderer styles and page boundaries" {
  let html = extract_html_body(
    "<style>.s1{font-weight:bold}</style><body><h1>Book</h1></body>",
  )
  assert_true(html.has_prefix("<style>"))
  assert_true(html.contains("<h1>Book</h1>"))
  let pages = word_surface(["<svg>A</svg>", "<svg>B</svg>"])
  assert_true(pages.contains("data-page=\"1\""))
  assert_true(pages.contains("data-page=\"2\""))
  assert_true(pages.find("<svg>A</svg>") is Some(_))
  assert_true(pages.find("<svg>B</svg>") is Some(_))
}

///|
test "viewer evaluates uncached Excel formulas" {
  let workbook = @mbtexcel.new_workbook()
  ignore(workbook.add_sheet("Data"))
  workbook.set_cell_int("Data", "A1", 20)
  workbook.set_cell_formula("Data", "A2", "A1*2")
  let bytes = @mbtexcel.write(workbook)
  match render_document("formula.xlsx", "/mock/office/formula.xlsx", bytes) {
    Ok(value) => {
      assert_true(value.surface_html.contains(">40</td>"))
      assert_true(value.surface_html.contains("office-formula-metadata"))
      assert_true(value.surface_html.contains("A1*2"))
      assert_false(
        value.surface_html.contains("<td class=\"formula-pending\">"),
      )
      assert_false(value.surface_html.contains("=A1*2"))
    }
    Err(error) => fail(error)
  }
}

///|
test "viewer normalizes cached formula text and preserves drawing metadata" {
  let bytes = @demos.demo_dashboard_bytes()
  match
    render_document("dashboard.xlsx", "/mock/office/dashboard.xlsx", bytes) {
    Ok(value) => {
      assert_true(value.surface_html.contains("office-object-metadata"))
      assert_true(value.surface_html.contains("\"anchor\":\"A1\""))
      assert_true(value.surface_html.contains("\"anchor\":\"B5\""))
      assert_true(value.surface_html.contains("width_px"))
      assert_true(value.surface_html.contains(">$28</td>"))
      assert_false(value.surface_html.contains("12 non-numeric values skipped"))
    }
    Err(error) => fail(error)
  }
}

///|
test "viewer renders composite charts instead of a placeholder" {
  let bytes = @demos.demo_combo_chart_bytes()
  match
    render_document("combo_chart.xlsx", "/mock/office/combo_chart.xlsx", bytes) {
    Ok(value) => {
      assert_true(value.surface_html.contains("viewer-combo-chart"))
      assert_true(value.surface_html.contains("<rect"))
      assert_true(value.surface_html.contains("<polyline"))
      assert_true(value.surface_html.contains("<path"))
      assert_true(value.surface_html.contains("secondary"))
      assert_true(value.surface_html.contains("#C00000"))
      assert_false(
        value.surface_html.contains("<figure class=\"chart-placeholder\""),
      )
    }
    Err(error) => fail(error)
  }
}
