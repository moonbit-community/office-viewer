///|
fn fidelity_json_script(html : String, class_name : String) -> Json raise {
  let marker = "<script type=\"application/json\" class=\"" + class_name + "\">"
  guard html.find(marker) is Some(start) else {
    fail("Missing JSON script " + class_name)
  }
  let content_start = start + marker.length()
  guard html[content_start:].find("</script>") is Some(length) else {
    fail("Missing script terminator")
  }
  @json.parse(html[content_start:content_start + length].to_owned())
}

///|
test "render initial window dimensions must be positive" {
  for (rows, columns) in [(0, 1), (1, 0), (-1, 1), (1, -1)] {
    match
      render_document(
        "bounds.xlsx",
        "/bounds.xlsx",
        b"not a workbook",
        max_rows=rows,
        max_cols=columns,
      ) {
      Err(message) => assert_true(message.contains("bounds must be positive"))
      Ok(_) => fail("Invalid bounds accepted")
    }
  }
}

///|
test "render initial windows accept large thresholds without truncating the workbook" {
  let workbook = @mbtexcel.new_workbook()
  ignore(workbook.add_sheet("Data"))
  workbook.set_cell_int("Data", "A1", 42)
  let bytes = @mbtexcel.write(workbook)
  for
    (rows, columns) in [
      (262144, 1),
      (1, 262144),
      (1024, 256),
      (1000, 256),
      (2147483647, 2147483647),
    ] {
    match
      render_document(
        "bounds.xlsx",
        "/bounds.xlsx",
        bytes,
        max_rows=rows,
        max_cols=columns,
      ) {
      Ok(rendered) => assert_true(rendered.surface_html.contains(">42</td>"))
      Err(message) => fail(message)
    }
  }
  match render_document("bounds.xlsx", "/bounds.xlsx", bytes) {
    Ok(rendered) => assert_true(rendered.surface_html.contains(">42</td>"))
    Err(message) => fail(message)
  }
}

///|
test "Word diagnostics escape hostile text and preserve original JSON with duplicate counts" {
  let hostile = "<b></script><img src=x onerror=alert(1)> & \"quote\""
  let html = render_warning_html([hostile, "nested table", hostile])
  assert_true(
    html.contains(
      "&lt;b&gt;&lt;/script&gt;&lt;img src=x onerror=alert(1)&gt; &amp;",
    ),
  )
  assert_false(html.contains("<img"))
  assert_false(html.contains(hostile))
  assert_true(html.contains("(2)</li>"))
  let diagnostics = fidelity_json_script(html, "office-render-diagnostics")
  assert_eq(
    annotation_test_field(diagnostics, "complete"),
    Json::boolean(false),
  )
  assert_eq(
    annotation_test_field(diagnostics, "truncated"),
    Json::boolean(false),
  )
  let items = json_array(Some(annotation_test_field(diagnostics, "items")))
  assert_eq(items.length(), 2)
  assert_eq(annotation_test_field(items[0], "feature"), Json::string(hostile))
  assert_eq(annotation_test_field(items[0], "count"), json_number(2))
  assert_eq(
    annotation_test_field(items[1], "feature"),
    Json::string("nested table"),
  )
}

///|
test "Word diagnostics remain empty for complete renders and bound many distinct warnings" {
  assert_eq(render_warning_html([]), "")
  let warnings = [ for index in 0..<65 => "feature-" + index.to_string() ]
  warnings.push("feature-0")
  let html = render_warning_html(warnings)
  let diagnostics = fidelity_json_script(html, "office-render-diagnostics")
  let items = json_array(Some(annotation_test_field(diagnostics, "items")))
  assert_eq(items.length(), 64)
  assert_eq(annotation_test_field(items[0], "count"), json_number(2))
  assert_eq(
    annotation_test_field(diagnostics, "truncated"),
    Json::boolean(true),
  )
  assert_true(html.contains("Additional limitations are present."))
  assert_false(html.contains("feature-64"))
  let exact = render_warning_html(
    [
      for index in 0..<64 => "feature-" + index.to_string()
    ],
  )
  assert_eq(
    annotation_test_field(
      fidelity_json_script(exact, "office-render-diagnostics"),
      "truncated",
    ),
    Json::boolean(false),
  )
}

///|
test "render_document exposes actual DOCX parser warnings while preserving readable pages" {
  let bytes = @docx2html.write_docx([
    @docx2html.paragraph([
      @docx2html.run([
        @docx2html.text("Readable text"),
        @document.image("image/gif", b"GIF89a\x01\x00\x01\x00\x00\x00\x00"),
      ]),
    ]),
  ])
  match render_document("image.docx", "/image.docx", bytes) {
    Ok(rendered) => {
      assert_true(rendered.item_count > 0)
      assert_true(rendered.surface_html.contains("Readable"))
      assert_true(rendered.surface_html.contains("office-render-warning"))
      assert_true(rendered.warning.contains(".gif"))
      let diagnostics = fidelity_json_script(
        rendered.surface_html,
        "office-render-diagnostics",
      )
      let items = json_array(Some(annotation_test_field(diagnostics, "items")))
      assert_true(items.length() > 0)
      assert_eq(
        annotation_test_field(diagnostics, "complete"),
        Json::boolean(false),
      )
    }
    Err(message) => fail(message)
  }
  let plain = @docx2html.write_docx(annotation_test_body("Plain document"))
  match render_document("plain.docx", "/plain.docx", plain) {
    Ok(rendered) => {
      assert_eq(rendered.warning, "")
      assert_false(rendered.surface_html.contains("office-render-warning"))
    }
    Err(message) => fail(message)
  }
}

///|
test "Excel window bounds retain all cells formulas and Notes anchors" {
  let workbook = @mbtexcel.new_workbook()
  ignore(workbook.add_sheet("Data"))
  for (cell, value) in [("A1", 11), ("B2", 22), ("C2", 32), ("B3", 23)] {
    workbook.set_cell_int("Data", cell, value)
    workbook.add_comment("Data", {
      cell,
      author: "Ada",
      author_id: None,
      text: cell,
      paragraph: [],
      width: None,
      height: None,
    })
  }
  workbook.set_cell_formula("Data", "B2", "A1*2")
  workbook.set_cell_formula("Data", "C2", "A1*3")
  workbook.set_cell_formula("Data", "B3", "A1*4")
  let bytes = @mbtexcel.write(workbook)
  match
    render_document("notes.xlsx", "/notes.xlsx", bytes, max_rows=2, max_cols=2) {
    Ok(rendered) => {
      let model = fidelity_json_script(
        rendered.surface_html,
        "office-virtual-workbook",
      )
      let data = annotation_test_field(
        annotation_test_field(model, "sheets"),
        "Data",
      )
      let cells = json_array(Some(annotation_test_field(data, "cells")))
      assert_true(
        cells.any(cell => {
          cell is Array([Number(2, ..), Number(2, ..), String("22"), ..])
        }),
      )
      assert_true(
        cells.any(cell => {
          cell is Array([Number(2, ..), Number(3, ..), String("33"), ..])
        }),
      )
      assert_true(
        cells.any(cell => {
          cell is Array([Number(3, ..), Number(2, ..), String("44"), ..])
        }),
      )
      let support = annotation_test_field(rendered.annotations, "support")
      assert_eq(
        annotation_test_field(support, "preview_max_rows"),
        json_number(1048576),
      )
      assert_eq(
        annotation_test_field(support, "preview_max_columns"),
        json_number(16384),
      )
      let items = json_array(
        Some(annotation_test_field(rendered.annotations, "items")),
      )
      assert_eq(items.length(), 4)
      for item in items {
        let anchor = annotation_test_field(item, "anchor")
        let expected = "within-preview"
        assert_eq(
          annotation_test_field(anchor, "preview"),
          Json::string(expected),
        )
      }
    }
    Err(message) => fail(message)
  }
}

///|
test "Excel metadata keeps hostile sheet names and formulas inside inert JSON" {
  let workbook = @mbtexcel.new_workbook()
  let sheet = "<b>"
  let formula = "\"<i></script><script>alert(1)</script>\""
  ignore(workbook.add_sheet(sheet))
  workbook.set_cell_formula(sheet, "A1", formula)
  let bytes = @mbtexcel.write(workbook)
  match render_document("hostile.xlsx", "/hostile.xlsx", bytes) {
    Ok(rendered) => {
      assert_false(rendered.surface_html.contains("<script>alert(1)</script>"))
      let metadata = fidelity_json_script(
        rendered.surface_html,
        "office-formula-metadata",
      )
      let formulas = annotation_test_field(
        annotation_test_field(metadata, "sheets"),
        sheet,
      )
      assert_eq(annotation_test_field(formulas, "A1"), Json::string(formula))
    }
    Err(message) => fail(message)
  }
}

///|
test "Excel render carries each worksheet freeze settings through actual XLSX serialization" {
  let workbook = @mbtexcel.new_workbook()
  for sheet in ["Frozen", "Plain", "Split"] {
    ignore(workbook.add_sheet(sheet))
    workbook.set_cell_int(sheet, "A1", 1)
  }
  workbook.set_panes(
    "Frozen",
    @xlsx.Panes::with_values(
      freeze=true,
      x_split=2,
      y_split=3,
      top_left_cell="C4",
    ),
  )
  workbook.set_panes(
    "Split",
    @xlsx.Panes::with_values(
      split=true,
      x_split=200,
      y_split=300,
      top_left_cell="C4",
    ),
  )
  match
    render_document("frozen.xlsx", "/frozen.xlsx", @mbtexcel.write(workbook)) {
    Ok(rendered) => {
      let metadata = annotation_test_field(
        fidelity_json_script(rendered.surface_html, "office-freeze-metadata"),
        "sheets",
      )
      for
        (sheet, rows, columns) in [
          ("Frozen", 3, 2),
          ("Plain", 0, 0),
          ("Split", 0, 0),
        ] {
        let setting = annotation_test_field(metadata, sheet)
        assert_eq(annotation_test_field(setting, "rows"), json_number(rows))
        assert_eq(
          annotation_test_field(setting, "columns"),
          json_number(columns),
        )
      }
    }
    Err(message) => fail(message)
  }
}
