///|
test "reader source index decorates a simple body paragraph" {
  let document = [
    @docx2html.paragraph([@docx2html.run([@docx2html.text("Hello world")])]),
  ]
  let bytes = @docx2html.write_docx(document)
  let rendered = render_document(
    "sample.docx", "/mock/office/sample.docx", bytes,
  )
  match rendered {
    Ok(value) => {
      assert_true(
        value.surface_html.contains("data-source-selector=\"/docx/body/p[1]\""),
      )
      assert_true(value.surface_html.contains("data-source-start=\"0\""))
      assert_true(value.surface_html.contains("data-source-end=\"11\""))
      assert_true(value.surface_html.contains("data-source-runs=\"p[1]/r[1]\""))
    }
    Err(error) => fail(error)
  }
}

///|
test "reader source index preserves run boundaries across a visual fragment" {
  let document = [
    @docx2html.paragraph([
      @docx2html.run([@docx2html.text("Hello ")]),
      @docx2html.run([@docx2html.text("world")]),
    ]),
  ]
  let bytes = @docx2html.write_docx(document)
  let rendered = render_document("runs.docx", "/mock/office/runs.docx", bytes)
  match rendered {
    Ok(value) => {
      assert_true(value.surface_html.contains("data-source-start=\"0\""))
      assert_true(value.surface_html.contains("data-source-end=\"6\""))
      assert_true(value.surface_html.contains("data-source-runs=\"p[1]/r[1]\""))
      assert_true(value.surface_html.contains("data-source-start=\"6\""))
      assert_true(value.surface_html.contains("data-source-end=\"11\""))
      assert_true(value.surface_html.contains("data-source-runs=\"p[1]/r[2]\""))
    }
    Err(error) => fail(error)
  }
}

///|
test "reader source index keeps astral and CJK offsets in UTF-16 units" {
  let text = "前😀后"
  let document = [
    @docx2html.paragraph([@docx2html.run([@docx2html.text(text)])]),
  ]
  let bytes = @docx2html.write_docx(document)
  let rendered = render_document(
    "unicode.docx", "/mock/office/unicode.docx", bytes,
  )
  match rendered {
    Ok(value) =>
      assert_true(value.surface_html.contains("data-source-end=\"4\""))
    Err(error) => fail(error)
  }
}

///|
test "reader source index declines a wrapper it cannot prove" {
  let document = [
    @docx2html.paragraph([
      @docx2html.hyperlink(
        [@docx2html.run([@docx2html.text("linked text")])],
        href=Some("https://example.com"),
      ),
    ]),
  ]
  let bytes = @docx2html.write_docx(document)
  let rendered = render_document(
    "wrapped.docx", "/mock/office/wrapped.docx", bytes,
  )
  match rendered {
    Ok(value) => {
      assert_false(value.surface_html.contains("data-source-selector="))
      assert_true(value.surface_html.contains("<text "))
    }
    Err(error) => fail(error)
  }
}
