///|
fn annotation_test_field(value : Json, key : String) -> Json {
  match value {
    Json::Object(fields) => fields.get(key).unwrap_or(Json::null())
    _ => Json::null()
  }
}

///|
fn annotation_test_body(text : String) -> Array[@document.DocumentElement] {
  [@docx2html.paragraph([@docx2html.run([@docx2html.text(text)])])]
}

///|
test "annotation contract preserves Word replies and source context without guessed ranges" {
  let bytes = @docx2html.write_docx_with_comments(
    annotation_test_body("Body"),
    [
      @docx2html.comment_spec(
        author="Ada",
        from=0,
        to=0,
        annotation_test_body("Check"),
      ),
      @docx2html.comment_reply(
        author="Bob",
        reply_to=0,
        done=true,
        annotation_test_body("Done"),
      ),
    ],
  )
  let envelope = match
    read_annotations("thread.docx", "/docs/thread.docx", bytes) {
    Ok(value) => value
    Err(error) => fail(error)
  }
  assert_eq(
    annotation_test_field(envelope, "schema"),
    Json::string("office.annotations.v1"),
  )
  assert_eq(
    annotation_test_field(envelope, "file"),
    Json::string("/docs/thread.docx"),
  )
  let items = json_array(Some(annotation_test_field(envelope, "items")))
  assert_eq(items.length(), 2)
  assert_eq(annotation_test_field(items[0], "text"), Json::string("Check\n\n"))
  assert_eq(
    annotation_test_field(items[1], "parent_id"),
    annotation_test_field(items[0], "id"),
  )
  assert_eq(annotation_test_field(items[1], "resolved"), Json::boolean(true))
  assert_eq(annotation_test_field(items[1], "date"), Json::null())
  assert_eq(annotation_test_field(items[1], "source_anchors"), Json::array([]))
  assert_eq(
    annotation_test_field(annotation_test_field(items[0], "anchor"), "ranges"),
    Json::array([]),
  )
  let identity = annotation_test_field(items[0], "document")
  assert_eq(
    annotation_test_field(identity, "file"),
    Json::string("/docs/thread.docx"),
  )
  assert_eq(
    annotation_test_field(identity, "sha256"),
    annotation_test_field(annotation_test_field(envelope, "document"), "sha256"),
  )
  match render_document("thread.docx", "/docs/thread.docx", bytes) {
    Ok(rendered) =>
      assert_true(rendered.surface_html.contains("data-source-path=\"p[1]\""))
    Err(error) => fail(error)
  }
}

///|
test "annotation Notes include other sheets and notes beyond preview" {
  let workbook = @mbtexcel.new_workbook()
  ignore(workbook.add_sheet("Visible"))
  ignore(workbook.add_sheet("Other"))
  for (sheet, cell) in [("Visible", "A1"), ("Other", "IW1001")] {
    workbook.add_comment(sheet, {
      cell,
      author: "Ada",
      author_id: None,
      text: "Note",
      paragraph: [],
      width: None,
      height: None,
    })
  }
  let envelope = excel_annotations("notes.xlsx", "/docs/notes.xlsx", workbook)
  let items = json_array(Some(annotation_test_field(envelope, "items")))
  assert_eq(items.length(), 2)
  let anchor = annotation_test_field(items[1], "anchor")
  assert_eq(annotation_test_field(anchor, "sheet"), Json::string("Other"))
  assert_eq(
    annotation_test_field(anchor, "preview"),
    Json::string("outside-preview"),
  )
  assert_eq(annotation_test_field(items[0], "resolved"), Json::null())
  assert_eq(annotation_test_field(items[0], "parent_id"), Json::null())
  assert_eq(
    annotation_test_field(
      annotation_test_field(envelope, "support"),
      "threaded_comments_presence",
    ),
    Json::string("unknown"),
  )
}

///|
test "annotation list survives a Word document with no proven visual join" {
  let body = [
    @docx2html.paragraph([
      @docx2html.hyperlink(
        [@docx2html.run([@docx2html.text("Link")])],
        href=Some("https://example.com"),
      ),
    ]),
  ]
  let bytes = @docx2html.write_docx_with_comments(body, [
    @docx2html.comment_spec(
      author="Ada",
      from=0,
      to=0,
      annotation_test_body("Still readable"),
    ),
  ])
  match render_document("wrapped.docx", "/docs/wrapped.docx", bytes) {
    Ok(value) => {
      assert_false(value.surface_html.contains("data-source-selector="))
      assert_eq(
        json_array(Some(annotation_test_field(value.annotations, "items"))).length(),
        1,
      )
      assert_true(value.surface_html.contains("office-annotations-metadata"))
    }
    Err(error) => fail(error)
  }
}

///|
test "annotation serialization escapes script terminators and preserves unknown booleans" {
  assert_eq(
    annotation_children_text([
      @docx2html.run([
        @docx2html.text("line one"),
        @docx2html.line_break(),
        @docx2html.text("line two"),
      ]),
    ]),
    "line one\nline two",
  )
  let script = annotations_script(
    Json::object({ "text": Json::string("</script><img src=x>") }),
  )
  assert_false(script.contains("</script><img"))
  assert_true(script.contains("\\u003c/script>"))
  assert_eq(annotation_bool(None), Json::null())
  assert_eq(annotation_bool(Some(false)), Json::boolean(false))
  assert_eq(annotation_body_paragraph("/body/p[3]/r[2]"), Some("p[3]"))
  assert_eq(annotation_body_paragraph("/headers/header[1]/p[3]"), None)
}
