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

///|
fn comment_source_test_package(
  body : String,
  furniture~ : Bool = false,
) -> Bytes raise {
  let builder = @opc.PackageBuilder::new()
  builder.add_default(
    "rels", "application/vnd.openxmlformats-package.relationships+xml",
  )
  let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  let r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  let content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml."
  let references = if furniture {
    "<w:headerReference w:type=\"default\" r:id=\"header\"/><w:footerReference w:type=\"default\" r:id=\"footer\"/>"
  } else {
    ""
  }
  builder.add_part(
    "word/document.xml",
    @utf8.encode(
      "<w:document xmlns:w=\"\{w}\" xmlns:r=\"\{r}\"><w:body>" +
      body +
      "<w:sectPr>" +
      references +
      "<w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>",
    ),
  )
  builder.add_override("/word/document.xml", content_type + "document.main+xml")
  builder.add_relationship(
    "document",
    r + "/officeDocument",
    "word/document.xml",
  )
  builder.add_part(
    "word/comments.xml",
    @utf8.encode(
      "<w:comments xmlns:w=\"\{w}\"><w:comment w:id=\"0\" w:author=\"Reviewer\"><w:p><w:r><w:t>Check</w:t></w:r></w:p></w:comment></w:comments>",
    ),
  )
  builder.add_override("/word/comments.xml", content_type + "comments+xml")
  builder.add_relationship(
    "comments",
    r + "/comments",
    "comments.xml",
    source="word/document.xml",
  )
  if furniture {
    for
      (kind, tag, text) in [
        ("header", "hdr", "Running header"),
        ("footer", "ftr", "Running footer"),
      ] {
      builder.add_part(
        "word/\{kind}.xml",
        @utf8.encode(
          "<w:\{tag} xmlns:w=\"\{w}\"><w:p><w:r><w:t>\{text}</w:t></w:r></w:p></w:\{tag}>",
        ),
      )
      builder.add_override("/word/\{kind}.xml", content_type + "\{kind}+xml")
      builder.add_relationship(
        kind,
        r + "/" + kind,
        "\{kind}.xml",
        source="word/document.xml",
      )
    }
  }
  builder.build()
}

///|
test "reader source index tolerates an interior comment-reference run" {
  let bytes = comment_source_test_package(
    "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t>Before</w:t></w:r><w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r><w:r><w:t xml:space=\"preserve\"> after</w:t></w:r></w:p>",
  )
  let rendered = render_document("comment.docx", "/demo/comment.docx", bytes).unwrap()
  assert_true(rendered.surface_html.contains("data-source-path=\"p[1]\""))
  assert_true(rendered.surface_html.contains("data-source-runs=\"p[1]/r[3]\""))
}

///|
test "running header and footer do not discard body comment anchors" {
  let bytes = comment_source_test_package(
    "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t>Body text</w:t></w:r><w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>" +
    "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p><w:p><w:r><w:t>Second page</w:t></w:r></w:p>",
    furniture=true,
  )
  let rendered = render_document(
    "furniture.docx", "/demo/furniture.docx", bytes,
  ).unwrap()
  assert_true(rendered.surface_html.contains("Running header"))
  assert_true(rendered.surface_html.contains("Running footer"))
  assert_true(rendered.surface_html.contains("data-source-path=\"p[1]\""))
  assert_true(rendered.surface_html.contains("data-source-path=\"p[3]\""))
  assert_eq(rendered.surface_html.split("data-source-path=").count(), 3)
  assert_eq(rendered.item_count, 2)
}

///|
test "proven comment context survives an unrelated unjoined paragraph" {
  let bytes = @docx2html.write_docx([
    @docx2html.paragraph([@docx2html.run([@docx2html.text("Commented text")])]),
    @docx2html.paragraph([
      @docx2html.hyperlink(
        [@docx2html.run([@docx2html.text("Unsupported wrapper")])],
        href=Some("https://example.com"),
      ),
    ]),
  ])
  let rendered = render_document("partial.docx", "/demo/partial.docx", bytes).unwrap()
  assert_true(rendered.surface_html.contains("data-comment-path=\"p[1]\""))
  assert_false(rendered.surface_html.contains("data-comment-path=\"p[2]\""))
  assert_false(rendered.surface_html.contains("data-source-selector="))
}

///|
test "a table cell matching later body text never becomes its comment anchor" {
  let paragraph = () => {
    @docx2html.paragraph([@docx2html.run([@docx2html.text("Same text")])])
  }
  let bytes = @docx2html.write_docx([
    @docx2html.paragraph([@docx2html.run([@docx2html.text("Before table")])]),
    @docx2html.table([
      @docx2html.table_row([@docx2html.table_cell([paragraph()])]),
    ]),
    paragraph(),
  ])
  let rendered = render_document("table.docx", "/demo/table.docx", bytes).unwrap()
  assert_true(rendered.surface_html.contains("data-comment-path=\"p[1]\""))
  assert_false(rendered.surface_html.contains("data-comment-path=\"p[2]\""))
  assert_false(rendered.surface_html.contains("data-source-selector="))
}
