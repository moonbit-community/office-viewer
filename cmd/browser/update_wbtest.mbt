///|
test "file changes retire pending annotation copies and PDF selections" {
  let previous = {
    ..initial_model(),
    generation: 7,
    annotations_copy_request: 12,
    copying_annotations: true,
    selection: "old selection",
    resolved_selection: "old resolved text",
  }
  let next = begin_read(previous)
  assert_eq(next.generation, 8)
  assert_eq(next.annotations_copy_request, 13)
  assert_false(next.copying_annotations)
  assert_eq(next.selection, "")
  assert_eq(next.resolved_selection, "")
  let (ignored, _) = update(next, AnnotationsCopied(12, true), _ => {
    @rabbita.none
  })
  assert_true(ignored == next)
}

///|
test "PDF browser dispatch keeps Office annotation extraction separate" {
  match parse_document({ name: "sample.pdf", bytes: b"%PDF-1.7", }) {
    Ok(document) => {
      assert_true(document.kind is Pdf)
      assert_eq(document.annotations, Json::null())
      assert_true(document.surface_html.contains("pdfViewer"))
    }
    Err(error) => fail(error)
  }
}
