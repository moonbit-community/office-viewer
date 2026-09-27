///|
test "document kind follows office file extensions" {
  assert_true(document_kind("notes.DOCX") is Some(Word))
  assert_true(document_kind("budget.xlsx") is Some(Excel))
  assert_true(document_kind("photo.pdf") is None)
}

///|
test "standalone selections use an explicit mock office file path" {
  assert_eq(mock_document_file("report.docx"), "/mock/office/report.docx")
}
