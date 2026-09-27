///|
test "document kind follows office file extensions" {
  assert_true(document_kind("notes.DOCX") is Some(Word))
  assert_true(document_kind("budget.xlsx") is Some(Excel))
  assert_true(document_kind("photo.pdf") is None)
}

test "standalone selections use an explicit mock office file path" {
  assert_eq(mock_document_file("report.docx"), "/mock/office/report.docx")
}

///|
test "spreadsheet body extraction keeps renderer styles" {
  let html = "<style>.s1{font-weight:bold}</style><body><h1>Book</h1></body>"
  let fragment = extract_html_body(html)
  assert_true(fragment.has_prefix("<style>"))
  assert_true(fragment.contains("<h1>Book</h1>"))
}

///|
test "word surface keeps page order" {
  let surface = word_surface(["<svg>A</svg>", "<svg>B</svg>"])
  let first = surface.find("data-page=\"1\"")
  let second = surface.find("data-page=\"2\"")
  assert_true(first is Some(_))
  assert_true(second is Some(_))
  assert_true(surface.find("<svg>A</svg>") is Some(_))
  assert_true(surface.find("<svg>B</svg>") is Some(_))
}
