///|
fn document_kind(name : String) -> DocumentKind? {
  let lower = name.to_lower()
  if lower.has_suffix(".docx") {
    Some(Word)
  } else if lower.has_suffix(".xlsx") {
    Some(Excel)
  } else {
    None
  }
}

/// The standalone browser has no host-owned document handle yet. Keep the
/// office CLI's `file` shape visible in the selection envelope with a stable
/// demo-only path; OpenSeek will replace this at the integration boundary.
fn mock_document_file(name : String) -> String {
  "/mock/office/" + name
}

///|
fn extract_html_body(document : String) -> String {
  let style = match (document.find("<style>"), document.find("</style>")) {
    (Some(start), Some(end)) => document[start:end + 8].to_owned()
    _ => ""
  }
  match (document.find("<body>"), document.find("</body>")) {
    (Some(start), Some(end)) => style + document[start + 6:end].to_owned()
    _ => style + document
  }
}

///|
fn word_surface(pages : Array[String]) -> String {
  let output = StringBuilder()
  output.write_string("<div class=\"word-pages\">")
  for index, page in pages {
    output.write_string(
      "<article class=\"word-page\" data-page=\"\{index + 1}\">",
    )
    output.write_string(page)
    output.write_string("</article>")
  }
  output.write_string("</div>")
  output.to_string()
}

///|
fn parse_document(payload : FilePayload) -> Result[LoadedDocument, String] {
  let kind = match document_kind(payload.name) {
    Some(kind) => kind
    None => return Err("Only .docx and .xlsx files are supported.")
  }
  try {
    match kind {
      Word => {
        let model = @docx.render_docx(payload.bytes)
        let pages = @svg.render_pages(model)
        Ok({
          kind,
          name: payload.name,
          file: mock_document_file(payload.name),
          surface_html: word_surface(pages),
          item_count: pages.length(),
          warning: "Word layout is read-only. Text selections retain source paths for Agent context.",
        })
      }
      Excel => {
        let workbook = @mbtexcel.read(payload.bytes.exact_view())
        let rendered = @xlsx2html.render_workbook(
          workbook,
          title=payload.name,
          max_rows=1000,
          max_cols=256,
          include_images=true,
          calc=false,
        )
        Ok({
          kind,
          name: payload.name,
          file: mock_document_file(payload.name),
          surface_html: "<div class=\"sheet-document\">" +
          extract_html_body(rendered) +
          "</div>",
          item_count: workbook.get_sheet_list().length(),
          warning: "Excel cells are read-only. Formula text is preserved when cached results are unavailable.",
        })
      }
    }
  } catch {
    error => Err(error.to_string())
  }
}
