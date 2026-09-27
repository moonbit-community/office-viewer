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

///|
/// The standalone browser has no host-owned document handle yet. Keep the
/// office CLI's `file` shape visible in the selection envelope with a stable
/// demo-only path; OpenSeek will replace this at the integration boundary.
fn mock_document_file(name : String) -> String {
  "/mock/office/" + name
}

///|
fn parse_document(payload : FilePayload) -> Result[LoadedDocument, String] {
  let kind = match document_kind(payload.name) {
    Some(kind) => kind
    None => return Err("Only .docx and .xlsx files are supported.")
  }
  let rendered = @viewer.render_document(
    payload.name,
    mock_document_file(payload.name),
    payload.bytes,
  )
  match rendered {
    Ok(document) =>
      Ok({
        kind,
        name: document.name,
        file: document.file,
        surface_html: document.surface_html,
        item_count: document.item_count,
        warning: document.warning,
      })
    Err(error) => Err(error)
  }
}
