///|
// Keep annotation provenance separate from the body. In particular, preview
// and deferred selections must never carry `text` or `body_candidates` here.
fn annotation_selection_reference(annotation : Json) -> Json {
  let output : Map[String, Json] = Map([])
  match annotation {
    Json::Object(fields) =>
      for
        key in [
          "id", "type", "source_id", "defined", "text_status", "author", "initials",
          "date", "resolved", "parent_id", "parent_source_id", "thread_status", "anchor",
          "source_anchors",
        ] {
        copy_optional_field(fields, output, key)
      }
    _ => ()
  }
  Json::object(output)
}

///|
/// Select a whole annotation body or an exact UTF-16 span within it. The item
/// must come from `read_annotations` or `RenderedDocument.annotations`; its
/// source-byte identity is retained. Off-preview and unanchored items are valid
/// selections, but missing/ambiguous bodies and split surrogate pairs are not.
/// `scope` is `whole` (default) or `text`; whole-body spans must cover the body.
pub fn annotation_selection(
  annotation : Json,
  start_utf16~ : Int32 = 0,
  end_utf16? : Int32,
  scope~ : String = "whole",
) -> Result[Json, String] {
  let fields = match annotation {
    Json::Object(fields) => fields
    _ => return Err("annotation must be an object")
  }
  let text = match fields.get("text") {
    Some(Json::String(text)) => text
    _ => return Err("annotation body is unavailable or ambiguous")
  }
  if json_string(fields.get("text_status")) != Some("available") {
    return Err("annotation body is unavailable or ambiguous")
  }
  let format = match json_string(fields.get("type")) {
    Some("docx.comment") => "docx"
    Some("xlsx.note") => "xlsx"
    _ => return Err("unsupported annotation type")
  }
  match json_string(fields.get("id")) {
    Some(id) if id != "" => ()
    _ => return Err("annotation.id must be a nonempty string")
  }
  let document = match fields.get("document") {
    Some(Json::Object(identity)) => identity
    _ => return Err("annotation has no document snapshot identity")
  }
  let file = match json_string(document.get("file")) {
    Some(file) => file
    None => return Err("annotation.document.file must be a string")
  }
  let sha256 = match json_string(document.get("sha256")) {
    Some(hash) if hash.length() == 64 => hash
    _ => return Err("annotation.document.sha256 must identify the source bytes")
  }
  let end = end_utf16.unwrap_or(text.length())
  if start_utf16 < 0 || end < start_utf16 || end > text.length() {
    return Err("annotation body UTF-16 range is out of bounds")
  }
  if scope != "whole" && scope != "text" {
    return Err("annotation selection scope must be whole or text")
  }
  if scope == "whole" && (start_utf16 != 0 || end != text.length()) {
    return Err("whole annotation selection must cover the entire body")
  }
  if scope == "text" && end == start_utf16 {
    return Err("annotation text selection must not be empty")
  }
  let selected = match text.get_view(start=start_utf16, end~) {
    Some(view) => view.to_owned()
    None =>
      return Err("annotation selection must not split a UTF-16 surrogate pair")
  }
  Ok(
    Json::object({
      "schema": Json::string("office.selection.v1"),
      "file": Json::string(file),
      "format": Json::string(format),
      "kind": Json::string("annotation"),
      "stability": Json::string("snapshot-relative"),
      "document": Json::object({
        "file": Json::string(file),
        "sha256": Json::string(sha256),
      }),
      "text": Json::string(selected),
      "annotation": annotation_selection_reference(annotation),
      "annotation_range": Json::object({
        "scope": Json::string(scope),
        "start_utf16": json_number(start_utf16),
        "end_utf16": json_number(end),
        "body_utf16_length": json_number(text.length()),
      }),
      // A comment's source context is not a document-text selection.
      "anchors": Json::array([]),
    }),
  )
}
