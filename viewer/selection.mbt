///|
// The materialization layer keeps the browser selection envelope stable while
// giving a host an explicit choice about how much selected text enters an
// Agent context. The source references are derived from the existing
// office.selection.v1 fields. PDF.js references are explicitly separate from
// office.mbt selectors and retain an untruncated text-range reference.

///|
pub(all) enum SelectionMaterializationMode {
  Full
  Preview
  Deferred
} derive(Eq)

///|
pub extend SelectionMaterializationMode with Eq::{equal, not_equal}

///|
/// Options for `materialize_selection`. Preview limits are UTF-16 code units,
/// matching browser ranges and office.mbt reader offsets. DOCX/XLSX source
/// anchors are bounded; PDF materializations carry only their complete range
/// reference and never emit per-item anchors.
pub(all) struct SelectionMaterializationOptions {
  mode : SelectionMaterializationMode
  max_preview_utf16 : Int32
  max_source_anchors : Int32
} derive(Eq)

///|
pub extend SelectionMaterializationOptions with Eq::{equal, not_equal}

///|
pub let default_selection_preview_utf16 : Int32 = 1000

///|
pub let default_selection_source_anchors : Int32 = 64

///|
priv struct SelectionCell {
  row : Int32?
  column : Int32?
  text : String
}

///|
priv struct CanonicalSelection {
  fields : Map[String, Json]
  file : String
  format : String
  kind : String
  text : String
  anchors : Array[Json]
  source_anchors : Array[Json]
  selectors : Array[String]
  pages : Array[Int32]
  cells : Array[SelectionCell]
  sheet : String?
  range : String?
  pdf_reference : Json?
}

///|
fn json_string(value : Json?) -> String? {
  match value {
    Some(Json::String(text)) => Some(text)
    _ => None
  }
}

///|
fn json_int(value : Json?) -> Int32? {
  match value {
    Some(Json::Number(number, ..)) => Some(number.to_int())
    _ => None
  }
}

///|
fn json_bool(value : Json?) -> Bool? {
  match value {
    Some(Json::True) => Some(true)
    Some(Json::False) => Some(false)
    _ => None
  }
}

///|
fn json_array(value : Json?) -> Array[Json] {
  match value {
    Some(Json::Array(values)) => values
    _ => []
  }
}

///|
fn json_number(value : Int32) -> Json {
  Json::number(value.to_double(), repr=value.to_string())
}

///|
fn json_string_array(values : Array[String]) -> Json {
  Json::array(values.map(value => Json::string(value)))
}

///|
fn push_unique(values : Array[String], value : String) -> Unit {
  if value != "" && !values.contains(value) {
    values.push(value)
  }
}

///|
fn compact_anchor(value : Json) -> (Json, String?) {
  match value {
    Json::Object(fields) => {
      let selector = json_string(fields.get("selector"))
      let path = json_string(fields.get("path"))
      let part = json_string(fields.get("part"))
      let stability = json_string(fields.get("stability"))
      let source_id = json_string(fields.get("source_id"))
      let para_id = json_string(fields.get("para_id"))
      let reason = json_string(fields.get("reason"))
      let start_utf16 = json_int(fields.get("start_utf16"))
      let end_utf16 = json_int(fields.get("end_utf16"))
      let page = json_int(fields.get("page"))
      let actionable = json_bool(fields.get("actionable"))
      let paragraph_anchor_status = json_string(
        fields.get("paragraph_anchor_status"),
      )
      let runs = json_array(fields.get("runs")).filter_map(value => {
        json_string(Some(value))
      })
      let compact_fields : Array[(String, Json)] = []
      match source_id {
        Some(value) => compact_fields.push(("source_id", Json::string(value)))
        None => ()
      }
      match part {
        Some(value) => compact_fields.push(("part", Json::string(value)))
        None => ()
      }
      match path {
        Some(value) => compact_fields.push(("path", Json::string(value)))
        None => ()
      }
      match selector {
        Some(value) => compact_fields.push(("selector", Json::string(value)))
        None => ()
      }
      match stability {
        Some(value) => compact_fields.push(("stability", Json::string(value)))
        None => ()
      }
      match para_id {
        Some(value) => compact_fields.push(("para_id", Json::string(value)))
        None => ()
      }
      match reason {
        Some(value) => compact_fields.push(("reason", Json::string(value)))
        None => ()
      }
      match paragraph_anchor_status {
        Some(value) =>
          compact_fields.push(("paragraph_anchor_status", Json::string(value)))
        None => ()
      }
      match start_utf16 {
        Some(value) => compact_fields.push(("start_utf16", json_number(value)))
        None => ()
      }
      match end_utf16 {
        Some(value) => compact_fields.push(("end_utf16", json_number(value)))
        None => ()
      }
      match page {
        Some(value) => compact_fields.push(("page", json_number(value)))
        None => ()
      }
      match actionable {
        Some(value) => compact_fields.push(("actionable", Json::boolean(value)))
        None => ()
      }
      if !runs.is_empty() {
        compact_fields.push(("runs", json_string_array(runs)))
      }
      for
        key in [
          "item_index", "page_ref", "font_name", "dir", "has_eol", "geometry", "marked_content_id",
          "semantics",
        ] {
        match fields.get(key) {
          Some(value) => compact_fields.push((key, value))
          None => ()
        }
      }
      (Json::object(Map(compact_fields)), selector)
    }
    _ => (Json::object(Map([])), None)
  }
}

///|
fn parse_cells(value : Json?) -> Array[SelectionCell] {
  let output : Array[SelectionCell] = []
  for cell in json_array(value) {
    match cell {
      Json::Object(fields) => {
        let text = json_string(fields.get("text")).unwrap_or("")
        let row = json_int(fields.get("row"))
        let column = json_int(fields.get("column"))
        output.push({ row, column, text, })
      }
      _ => ()
    }
  }
  output
}

///|
fn cells_text(cells : Array[SelectionCell]) -> String {
  let output = StringBuilder()
  let mut has_cell = false
  let mut previous_row : Int32? = None
  let mut previous_column : Int32? = None
  for cell in cells {
    match (previous_row, cell.row) {
      (Some(previous), Some(row)) if previous == row => {
        let tabs = match (previous_column, cell.column) {
          (Some(previous), Some(column)) if column > previous =>
            column - previous
          _ => 1
        }
        for _ in 0..<tabs {
          output.write_char('\t')
        }
      }
      (Some(_), _) => output.write_char('\n')
      (None, _) if has_cell => output.write_char('\n')
      _ => ()
    }
    output.write_string(cell.text)
    has_cell = true
    previous_row = cell.row
    previous_column = cell.column
  }
  output.to_string()
}

///|
fn parse_pages(value : Json?) -> Array[Int32] {
  let output : Array[Int32] = []
  for page in json_array(value) {
    match page {
      Json::Number(value, ..) => output.push(value.to_int())
      _ => ()
    }
  }
  output
}

///|
fn pdf_reference(fields : Map[String, Json], file : String) -> Json? {
  guard json_string(fields.get("format")) == Some("pdf") &&
    json_string(fields.get("kind")) == Some("pdf") else {
    return None
  }
  guard fields.get("reference") is Some(reference) else { return None }
  guard reference
    is {
      "schema": String("office.pdf.text-range.v1"),
      "file": String(reference_file),
      "document": identity,
      "join": String("pdfjs-items-eol-v1"),
      "text_content": {
        "include_marked_content": True,
        "disable_normalization": True,
        ..
      },
      "start": {
        "page": Number(start_page, ..),
        "item_index": Number(start_item, ..),
        "offset_utf16": Number(start_offset, ..),
        ..
      },
      "end": {
        "page": Number(end_page, ..),
        "item_index": Number(end_item, ..),
        "offset_utf16": Number(end_offset, ..),
        ..
      },
      ..
    } else {
    return None
  }
  guard reference_file == file && fields.get("document") == Some(identity) else {
    return None
  }
  guard identity
    is {
      "sha256": String(hash),
      "byte_length": Number(byte_length, ..),
      "page_count": Number(page_count, ..),
      "parser": { "name": String("pdfjs"), "version": String(version), .. },
      ..
    } else {
    return None
  }
  guard hash.length() == 64 && version != "" else { return None }
  for unit in hash.code_units() {
    guard (unit >= 48 && unit <= 57) || (unit >= 97 && unit <= 102) else {
      return None
    }
  }
  for
    number in [
      start_page, end_page, start_item, end_item, start_offset, end_offset, page_count,
      byte_length,
    ] {
    guard number >= 0 && number == number.to_int().to_double() else {
      return None
    }
  }
  guard byte_length > 0 else { return None }
  guard fields.get("text") is Some(String(text)) && text.length() > 0 else {
    return None
  }
  if reference is Object(reference_fields) {
    match reference_fields.get("text_utf16_length") {
      None => ()
      Some(Number(length, ..)) if length == text.length().to_double() => ()
      _ => return None
    }
  }
  guard start_page >= 1 && end_page <= page_count else { return None }
  guard start_page < end_page ||
    (start_page == end_page && start_item < end_item) ||
    (
      start_page == end_page &&
      start_item == end_item &&
      start_offset < end_offset
    ) else {
    return None
  }
  Some(reference)
}

///|
fn parse_selection(
  selection_json : String,
) -> Result[CanonicalSelection, String] {
  let parsed = @json.parse(selection_json) catch {
    error => {
      let message = error.to_string()
      return Err("invalid selection JSON: \{message}")
    }
  }
  let fields = match parsed {
    Json::Object(fields) => fields
    _ => return Err("selection JSON must be an object")
  }
  let schema = match json_string(fields.get("schema")) {
    Some(value) => value
    None => return Err("selection.schema must be a JSON string")
  }
  if schema != "office.selection.v1" {
    return Err("unsupported selection schema: \{schema}")
  }
  let file = match json_string(fields.get("file")) {
    Some(value) => value
    None => return Err("selection.file must be a JSON string")
  }
  let format = match json_string(fields.get("format")) {
    Some(value) => value
    None => return Err("selection.format must be a JSON string")
  }
  let kind = match json_string(fields.get("kind")) {
    Some(value) => value
    None => return Err("selection.kind must be a JSON string")
  }
  let cells = parse_cells(fields.get("cells"))
  let text = match fields.get("text") {
    None => cells_text(cells)
    Some(Json::String(value)) => value
    Some(_) => return Err("selection.text must be a JSON string when present")
  }
  let anchors = json_array(fields.get("anchors"))
  let source_anchors : Array[Json] = []
  let selectors : Array[String] = []
  let selector = if format == "pdf" {
    None
  } else {
    json_string(fields.get("selector"))
  }
  match selector {
    Some(value) => {
      push_unique(selectors, value)
      let stability = json_string(fields.get("stability"))
      let compact_fields : Array[(String, Json)] = [
        ("selector", Json::string(value)),
      ]
      match stability {
        Some(value) => compact_fields.push(("stability", Json::string(value)))
        None => ()
      }
      source_anchors.push(Json::object(Map(compact_fields)))
    }
    None => ()
  }
  for anchor in anchors {
    let (compact, anchor_selector) = compact_anchor(anchor)
    if format == "pdf" && compact is Object(anchor_fields) {
      anchor_fields.remove("selector")
    }
    source_anchors.push(compact)
    match anchor_selector {
      Some(value) if format != "pdf" => push_unique(selectors, value)
      Some(_) => ()
      None => ()
    }
  }
  Ok({
    fields,
    file,
    format,
    kind,
    text,
    anchors,
    source_anchors,
    selectors,
    pages: parse_pages(fields.get("pages")),
    cells,
    sheet: json_string(fields.get("sheet")),
    range: json_string(fields.get("range")),
    pdf_reference: pdf_reference(fields, file),
  })
}

///|
fn safe_prefix(text : String, limit : Int32) -> String {
  if limit <= 0 {
    return ""
  }
  let (prefix, _) = text.split_at(limit)
  prefix.to_owned()
}

///|
fn safe_suffix(text : String, limit : Int32) -> String {
  if limit <= 0 {
    return ""
  }
  let mut requested = if limit > text.length() { text.length() } else { limit }
  while requested > 0 {
    let (_, suffix) = text.split_at(text.length() - requested)
    if suffix.length() <= limit {
      return suffix.to_owned()
    }
    requested -= 1
  }
  ""
}

///|
fn bounded_preview(text : String, limit : Int32) -> (String, Int32, Int32) {
  let head_limit = limit / 2
  let tail_limit = limit - head_limit
  let head = safe_prefix(text, head_limit)
  let tail = safe_suffix(text, tail_limit)
  let marker = "[...]"
  (head + marker + tail, head.length() + tail.length(), marker.length())
}

///|
fn copy_optional_field(
  source : Map[String, Json],
  target : Map[String, Json],
  name : String,
) -> Unit {
  match source.get(name) {
    Some(value) => target.set(name, value)
    None => ()
  }
}

///|
fn projected_selection(
  selection : CanonicalSelection,
  mode : SelectionMaterializationMode,
  materialized_text : String?,
) -> Json {
  if selection.format == "pdf" {
    Json::object(
      Map([
        ("schema", Json::string("office.selection.v1")),
        ("format", Json::string(selection.format)),
        ("kind", Json::string(selection.kind)),
        (
          "text",
          match materialized_text {
            Some(value) => Json::string(value)
            None => Json::null()
          },
        ),
      ]),
    )
  } else if mode is Full {
    let fields = selection.fields.copy()
    if selection.kind == "annotation" {
      fields.set(
        "annotation",
        annotation_selection_reference(
          selection.fields.get("annotation").unwrap_or(Json::null()),
        ),
      )
    }
    match materialized_text {
      Some(value) => fields.set("text", Json::string(value))
      None => ()
    }
    Json::object(fields)
  } else {
    let fields : Map[String, Json] = Map([
      ("schema", Json::string("office.selection.v1")),
      ("file", Json::string(selection.file)),
      ("format", Json::string(selection.format)),
      ("kind", Json::string(selection.kind)),
      ("anchor_count", json_number(selection.anchors.length())),
      ("cell_count", json_number(selection.cells.length())),
      ("selector_count", json_number(selection.selectors.length())),
    ])
    match materialized_text {
      Some(value) => fields.set("text", Json::string(value))
      None => fields.set("text", Json::null())
    }
    copy_optional_field(selection.fields, fields, "sheet")
    copy_optional_field(selection.fields, fields, "range")
    copy_optional_field(selection.fields, fields, "selector")
    copy_optional_field(selection.fields, fields, "stability")
    copy_optional_field(selection.fields, fields, "pages")
    if selection.kind == "annotation" {
      copy_optional_field(selection.fields, fields, "document")
      copy_optional_field(selection.fields, fields, "annotation_range")
      fields.set(
        "annotation",
        annotation_selection_reference(
          selection.fields.get("annotation").unwrap_or(Json::null()),
        ),
      )
    }
    for name in ["document", "reference", "page_info"] {
      copy_optional_field(selection.fields, fields, name)
    }
    Json::object(fields)
  }
}

///|
fn source_readable(selection : CanonicalSelection) -> Bool {
  selection.kind != "annotation" &&
  selection.file != "" &&
  (if selection.format == "pdf" {
    selection.pdf_reference is Some(_)
  } else {
    !selection.selectors.is_empty()
  })
}

///|
fn source_reason(selection : CanonicalSelection) -> String? {
  if selection.kind == "annotation" {
    Some(
      "annotation bodies require read_annotations on the same SHA-256 snapshot; a document selector would address the surrounding text, not the comment",
    )
  } else if source_readable(selection) {
    None
  } else if selection.file == "" {
    Some("selection.file is empty")
  } else if selection.format == "pdf" {
    Some(
      "no complete PDF text-range reference with matching file, document identity, and parser was present",
    )
  } else {
    Some(
      "no reader-friendly selector was present; physical-only anchors cannot be read back",
    )
  }
}

///|
fn materialized_json(
  selection : CanonicalSelection,
  options : SelectionMaterializationOptions,
) -> Result[String, String] {
  if options.max_preview_utf16 <= 0 {
    return Err("max_preview_utf16 must be greater than zero")
  }
  if options.max_source_anchors <= 0 {
    return Err("max_source_anchors must be greater than zero")
  }
  let text_length = selection.text.length()
  let (materialized_text, included_utf16, marker_utf16, representation) = match
    options.mode {
    Full => (Some(selection.text), text_length, 0, "exact")
    Preview if text_length <= options.max_preview_utf16 =>
      (Some(selection.text), text_length, 0, "exact")
    Preview => {
      let (preview, included, marker) = bounded_preview(
        selection.text,
        options.max_preview_utf16,
      )
      (Some(preview), included, marker, "head-tail")
    }
    Deferred => (None, 0, 0, "reference")
  }
  let omitted_utf16 = text_length - included_utf16
  let truncated = omitted_utf16 > 0
  let source_anchors : Array[Json] = []
  let source_limit = if selection.format == "pdf" {
    0
  } else if options.max_source_anchors > selection.source_anchors.length() {
    selection.source_anchors.length()
  } else {
    options.max_source_anchors
  }
  for index in 0..<source_limit {
    match selection.source_anchors.get(index) {
      Some(value) => source_anchors.push(value)
      None => ()
    }
  }
  let source_fields : Array[(String, Json)] = if selection.format == "pdf" {
    let fields : Array[(String, Json)] = [
      ("readable", Json::boolean(source_readable(selection))),
    ]
    match selection.pdf_reference {
      Some(reference) => fields.push(("reference", reference))
      None =>
        match source_reason(selection) {
          Some(reason) => fields.push(("reason", Json::string(reason)))
          None => ()
        }
    }
    fields
  } else {
    [
      ("file", Json::string(selection.file)),
      ("format", Json::string(selection.format)),
      ("kind", Json::string(selection.kind)),
      (
        "selector_schema",
        if selection.format == "pdf" || selection.kind == "annotation" {
          Json::null()
        } else {
          Json::string("office.selector/1")
        },
      ),
      ("selectors", json_string_array(selection.selectors)),
      ("anchors", Json::array(source_anchors)),
      ("anchor_count", json_number(selection.source_anchors.length())),
      (
        "anchors_truncated",
        Json::boolean(source_limit < selection.source_anchors.length()),
      ),
      ("readable", Json::boolean(source_readable(selection))),
    ]
  }
  if selection.kind == "annotation" {
    source_fields.push(
      ("annotation_schema", Json::string("office.annotations.v1")),
    )
    source_fields.push(("reference_kind", Json::string("annotation")))
    source_fields.push(("resolve_with", Json::string("read_annotations")))
    source_fields.push(
      ("document", selection.fields.get("document").unwrap_or(Json::null())),
    )
    source_fields.push(
      (
        "annotation",
        annotation_selection_reference(
          selection.fields.get("annotation").unwrap_or(Json::null()),
        ),
      ),
    )
    source_fields.push(
      (
        "annotation_range",
        selection.fields.get("annotation_range").unwrap_or(Json::null()),
      ),
    )
  }
  if selection.format != "pdf" {
    match source_reason(selection) {
      Some(reason) => source_fields.push(("reason", Json::string(reason)))
      None => ()
    }
  }
  let structure_fields : Array[(String, Json)] = [
    ("anchor_count", json_number(selection.anchors.length())),
    ("page_count", json_number(selection.pages.length())),
    ("cell_count", json_number(selection.cells.length())),
    ("selector_count", json_number(selection.selectors.length())),
  ]
  match selection.sheet {
    Some(value) => structure_fields.push(("sheet", Json::string(value)))
    None => ()
  }
  match selection.range {
    Some(value) => structure_fields.push(("range", Json::string(value)))
    None => ()
  }
  let content_fields : Array[(String, Json)] = [
    ("text_utf16_length", json_number(text_length)),
    ("included_utf16", json_number(included_utf16)),
    ("omitted_utf16", json_number(omitted_utf16)),
    ("marker_utf16", json_number(marker_utf16)),
    ("truncated", Json::boolean(truncated)),
    ("text_available", Json::boolean(materialized_text is Some(_))),
    ("representation", Json::string(representation)),
  ]
  let output_fields : Array[(String, Json)] = [
    ("schema", Json::string("office.selection.materialized.v1")),
    (
      "mode",
      Json::string(
        match options.mode {
          Full => "full"
          Preview => "preview"
          Deferred => "deferred"
        },
      ),
    ),
    (
      "selection",
      projected_selection(selection, options.mode, materialized_text),
    ),
    ("content", Json::object(Map(content_fields))),
    ("source", Json::object(Map(source_fields))),
  ]
  if selection.format != "pdf" {
    output_fields.push(("structure", Json::object(Map(structure_fields))))
  }
  let output = Json::object(Map(output_fields))
  Ok(output.stringify(indent=2))
}

///|
/// Materialize an `office.selection.v1` envelope for an Agent-facing host.
/// `Full` carries the exact selected text, `Preview` carries a deterministic
/// head/tail excerpt bounded by UTF-16 units, and `Deferred` carries only
/// reader-friendly source references. The returned string is JSON with schema
/// `office.selection.materialized.v1`.
pub fn materialize_selection(
  selection_json : String,
  options : SelectionMaterializationOptions,
) -> Result[String, String] {
  match parse_selection(selection_json) {
    Ok(selection) => materialized_json(selection, options)
    Err(error) => Err(error)
  }
}

///|
/// Return the exact selection and its source references.
pub fn materialize_selection_full(
  selection_json : String,
) -> Result[String, String] {
  materialize_selection(selection_json, {
    mode: Full,
    max_preview_utf16: default_selection_preview_utf16,
    max_source_anchors: default_selection_source_anchors,
  })
}

///|
/// Return a deterministic exact head/tail preview. The limit is measured in
/// UTF-16 code units and excludes the `[ ... ]` marker inserted between sides.
pub fn materialize_selection_preview(
  selection_json : String,
  max_preview_utf16 : Int32,
) -> Result[String, String] {
  materialize_selection(selection_json, {
    mode: Preview,
    max_preview_utf16,
    max_source_anchors: default_selection_source_anchors,
  })
}

///|
/// Return source references for an office.mbt read or PDF.js resolution.
pub fn materialize_selection_deferred(
  selection_json : String,
) -> Result[String, String] {
  materialize_selection(selection_json, {
    mode: Deferred,
    max_preview_utf16: default_selection_preview_utf16,
    max_source_anchors: default_selection_source_anchors,
  })
}
