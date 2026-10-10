///|
// Read-only annotation projection. The JSON contract is independent of DOM,
// pagination, and the standalone inspector. Missing source facts stay null.
fn annotation_string(value : String?) -> Json {
  match value {
    Some(text) => Json::string(text)
    None => Json::null()
  }
}

///|
fn annotation_bool(value : Bool?) -> Json {
  match value {
    Some(flag) => Json::boolean(flag)
    None => Json::null()
  }
}

///|
// Preserve explicit line breaks as well as paragraph separators. The general
// reader's raw-text helper omits Break nodes, which would join comment words.
fn annotation_text(node : @document.DocumentElement) -> String {
  match node {
    @document.DocumentElement::Paragraph(children~, ..) =>
      annotation_children_text(children) + "\n\n"
    @document.DocumentElement::Document(children~, ..)
    | @document.DocumentElement::Run(children~, ..)
    | @document.DocumentElement::Hyperlink(children~, ..)
    | @document.DocumentElement::Table(children~, ..)
    | @document.DocumentElement::TableRow(children~, ..)
    | @document.DocumentElement::TableCell(children~, ..) =>
      annotation_children_text(children)
    @document.DocumentElement::Break(_) => "\n"
    _ => node.to_raw_text()
  }
}

///|
fn annotation_children_text(nodes : Array[@document.DocumentElement]) -> String {
  let text = StringBuilder()
  for node in nodes {
    text.write_string(annotation_text(node))
  }
  text.to_string()
}

///|
fn annotation_envelope(
  name : String,
  file : String,
  format : String,
  items : Array[Json],
  support : Json,
  warnings : Array[String],
) -> Json {
  Json::object({
    "schema": Json::string("office.annotations.v1"),
    "file": Json::string(file),
    "document": Json::object({
      "name": Json::string(name),
      "format": Json::string(format),
      "identity": Json::string("host-file"),
    }),
    "read_only": Json::boolean(true),
    "items": Json::array(items),
    "support": support,
    "warnings": json_string_array(warnings),
  })
}

///|
// Bind every item to the exact input snapshot. Ordinal IDs alone repeat in
// other files and revisions and must never be sufficient for UI navigation.
fn bind_annotation_snapshot(envelope : Json, bytes : Bytes) -> Json {
  let sha256 = @crypto.bytes_to_hex_string(@crypto.sha256(bytes))
  match envelope {
    Json::Object(fields) => {
      let identity = Json::object({
        "file": fields.get("file").unwrap_or(Json::null()),
        "sha256": Json::string(sha256),
      })
      match fields.get("document") {
        Some(Json::Object(document)) => {
          document["identity"] = Json::string("host-file+sha256")
          document["sha256"] = Json::string(sha256)
        }
        _ => ()
      }
      for item in json_array(fields.get("items")) {
        match item {
          Json::Object(item_fields) => item_fields["document"] = identity
          _ => ()
        }
      }
    }
    _ => ()
  }
  envelope
}

///|
// Only direct main-story paragraphs are rendered with a proven source join.
// A source node denotes context, never an invented character-level range.
fn annotation_body_paragraph(path : String) -> String? {
  if !path.has_prefix("/body/p[") {
    return None
  }
  match path.find("]") {
    Some(end) => {
      let paragraph = path[6:end + 1].to_owned()
      if path.length() == end + 1 || path[end + 1:].has_prefix("/") {
        Some(paragraph)
      } else {
        None
      }
    }
    None => None
  }
}

///|
fn word_annotations(
  name : String,
  file : String,
  annotated : @reader.DocxAnnotatedResult,
) -> Json {
  let comments = annotated.annotations().comments()
  let bodies : Map[String, Array[Json]] = Map([])
  for (id, body) in annotated.comment_bodies() {
    let candidate = Json::string(annotation_children_text(body))
    match bodies.get(id) {
      Some(candidates) => candidates.push(candidate)
      None => bodies[id] = [candidate]
    }
  }
  let ids : Map[String, Array[String]] = Map([])
  for ordinal, comment in comments {
    let id = "docx.comment:\{ordinal}"
    match ids.get(comment.id()) {
      Some(matches) => matches.push(id)
      None => ids[comment.id()] = [id]
    }
  }
  let items : Array[Json] = []
  for ordinal, comment in comments {
    let candidates = bodies.get(comment.id()).unwrap_or([])
    let body_unique = ids.get(comment.id()).unwrap_or([]).length() == 1 &&
      candidates.length() == 1
    let parent_id = match comment.parent_id() {
      Some(source_id) =>
        match ids.get(source_id) {
          Some([id]) => Some(id)
          _ => None
        }
      None => None
    }
    let source_anchors : Array[Json] = []
    let paths : Array[String] = []
    for anchor in comment.anchors() {
      source_anchors.push(
        Json::object({
          "story": Json::string(anchor.story()),
          "start": annotation_string(anchor.start()),
          "end": annotation_string(anchor.end()),
          "start_boundary": annotation_string(
            anchor.start_boundary().map(value => value.to_string()),
          ),
          "end_boundary": annotation_string(
            anchor.end_boundary().map(value => value.to_string()),
          ),
          "references": json_string_array(anchor.references()),
        }),
      )
      // Keep every explicit endpoint/reference, without filling guessed
      // intermediate paragraphs or borrowing a parent's anchor for a reply.
      let endpoints = anchor.references()
      match anchor.start() {
        Some(value) => endpoints.push(value)
        None => ()
      }
      match anchor.end() {
        Some(value) => endpoints.push(value)
        None => ()
      }
      for endpoint in endpoints {
        match annotation_body_paragraph(endpoint) {
          Some(path) => push_unique(paths, path)
          None => ()
        }
      }
    }
    items.push(
      Json::object({
        "id": Json::string("docx.comment:\{ordinal}"),
        "type": Json::string("docx.comment"),
        "source_id": Json::string(comment.id()),
        "defined": Json::boolean(comment.defined()),
        "text": if body_unique {
          candidates[0]
        } else {
          Json::null()
        },
        "text_status": Json::string(
          if body_unique {
            "available"
          } else if candidates.is_empty() {
            "missing"
          } else {
            "ambiguous"
          },
        ),
        "body_candidates": Json::array(
          if body_unique {
            []
          } else {
            candidates
          },
        ),
        "author": annotation_string(comment.author()),
        "initials": annotation_string(comment.initials()),
        "date": annotation_string(comment.date()),
        "resolved": annotation_bool(comment.done()),
        "parent_id": annotation_string(parent_id),
        "parent_source_id": annotation_string(comment.parent_id()),
        "thread_status": Json::string(
          if parent_id is Some(_) {
            "reply"
          } else if comment.parent_id() is Some(_) {
            "unresolved-parent"
          } else {
            "root-or-unknown"
          },
        ),
        "source_anchors": Json::array(source_anchors),
        "anchor": Json::object({
          "kind": Json::string("docx"),
          "part": if paths.is_empty() {
            Json::null()
          } else {
            Json::string(annotated.main_story_part())
          },
          "granularity": Json::string("paragraph-context"),
          "confidence": Json::string(
            if comment.anchors().is_empty() {
              "unanchored"
            } else {
              "source-node"
            },
          ),
          "preview": Json::string("source-only"),
          "ranges": Json::array([]),
          "locations": Json::array(
            paths.map(path => {
              Json::object({
                "path": Json::string(path),
                "part": Json::string(annotated.main_story_part()),
              })
            }),
          ),
        }),
      }),
    )
  }
  let sidecars = annotated.annotation_sidecars()
  let warnings = annotated.annotations().warnings()
  if !sidecars.is_empty() {
    warnings.push(
      "Modern Word comment sidecars are present but not interpreted; the comment/thread projection may be incomplete.",
    )
  }
  annotation_envelope(
    name,
    file,
    "docx",
    items,
    Json::object({
      "scope": json_string_array(["docx.comment", "docx.w15-thread"]),
      "completeness": Json::string(
        if warnings.is_empty() {
          "complete-supported-kinds"
        } else {
          "partial"
        },
      ),
      "modern_comments": Json::string("unsupported"),
      "detected_unsupported_sidecar_types": json_string_array(sidecars),
      "exact_text_ranges": Json::string("unavailable"),
    }),
    warnings,
  )
}

///|
fn excel_annotations(
  name : String,
  file : String,
  workbook : @xlsx.Workbook,
) -> Json raise {
  let items : Array[Json] = []
  for sheet_index, sheet in workbook.get_sheet_list() {
    for ordinal, note in workbook.get_comments(sheet) {
      let (column, row) = @mbtexcel.cell_name_to_coordinates(note.cell)
      items.push(
        Json::object({
          "id": Json::string("xlsx.note:\{sheet_index}:\{ordinal}"),
          "type": Json::string("xlsx.note"),
          "source_id": Json::string(note.cell),
          "defined": Json::boolean(true),
          "text": Json::string(note.text),
          "text_status": Json::string("available"),
          "author": Json::string(note.author),
          "author_id": match note.author_id {
            Some(value) => json_number(value)
            None => Json::null()
          },
          "initials": Json::null(),
          "date": Json::null(),
          "resolved": Json::null(),
          "parent_id": Json::null(),
          "parent_source_id": Json::null(),
          "thread_status": Json::string("not-applicable"),
          "anchor": Json::object({
            "kind": Json::string("xlsx"),
            "sheet": Json::string(sheet),
            "cell": Json::string(note.cell),
            "row": json_number(row),
            "column": json_number(column),
            "confidence": Json::string("exact-cell"),
            "preview": Json::string(
              if row <= 1000 && column <= 256 {
                "within-preview"
              } else {
                "outside-preview"
              },
            ),
          }),
        }),
      )
    }
  }
  annotation_envelope(
    name,
    file,
    "xlsx",
    items,
    Json::object({
      "scope": json_string_array(["xlsx.note"]),
      "completeness": Json::string("partial"),
      "traditional_notes": Json::string("supported"),
      "threaded_comments": Json::string("unsupported"),
      "threaded_comments_presence": Json::string("unknown"),
      "preview_max_rows": json_number(1000),
      "preview_max_columns": json_number(256),
    }),
    [
      "Only traditional Excel Notes are read. Modern threaded comments are unsupported and their presence is not detected. An empty Notes list does not prove the workbook has no comments.",
    ],
  )
}

///|
/// Read existing DOCX comments/w15 replies or XLSX traditional Notes without
/// rendering a document, requiring a browser, or modifying source bytes.
/// Returns the versioned office.annotations.v1 JSON value. Item IDs are
/// snapshot-local; each item also carries the host file identity and source
/// byte SHA-256 so a stale item cannot navigate a different file or revision.
pub fn read_annotations(
  name : String,
  file : String,
  bytes : Bytes,
) -> Result[Json, String] {
  try {
    match document_kind(name) {
      Some(Word) =>
        Ok(
          bind_annotation_snapshot(
            word_annotations(name, file, @docx2html.read_docx_annotated(bytes)),
            bytes,
          ),
        )
      Some(Excel) =>
        Ok(
          bind_annotation_snapshot(
            excel_annotations(name, file, @mbtexcel.read(bytes.exact_view())),
            bytes,
          ),
        )
      None => Err("Only .docx and .xlsx files are supported.")
    }
  } catch {
    error => Err(error.to_string())
  }
}

///|
// Safe inside the renderer's inert application/json data block.
fn annotations_script(annotations : Json) -> String {
  "<script type=\"application/json\" class=\"office-annotations-metadata\">" +
  annotations.stringify().replace(old="<", new="\\u003c") +
  "</script>"
}
