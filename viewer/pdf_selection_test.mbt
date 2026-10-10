///|
fn pdf_fixture() -> Json {
  let identity : Json = {
    "sha256": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
    "byte_length": 1000,
    "page_count": 4,
    "parser": { "name": "pdfjs", "version": "6.4.299" },
  }
  {
    "schema": "office.selection.v1",
    "file": "/mock/office/report.pdf",
    "format": "pdf",
    "kind": "pdf",
    "text": "abcdefghij",
    "browser_text": "abcdefghij",
    "document": identity,
    "reference": {
      "schema": "office.pdf.text-range.v1",
      "file": "/mock/office/report.pdf",
      "document": identity,
      "join": "pdfjs-items-eol-v1",
      "text_content": {
        "include_marked_content": true,
        "disable_normalization": true,
      },
      "start": { "page": 1, "item_index": 2, "offset_utf16": 3 },
      "end": { "page": 2, "item_index": 8, "offset_utf16": 5 },
      "text_utf16_length": 10,
    },
    "anchors": [
      {
        "page": 1,
        "page_ref": { "num": 3, "gen": 0 },
        "item_index": 2,
        "start_utf16": 3,
        "end_utf16": 13,
        "text": "abcdefghij",
        "stability": "document-and-parser-relative",
        "dir": "ltr",
        "geometry": {
          "space": "pdf-page-user-space",
          "precision": "selected-dom-range",
          "bbox": [1, 2, 30, 40],
        },
        "semantics": { "role_path": ["Root", "H1"], "lang": "en-US" },
      },
    ],
    "pages": [1, 2],
    "page_info": [{ "page": 1, "user_unit": 2, "view_box": [20, 10, 380, 280] }],
  }
}

///|
fn pdf_object(value : Json) -> Map[String, Json] raise {
  guard value is Object(fields) else { fail("expected a JSON object") }
  fields
}

///|
fn pdf_source(
  selection : Json,
  mode : @viewer.SelectionMaterializationMode,
) -> Json raise {
  let result = @viewer.materialize_selection(selection.stringify(), {
    mode,
    max_preview_utf16: 4,
    max_source_anchors: 64,
  })
  match result {
    Ok(value) => @json.parse(value)
    Err(error) => fail(error)
  }
}

///|
test "PDF materialization keeps a native reference in all three modes" {
  for mode in [@viewer.Full, @viewer.Preview, @viewer.Deferred] {
    let selection = pdf_fixture()
    let materialized = pdf_source(selection, mode)
    guard materialized
      is {
        "source": {
          "readable": True,
          "reference_complete": True,
          "selector_schema": Null,
          "selectors": Array(selectors),
          "reference": reference,
          "anchors": Array(anchors),
          ..
        },
        "selection": projected,
        ..
      } else {
      fail(materialized.stringify())
    }
    assert_true(selectors.is_empty())
    assert_eq(pdf_object(selection)["reference"], reference)
    assert_true(
      anchors[0]
      is {
        "item_index": Number(2, ..),
        "semantics": { "role_path": Array(_), .. },
        "geometry": { "space": String("pdf-page-user-space"), .. },
        ..
      },
    )
    assert_true(
      projected is { "page_info": Array(_), "document": Object(_), .. },
    )
    match mode {
      @viewer.Full =>
        assert_true(
          projected
          is { "text": String("abcdefghij"), "browser_text": String(_), .. },
        )
      @viewer.Preview => {
        assert_true(projected is { "text": String("ab[...]ij"), .. })
        assert_false(materialized.stringify().contains("abcdefghij"))
      }
      @viewer.Deferred => {
        assert_true(projected is { "text": Null, .. })
        assert_false(materialized.stringify().contains("abcdefghij"))
      }
    }
  }
}

///|
test "PDF anchor truncation retains the complete cross-page range" {
  let selection = pdf_fixture()
  let fields = pdf_object(selection)
  let anchor = pdf_object(selection)["anchors"]
  guard anchor is Array(anchors) else { fail("missing fixture anchors") }
  fields.set("anchors", Json::array(Array::make(100, anchors[0])))
  let result = pdf_source(selection, @viewer.Deferred)
  guard result
    is {
      "source": {
        "readable": True,
        "reference_complete": True,
        "anchor_count": Number(100, ..),
        "anchors_truncated": True,
        "anchors": Array(compact),
        "reference": {
          "start": { "page": Number(1, ..), .. },
          "end": { "page": Number(2, ..), .. },
          ..
        },
        ..
      },
      ..
    } else {
    fail(result.stringify())
  }
  assert_eq(compact.length(), 64)
}

///|
test "malformed PDF references fail closed without discarding selected text" {
  for
    (key, replacement) in [
      ("schema", Json::string("unknown")),
      ("file", Json::string("/different.pdf")),
      ("document", Json::object({})),
      ("join", Json::string("unknown")),
      (
        "text_content",
        (
          { "include_marked_content": true, "disable_normalization": false } :
          Json),
      ),
      ("start", ({ "page": 0, "item_index": 2, "offset_utf16": 3 } : Json)),
      ("end", ({ "page": 5, "item_index": 8, "offset_utf16": 5 } : Json)),
      ("end", ({ "page": 1, "item_index": 2, "offset_utf16": 3 } : Json)),
      ("end", ({ "page": 1, "item_index": 2, "offset_utf16": 2 } : Json)),
      ("end", ({ "page": 2, "item_index": 8.5, "offset_utf16": 5 } : Json)),
      ("end", ({ "page": 2, "item_index": 8, "offset_utf16": -1 } : Json)),
      ("text_utf16_length", Json::number(9)),
      ("text_utf16_length", Json::boolean(true)),
    ] {
    let selection = pdf_fixture()
    pdf_object(pdf_object(selection)["reference"]).set(key, replacement)
    let result = pdf_source(selection, @viewer.Full)
    assert_true(
      result
      is {
        "source": { "readable": False, "reference_complete": False, .. },
        "selection": { "text": String("abcdefghij"), .. },
        ..
      },
    )
  }
}

///|
test "PDF identities require a valid hash byte count and parser" {
  for
    (key, replacement) in [
      ("sha256", Json::string("not-a-hash")),
      (
        "sha256",
        Json::string(
          "G123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
        ),
      ),
      ("byte_length", Json::number(0)),
      ("byte_length", Json::number(1.5)),
      ("page_count", Json::number(0)),
      ("page_count", Json::number(4.5)),
      ("parser", ({ "name": "pdflite", "version": "6.4.299" } : Json)),
      ("parser", ({ "name": "pdfjs", "version": "" } : Json)),
    ] {
    let selection = pdf_fixture()
    let fields = pdf_object(selection)
    let identity = pdf_object(fields["document"])
    identity.set(key, replacement)
    fields.set("document", Json::object(identity))
    pdf_object(fields["reference"]).set("document", Json::object(identity))
    assert_true(
      pdf_source(selection, @viewer.Deferred)
      is {
        "source": { "readable": False, "reference_complete": False, .. },
        ..
      },
    )
  }
}

///|
test "PDF sources do not claim an office.mbt selector" {
  let selection = pdf_fixture()
  let fields = pdf_object(selection)
  fields.set("selector", Json::string("/docx/body/p[1]"))
  guard fields["anchors"] is Array(anchors) else {
    fail("missing fixture anchors")
  }
  pdf_object(anchors[0]).set("selector", Json::string("/docx/body/p[2]"))
  let result = pdf_source(selection, @viewer.Deferred)
  guard result
    is {
      "source": {
        "selector_schema": Null,
        "selectors": Array(selectors),
        "anchors": Array(compact),
        ..
      },
      ..
    } else {
    fail(result.stringify())
  }
  assert_true(selectors.is_empty())
  assert_false(pdf_object(compact[0]).contains("selector"))
}

///|
test "PDF dispatch uses an asynchronous surface" {
  assert_true(@viewer.document_kind("DOCUMENT.PDF") == Some(@viewer.Pdf))
  assert_true(@viewer.pdf_surface_html().contains("pdf-container"))
  guard @viewer.render_document("document.pdf", "/document.pdf", b"%PDF-")
    is Err(message) else {
    fail("PDF must use the asynchronous loader")
  }
  assert_true(message.contains("load_pdf_command"))
}

///|
test "PDF uses its own runtime rather than Office annotation readers" {
  match
    @viewer.read_annotations("sample.pdf", "/docs/sample.pdf", b"%PDF-1.7") {
    Err(_) => ()
    Ok(_) => fail("PDF annotations are displayed by PDF.js")
  }
}
