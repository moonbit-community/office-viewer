///|
let word_selection_fixture = "{\"schema\":\"office.selection.v1\",\"file\":\"/mock/office/report.docx\",\"format\":\"docx\",\"kind\":\"word\",\"text\":\"0123456789ABCDEFGHIJ\",\"anchors\":[{\"source_id\":\"docx\",\"part\":\"word/document.xml\",\"path\":\"p[3]\",\"selector\":\"/docx/body/p[3]\",\"stability\":\"snapshot-relative\",\"start_utf16\":4,\"end_utf16\":24,\"page\":2,\"text\":\"0123456789ABCDEFGHIJ\",\"rect\":{\"left\":1,\"top\":2,\"width\":3,\"height\":4},\"runs\":[\"p[3]/r[2]\"],\"actionable\":true}],\"pages\":[2]}"

///|
let excel_selection_fixture = "{\"schema\":\"office.selection.v1\",\"file\":\"/mock/office/budget.xlsx\",\"format\":\"xlsx\",\"kind\":\"excel\",\"sheet\":\"Data\",\"range\":\"B4:C5\",\"selector\":\"/xlsx/sheet[name=\\\"Data\\\"]/range[B4:C5]\",\"stability\":\"snapshot-relative\",\"cells\":[{\"address\":\"B4\",\"row\":4,\"column\":2,\"text\":\"Revenue\"},{\"address\":\"C4\",\"row\":4,\"column\":3,\"text\":\"100\"},{\"address\":\"B5\",\"row\":5,\"column\":2,\"text\":\"Cost\"},{\"address\":\"C5\",\"row\":5,\"column\":3,\"text\":\"40\"}]}"

///|
let physical_only_selection_fixture = "{\"schema\":\"office.selection.v1\",\"file\":\"/mock/office/wrapped.docx\",\"format\":\"docx\",\"kind\":\"word\",\"text\":\"linked text\",\"anchors\":[{\"part\":\"word/document.xml\",\"path\":\"p[1]\",\"stability\":\"physical-only\",\"page\":1,\"start_utf16\":0,\"end_utf16\":11,\"text\":\"linked text\",\"rect\":{\"left\":1,\"top\":2,\"width\":3,\"height\":4}}],\"pages\":[1]}"

///|
test "full materialization keeps exact Word text and reader anchor" {
  let result = materialize_selection_full(word_selection_fixture)
  match result {
    Ok(value) => {
      assert_true(value.contains("office.selection.materialized.v1"))
      assert_true(value.contains("\"mode\": \"full\""))
      assert_true(value.contains("\"text\": \"0123456789ABCDEFGHIJ\""))
      assert_true(value.contains("/docx/body/p[3]"))
      assert_true(value.contains("\"text_utf16_length\": 20"))
      assert_true(value.contains("\"truncated\": false"))
    }
    Err(error) => fail(error)
  }
}

///|
test "preview materialization is bounded and reports omitted UTF-16 units" {
  let result = materialize_selection_preview(word_selection_fixture, 8)
  match result {
    Ok(value) => {
      assert_true(value.contains("\"mode\": \"preview\""))
      assert_true(value.contains("\"text\": \"0123[...]GHIJ\""))
      assert_true(value.contains("\"text_utf16_length\": 20"))
      assert_true(value.contains("\"included_utf16\": 8"))
      assert_true(value.contains("\"omitted_utf16\": 12"))
      assert_true(value.contains("\"representation\": \"head-tail\""))
    }
    Err(error) => fail(error)
  }
}

///|
test "deferred materialization omits text but keeps Excel selector and shape" {
  let result = materialize_selection_deferred(excel_selection_fixture)
  match result {
    Ok(value) => {
      assert_true(value.contains("\"mode\": \"deferred\""))
      assert_true(value.contains("\"text\": null"))
      assert_true(value.contains("/xlsx/sheet[name=\\\"Data\\\"]/range[B4:C5]"))
      assert_true(value.contains("\"cell_count\": 4"))
      assert_true(value.contains("\"text_utf16_length\": 19"))
      assert_true(value.contains("\"text_available\": false"))
      assert_true(value.contains("\"representation\": \"reference\""))
      assert_false(value.contains("Revenue"))
    }
    Err(error) => fail(error)
  }
}

///|
test "physical-only selections stay non-readable and keep physical source facts" {
  let result = materialize_selection_deferred(physical_only_selection_fixture)
  match result {
    Ok(value) => {
      assert_true(value.contains("\"readable\": false"))
      assert_true(value.contains("physical-only"))
      assert_true(value.contains("\"page\": 1"))
      assert_true(value.contains("no reader-friendly selector"))
      assert_false(value.contains("\"rect\""))
      assert_false(value.contains("linked text"))
    }
    Err(error) => fail(error)
  }
}

///|
test "selection materialization rejects malformed JSON and bad limits" {
  match materialize_selection_full("[]") {
    Err(error) => assert_true(error.contains("must be an object"))
    Ok(_) => fail("expected malformed selection error")
  }
  match materialize_selection_full("{") {
    Err(error) => assert_true(error.contains("invalid selection JSON"))
    Ok(_) => fail("expected invalid JSON error")
  }
  match materialize_selection_preview(word_selection_fixture, 0) {
    Err(error) => assert_true(error.contains("max_preview_utf16"))
    Ok(_) => fail("expected invalid preview limit error")
  }
}

///|
test "preview boundaries do not split an astral UTF-16 pair" {
  let json = "{\"schema\":\"office.selection.v1\",\"file\":\"/mock/office/u.docx\",\"format\":\"docx\",\"kind\":\"word\",\"text\":\"ab😀cd\",\"anchors\":[]}"
  let result = materialize_selection_preview(json, 3)
  match result {
    Ok(value) => {
      assert_true(value.contains("\"text\": \"a[...]cd\""))
      assert_true(value.contains("\"text_utf16_length\": 6"))
      assert_true(value.contains("\"included_utf16\": 3"))
    }
    Err(error) => fail(error)
  }
}

///|
fn selected_annotation_fixture() -> Json {
  Json::object({
    "id": Json::string("docx.comment:0"),
    "type": Json::string("docx.comment"),
    "source_id": Json::string("17"),
    "text": Json::string("Before 😀 selected after"),
    "text_status": Json::string("available"),
    "author": Json::string("Ada"),
    "resolved": Json::boolean(false),
    "parent_id": Json::null(),
    "thread_status": Json::string("root-or-unknown"),
    "body_candidates": Json::array([Json::string("never carry this body")]),
    "document": Json::object({
      "file": Json::string("/docs/comments.docx"),
      "sha256": Json::string(
        "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
      ),
    }),
    "anchor": Json::object({
      "kind": Json::string("docx"),
      "confidence": Json::string("source-node"),
      "granularity": Json::string("paragraph-context"),
      "locations": Json::array([
        Json::object({
          "part": Json::string("word/document.xml"),
          "path": Json::string("p[3]"),
        }),
      ]),
      "ranges": Json::array([]),
    }),
  })
}

///|
test "annotation selection distinguishes its body from paragraph context" {
  let selection = match
    annotation_selection(
      selected_annotation_fixture(),
      start_utf16=10,
      end_utf16=18,
      scope="text",
    ) {
    Ok(value) => value
    Err(error) => fail(error)
  }
  assert_eq(
    annotation_test_field(selection, "kind"),
    Json::string("annotation"),
  )
  assert_eq(annotation_test_field(selection, "text"), Json::string("selected"))
  assert_eq(annotation_test_field(selection, "anchors"), Json::array([]))
  let reference = annotation_test_field(selection, "annotation")
  assert_eq(annotation_test_field(reference, "resolved"), Json::boolean(false))
  assert_eq(annotation_test_field(reference, "parent_id"), Json::null())
  assert_eq(annotation_test_field(reference, "text"), Json::null())
  assert_false(selection.stringify().contains("Before"))
  assert_false(selection.stringify().contains("never carry this body"))
  let range = annotation_test_field(selection, "annotation_range")
  assert_eq(annotation_test_field(range, "start_utf16"), json_number(10))
  assert_eq(annotation_test_field(range, "end_utf16"), json_number(18))
  assert_eq(annotation_test_field(range, "body_utf16_length"), json_number(24))
}

///|
test "annotation preview and deferred keep identity without the full body" {
  let selection = match annotation_selection(selected_annotation_fixture()) {
    Ok(value) => value.stringify()
    Err(error) => fail(error)
  }
  let preview = match materialize_selection_preview(selection, 6) {
    Ok(value) => value
    Err(error) => fail(error)
  }
  assert_true(preview.contains("Bef[...]ter"))
  assert_false(preview.contains("selected after"))
  assert_true(preview.contains("docx.comment:0"))
  let deferred = match materialize_selection_deferred(selection) {
    Ok(value) => value
    Err(error) => fail(error)
  }
  assert_false(deferred.contains("Before"))
  assert_false(deferred.contains("never carry this body"))
  assert_true(deferred.contains("\"readable\": false"))
  assert_true(deferred.contains("office.annotations.v1"))
  assert_true(deferred.contains("paragraph-context"))
  assert_true(deferred.contains("sha256"))
  assert_true(deferred.contains("annotation_range"))
  assert_false(deferred.contains("office.selector/1\""))
}

///|
test "annotation selection rejects missing bodies and invalid UTF-16 spans" {
  for
    (start, end, scope) in [
      (8, 9, "text"),
      (-1, 2, "text"),
      (0, 30, "text"),
      (0, 2, "whole"),
    ] {
    match
      annotation_selection(
        selected_annotation_fixture(),
        start_utf16=start,
        end_utf16=end,
        scope~,
      ) {
      Err(_) => ()
      Ok(_) => fail("expected invalid annotation span")
    }
  }
  match selected_annotation_fixture() {
    Json::Object(fields) => {
      fields["text"] = Json::null()
      match annotation_selection(Json::object(fields)) {
        Err(_) => ()
        Ok(_) => fail("expected unavailable body")
      }
    }
    _ => fail("fixture must be an object")
  }
}

///|
test "off-preview Excel Note remains selectable with an exact cell reference" {
  let workbook = @mbtexcel.new_workbook()
  ignore(workbook.add_sheet("Data"))
  workbook.add_comment("Data", {
    cell: "A1001",
    author: "Ada",
    author_id: None,
    text: "Outside note",
    paragraph: [],
    width: None,
    height: None,
  })
  let envelope = bind_annotation_snapshot(
    excel_annotations("notes.xlsx", "/docs/notes.xlsx", workbook),
    b"snapshot",
  )
  let items = json_array(Some(annotation_test_field(envelope, "items")))
  let selection = match annotation_selection(items[0]) {
    Ok(value) => value
    Err(error) => fail(error)
  }
  assert_eq(annotation_test_field(selection, "format"), Json::string("xlsx"))
  assert_eq(
    annotation_test_field(selection, "text"),
    Json::string("Outside note"),
  )
  let anchor = annotation_test_field(
    annotation_test_field(selection, "annotation"),
    "anchor",
  )
  assert_eq(annotation_test_field(anchor, "cell"), Json::string("A1001"))
  assert_eq(
    annotation_test_field(anchor, "preview"),
    Json::string("outside-preview"),
  )
}
