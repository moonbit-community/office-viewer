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
