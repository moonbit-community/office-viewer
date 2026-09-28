///|
// A source run that survived the reader/layout join. Offsets are UTF-16
// code-unit offsets in the paragraph projection, exactly as office.mbt's
// `find` and formatting APIs define them.
priv struct WordSourceRun {
  path : String
  start_utf16 : Int
  end_utf16 : Int
  text : String
}

///|
// The body paragraph address understood by office.mbt. `path` is the reader
// path (`p[3]`); `selector` is the canonical facade selector
// (`/docx/body/p[3]` or `/docx/body/p[id="..."]`).
priv struct WordSourceParagraph {
  path : String
  selector : String
  part : String
  text : String
  stability : String
  para_id : String?
  paragraph_anchor_status : String
  runs : Array[WordSourceRun]
  actionable : Bool
  reason : String
  joined : Bool
}

///|
priv struct WordSourceIndex {
  paragraphs : Array[WordSourceParagraph]
}

///|
// Metadata for one emitted SVG text fragment. This is deliberately separate
// from the source index: a paragraph can be split across lines and pages, and
// each visual fragment gets its own paragraph-relative range.
priv struct WordVisualSource {
  path : String
  selector : String
  part : String
  stability : String
  para_id : String?
  paragraph_anchor_status : String
  actionable : Bool
  reason : String
  start_utf16 : Int
  end_utf16 : Int
  runs : Array[String]
}

///|
fn source_xml_escape(value : String) -> String {
  let output = StringBuilder()
  for character in value {
    match character {
      '&' => output.write_string("&amp;")
      '<' => output.write_string("&lt;")
      '>' => output.write_string("&gt;")
      '"' => output.write_string("&quot;")
      '\'' => output.write_string("&apos;")
      _ => output.write_char(character)
    }
  }
  output.to_string()
}

///|
fn source_join_paths(paths : Array[String]) -> String {
  let output = StringBuilder()
  for index, path in paths {
    if index > 0 {
      output.write_char(',')
    }
    output.write_string(path)
  }
  output.to_string()
}

///|
// Keep UTF-16 slicing in one place. `substring` preserves a reader/layout
// offset even when an input document contains an unusual surrogate boundary.
#warnings("-deprecated")
fn source_slice(text : String, start : Int, end : Int) -> String {
  text.substring(start~, end~)
}

///|
fn source_selector(
  path : String,
  anchor : @reader.DocxParagraphAnchor?,
) -> (String, String, String?) {
  match anchor {
    Some(value) if value.status() == "unique" && value.para_id() is Some(id) =>
      ("/docx/body/p[id=\"\{id}\"]", "stable", Some(id))
    _ => ("/docx/body/" + path, "snapshot-relative", None)
  }
}

///|
fn direct_body_paragraphs(
  document : @document.DocumentElement,
) -> Array[(@document.DocumentElement, Int)] {
  let output : Array[(@document.DocumentElement, Int)] = []
  match document {
    @document.DocumentElement::Document(children~, ..) => {
      let mut ordinal = 0
      for child in children {
        match child {
          @document.DocumentElement::Paragraph(..) => {
            ordinal += 1
            output.push((child, ordinal))
          }
          _ => ()
        }
      }
    }
    _ => ()
  }
  output
}

///|
fn direct_run_entries(
  paragraph : @document.DocumentElement,
  paragraph_path : String,
) -> Array[WordSourceRun] {
  let output : Array[WordSourceRun] = []
  match paragraph {
    @document.DocumentElement::Paragraph(children~, ..) => {
      let mut ordinal = 0
      let mut offset = 0
      for child in children {
        match child {
          @document.DocumentElement::Run(..) => {
            ordinal += 1
            let path = "\{paragraph_path}/r[\{ordinal}]"
            let text = child.to_raw_text()
            let end = offset + text.length()
            output.push({ path, start_utf16: offset, end_utf16: end, text, })
            offset = end
          }
          // A hyperlink or another wrapper changes the canonical ancestry.
          // The first milestone intentionally declines that join until the
          // reader/layout mapping can prove the nested path as well.
          _ => ()
        }
      }
    }
    _ => ()
  }
  output
}

///|
fn find_full_paragraph_match(
  annotated : @reader.DocxAnnotatedResult,
  story : @reader.DocxStoryPartSource,
  path : String,
  text : String,
) -> @reader.DocxMatch? {
  if text == "" {
    return None
  }
  let matches = @reader.find_docx_matches(
    annotated,
    story,
    needle=text,
    within=path,
    limit=2,
  ) catch {
    _ => return None
  }
  for hit in matches.matches() {
    if hit.path() == path &&
      hit.start() == 0 &&
      hit.end() == text.length() &&
      hit.text() == text {
      return Some(hit)
    }
  }
  None
}

///|
fn source_paragraph(
  annotated : @reader.DocxAnnotatedResult,
  story : @reader.DocxStoryPartSource,
  paragraph : @document.DocumentElement,
  path : String,
  anchor_index : @reader.DocxParagraphAnchorIndex?,
) -> WordSourceParagraph {
  let projection = @reader.docx_paragraph_projection(annotated, story, path~)
  let text = match projection {
    Some(value) => value
    None => ""
  }
  let parsed_runs = direct_run_entries(paragraph, path)
  let mut assembled = ""
  for run in parsed_runs {
    assembled = assembled + run.text
  }
  let match_result = if projection is Some(_) && assembled == text {
    find_full_paragraph_match(annotated, story, path, text)
  } else {
    None
  }
  let anchor = match anchor_index {
    Some(index) => index.anchor_at(path)
    None => None
  }
  let (selector, stability, para_id) = source_selector(path, anchor)
  let joined = projection is Some(_) &&
    assembled == text &&
    (text == "" || match_result is Some(_))
  let actionable = match match_result {
    Some(hit) => hit.actionable()
    None => false
  }
  let reason = match match_result {
    Some(hit) => hit.reason()
    None => if joined { "" } else { "reader-layout-unjoined" }
  }
  let paragraph_status = match anchor {
    Some(value) => value.status()
    None => "missing"
  }
  let runs = if joined {
    match match_result {
      Some(hit) => {
        // The reader is authoritative about which runs carry the projection.
        // Keep only a one-to-one run-path agreement with the parsed tree;
        // otherwise a partial/flattened construct remains physical-only.
        let reader_runs = hit.runs()
        if reader_runs == parsed_runs.map(run => run.path) {
          parsed_runs
        } else {
          []
        }
      }
      None => []
    }
  } else {
    []
  }
  {
    path,
    selector,
    part: story.part(),
    text,
    stability,
    para_id,
    paragraph_anchor_status: paragraph_status,
    runs,
    actionable,
    reason,
    joined: joined && (text == "" || !runs.is_empty()),
  }
}

///|
// Build the source index with office.mbt's reader as the semantic authority.
// The parsed DocumentElement tree is used only to establish run boundaries;
// its concatenated text must equal the reader projection before a join is
// accepted.
fn build_word_source_index(bytes : Bytes) -> WordSourceIndex? {
  try {
    let annotated = @docx2html.read_docx_annotated(bytes)
    let story = annotated.main_story_source()
    let document = @docx2html.read_docx(bytes)
    let anchor_index = @reader.docx_paragraph_anchor_index(annotated, story)
    let paragraphs = direct_body_paragraphs(document)
    let count = @reader.docx_direct_body_paragraph_count(annotated)
    if paragraphs.length() != count {
      return None
    }
    let output : Array[WordSourceParagraph] = []
    for (paragraph, ordinal) in paragraphs {
      let path = "p[\{ordinal}]"
      output.push(
        source_paragraph(annotated, story, paragraph, path, Some(anchor_index)),
      )
    }
    Some({ paragraphs: output, })
  } catch {
    _ => None
  }
}

///|
fn overlapping_source_runs(
  paragraph : WordSourceParagraph,
  start : Int,
  end : Int,
) -> Array[String] {
  let output : Array[String] = []
  for run in paragraph.runs {
    if run.start_utf16 < end && run.end_utf16 > start {
      output.push(run.path)
    }
  }
  output
}

///|
fn source_fragment_attrs(source : WordVisualSource) -> String {
  let base = "data-source-id=\"docx\" " +
    "data-source-part=\"\{source_xml_escape(source.part)}\" " +
    "data-source-path=\"\{source_xml_escape(source.path)}\" " +
    "data-source-selector=\"\{source_xml_escape(source.selector)}\" " +
    "data-source-start=\"\{source.start_utf16}\" " +
    "data-source-end=\"\{source.end_utf16}\" " +
    "data-source-runs=\"\{source_xml_escape(source_join_paths(source.runs))}\" " +
    "data-source-stability=\"\{source.stability}\" " +
    "data-source-anchor-status=\"\{source.paragraph_anchor_status}\" " +
    "data-source-actionable=\"\{source.actionable}\" " +
    "data-source-reason=\"\{source_xml_escape(source.reason)}\""
  let suffix = match source.para_id {
    Some(id) => " data-source-para-id=\"\{source_xml_escape(id)}\""
    None => ""
  }
  base + suffix
}

///|
fn decorate_svg_text(
  svg : String,
  metadata : Array[WordVisualSource?],
) -> String {
  let output = StringBuilder()
  let mut cursor = 0
  let mut metadata_index = 0
  while cursor < svg.length() {
    let tail = source_slice(svg, cursor, svg.length())
    match tail.find("<text ") {
      None => {
        output.write_string(tail)
        cursor = svg.length()
      }
      Some(relative) => {
        let start = cursor + relative
        output.write_string(source_slice(svg, cursor, start + 6))
        match metadata.get(metadata_index) {
          Some(Some(source)) => {
            output.write_string(source_fragment_attrs(source))
            output.write_char(' ')
          }
          _ => ()
        }
        metadata_index += 1
        cursor = start + 6
      }
    }
  }
  output.to_string()
}

///|
fn decorate_word_pages(
  pages : Array[String],
  model : @pagelayout.PageModel,
  index : WordSourceIndex,
) -> Array[String] {
  let output : Array[String] = []
  let metadata_by_page : Array[Array[WordVisualSource?]] = []
  let mut paragraph_index = 0
  let mut paragraph_offset = 0
  let mut join_broken = false
  let skip_empty_paragraphs = () => {
    while paragraph_index < index.paragraphs.length() &&
          index.paragraphs[paragraph_index].text == "" {
      paragraph_index += 1
      paragraph_offset = 0
    }
  }
  skip_empty_paragraphs()
  for page in model.pages {
    let page_metadata : Array[WordVisualSource?] = []
    for item in page.items {
      match item {
        @pagelayout.PageItem::Text(run) if run.text.length() > 0 => {
          let metadata : WordVisualSource? = if join_broken {
            None
          } else {
            skip_empty_paragraphs()
            if paragraph_index >= index.paragraphs.length() {
              join_broken = true
              None
            } else {
              let paragraph = index.paragraphs[paragraph_index]
              let run_length = run.text.length()
              let end = paragraph_offset + run_length
              let expected = if end <= paragraph.text.length() {
                source_slice(paragraph.text, paragraph_offset, end)
              } else {
                ""
              }
              if !paragraph.joined ||
                end > paragraph.text.length() ||
                expected != run.text {
                join_broken = true
                None
              } else {
                let source = Some({
                  path: paragraph.path,
                  selector: paragraph.selector,
                  part: paragraph.part,
                  stability: paragraph.stability,
                  para_id: paragraph.para_id,
                  paragraph_anchor_status: paragraph.paragraph_anchor_status,
                  actionable: paragraph.actionable,
                  reason: paragraph.reason,
                  start_utf16: paragraph_offset,
                  end_utf16: end,
                  runs: overlapping_source_runs(
                    paragraph, paragraph_offset, end,
                  ),
                })
                paragraph_offset = end
                if paragraph_offset == paragraph.text.length() {
                  paragraph_index += 1
                  paragraph_offset = 0
                  skip_empty_paragraphs()
                }
                source
              }
            }
          }
          page_metadata.push(metadata)
        }
        _ => ()
      }
    }
    metadata_by_page.push(page_metadata)
  }
  // A complete join must consume every source paragraph and every visual text
  // fragment. Any trailing reader paragraph or hidden/layout-only text makes
  // the map unsafe, so discard all metadata for this render.
  skip_empty_paragraphs()
  if join_broken || paragraph_index != index.paragraphs.length() {
    for page in pages {
      output.push(decorate_svg_text(page, []))
    }
    return output
  }
  for page_index, page in pages {
    output.push(decorate_svg_text(page, metadata_by_page[page_index]))
  }
  output
}
