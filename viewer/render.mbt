///|
pub(all) enum DocumentKind {
  Word
  Excel
} derive(Eq)

///|
pub extend DocumentKind with Eq::{equal, not_equal}

///|
/// The renderer output consumed by both the standalone shell and OpenSeek.
/// `file` is owned by the host. The standalone demo uses a mock path; an
/// embedded host should pass the real path used by `office.mbt`.
pub(all) struct RenderedDocument {
  kind : DocumentKind
  name : String
  file : String
  surface_html : String
  item_count : Int
  warning : String
} derive(Eq)

///|
pub extend RenderedDocument with Eq::{equal, not_equal}

///|
pub fn document_kind(name : String) -> DocumentKind? {
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
/// Extract the body while preserving the renderer's generated style block.
/// `xlsx2html` returns a complete HTML document; a host surface only needs its
/// body fragment because the host owns the page shell and stylesheet.
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
/// Keep page boundaries as DOM elements so native browser text selection can
/// report the physical page containing each selected fragment.
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
/// Render one Office package with the same layout engines used by the
/// standalone demo. This function has no browser or OpenSeek state, so an
/// embedded JS frontend can call it after its native host supplies bytes.
pub fn render_document(
  name : String,
  file : String,
  bytes : Bytes,
) -> Result[RenderedDocument, String] {
  let kind = match document_kind(name) {
    Some(kind) => kind
    None => return Err("Only .docx and .xlsx files are supported.")
  }
  try {
    match kind {
      Word => {
        let model = @docx.render_docx(bytes)
        let rendered_pages = @svg.render_pages(model)
        let pages = match build_word_source_index(bytes) {
          Some(index) => decorate_word_pages(rendered_pages, model, index)
          None => rendered_pages
        }
        Ok({
          kind,
          name,
          file,
          surface_html: word_surface(pages),
          item_count: pages.length(),
          warning: "Word layout is read-only. Reader-backed paragraph selectors are emitted only when the source and PageModel text join completely; unsupported or ambiguous text remains physical-only.",
        })
      }
      Excel => {
        let workbook = @mbtexcel.read(bytes.exact_view())
        let rendered = @xlsx2html.render_workbook(
          workbook,
          title=name,
          max_rows=1000,
          max_cols=256,
          include_images=true,
          calc=false,
        )
        Ok({
          kind,
          name,
          file,
          surface_html: "<div class=\"sheet-document\">" +
          extract_html_body(rendered) +
          "</div>",
          item_count: workbook.get_sheet_list().length(),
          warning: "Excel cells are read-only. Cell addresses are derived from the rendered table coordinates and use the canonical office.xlsx selector shape.",
        })
      }
    }
  } catch {
    error => Err(error.to_string())
  }
}
