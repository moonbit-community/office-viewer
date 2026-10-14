///|
// The banner is part of the reusable surface so embedded hosts cannot
// accidentally hide incomplete previews by omitting the standalone inspector.
fn render_warning_html(warnings : Array[String]) -> String {
  if warnings.is_empty() {
    return ""
  }
  let counts : Map[String, Int32] = Map([])
  let ordered : Array[String] = []
  for warning in warnings {
    if !counts.contains(warning) {
      ordered.push(warning)
    }
    counts[warning] = counts.get(warning).unwrap_or(0) + 1
  }
  let output = StringBuilder()
  output.write_string(
    "<aside class=\"office-render-warning\" role=\"status\" style=\"padding:12px;margin:0 0 16px;border:1px solid #d4a535;background:#fff8df;color:#634d16;border-radius:6px;font:13px/1.5 sans-serif\"><strong>This preview has layout limitations</strong><ul>",
  )
  let entries : Array[Json] = []
  for index, warning in ordered {
    if index >= 64 {
      break
    }
    let count = counts[warning]
    output.write_string(
      "<li>" +
      @xlsx2html.escape_html(warning) +
      " (" +
      count.to_string() +
      ")</li>",
    )
    entries.push(
      Json::object({
        "feature": Json::string(warning),
        "count": json_number(count),
      }),
    )
  }
  if ordered.length() > 64 {
    output.write_string("<li>Additional limitations are present.</li>")
  }
  output.write_string("</ul></aside>")
  let diagnostics = Json::object({
      "schema": Json::string("office.render.diagnostics.v1"),
      "complete": Json::boolean(false),
      "items": Json::array(entries),
      "truncated": Json::boolean(ordered.length() > 64),
    })
    .stringify()
    .split("<")
    .join("\\u003c")
  output.write_string(
    "<script type=\"application/json\" class=\"office-render-diagnostics\">" +
    diagnostics +
    "</script>",
  )
  output.to_string()
}
