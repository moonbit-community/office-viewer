///|
/// A small browser-oriented chart renderer for the chart combinations that
/// xlsx2html deliberately leaves as placeholders. It consumes the same
/// reader-owned chart XML and cell ranges as office.mbt's inspect package, so
/// the viewer does not invent a second workbook model.
let viewer_chart_width : Int = 640

///|
let viewer_chart_height : Int = 360

///|
let viewer_plot_left : Int = 54

///|
let viewer_plot_top : Int = 30

///|
let viewer_plot_width : Int = 548

///|
let viewer_plot_height : Int = 250

///|
let viewer_chart_palette : Array[String] = [
  "#4472C4", "#ED7D31", "#70AD47", "#A5A5A5", "#FFC000", "#5B9BD5",
]

///|
fn viewer_parse_double_opt(text : String) -> Double? {
  let number = @string.parse_double(text) catch { _ => return None }
  Some(number)
}

///|
priv struct ViewerChartSeries {
  kind : String
  name : String
  values : Array[Double]
  present : Array[Bool]
  categories : Array[String]
}

///|
/// Find a plot block in document order. `inspect.chart_kinds` intentionally
/// returns a stable set order; the viewer needs the XML order to know which
/// series belongs to a bar, line, or area layer.
fn viewer_plot_block(xml : String, from : Int) -> (String, Int, Int)? {
  let candidates : Array[(Int, String)] = []
  for kind in ["barChart", "lineChart", "areaChart"] {
    let open = "<c:" + kind
    match xml[from:].find(open) {
      Some(offset) => {
        let start = from + offset
        let after = start + open.length()
        if after >= xml.length() ||
          xml[after] is (' ' | '\t' | '\r' | '\n' | '/' | '>') {
          candidates.push((start, kind))
        }
      }
      None => ()
    }
  }
  guard candidates.length() > 0 else { return None }
  let mut chosen = candidates[0]
  for candidate in candidates[1:] {
    if candidate.0 < chosen.0 {
      chosen = candidate
    }
  }
  let (start, kind) = chosen
  let body_start = match xml[start:].find(">") {
    Some(offset) => start + offset + 1
    None => return None
  }
  let close = "</c:" + kind + ">"
  match xml[body_start:].find(close) {
    Some(offset) => Some((kind, start, body_start + offset + close.length()))
    None => None
  }
}

///|
fn viewer_count_tag(block : String, tag : String) -> Int {
  let open = "<c:" + tag
  let mut from = 0
  let mut count = 0
  while from < block.length() {
    match block[from:].find(open) {
      Some(offset) => {
        let start = from + offset
        let after = start + open.length()
        if after < block.length() &&
          block[after] is (' ' | '\t' | '\r' | '\n' | '/' | '>') {
          count += 1
        }
        from = after
      }
      None => break
    }
  }
  count
}

///|
fn viewer_series_kinds(xml : String) -> Array[String] {
  let result : Array[String] = []
  let mut from = 0
  while from < xml.length() {
    match viewer_plot_block(xml, from) {
      Some((kind, _, end)) => {
        let block = xml[from:end].to_owned()
        let count = viewer_count_tag(block, "ser")
        for _ in 0..<count {
          result.push(kind)
        }
        from = end
      }
      None => break
    }
  }
  result
}

///|
fn viewer_chart_ref_parts(reference : String) -> (String, String)? {
  let mut quote = false
  let mut bang = None
  let mut index = 0
  while index < reference.length() {
    let character = reference[index]
    if character == '\'' {
      if quote && index + 1 < reference.length() && reference[index + 1] == '\'' {
        index += 2
        continue
      }
      quote = !quote
    } else if character == '!' && !quote {
      bang = Some(index)
      break
    }
    index += 1
  }
  let bang = match bang {
    Some(value) => value
    None => return None
  }
  let sheet_raw = reference[0:bang].to_owned()
  let range_raw = reference[bang + 1:reference.length()].to_owned()
  let sheet = if sheet_raw.length() >= 2 &&
    sheet_raw[0] == '\'' &&
    sheet_raw[sheet_raw.length() - 1] == '\'' {
    sheet_raw[1:sheet_raw.length() - 1]
    .to_owned()
    .replace_all(old="''", new="'")
  } else {
    sheet_raw
  }
  let range = range_raw.replace_all(old="$", new="")
  if sheet == "" || range == "" {
    None
  } else {
    Some((sheet, range))
  }
}

///|
fn viewer_range_cells(reference : String) -> (String, Array[String])? {
  guard viewer_chart_ref_parts(reference) is Some((sheet, range)) else {
    return None
  }
  let cells = @inspect.parse_range(range).cells() catch { _ => [] }
  Some((sheet, cells))
}

///|
fn viewer_cell_number(
  workbook : @xlsx.Workbook,
  sheet : String,
  cell : String,
) -> Double? raise {
  match workbook.get_cell_value_raw(sheet, cell) {
    Some(@xlsx.Numeric(value)) => Some(value)
    Some(@xlsx.String(text)) =>
      match viewer_parse_double_opt(text) {
        Some(value) if !value.is_nan() && !value.is_inf() => Some(value)
        _ => None
      }
    _ =>
      match workbook.calc_cell_typed(sheet, cell) {
        Some(@xlsx.Numeric(value)) if !value.is_nan() && !value.is_inf() =>
          Some(value)
        _ => None
      }
  }
}

///|
fn viewer_cell_label(
  workbook : @xlsx.Workbook,
  sheet : String,
  cell : String,
) -> String raise {
  match workbook.get_cell_value_raw(sheet, cell) {
    Some(@xlsx.String(value)) => value
    Some(@xlsx.Numeric(value)) => value.to_string()
    Some(@xlsx.Bool(value)) => if value { "TRUE" } else { "FALSE" }
    Some(@xlsx.Error(value)) => value
    _ => ""
  }
}

///|
fn viewer_series_name(
  workbook : @xlsx.Workbook,
  reference : String,
) -> String raise {
  match viewer_range_cells(reference) {
    Some((sheet, cells)) =>
      match cells.get(0) {
        Some(cell) => {
          let label = viewer_cell_label(workbook, sheet, cell)
          if label == "" {
            reference
          } else {
            label
          }
        }
        None => reference
      }
    None => reference
  }
}

///|
fn viewer_chart_series(
  workbook : @xlsx.Workbook,
  xml : String,
) -> Array[ViewerChartSeries] raise {
  let metas = @inspect.chart_series(xml)
  let kinds = viewer_series_kinds(xml)
  let series : Array[ViewerChartSeries] = []
  for index, meta in metas {
    guard meta.values is Some(values_reference) else { continue }
    guard viewer_range_cells(values_reference) is Some((sheet, cells)) else {
      continue
    }
    let values : Array[Double] = []
    let present : Array[Bool] = []
    for cell in cells {
      if values.length() >= 256 {
        break
      }
      match viewer_cell_number(workbook, sheet, cell) {
        Some(value) => {
          values.push(value)
          present.push(true)
        }
        None => {
          values.push(0.0)
          present.push(false)
        }
      }
    }
    let categories = match meta.categories {
      Some(categories_reference) =>
        match viewer_range_cells(categories_reference) {
          Some((category_sheet, category_cells)) =>
            category_cells.map(cell => {
              viewer_cell_label(workbook, category_sheet, cell)
            })
          None => []
        }
      None => []
    }
    let name = match meta.name {
      Some(reference) => viewer_series_name(workbook, reference)
      None => "series \{index + 1}"
    }
    let kind = if index < kinds.length() { kinds[index] } else { "lineChart" }
    if kind is ("barChart" | "lineChart" | "areaChart") && values.length() > 0 {
      series.push({ kind, name, values, present, categories, })
    }
  }
  series
}

///|
fn viewer_svg_num(value : Double) -> String {
  if value.is_nan() || value.is_inf() {
    "0"
  } else {
    value.to_string()
  }
}

///|
fn viewer_chart_point_count(series : Array[ViewerChartSeries]) -> Int {
  let mut count = 0
  for entry in series {
    if entry.values.length() > count {
      count = entry.values.length()
    }
  }
  count
}

///|
fn viewer_chart_scale_y(value : Double, low : Double, high : Double) -> Double {
  let span = high - low
  let ratio = if span == 0.0 { 0.5 } else { (value - low) / span }
  (viewer_plot_top + viewer_plot_height).to_double() -
  ratio * viewer_plot_height.to_double()
}

///|
fn viewer_chart_colour(entry : ViewerChartSeries, index : Int) -> String {
  match entry.kind {
    "lineChart" =>
      if index == 0 {
        "#C00000"
      } else {
        viewer_chart_palette[(index + 1) % viewer_chart_palette.length()]
      }
    "areaChart" => "#9DC3E6"
    _ => viewer_chart_palette[index % viewer_chart_palette.length()]
  }
}

///|
fn viewer_chart_area(
  svg : StringBuilder,
  entry : ViewerChartSeries,
  low : Double,
  high : Double,
  colour : String,
) -> Unit {
  let count = entry.values.length()
  guard count > 0 else { return }
  let step = viewer_plot_width.to_double() / count.to_double()
  let baseline = viewer_chart_scale_y(0.0, low, high)
  let points = StringBuilder()
  let mut first_x = 0.0
  let mut last_x = 0.0
  let mut plotted = 0
  for index, value in entry.values {
    if entry.present[index] == false {
      continue
    }
    let x = viewer_plot_left.to_double() + step * (index.to_double() + 0.5)
    let y = viewer_chart_scale_y(value, low, high)
    if plotted == 0 {
      first_x = x
    }
    last_x = x
    if plotted > 0 {
      points.write_string(" ")
    }
    points.write_string("\{viewer_svg_num(x)},\{viewer_svg_num(y)}")
    plotted += 1
  }
  if plotted > 1 {
    svg.write_string(
      "<path d=\"M \{viewer_svg_num(first_x)} \{viewer_svg_num(baseline)} L \{points.to_string()} L \{viewer_svg_num(last_x)} \{viewer_svg_num(baseline)} Z\" fill=\"\{colour}\" opacity=\"0.28\"/>",
    )
  }
}

///|
fn viewer_chart_bars(
  svg : StringBuilder,
  series : Array[ViewerChartSeries],
  low : Double,
  high : Double,
) -> Unit {
  let count = viewer_chart_point_count(series)
  let bar_series : Array[ViewerChartSeries] = []
  for entry in series {
    if entry.kind == "barChart" {
      bar_series.push(entry)
    }
  }
  guard count > 0 && bar_series.length() > 0 else { return }
  let group_width = viewer_plot_width.to_double() / count.to_double()
  let bar_width = group_width * 0.78 / bar_series.length().to_double()
  let baseline = viewer_chart_scale_y(0.0, low, high)
  for series_index, entry in bar_series {
    let colour = viewer_chart_colour(entry, series_index)
    for point_index, value in entry.values {
      if entry.present[point_index] == false {
        continue
      }
      let x = viewer_plot_left.to_double() +
        group_width * point_index.to_double() +
        group_width * 0.11 +
        bar_width * series_index.to_double()
      let y = viewer_chart_scale_y(value, low, high)
      let top = if y < baseline { y } else { baseline }
      let height = if y < baseline { baseline - y } else { y - baseline }
      svg.write_string(
        "<rect x=\"\{viewer_svg_num(x)}\" y=\"\{viewer_svg_num(top)}\" width=\"\{viewer_svg_num(bar_width)}\" height=\"\{viewer_svg_num(height)}\" fill=\"\{colour}\" rx=\"1\"/>",
      )
    }
  }
}

///|
fn viewer_chart_lines(
  svg : StringBuilder,
  series : Array[ViewerChartSeries],
  low : Double,
  high : Double,
) -> Unit {
  let count = viewer_chart_point_count(series)
  guard count > 0 else { return }
  let step = viewer_plot_width.to_double() / count.to_double()
  let mut line_index = 0
  for entry in series {
    if entry.kind != "lineChart" {
      continue
    }
    let colour = viewer_chart_colour(entry, line_index)
    line_index += 1
    let points = StringBuilder()
    let mut plotted = 0
    for point_index, value in entry.values {
      if entry.present[point_index] == false {
        continue
      }
      let x = viewer_plot_left.to_double() +
        step * (point_index.to_double() + 0.5)
      let y = viewer_chart_scale_y(value, low, high)
      if plotted > 0 {
        points.write_string(" ")
      }
      points.write_string("\{viewer_svg_num(x)},\{viewer_svg_num(y)}")
      plotted += 1
      svg.write_string(
        "<circle cx=\"\{viewer_svg_num(x)}\" cy=\"\{viewer_svg_num(y)}\" r=\"3\" fill=\"#fff\" stroke=\"\{colour}\" stroke-width=\"2\"/>",
      )
    }
    if plotted > 1 {
      svg.write_string(
        "<polyline fill=\"none\" stroke=\"\{colour}\" stroke-width=\"2.5\" points=\"\{points.to_string()}\"/>",
      )
    }
  }
}

///|
fn viewer_chart_axes(
  svg : StringBuilder,
  series : Array[ViewerChartSeries],
  primary_low : Double,
  primary_high : Double,
  secondary_low : Double,
  secondary_high : Double,
) -> Unit {
  let left = viewer_plot_left
  let top = viewer_plot_top
  let right = left + viewer_plot_width
  let bottom = top + viewer_plot_height
  svg.write_string(
    "<line class=\"axis\" x1=\"\{left}\" y1=\"\{top}\" x2=\"\{left}\" y2=\"\{bottom}\"/><line class=\"axis\" x1=\"\{left}\" y1=\"\{bottom}\" x2=\"\{right}\" y2=\"\{bottom}\"/>",
  )
  svg.write_string(
    "<text class=\"tick\" x=\"\{left - 7}\" y=\"\{top + 4}\" text-anchor=\"end\">\{viewer_svg_num(primary_high)}</text><text class=\"tick\" x=\"\{left - 7}\" y=\"\{bottom + 4}\" text-anchor=\"end\">\{viewer_svg_num(primary_low)}</text>",
  )
  let has_secondary = series.exists(entry => entry.kind == "lineChart")
  if has_secondary {
    svg.write_string(
      "<line class=\"axis secondary\" x1=\"\{right}\" y1=\"\{top}\" x2=\"\{right}\" y2=\"\{bottom}\"/><text class=\"tick\" x=\"\{right + 7}\" y=\"\{top + 4}\">\{viewer_svg_num(secondary_high)}</text><text class=\"tick\" x=\"\{right + 7}\" y=\"\{bottom + 4}\">\{viewer_svg_num(secondary_low)}</text>",
    )
  }
  let count = viewer_chart_point_count(series)
  let categories = match series.get(0) {
    Some(entry) => entry.categories
    None => []
  }
  if count > 0 && categories.length() > 0 {
    let step = viewer_plot_width.to_double() / count.to_double()
    let stride = if count <= 8 { 1 } else { (count + 7) / 8 }
    for index in 0..<count {
      if index % stride != 0 || index >= categories.length() {
        continue
      }
      let x = left.to_double() + step * (index.to_double() + 0.5)
      svg.write_string(
        "<text class=\"tick\" x=\"\{viewer_svg_num(x)}\" y=\"\{bottom + 18}\" text-anchor=\"middle\">\{@xlsx2html.escape_html(categories[index])}</text>",
      )
    }
  }
}

///|
fn viewer_chart_figure(
  workbook : @xlsx.Workbook,
  xml : String,
  anchor : String,
) -> String? raise {
  let kinds = @inspect.chart_kinds(xml)
  let has_viewer_kind = kinds.contains("areaChart") || kinds.length() > 1
  if !has_viewer_kind {
    return None
  }
  let series = viewer_chart_series(workbook, xml)
  guard series.length() > 0 else { return None }
  let mut primary_low = 0.0
  let mut primary_high = 0.0
  let mut secondary_low = 0.0
  let mut secondary_high = 0.0
  let mut has_primary = false
  let mut has_secondary = false
  for entry in series {
    for index, value in entry.values {
      if entry.present[index] {
        if entry.kind == "lineChart" {
          has_secondary = true
          if value < secondary_low {
            secondary_low = value
          }
          if value > secondary_high {
            secondary_high = value
          }
        } else {
          has_primary = true
          if value < primary_low {
            primary_low = value
          }
          if value > primary_high {
            primary_high = value
          }
        }
      }
    }
  }
  if !has_primary {
    primary_low = secondary_low
    primary_high = secondary_high
  }
  if !has_secondary {
    secondary_low = primary_low
    secondary_high = primary_high
  }
  if primary_low == primary_high {
    primary_high = primary_low + 1.0
  }
  if secondary_low == secondary_high {
    secondary_high = secondary_low + 1.0
  }
  let title = @inspect.chart_title(xml).unwrap_or("Chart")
  let description = title + " with " + series.length().to_string() + " series"
  let svg = StringBuilder()
  svg.write_string(
    "<svg class=\"viewer-combo-chart\" role=\"img\" aria-label=\"\{@xlsx2html.escape_html_attribute(description)}\" viewBox=\"0 0 \{viewer_chart_width} \{viewer_chart_height}\" width=\"\{viewer_chart_width}\" height=\"\{viewer_chart_height}\" xmlns=\"http://www.w3.org/2000/svg\"><title>\{@xlsx2html.escape_html(title)}</title>",
  )
  viewer_chart_axes(
    svg, series, primary_low, primary_high, secondary_low, secondary_high,
  )
  for entry in series {
    if entry.kind == "areaChart" {
      viewer_chart_area(svg, entry, primary_low, primary_high, "#9DC3E6")
    }
  }
  viewer_chart_bars(svg, series, primary_low, primary_high)
  viewer_chart_lines(svg, series, secondary_low, secondary_high)
  let legend_y = viewer_plot_top + viewer_plot_height + 43
  let mut bar_legend_index = 0
  let mut line_legend_index = 0
  let mut area_legend_index = 0
  for index, entry in series {
    let x = viewer_plot_left + index % 2 * 270
    let y = legend_y + index / 2 * 16
    let layer_index = match entry.kind {
      "barChart" => {
        let value = bar_legend_index
        bar_legend_index += 1
        value
      }
      "lineChart" => {
        let value = line_legend_index
        line_legend_index += 1
        value
      }
      _ => {
        let value = area_legend_index
        area_legend_index += 1
        value
      }
    }
    let colour = viewer_chart_colour(entry, layer_index)
    svg.write_string(
      "<rect x=\"\{x}\" y=\"\{y - 9}\" width=\"10\" height=\"10\" fill=\"\{colour}\"/><text class=\"legend\" x=\"\{x + 15}\" y=\"\{y}\">\{@xlsx2html.escape_html(entry.name)}</text>",
    )
  }
  svg.write_string("</svg>")
  Some(
    "<figure class=\"chart viewer-combo-chart-figure\" data-anchor=\"" +
    @xlsx2html.escape_html_attribute(anchor) +
    "\">" +
    svg.to_string() +
    "<figcaption>" +
    @xlsx2html.escape_html(description) +
    "</figcaption></figure>\n",
  )
}
