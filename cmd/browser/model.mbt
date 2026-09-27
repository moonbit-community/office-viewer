///|
enum DocumentKind {
  Word
  Excel
} derive(Eq)

///|
struct FilePayload {
  name : String
  bytes : Bytes
}

///|
struct LoadedDocument {
  kind : DocumentKind
  name : String
  surface_html : String
  item_count : Int
  warning : String
} derive(Eq)

///|
struct Model {
  kind : DocumentKind?
  file_name : String
  surface_html : String
  item_count : Int
  warning : String
  status : String
  zoom : Int
  query : String
  selection : String
  inspector_open : Bool
  prefer_system_fonts : Bool
  loading : Bool
} derive(Eq)

///|
fn initial_model() -> Model {
  {
    kind: None,
    file_name: "",
    surface_html: "",
    item_count: 0,
    warning: "",
    status: "Open a .docx or .xlsx file to begin.",
    zoom: 100,
    query: "",
    selection: "",
    inspector_open: false,
    prefer_system_fonts: false,
    loading: false,
  }
}

///|
enum Msg {
  OpenRequested
  FileSelected
  FileReady(FilePayload?)
  DropRequested(@dom.DragEvent)
  DropReady(FilePayload?)
  ZoomChanged(String)
  SearchChanged(String)
  SelectionChanged(String)
  ToggleSystemFonts
  ToggleInspector
  CopySelection
  ClearSelection
}
