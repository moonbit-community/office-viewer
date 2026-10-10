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
  file : String
  surface_html : String
  item_count : Int
  warning : String
  annotations : Json
} derive(Eq)

///|
struct Model {
  kind : DocumentKind?
  file_name : String
  document_file : String
  surface_html : String
  item_count : Int
  warning : String
  status : String
  zoom : Int
  query : String
  selection : String
  inspector_open : Bool
  annotations : Json
  comments_open : Bool
  copying_annotations : Bool
  annotations_copy_request : Int
  annotations_copy_status : String
  prefer_system_fonts : Bool
  loading : Bool
} derive(Eq)

///|
fn initial_model() -> Model {
  {
    kind: None,
    file_name: "",
    document_file: "",
    surface_html: "",
    item_count: 0,
    warning: "",
    status: "Open a .docx or .xlsx file to begin.",
    zoom: 100,
    query: "",
    selection: "",
    inspector_open: false,
    annotations: Json::null(),
    comments_open: false,
    copying_annotations: false,
    annotations_copy_request: 0,
    annotations_copy_status: "",
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
  ToggleComments
  FocusAnnotation(Json)
  AnnotationFocused(Bool)
  CopyAnnotations
  AnnotationsCopied(Int, Bool)
  CopySelection
  ClearSelection
}
