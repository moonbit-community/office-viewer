///|
enum DocumentKind {
  Word
  Excel
  PowerPoint
  Pdf
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
  item_count : Int32
  warning : String
  annotations : Json
} derive(Eq)

///|
struct Model {
  kind : DocumentKind?
  file_name : String
  document_file : String
  surface_html : String
  item_count : Int32
  warning : String
  status : String
  zoom : Int32
  fit_width : Bool
  query : String
  selection : String
  inspector_open : Bool
  annotations : Json
  comments_open : Bool
  copying_annotations : Bool
  annotations_copy_request : Int32
  annotations_copy_status : String
  prefer_system_fonts : Bool
  loading : Bool
  generation : Int32
  selection_mode : @viewer.SelectionMaterializationMode
  resolved_selection : String
  resolving : Bool
  page : Int32
  rotation : Int32
  search_current : Int32
  search_total : Int32
  password_required : Bool
  password : String
  password_incorrect : Bool
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
    status: "No document loaded.",
    zoom: 100,
    fit_width: false,
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
    generation: 0,
    selection_mode: @viewer.Preview,
    resolved_selection: "",
    resolving: false,
    page: 1,
    rotation: 0,
    search_current: 0,
    search_total: 0,
    password_required: false,
    password: "",
    password_incorrect: false,
  }
}

///|
enum Msg {
  OpenRequested
  FileSelected
  FileReady(Int32, FilePayload?)
  DropRequested(@dom.DragEvent)
  DropReady(Int32, FilePayload?)
  ZoomChanged(String)
  WidthFitted(Int32, Int32)
  SearchChanged(String)
  SelectionChanged(String)
  PdfSelectionChanged(Int32, String)
  PdfStateChanged(Int32, String)
  SelectionModeChanged(@viewer.SelectionMaterializationMode)
  ResolveSelection
  SelectionResolved(Int32, String, Result[String, String])
  PageChanged(String)
  RotatePdf
  FitWidth
  FindPdf(Bool)
  PasswordChanged(String)
  SubmitPassword
  CancelPassword
  ToggleSystemFonts
  ToggleInspector
  ToggleComments
  FocusAnnotation(Json)
  SelectAnnotation(Json)
  AnnotationSelected(Bool)
  AnnotationFocused(Bool)
  CopyAnnotations
  AnnotationsCopied(Int32, Bool)
  CopySelection
  ClearSelection
}
