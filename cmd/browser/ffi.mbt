///|
// Browser-only helpers are kept behind a small boundary so the standalone
// shell remains separate from the reusable viewer surface.

///|
#cfg(target="js")
extern "js" fn selected_file_blob(input_id : String) -> @js.Nullable[@dom.Blob] =
  #| (id) => {
  #|   const input = document.getElementById(id);
  #|   return input && input.files && input.files.length > 0 ? input.files[0] : null;
  #| }

///|
#cfg(target="js")
extern "js" fn selected_file_name(input_id : String) -> String =
  #| (id) => {
  #|   const input = document.getElementById(id);
  #|   return input && input.files && input.files.length > 0 ? input.files[0].name : "";
  #| }

///|
#cfg(target="js")
extern "js" fn dropped_file_blob(
  event : @dom.DragEvent,
) -> @js.Nullable[@dom.Blob] =
  #| (event) => {
  #|   const files = event && event.dataTransfer && event.dataTransfer.files;
  #|   return files && files.length > 0 ? files[0] : null;
  #| }

///|
#cfg(target="js")
extern "js" fn dropped_file_name(event : @dom.DragEvent) -> String =
  #| (event) => {
  #|   const files = event && event.dataTransfer && event.dataTransfer.files;
  #|   return files && files.length > 0 ? files[0].name : "";
  #| }

///|
async fn read_selected_file(input_id : String) -> FilePayload? noraise {
  try {
    let blob = selected_file_blob(input_id).to_option()
    match blob {
      Some(blob) =>
        Some({ name: selected_file_name(input_id), bytes: blob.bytes(), })
      None => None
    }
  } catch {
    _ => None
  }
}

///|
async fn read_dropped_file(event : @dom.DragEvent) -> FilePayload? noraise {
  try {
    event.prevent_default()
    let blob = dropped_file_blob(event).to_option()
    match blob {
      Some(blob) =>
        Some({ name: dropped_file_name(event), bytes: blob.bytes(), })
      None => None
    }
  } catch {
    _ => None
  }
}

///|
#cfg(target="js")
extern "js" fn open_file_picker(input_id : String) =
  #| (id) => {
  #|   const input = document.getElementById(id);
  #|   if (input) input.click();
  #| }

///|
#cfg(target="js")
extern "js" fn copy_to_clipboard(text : String) =
  #| (text) => {
  #|   if (text && navigator.clipboard && navigator.clipboard.writeText) {
  #|     navigator.clipboard.writeText(text).catch(() => {});
  #|   }
  #| }
