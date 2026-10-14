# Read-only annotations (`office.annotations.v1`)

The reusable viewer exposes annotations independently of its browser UI:

```mbt
match @viewer.read_annotations("review.docx", "/docs/review.docx", bytes) {
  Ok(envelope) => envelope.stringify(indent=2)
  Err(error) => error
}
```

`@viewer.render_document(name, file, bytes)` returns the same projection in
`RenderedDocument.annotations`. It reuses the parsed reader/workbook rather
than calling the standalone extractor again. Rendering also embeds the envelope
as inert, escaped JSON for the surface interactions. Hosts do not need to scrape
the inspector or DOM to obtain annotation data.

## Envelope and identity

```json
{
  "schema": "office.annotations.v1",
  "file": "/docs/review.xlsx",
  "document": {
    "name": "review.xlsx",
    "format": "xlsx",
    "identity": "host-file+sha256",
    "sha256": "<SHA-256 of the source bytes>"
  },
  "read_only": true,
  "items": [{
    "id": "xlsx.note:0:0",
    "document": {"file": "/docs/review.xlsx", "sha256": "<same SHA-256>"},
    "type": "xlsx.note",
    "source_id": "B3",
    "defined": true,
    "text": "Please check the total.",
    "text_status": "available",
    "author": "Ada",
    "author_id": 0,
    "initials": null,
    "date": null,
    "resolved": null,
    "parent_id": null,
    "parent_source_id": null,
    "thread_status": "not-applicable",
    "anchor": {
      "kind": "xlsx",
      "sheet": "Budget",
      "cell": "B3",
      "row": 3,
      "column": 2,
      "confidence": "exact-cell",
      "preview": "within-preview"
    }
  }],
  "support": {
    "scope": ["xlsx.note"],
    "completeness": "partial",
    "traditional_notes": "supported",
    "threaded_comments": "unsupported",
    "threaded_comments_presence": "unknown",
    "preview_max_rows": 1000,
    "preview_max_columns": 256
  },
  "warnings": ["Only traditional Excel Notes are read. Modern threaded comments are unsupported and their presence is not detected. An empty Notes list does not prove the workbook has no comments."]
}
```

`file` is the host-owned document identity already used by `office.selection.v1`.
The standalone app passes `/mock/office/<filename>`; that path is not a real
filesystem handle or a globally unique identity. Hosts must supply their own
stable file identity. `document.sha256` binds the projection to the exact input
bytes; every item carries the same file/hash identity. Focus rejects requests
from another file or revision even if their ordinal IDs coincide. Item IDs are
unique only inside this loaded snapshot, and must not be persisted as
revision-stable IDs.
`source_id` preserves the underlying comment ID or Note cell address.

`null` means absent/unknown (or not applicable for Notes), never `false`.
An empty, successfully parsed body is a string; an unavailable or ambiguously
identified Word body is `null` with `text_status: "missing" | "ambiguous"`.
Ambiguous bodies are preserved in `body_candidates`, not arbitrarily joined to
a duplicate ID. Plain text is exposed; rich comment formatting is not rendered.

## Word comments and w15 replies

- `type` is `docx.comment`. Author, initials, date, definition presence, raw
  source ID, and optional `resolved` (`w15:done`) are preserved.
- `parent_source_id` preserves the reader's parent ID; `parent_id` links to a
  viewer item only when that parent source ID is unique. No recorded parent
  is labeled `root-or-unknown`, rather than proving a comment is a thread root.
- Replies retain their own anchors. A parent's anchor is not silently copied
  onto an unanchored reply. Orphan, unresolved, and ambiguous items stay listed.
- `source_anchors` preserves each reader anchor's story, start/end paths,
  start/end boundary kind, and reference paths. Boundaries describe source
  nodes, not exact character offsets. In particular, `inside_end` must not be
  interpreted as the end of all text in a containing paragraph.
- `anchor.locations` contains only explicit main-story endpoint/reference
  paragraphs suitable for contextual navigation. It does not fill intervening
  paragraphs. `granularity` is `paragraph-context`; `ranges` is empty until an
  upstream API can prove exact offsets.
- The browser navigates only to DOM text with a matching source part/path from
  the existing all-or-nothing reader/layout join. If any layout/source join
  fails, the full comments list remains available without guessed highlights.
  A context highlight is never presented as the exact commented text range.
- Modern `commentsIds`, `people`, and `commentsExtensible` sidecars are detected
  by the reader and reported as unsupported. Reader diagnostics are retained.
  `complete-supported-kinds` only describes the supported comment projection;
  it is not a claim of full Office annotation or layout fidelity. A failed
  annotation parse during rendering yields `completeness: "unavailable"` and
  a warning instead of making an empty list look authoritative.

## Excel traditional Notes

All sheets and all parsed Notes are included, independently of the preview's
1000-row/256-column bounds. `outside-preview` indicates a note beyond those
bounds. `within-preview` does not guarantee an actual DOM cell: a Note-only
cell, a merged-away cell, or a non-rendered cell may still have no visible node.
Navigation requires the exact sheet and rendered cell. It switches sheets only
when that exact target exists, and returns false otherwise.

Modern Excel threaded comments, reply identity, and resolved state are not
supported or detected. An empty Notes list must not be described as “no
comments.” The envelope makes this limitation explicit even for an empty file.

## Browser integration

Call the existing `@viewer.interactions_command` after rendering the surface.
It marks visible cells and mapped Word paragraph contexts, using the same JSON.
Comment/reply bodies are rendered automatically in the document gutter on file
load; markers only aid selection and source navigation. No Comments-panel or
marker click is required to read or natively select the text. Multiple cards
are stacked without covering the document. Unmapped Word comments, Notes on
other sheets, and Notes outside the current view are automatically listed with
explicit source-location labels. Very large files initially show 200 bodies,
with an explicit count and a gutter control to show the next 200. On narrow
screens the labeled cards flow below the document viewport. The page SVG
and worksheet source content are unchanged. Word placement remains paragraph
context, not an exact commented range. Linked replies shown with a parent's
context are labeled as such; their original anchors are not modified.
Hosts can invoke:

```mbt
@viewer.focus_annotation_command(
  "document-surface",
  item,
  on_result=found => emit(AnnotationFocused(found)),
)
```

The callback reports whether a current rendered target was found. To select an
available body (also for off-preview or unanchored items), hosts can invoke
`@viewer.select_annotation_command(surface_id, item, on_result~)`. It emits the
normal interactions callback with `office.selection.v1`, `kind: "annotation"`.
Native highlighting within the contextual body emits the selected substring and
its exact UTF-16 range; the **Select comment** action emits the whole body. Both
commands require matching current file/hash/item identity. See
[`selection-model.md`](selection-model.md#annotation-body-selections) for the
contract and bounded preview/deferred behavior. The
standalone Comments panel keeps the list available on failure. It also exposes
the whole JSON and reports clipboard failures honestly. All comment content is
rendered as text; metadata escapes `<` before insertion in an inert data block.
There are no editing, replying, resolving, writing, or save-back operations.
Reinstalling interactions restores the canonical surface HTML before deriving
coordinates, so reopening the same file does not count generated grid headers
as source rows or duplicate wrappers/markers.

## Reproducible dependency build

The annotation APIs are tested with published `docx2html 0.7.0`,
`mbtexcel 0.3.0`, `pagelayout 0.8.0`, and Rabbita `0.16.3` from the
viewer manifest. Sibling checkouts are not required. Use MoonBit
`0.10.15+943464bfa`, run `moon update`, then `moon check --target js`,
`moon test --target js`, and `moon build --target js` in this repository.
See [the README](../README.md#run-locally) for browser dependencies and
[rendering support](rendering-support.md) for upstream fixes awaiting release.
