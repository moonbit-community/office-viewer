/// Shared component styles remain part of the reusable runtime.

///|
let annotation_styles : String =
  #| .office-viewer-surface .office-annotation-avatar { display: inline-flex; align-items: center; justify-content: center; flex: 0 0 28px; width: 28px; height: 28px; box-sizing: border-box; border-radius: 50%; background: var(--office-avatar-background, #e2e8f0); color: var(--office-avatar-color, #475569); font-family: system-ui, sans-serif; font-size: 11px; font-weight: 650; line-height: 1; user-select: none; overflow: hidden; vertical-align: middle; }
  #| .office-viewer-surface .office-context-person { display: flex; align-items: center; gap: 8px; min-width: 0; }
  #| .office-viewer-surface .office-context-person > .office-context-author { min-width: 0; }
  #| .office-viewer-surface .office-context-quote { display: -webkit-box; -webkit-box-orient: vertical; -webkit-line-clamp: 2; overflow: hidden; max-height: 40px; margin: 8px 0 0; padding: 0 0 0 8px; border-left: 2px solid #cbd5e1; color: #64748b; font-size: 12px; line-height: 1.6; }
  #| .office-viewer-surface .office-comment-source-marker { display: grid; place-items: center; }
  #| .office-viewer-surface .office-comment-source-marker::before { display: none; }
  #| .office-viewer-surface .office-comment-source-marker[hidden],
  #| .office-viewer-surface .office-context-rail[open] .office-comment-source-marker { display: none; }
  #| @media (max-width: 760px) {
  #|   .office-viewer-surface .office-comment-source-marker { display: none; }
  #| }
  #| @media (min-width: 761px) {
  #|   .office-viewer-surface .office-context-rail:not([open]) { flex-basis: 36px; width: 36px; }
  #|   .office-viewer-surface .office-context-rail:not([open]) > .office-context-heading { padding-left: 8px; padding-right: 8px; }
  #|   .office-viewer-surface .office-context-rail:not([open]) .office-context-heading-label { display: none; }
  #|   .office-viewer-surface .office-context-word { justify-content: center; gap: 12px; }
  #|   .office-viewer-surface .office-context-word > .office-context-document { flex: 0 1 auto; width: max-content; }
  #|   .office-viewer-surface .office-context-word > .office-context-rail[open] > .office-context-placed > .office-context-card { left: 0; right: 0; }
  #|   .office-viewer-surface .office-context-word > .office-context-rail[open] .office-context-unplaced-card,
  #|   .office-viewer-surface .office-context-word > .office-context-rail[open] .office-context-unplaced-heading,
  #|   .office-viewer-surface .office-context-word > .office-context-rail[open] .office-context-more { margin-left: 0; }
  #| }
