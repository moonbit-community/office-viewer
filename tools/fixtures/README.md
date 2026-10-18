`large-workbook.xlsx` is an original minimal OOXML fixture generated with Python's
standard-library `zipfile` module. It contains a sparse worksheet spanning A1 to
XFD1048576, a three-row/two-column merge at A1000, a cached formula at B1200,
text beyond the previous row/column limits, authored frozen panes, a Note beyond column 256, and a second sparse worksheet. It has no
third-party or personal content. The fixture verifies complete scroll reachability
without requiring millions of stored cells.
