# Modified fonts distributed with Office Viewer

These fonts remain licensed under the SIL Open Font License 1.1. The original
copyright statements, Reserved Font Name declarations and license texts are
preserved in the accompanying files and referenced in font metadata. Font
licensing is independent of the viewer application's license.

The pagelayout font programs omit naming metadata. Office Viewer adds naming,
copyright and license records and rebuilds the SFNT directory and checksums.
Glyph outlines, character maps and metric tables are unchanged.

The modified fonts use distinct internal family and PostScript names:

| Original program / renderer lookup | Modified internal family |
| --- | --- |
| Carlito | Office Viewer Sans |
| Liberation Sans | Office Viewer Sans Alt |
| Liberation Serif | Office Viewer Serif |
| Liberation Mono | Office Viewer Mono |
| Noto Sans SC | Office Viewer CJK |

Filenames and CSS lookup aliases retain the renderer's substitution keys for
compatibility with existing page layouts; they do not rename the internal font
families or identify the derivatives as original releases. The OFL FAQ explains
that font substitution and matching mechanisms are outside the naming clause:
https://openfontlicense.org/ofl-faq/ (section 5.2)

The complete copyright and license notices are in `Carlito-OFL.txt`,
`Liberation-LICENSE.txt`, and `NotoSansSC-OFL.txt`. Redistribution must preserve
the applicable copyright and OFL notices. Office Viewer's packaging includes
this modification record alongside those notices.
