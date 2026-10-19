///|
fn pptx_fixture() -> Bytes raise {
  let archive = @zip.Archive::Archive()
  archive.add(
    "[Content_Types].xml",
    @utf8.encode(
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"png\" ContentType=\"image/png\"/><Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/><Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/><Override PartName=\"/ppt/slides/slide2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/></Types>",
    ),
  )
  archive.add(
    "_rels/.rels",
    @utf8.encode(
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/></Relationships>",
    ),
  )
  archive.add(
    "ppt/presentation.xml",
    @utf8.encode(
      "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:sldIdLst><p:sldId id=\"256\" r:id=\"first\"/><p:sldId id=\"257\" r:id=\"second\"/></p:sldIdLst><p:sldSz cx=\"9144000\" cy=\"5143500\"/></p:presentation>",
    ),
  )
  archive.add(
    "ppt/_rels/presentation.xml.rels",
    @utf8.encode(
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"first\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide1.xml\"/><Relationship Id=\"second\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide2.xml\"/></Relationships>",
    ),
  )
  archive.add(
    "ppt/slides/slide1.xml",
    @utf8.encode(
      "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val=\"142F55\"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Text 2\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"571500\" y=\"571500\"/><a:ext cx=\"7874000\" cy=\"889000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr lIns=\"0\" rIns=\"0\" tIns=\"0\" bIns=\"0\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"3600\"><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill><a:latin typeface=\"Arial\"/></a:rPr><a:t>PPTX preview</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Text 3\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"571500\" y=\"1714500\"/><a:ext cx=\"7747000\" cy=\"762000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr lIns=\"0\" rIns=\"0\" tIns=\"0\" bIns=\"0\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"2000\"><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill><a:latin typeface=\"Arial\"/></a:rPr><a:t>Read-only slides. Selectable text.</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id=\"5\" name=\"Synthetic checkerboard\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"image\"/><a:srcRect l=\"10000\" r=\"10000\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"571500\" y=\"2857500\"/><a:ext cx=\"3429000\" cy=\"1206500\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"55B6A4\"/></a:solidFill></a:ln></p:spPr></p:pic></p:spTree></p:cSld></p:sld>",
    ),
  )
  archive.add(
    "ppt/slides/slide2.xml",
    @utf8.encode(
      "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val=\"F1F5F9\"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Text 2\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"571500\" y=\"571500\"/><a:ext cx=\"7874000\" cy=\"889000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr lIns=\"0\" rIns=\"0\" tIns=\"0\" bIns=\"0\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"3200\"><a:solidFill><a:srgbClr val=\"142F55\"/></a:solidFill><a:latin typeface=\"Arial\"/></a:rPr><a:t>Second slide</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Text 3\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"571500\" y=\"1651000\"/><a:ext cx=\"7747000\" cy=\"762000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr lIns=\"0\" rIns=\"0\" tIns=\"0\" bIns=\"0\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"2000\"><a:solidFill><a:srgbClr val=\"142F55\"/></a:solidFill><a:latin typeface=\"Arial\"/></a:rPr><a:t>Search, zoom and page navigation</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id=\"5\" name=\"Synthetic checkerboard\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"image\"/><a:srcRect l=\"10000\" r=\"10000\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=\"4635500\" y=\"2857500\"/><a:ext cx=\"3810000\" cy=\"1397000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"55B6A4\"/></a:solidFill></a:ln></p:spPr></p:pic></p:spTree></p:cSld></p:sld>",
    ),
  )
  archive.add(
    "ppt/slides/_rels/slide1.xml.rels",
    @utf8.encode(
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"image\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/checker.png\"/></Relationships>",
    ),
  )
  archive.add(
    "ppt/slides/_rels/slide2.xml.rels",
    @utf8.encode(
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"image\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/checker.png\"/></Relationships>",
    ),
  )
  archive.add(
    "ppt/media/checker.png", b"\x89\x50\x4e\x47\x0d\x0a\x1a\x0a\x00\x00\x00\x0d\x49\x48\x44\x52\x00\x00\x00\x50\x00\x00\x00\x28\x08\x02\x00\x00\x00\xe6\x02\xea\x56\x00\x00\x00\x6d\x49\x44\x41\x54\x78\x9c\xed\xcf\xb1\x09\x80\x50\x10\x44\xc1\x5f\x9b\x3d\x58\x98\x45\x58\x8e\xfd\x18\xfe\xd0\xc8\xcb\x0e\x6e\x61\x60\xd3\x85\x37\xeb\x7d\xce\x62\xc7\x7d\x15\x4b\xfc\xae\xc4\x68\x60\x60\x60\xe0\xa8\x68\x60\x60\xe0\x0f\x3c\x33\xab\xef\x0b\x9c\x10\x0d\x0c\x0c\x0c\x1c\x15\x0d\x0c\x5c\x80\x67\x66\xf5\x7d\x81\x13\xa2\x81\x81\x81\x81\xa3\xa2\x81\x81\x0b\xf0\xcc\xac\xbe\x2f\x70\x42\x34\x30\x30\x30\x70\x54\x34\x30\xf0\xff\x36\x7c\x4a\x51\xcb\x93\xab\x6b\xfa\x00\x00\x00\x00\x49\x45\x4e\x44\xae\x42\x60\x82",
  )
  @zip.write(archive)
}

///|
test "PowerPoint slides retain text, page boundaries and independent clips" {
  let original = pptx_fixture()
  let result = render_document("slides.PPTX", "/demo/slides.pptx", original)
  guard result is Ok(document) else {
    fail("could not render synthetic slides")
  }
  assert_true(document.kind is PowerPoint)
  assert_eq(document.item_count, 2)
  assert_true(document.surface_html.contains("PPTX preview"))
  assert_true(document.surface_html.contains("Second slide"))
  assert_true(document.surface_html.contains("data-page=\"2\""))
  assert_true(document.surface_html.contains("page1-c1"))
  assert_true(document.surface_html.contains("page2-c1"))
  assert_true(document.warning.contains("Arial"))
  assert_eq(original, pptx_fixture())
}
