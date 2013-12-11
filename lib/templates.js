var esc, utils, xml;

utils = require('./utils');

xml = utils.compress;

esc = utils.escapeXML;

module.exports = {
  worksheet: {
    header: xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \n xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" \n xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" \n mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">\n  <sheetViews>\n    <sheetView workbookViewId=\"0\"/>\n  </sheetViews>\n  <sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\"/>\n  <sheetData>"),
    footer: xml("  </sheetData>\n</worksheet>")
  },
  sheet_related: {
    "[Content_Types].xml": {
      header: xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n  <Default Extension=\"rels\" \n   ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n  <Default Extension=\"xml\" \n   ContentType=\"application/xml\"/>\n  <Override PartName=\"/xl/workbook.xml\" \n   ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n  <Override PartName=\"/xl/styles.xml\" \n   ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n  <Override PartName=\"/xl/sharedStrings.xml\" \n   ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"),
      sheet: function(sheet) {
        return "<Override PartName=\"/" + (esc(sheet.path)) + "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
      },
      footer: xml("</Types>")
    },
    "xl/_rels/workbook.xml.rels": {
      header: xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"),
      sheet: function(sheet) {
        return "<Relationship Id=\"rSheet" + (esc(sheet.index)) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"" + (esc(sheet.rel)) + "\"/>";
      },
      footer: xml("  <Relationship Id=\"rId2\" \n   Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" \n   Target=\"sharedStrings.xml\"/>\n  <Relationship Id=\"rId3\" \n   Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" \n   Target=\"styles.xml\"/>\n</Relationships>")
    },
    "xl/workbook.xml": {
      header: xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \n xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n  <fileVersion appName=\"xl\" lastEdited=\"5\" lowestEdited=\"5\" rupBuild=\"9303\"/>\n  <workbookPr defaultThemeVersion=\"124226\"/>\n  <bookViews>\n  <workbookView xWindow=\"480\" yWindow=\"60\" windowWidth=\"18195\" windowHeight=\"8505\"/>\n  </bookViews>\n  <sheets>"),
      sheet: function(sheet) {
        return xml("<sheet name=\"" + (esc(sheet.name)) + "\" sheetId=\"" + (esc(sheet.index)) + "\" r:id=\"rSheet" + (esc(sheet.index)) + "\"/>");
      },
      footer: xml("  </sheets>\n  <calcPr calcId=\"145621\"/>\n</workbook>")
    }
  },
  statics: {
    "_rels/.rels": xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n  <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  <Relationship Id=\"rId1\" \n   Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" \n   Target=\"xl/workbook.xml\"/>\n</Relationships>"),
    "xl/styles.xml": xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet \n xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \n xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" \n mc:Ignorable=\"x14ac\" \n xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">\n  <fonts count=\"1\" x14ac:knownFonts=\"1\">\n    <font>\n      <sz val=\"11\"/>\n      <color theme=\"1\"/>\n      <name val=\"Calibri\"/>\n      <family val=\"2\"/>\n      <scheme val=\"minor\"/>\n    </font>\n  </fonts>\n  <fills count=\"2\">\n    <fill>\n      <patternFill patternType=\"none\"/>\n    </fill>\n    <fill>\n      <patternFill patternType=\"gray125\"/>\n    </fill>\n  </fills>\n  <borders count=\"1\">\n    <border>\n      <left/>\n      <right/>\n      <top/>\n      <bottom/>\n      <diagonal/>\n    </border>\n  </borders>\n  <cellStyleXfs count=\"1\">\n    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n  </cellStyleXfs>\n  <cellXfs count=\"1\">\n    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" />\n  </cellXfs>\n  <cellStyles count=\"1\">\n    <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\n  </cellStyles>\n  <dxfs count=\"0\"/>\n  <tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>\n  <extLst>\n    <ext uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\" \n     xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">\n      <x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"/>\n    </ext>\n  </extLst>\n</styleSheet>"),
    "xl/sharedStrings.xml": xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"/>")
  }
};
