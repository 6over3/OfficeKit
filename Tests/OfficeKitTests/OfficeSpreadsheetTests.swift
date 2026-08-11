import Foundation
import Testing
import UniformTypeIdentifiers

@testable import OfficeKit

@Test func strictSpreadsheetCanonicalizesWorkbookWorksheetAndRelationshipNamespaces() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/O14ISOStrict/Excel/filter_type.xlsx"
    )
  )
  let worksheet = try workbook.worksheet(at: 0)

  #expect(workbook.document.conformance == .strict)
  #expect(!workbook.sheets.isEmpty)
  #expect(!worksheet.rows.isEmpty)
  #expect(
    workbook.attachments.contains {
      $0.relationship.type.isEquivalent(to: .worksheet)
    })
}

@Test(arguments: [
  ("ChartSheet.xlsx", OfficeSheetKind.chart),
  ("Dialogsheet.xlsx", OfficeSheetKind.dialog),
])
func workbookPreservesNonGridSheetsWithoutParsingThemAsWorksheets(
  filename: String,
  expectedKind: OfficeSheetKind
) throws {
  let url = try FixtureCatalog.url(
    for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/spreadsheet/\(filename)"
  )
  let workbook = try OfficeWorkbook(contentsOf: url)

  #expect(workbook.sheets.contains { $0.kind == expectedKind })
  #expect(workbook.sheets.filter { $0.kind == .worksheet }.count == workbook.worksheets.count)
  let sheet = try #require(workbook.sheets.first { $0.kind == expectedKind })
  #expect(sheet.part != nil)
  #expect(try sheet.attachment.url().isFileURL)
}

@Test func worksheetCollectionLimitsDoNotRestrictStreamingRows() throws {
  let limits = OfficeParsingLimits(
    maximumRetainedWorksheetRows: 0,
    maximumRetainedWorksheetCells: 0
  )
  let url = try FixtureCatalog.url(
    for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/spreadsheet/smallset/SheetData.xlsx"
  )
  let workbook = try OfficeWorkbook(contentsOf: url, limits: limits)

  #expect(throws: OfficeKitError.self) {
    try workbook.worksheet(at: 0)
  }
  var streamedRows = 0
  try workbook.streamRows(inWorksheetAt: 0) { _ in streamedRows += 1 }
  #expect(streamedRows > 0)
}

@Test func spreadsheetAddressesRespectTheExcelGridBoundary() {
  #expect(OfficeCellReference(rawValue: "XFD1048576") != nil)
  #expect(OfficeCellReference(rawValue: "XFE1") == nil)
  #expect(OfficeCellReference(rawValue: "A1048577") == nil)
}

@Test func sharedStringRetentionHonorsTheConfiguredCountLimit() throws {
  let limits = OfficeParsingLimits(maximumSharedStringCount: 0)
  let url = try FixtureCatalog.url(
    for: "Open-XML-SDK/TestFiles/basicspreadsheet.xlsx"
  )
  let workbook = try OfficeWorkbook(contentsOf: url, limits: limits)

  #expect(throws: OfficeKitError.self) {
    try workbook.worksheet(at: 0)
  }
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/OpenXmlDomTest/DocumentTraverseTest.cs,
// TraverseSpreadSheetDocument.
@Test(arguments: [
  "SharedWorkbook.xlsx",
  "SheetData.xlsx",
  "SheetViewsFSB.xlsx",
])
func traverseSpreadsheetDocument(filename: String) throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/spreadsheet/smallset/" + filename
    )
  )
  #expect(workbook.worksheets.count == 3)

  for index in workbook.worksheets.indices {
    let worksheet = try workbook.worksheet(at: index)
    #expect(worksheet.reference.index == index)
    for row in worksheet.rows {
      for cell in row.cells {
        #expect(cell.reference.rowIndex == row.index)
      }
    }
  }
}

@Test func workbookDefinedNamesCalculationAndRelationshipsRemainInspectable() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/spreadsheet/smallset/SharedWorkbook.xlsx"
    )
  )
  let name = try #require(workbook.definedNames.first)

  #expect(workbook.properties.dateSystem == .nineteenHundred)
  #expect(name.name == "defindedName")
  #expect(name.formula == "Sheet1!$A$17")
  #expect(name.localSheetIndex == nil)
  #expect(!name.isHidden)
  #expect(workbook.calculation.calculationIdentifier == 124_519)
  #expect(
    workbook.attachments.contains {
      $0.relationship.type.isEquivalent(to: .worksheet)
    })
  #expect(
    try workbook.attachments.filter { $0.part != nil }.allSatisfy {
      try $0.url().isFileURL
    })
}

@Test func worksheetColumnsViewsPrintSettingsAndHeadersPreserveAuthoredMetadata() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/basicspreadsheet.xlsx")
  )
  let sheet = try workbook.worksheet(at: 0)
  let view = try #require(sheet.views.first)
  let selection = try #require(view.selections.first)
  let firstColumn = try #require(sheet.columns.first)
  let groupedColumns = try #require(sheet.columns.first { $0.firstIndex == 2 })

  #expect(sheet.columns.count == 7)
  #expect(firstColumn.firstIndex == 0)
  #expect(firstColumn.lastIndex == 0)
  #expect(firstColumn.width == 16)
  #expect(groupedColumns.lastIndex == 3)
  #expect(groupedColumns.width == 11.625)
  #expect(view.workbookViewIndex == 0)
  #expect(view.isTabSelected)
  #expect(view.topLeftCell == OfficeCellReference(rawValue: "A7"))
  #expect(selection.activeCell == OfficeCellReference(rawValue: "F38"))
  #expect(selection.ranges == ["F38:H43"])
  #expect(sheet.pageMargins?.left == 0.7)
  #expect(sheet.pageMargins?.header == 0.3)
  #expect(sheet.pageSetup?.orientation == "portrait")
  #expect(sheet.pageSetup?.relationshipID == OfficeRelationshipID(rawValue: "rId3"))
  #expect(sheet.headerFooter?.oddHeader?.contains("Themed Heading Font header") == true)
}

@Test func sharedStringRichTextPreservesRunBoundariesAndFormatting() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/basicspreadsheet.xlsx")
  )
  let sheet = try workbook.worksheet(at: 0)
  let reference = try #require(OfficeCellReference(rawValue: "A3"))
  let cell = try #require(sheet.cell(at: reference))
  let richText = try #require(cell.richText)

  #expect(cell.value == .string("Rich Text"))
  #expect(richText.text == "Rich Text")
  #expect(richText.runs.map(\.text) == ["Ri", "ch", " ", "Te", "xt"])
  #expect(richText.runs[0].properties == .none)
  #expect(richText.runs[1].properties.fontName == "宋体")
  #expect(richText.runs[1].properties.color?.theme == 4)
  #expect(richText.runs[2].properties.isItalic == true)
  #expect(richText.runs[2].properties.underline == "single")
  #expect(richText.runs[3].properties.isBold == true)
  #expect(richText.runs[4].properties.color?.theme == 1)
}

@Test func inlineStringRichTextPreservesRunBoundariesAndFormatting() throws {
  let url = try makeInlineRichTextWorkbook()
  defer { try? FileManager.default.removeItem(at: url) }
  let cell = try #require(OfficeWorkbook(contentsOf: url).worksheet(at: 0).rows.first?.cells.first)
  let richText = try #require(cell.richText)

  #expect(cell.value == .string("Plain Bold Red"))
  #expect(richText.text == "Plain Bold Red")
  #expect(richText.runs.map(\.text) == ["Plain ", "Bold", " Red"])
  #expect(richText.runs[0].properties == .none)
  #expect(richText.runs[1].properties.isBold == true)
  #expect(richText.runs[1].properties.fontName == "Aptos")
  #expect(richText.runs[2].properties.color?.argb == "FFFF0000")
}

@Test func conditionalFormattingPreservesRangesRulesFormulasAndVisuals() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/basicspreadsheet.xlsx")
  )
  let formatting = try workbook.worksheet(at: 0).conditionalFormatting

  #expect(formatting.count == 5)
  #expect(
    formatting.map { $0.ranges.map(\.rawValue) } == [
      ["G1:G3"], ["H1:H3"], ["I1:I3"], ["J1:J3"], ["K1:K3"],
    ])
  #expect(
    formatting.flatMap(\.rules).map(\.type)
      == ["dataBar", "colorScale", "cellIs", "cellIs", "top10"])
  #expect(formatting[0].rules[0].formulas.first?.hasPrefix("MAX(IF(ISBLANK") == true)
  #expect(formatting[0].rules[0].dataBar?.thresholds.map(\.kind) == ["min", "max"])
  #expect(formatting[0].rules[0].dataBar?.color?.theme == 4)
  #expect(formatting[1].rules[0].colorScale?.colors.map(\.theme) == [4, 6])
  #expect(formatting[2].rules[0].comparisonOperator == "greaterThan")
  #expect(formatting[2].rules[0].differentialStyleIdentifier == 2)
  #expect(formatting[3].rules[0].formulas == ["2", "3"])
  #expect(formatting[4].rules[0].rank == 10)
}

@Test func worksheetObjectsRetainExactAnchorsAndURLBackedInertPayloads() throws {
  let url = try makeInlineRichTextWorkbook()
  defer { try? FileManager.default.removeItem(at: url) }
  let workbook = try OfficeWorkbook(contentsOf: url)
  let objects = try workbook.worksheet(at: 0).objects
  let ole = try #require(objects.first { $0.kind == .ole })
  let control = try #require(objects.first { $0.kind == .control })

  #expect(objects.count == 2)
  #expect(ole.shapeIdentifier == 1025)
  #expect(ole.programIdentifier == "PowerPoint.Show.12")
  #expect(ole.usesDefaultSize == false)
  #expect(ole.movesWithCells == true)
  #expect(ole.sizesWithCells == false)
  guard case .twoCell(let from, let to, _) = ole.anchor else {
    Issue.record("Expected an exact two-cell OLE anchor.")
    return
  }
  #expect(from.columnIndex == 1)
  #expect(from.columnOffset.emu == 9_525)
  #expect(to.rowIndex == 4)
  #expect(to.rowOffset.emu == 19_050)
  #expect(ole.previewImage?.part?.name.rawValue == "/xl/media/preview.png")
  #expect(try ole.attachments.allSatisfy { try $0.url().isFileURL })
  #expect(control.name == "CommandButton1")
  #expect(control.attachment.part?.name.rawValue == "/xl/activeX/control.xml")
  #expect(
    control.attachments.contains {
      $0.part?.name.rawValue == "/xl/activeX/control.bin"
    })
  #expect(try control.attachments.allSatisfy { try $0.url().isFileURL })
  let controlPart = try #require(control.attachment.part)
  let cycle = try #require(
    try workbook.document.package.relationships(
      from: .part(controlPart.name)
    ).first { $0.id == OfficeRelationshipID(rawValue: "rIdCycle") })
  #expect(cycle.target == .internalPart(workbook.worksheets[0].part.name, fragment: nil))
}

private struct RowStreamingStopped: Error, Equatable {}

@Test func worksheetRowsStreamWithoutRetentionAndPropagateConsumerErrors() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/spreadsheet/smallset/SheetData.xlsx"
    )
  )
  let retainedRows = try workbook.worksheet(at: 0).rows
  var streamedRows: [OfficeWorksheetRow] = []
  try workbook.streamRows(inWorksheetAt: 0) { streamedRows.append($0) }

  #expect(streamedRows == retainedRows)
  #expect(streamedRows.count == 109)

  var deliveredRowCount = 0
  #expect(throws: RowStreamingStopped.self) {
    try workbook.streamRows(inWorksheetAt: 0) { _ in
      deliveredRowCount += 1
      if deliveredRowCount == 3 { throw RowStreamingStopped() }
    }
  }
  #expect(deliveredRowCount == 3)
}

@Test func cellStylesResolveFontsFillsBordersAlignmentAndColors() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Complex01.xlsx")
  )
  let sheet = try workbook.worksheet(at: 0)
  let styledCells = sheet.rows.flatMap(\.cells).filter { $0.styleIndex != nil }
  let redBold = try #require(styledCells.first { $0.styleIndex == 1 }?.style)
  let wrapped = try #require(styledCells.first { $0.styleIndex == 2 }?.style)
  let hyperlink = try #require(styledCells.first { $0.styleIndex == 3 }?.style)

  #expect(redBold.fontIndex == 1)
  #expect(redBold.font.name == "Calibri")
  #expect(redBold.font.sizeInPoints == 11)
  #expect(redBold.font.isBold)
  #expect(redBold.font.underline == "single")
  #expect(redBold.font.color?.argb == "FFFF0000")
  #expect(redBold.fill.patternType == "none")
  #expect(redBold.border.leading.style == nil)
  #expect(wrapped.alignment?.wrapsText == true)
  #expect(hyperlink.font.underline == "single")
  #expect(hyperlink.font.color?.theme == 10)
}

@Test func worksheetResolvesSparseSharedStringsFormulasAndCachedValues() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/spreadsheet/smallset/SheetData.xlsx"
    )
  )
  let sheet = try workbook.worksheet(at: 0)
  let b1Reference = try #require(OfficeCellReference(rawValue: "B1"))
  let a2Reference = try #require(OfficeCellReference(rawValue: "A2"))
  let b2Reference = try #require(OfficeCellReference(rawValue: "B2"))
  let b3Reference = try #require(OfficeCellReference(rawValue: "B3"))
  let b1 = try #require(sheet.cell(at: b1Reference))
  let a2 = try #require(sheet.cell(at: a2Reference))
  let b2 = try #require(sheet.cell(at: b2Reference))
  let b3 = try #require(sheet.cell(at: b3Reference))

  #expect(workbook.worksheets.map(\.name) == ["Sheet1", "Sheet2", "Sheet3"])
  #expect(workbook.worksheets.map(\.identifier) == [1, 2, 3])
  #expect(workbook.worksheets.allSatisfy { $0.visibility == .visible })
  #expect(sheet.dimensionReference == "A1:H109")
  #expect(sheet.rows.count == 109)
  #expect(sheet.rows[0].cells.map(\.reference.rawValue) == ["B1"])
  #expect(b1.rawValue == "0")
  #expect(b1.value == .string("Values"))
  #expect(a2.value == .string("Row Labels"))
  #expect(b2.value == .string("Profit"))
  #expect(b2.formula?.text == "CUBEMEMBER(\"FoodMart 2000 Sales\",\"[Measures].[Profit]\")")
  #expect(b3.rawValue == "339610.89640000137")
  #expect(b3.value == .number(339_610.896_400_001_37))
  #expect(b3.formula?.text == "CUBEVALUE(\"FoodMart 2000 Sales\",$A3,B$2)")
}

@Test func numericDateStyleResolvesWorkbookEpochAndCachedFormulaDate() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Spreadsheet.xlsx")
  )
  let sheet = try workbook.worksheet(at: 1)
  let reference = try #require(OfficeCellReference(rawValue: "B10"))
  let cell = try #require(sheet.cell(at: reference))
  let date = try #require(cell.dateValue)

  #expect(workbook.dateSystem == .nineteenHundred)
  #expect(cell.rawValue == "42189")
  #expect(cell.value == .number(42_189))
  #expect(cell.formula?.text == "TODAY()")
  #expect(cell.styleIndex == 4)
  #expect(cell.style?.numberFormat.identifier == 14)
  #expect(cell.style?.numberFormat.code == "m/d/yy")
  #expect(cell.style?.numberFormat.isDate == true)
  #expect(date.timeIntervalSince1970 == 1_435_968_000)
  #expect(OfficeSpreadsheetDateSystem.nineteenHundred.date(fromSerial: 60) == nil)
  #expect(
    OfficeSpreadsheetDateSystem.nineteenOhFour.date(fromSerial: 0)?.timeIntervalSince1970
      == -2_082_844_800
  )
}

@Test func worksheetRetainsMergedCellRanges() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/basicspreadsheet.xlsx")
  )
  let sheet = try workbook.worksheet(at: 0)
  let range = try #require(sheet.mergedRanges.first)

  #expect(sheet.mergedRanges.count == 1)
  #expect(range.rawValue == "G6:K6")
  #expect(range.start.columnIndex == 6)
  #expect(range.start.rowIndex == 5)
  #expect(range.end.columnIndex == 10)
  #expect(range.end.rowIndex == 5)
}

// Upstream fixture generated by DocumentFormat.OpenXml.Tests/GeneratedClass002.cs.
@Test func worksheetHyperlinkResolvesExternalURLWithoutFetching() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Complex01.xlsx")
  )
  let sheet = try workbook.worksheet(at: 0)
  let hyperlink = try #require(sheet.hyperlinks.first)
  let attachment = try #require(hyperlink.attachment)
  let q5 = try #require(OfficeCellReference(rawValue: "Q5"))

  #expect(sheet.hyperlinks.count == 1)
  #expect(hyperlink.reference == OfficeCellRange(rawValue: "Q5"))
  #expect(sheet.cell(at: q5)?.value == .string("www.ericwhite.com"))
  #expect(attachment.relationship.type.isEquivalent(to: .hyperlink))
  #expect(attachment.part == nil)
  #expect(try attachment.url().absoluteString == "http://www.ericwhite.com/")
}

@Test func worksheetDrawingRetainsTwoCellAnchorsChartsAndImageURLs() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Spreadsheet.xlsx")
  )
  let sheet = try workbook.worksheet(at: 0)
  let drawing = try #require(sheet.drawings.first)
  let chartElement = drawing.elements[0]
  let pictureElement = drawing.elements[1]
  guard case .twoCell(let chartFrom, let chartTo, let chartEditBehavior) = chartElement.anchor else {
    Issue.record("Expected a two-cell chart anchor.")
    return
  }
  guard case .chart(let chartReference) = chartElement.graphicContent else {
    Issue.record("Expected a typed chart reference.")
    return
  }
  guard
    case .twoCell(let pictureFrom, let pictureTo, let pictureEditBehavior) =
      pictureElement.anchor else {
    Issue.record("Expected a two-cell picture anchor.")
    return
  }
  let picture = try #require(pictureElement.picture)
  let image = try #require(picture.primaryImage)

  #expect(drawing.part.name.rawValue == "/xl/drawings/drawing1.xml")
  #expect(drawing.elements.count == 2)
  #expect(drawing.elements.map(\.spatialInfo.zIndex) == [0, 1])
  #expect(chartElement.identifier == 2)
  #expect(chartElement.name == "Chart 1")
  #expect(chartFrom.columnIndex == 5)
  #expect(chartFrom.columnOffset.emu == 552_450)
  #expect(chartFrom.rowIndex == 9)
  #expect(chartFrom.rowOffset.emu == 33_337)
  #expect(chartTo.columnIndex == 13)
  #expect(chartTo.columnOffset.emu == 247_650)
  #expect(chartTo.rowIndex == 23)
  #expect(chartTo.rowOffset.emu == 109_537)
  #expect(chartEditBehavior == nil)
  #expect(chartElement.spatialInfo.frame == nil)
  #expect(chartReference.part.name.rawValue == "/xl/charts/chart1.xml")
  #expect(try chartReference.attachment.url().isFileURL)

  #expect(pictureElement.identifier == 3)
  #expect(pictureElement.name == "Picture 2")
  #expect(pictureFrom.columnIndex == 15)
  #expect(pictureFrom.rowIndex == 2)
  #expect(pictureTo.columnIndex == 19)
  #expect(pictureTo.rowIndex == 13)
  #expect(pictureEditBehavior == "oneCell")
  #expect(pictureElement.spatialInfo.sourceTransform?.x.emu == 9_182_100)
  #expect(pictureElement.spatialInfo.sourceTransform?.y.emu == 552_450)
  #expect(pictureElement.spatialInfo.sourceTransform?.width.emu == 2_847_975)
  #expect(pictureElement.spatialInfo.sourceTransform?.height.emu == 2_132_845)
  #expect(pictureElement.spatialInfo.resolution == .exact)
  #expect(image.part?.name.rawValue == "/xl/media/image1.jpg")
  #expect(image.contentType == .jpeg)
  #expect(try image.url().isFileURL)
  #expect(sheet.attachments.contains { $0.part?.name.rawValue == "/xl/media/image1.jpg" })
}

@Test func worksheetDrawingRetainsOneCellExtentAndMultipleImageSources() throws {
  let basicWorkbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/basicspreadsheet.xlsx")
  )
  let basicSheet = try basicWorkbook.worksheet(at: 0)
  let chart = try #require(basicSheet.drawings.first?.elements.first)
  guard case .oneCell(let origin, let width, let height) = chart.anchor else {
    Issue.record("Expected a one-cell chart anchor.")
    return
  }
  #expect(origin.columnIndex == 0)
  #expect(origin.columnOffset.emu == 38_100)
  #expect(origin.rowIndex == 12)
  #expect(origin.rowOffset.emu == 38_100)
  #expect(width.emu == 4_572_000)
  #expect(height.emu == 2_743_200)
  #expect(chart.spatialInfo.frame == nil)

  let complexWorkbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Complex01.xlsx")
  )
  let complexSheet = try complexWorkbook.worksheet(at: 1)
  let pictureElement = try #require(
    complexSheet.drawings.first?.elements.first {
      $0.kind == .picture
    })
  let picture = try #require(pictureElement.picture)

  #expect(pictureElement.alternativeText == "Screen Clipping")
  #expect(picture.images.count == 2)
  #expect(picture.primaryImage?.part?.name.rawValue == "/xl/media/image1.png")
  #expect(
    Set(picture.images.compactMap { $0.part?.name.rawValue })
      == Set([
        "/xl/media/image1.png",
        "/xl/media/hdphoto1.wdp",
      ]))
  #expect(try picture.images.allSatisfy { try $0.url().isFileURL })
}

@Test func worksheetTableReferenceLazilyParsesRangeColumnsAndStyle() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Spreadsheet.xlsx")
  )
  let sheet = try workbook.worksheet(at: 0)
  let reference = try #require(sheet.tables.first)

  #expect(sheet.tables.count == 1)
  #expect(reference.part.name.rawValue == "/xl/tables/table1.xml")
  #expect(try reference.attachment.url().isFileURL)

  let table = try reference.table()
  #expect(table.sourcePart == reference.part)
  #expect(table.identifier == 1)
  #expect(table.name == "Funny")
  #expect(table.displayName == "Funny")
  #expect(table.range == OfficeCellRange(rawValue: "A1:C3"))
  #expect(table.headerRowCount == 1)
  #expect(table.totalsRowCount == 0)
  #expect(table.columns.map(\.identifier) == [1, 2, 3])
  #expect(table.columns.map(\.name) == ["a", "b", "c"])
  #expect(table.style?.name == "TableStyleMedium2")
  #expect(table.style?.showsFirstColumn == false)
  #expect(table.style?.showsLastColumn == false)
  #expect(table.style?.showsRowStripes == true)
  #expect(table.style?.showsColumnStripes == false)
}

@Test func worksheetCommentsResolveAuthorsRichTextAndSourceParts() throws {
  let workbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Spreadsheet.xlsx")
  )
  let sheet = try workbook.worksheet(at: 1)
  let comment = try #require(sheet.comments.first)
  let shape = try #require(comment.shape)
  let anchor = try #require(shape.anchor)
  let vml = try #require(
    sheet.attachments.first {
      $0.relationship.type.isEquivalent(to: .vmlDrawing)
    })

  #expect(sheet.comments.count == 1)
  #expect(comment.reference == OfficeCellReference(rawValue: "B3"))
  #expect(comment.authorIndex == 0)
  #expect(comment.author == "Author")
  #expect(comment.shapeIdentifier == 0)
  #expect(comment.text == "Author:\nAdd a comment")
  #expect(comment.sourcePart.name.rawValue == "/xl/comments1.xml")
  #expect(shape.identifier == "_x0000_s2049")
  #expect(shape.isVisible == false)
  #expect(shape.movesWithCells)
  #expect(shape.sizesWithCells)
  #expect(shape.zIndex == 1)
  #expect(anchor.from.columnIndex == 2)
  #expect(anchor.from.columnOffset == 15)
  #expect(anchor.from.rowIndex == 1)
  #expect(anchor.from.rowOffset == 10)
  #expect(anchor.to.columnIndex == 4)
  #expect(anchor.to.columnOffset == 31)
  #expect(anchor.to.rowIndex == 5)
  #expect(anchor.to.rowOffset == 9)
  #expect(shape.spatialInfo.frame == OfficeRect(x: 107.25, y: 22.5, width: 108, height: 59.25))
  #expect(shape.spatialInfo.resolution == .exact)
  #expect(shape.sourcePart.name.rawValue == "/xl/drawings/vmlDrawing1.vml")
  #expect(vml.part?.name.rawValue == "/xl/drawings/vmlDrawing1.vml")
  #expect(try vml.url().isFileURL)

  let complexWorkbook = try OfficeWorkbook(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Complex01.xlsx")
  )
  let complexSheet = try complexWorkbook.worksheet(at: 0)
  let complexComment = try #require(complexSheet.comments.first)
  let complexShape = try #require(complexComment.shape)
  let complexAnchor = try #require(complexShape.anchor)
  #expect(complexComment.reference == OfficeCellReference(rawValue: "V10"))
  #expect(complexComment.text == "Author:\nThis is a comment")
  #expect(complexShape.identifier == "_x0000_s1025")
  #expect(complexShape.isVisible)
  #expect(complexAnchor.from.columnIndex == 22)
  #expect(complexAnchor.from.columnOffset == 15)
  #expect(complexAnchor.from.rowIndex == 8)
  #expect(complexAnchor.from.rowOffset == 10)
  #expect(complexAnchor.to.columnIndex == 24)
  #expect(complexAnchor.to.columnOffset == 31)
  #expect(complexAnchor.to.rowIndex == 12)
  #expect(complexAnchor.to.rowOffset == 9)
  #expect(
    complexShape.spatialInfo.frame
      == OfficeRect(x: 1067.25, y: 247.5, width: 108, height: 59.25)
  )
  #expect(complexShape.spatialInfo.resolution == .exact)
}

private func makeInlineRichTextWorkbook() throws -> URL {
  let entries = [
    "[Content_Types].xml": """
    <?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="bin" ContentType="application/octet-stream"/>
      <Default Extension="png" ContentType="image/png"/>
      <Override PartName="/xl/embeddings/object.pptx" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    </Types>
    """,
    "_rels/.rels": """
    <?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    </Relationships>
    """,
    "xl/workbook.xml": """
    <?xml version="1.0" encoding="UTF-8"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
    </workbook>
    """,
    "xl/_rels/workbook.xml.rels": """
    <?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    </Relationships>
    """,
    "xl/worksheets/sheet1.xml": """
    <?xml version="1.0" encoding="UTF-8"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
      <sheetData><row r="1"><c r="A1" t="inlineStr"><is>
        <t xml:space="preserve">Plain </t>
        <r><rPr><rFont val="Aptos"/><b/></rPr><t>Bold</t></r>
        <r><rPr><color rgb="FFFF0000"/></rPr><t xml:space="preserve"> Red</t></r>
        <rPh sb="0" eb="1"><t>Ignored phonetic guide</t></rPh>
      </is></c></row></sheetData>
      <oleObjects><oleObject progId="PowerPoint.Show.12" shapeId="1025" r:id="rIdOle">
        <objectPr defaultSize="0" r:id="rIdPreview"><anchor moveWithCells="1" sizeWithCells="0">
          <from><xdr:col>1</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></from>
          <to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>19050</xdr:rowOff></to>
        </anchor></objectPr>
      </oleObject></oleObjects>
      <controls><control shapeId="1026" name="CommandButton1" r:id="rIdControl"/></controls>
    </worksheet>
    """,
    "xl/worksheets/_rels/sheet1.xml.rels": """
    <?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/object.pptx"/>
      <Relationship Id="rIdPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/preview.png"/>
      <Relationship Id="rIdControl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="../activeX/control.xml"/>
    </Relationships>
    """,
    "xl/activeX/_rels/control.xml.rels": """
    <?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="control.bin"/>
      <Relationship Id="rIdCycle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../worksheets/sheet1.xml"/>
    </Relationships>
    """,
    "xl/activeX/control.xml": "<ocx xmlns=\"http://schemas.microsoft.com/office/2006/activeX\"/>",
    "xl/activeX/control.bin": "inert ActiveX bytes",
    "xl/embeddings/object.pptx": "inert embedded package bytes",
    "xl/media/preview.png": "not a decoded image",
  ]
  return try makeSyntheticOfficePackage(entries: entries, pathExtension: "xlsx")
}
