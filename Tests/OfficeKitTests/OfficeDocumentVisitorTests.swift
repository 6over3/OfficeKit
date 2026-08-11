import Testing

@testable import OfficeKit

private struct RecordingOfficeVisitor: OfficeDocumentVisitor {
  var stack: [String] = []
  var mismatch = false
  var begins: [String: Int] = [:]
  var ends: [String: Int] = [:]
  var cellCount = 0
  var runCount = 0
  var slideTextRunCount = 0
  var inlineContentCount = 0
  var attachmentCount = 0

  mutating func begin(_ name: String) {
    stack.append(name)
    begins[name, default: 0] += 1
  }

  mutating func end(_ name: String) {
    if stack.popLast() != name { mismatch = true }
    ends[name, default: 0] += 1
  }

  mutating func willVisitDocument(_: OfficeDocument) { begin("document") }
  mutating func didVisitDocument(_: OfficeDocument) { end("document") }
  mutating func willVisitPresentation(_: OfficePresentation) { begin("presentation") }
  mutating func didVisitPresentation(_: OfficePresentation) { end("presentation") }
  mutating func willVisitSlide(_: OfficeSlide) { begin("slide") }
  mutating func didVisitSlide(_: OfficeSlide) { end("slide") }
  mutating func willVisitSlideLayer(_: OfficeSlideLayer) { begin("slideLayer") }
  mutating func didVisitSlideLayer(_: OfficeSlideLayer) { end("slideLayer") }
  mutating func willVisitSlideNotes(_: OfficeSlideNotes) { begin("slideNotes") }
  mutating func didVisitSlideNotes(_: OfficeSlideNotes) { end("slideNotes") }
  mutating func willVisitSlideElement(_: OfficeSlideElement) { begin("slideElement") }
  mutating func didVisitSlideElement(_: OfficeSlideElement) { end("slideElement") }
  mutating func visitSlideTextRun(_: OfficeSlideTextRun) { slideTextRunCount += 1 }
  mutating func willVisitWorkbook(_: OfficeWorkbook) { begin("workbook") }
  mutating func didVisitWorkbook(_: OfficeWorkbook) { end("workbook") }
  mutating func willVisitWorksheet(_: OfficeWorksheet) { begin("worksheet") }
  mutating func didVisitWorksheet(_: OfficeWorksheet) { end("worksheet") }
  mutating func willVisitWorksheetRow(_: OfficeWorksheetRow) { begin("worksheetRow") }
  mutating func didVisitWorksheetRow(_: OfficeWorksheetRow) { end("worksheetRow") }
  mutating func visitCell(_: OfficeCell) { cellCount += 1 }
  mutating func willVisitWordDocument(_: OfficeWordDocument) { begin("wordDocument") }
  mutating func didVisitWordDocument(_: OfficeWordDocument) { end("wordDocument") }
  mutating func willVisitWordStory(_: OfficeWordStory) { begin("wordStory") }
  mutating func didVisitWordStory(_: OfficeWordStory) { end("wordStory") }
  mutating func willVisitWordParagraph(_: OfficeWordParagraph) { begin("wordParagraph") }
  mutating func didVisitWordParagraph(_: OfficeWordParagraph) { end("wordParagraph") }
  mutating func visitWordRun(_: OfficeWordRun) { runCount += 1 }
  mutating func visitWordInlineContent(_: OfficeWordInlineContent) { inlineContentCount += 1 }
  mutating func willVisitWordTable(_: OfficeWordTable) { begin("wordTable") }
  mutating func didVisitWordTable(_: OfficeWordTable) { end("wordTable") }
  mutating func willVisitWordTableRow(_: OfficeWordTableRow) { begin("wordTableRow") }
  mutating func didVisitWordTableRow(_: OfficeWordTableRow) { end("wordTableRow") }
  mutating func willVisitWordTableCell(_: OfficeWordTableCell) { begin("wordTableCell") }
  mutating func didVisitWordTableCell(_: OfficeWordTableCell) { end("wordTableCell") }
  mutating func visitAttachment(_: OfficeAttachment) { attachmentCount += 1 }
}

@Test func semanticVisitorBalancesContainersAcrossAllDocumentFamilies() throws {
  let fixturePaths = [
    "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/smallset/"
      + "Text_withExtrusion_200chars.pptx",
    "Open-XML-SDK/TestDataStorage/v2FxTestFiles/spreadsheet/smallset/SheetData.xlsx",
    "Open-XML-SDK/TestFiles/Complex01.docx",
  ]

  for path in fixturePaths {
    let document = try OfficeDocument(contentsOf: FixtureCatalog.url(for: path))
    var visitor = RecordingOfficeVisitor()
    try document.traverse(using: &visitor)

    #expect(visitor.stack.isEmpty)
    #expect(!visitor.mismatch)
    #expect(visitor.begins == visitor.ends)
    #expect(visitor.begins["document"] == 1)
    switch document.kind {
    case .presentation:
      #expect(visitor.begins["presentation"] == 1)
      #expect(visitor.begins["slide", default: 0] > 0)
      #expect(visitor.begins["slideElement", default: 0] > 0)
      #expect(visitor.begins["slideLayer", default: 0] > 0)
      #expect(visitor.begins["slideNotes", default: 0] > 0)
      #expect(visitor.slideTextRunCount > 0)
    case .spreadsheet:
      #expect(visitor.begins["workbook"] == 1)
      #expect(visitor.begins["worksheet"] == 3)
      #expect(visitor.cellCount > 0)
    case .wordProcessing:
      #expect(visitor.begins["wordDocument"] == 1)
      #expect(visitor.begins["wordStory"] == 18)
      #expect(visitor.runCount > 0)
      #expect(visitor.attachmentCount > 0)
      #expect(visitor.inlineContentCount > 0)
    }
  }
}
