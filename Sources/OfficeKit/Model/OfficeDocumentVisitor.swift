/// A synchronous, read-only visitor over OfficeKit's semantic document models.
///
/// Container callbacks are balanced on every successful traversal. Semantic objects retain their
/// source parts, spatial information, and lazy attachments, so a visitor does not need a parallel
/// package lookup table. Default implementations are no-ops.
public protocol OfficeDocumentVisitor {
  /// Begins traversal of an opened document.
  mutating func willVisitDocument(_ document: OfficeDocument)
  /// Ends traversal of an opened document.
  mutating func didVisitDocument(_ document: OfficeDocument)

  /// Begins traversal of a presentation.
  mutating func willVisitPresentation(_ presentation: OfficePresentation)
  /// Ends traversal of a presentation.
  mutating func didVisitPresentation(_ presentation: OfficePresentation)
  /// Begins traversal of one slide.
  mutating func willVisitSlide(_ slide: OfficeSlide)
  /// Ends traversal of one slide.
  mutating func didVisitSlide(_ slide: OfficeSlide)
  /// Begins traversal of a layout or master layer.
  mutating func willVisitSlideLayer(_ layer: OfficeSlideLayer)
  /// Ends traversal of a layout or master layer.
  mutating func didVisitSlideLayer(_ layer: OfficeSlideLayer)
  /// Begins traversal of a speaker-notes page.
  mutating func willVisitSlideNotes(_ notes: OfficeSlideNotes)
  /// Ends traversal of a speaker-notes page.
  mutating func didVisitSlideNotes(_ notes: OfficeSlideNotes)
  /// Begins traversal of a slide shape-tree element.
  mutating func willVisitSlideElement(_ element: OfficeSlideElement)
  /// Ends traversal of a slide shape-tree element.
  mutating func didVisitSlideElement(_ element: OfficeSlideElement)
  /// Visits one structured PresentationML text paragraph.
  mutating func visitSlideTextParagraph(_ paragraph: OfficeSlideTextParagraph)
  /// Visits one PresentationML text run or field.
  mutating func visitSlideTextRun(_ run: OfficeSlideTextRun)
  /// Visits one presentation comment.
  mutating func visitPresentationComment(_ comment: OfficePresentationComment)

  /// Begins traversal of an Excel workbook.
  mutating func willVisitWorkbook(_ workbook: OfficeWorkbook)
  /// Ends traversal of an Excel workbook.
  mutating func didVisitWorkbook(_ workbook: OfficeWorkbook)
  /// Visits one authored workbook sheet reference.
  mutating func visitWorkbookSheet(_ sheet: OfficeSheetReference)
  /// Visits one workbook or sheet-scoped defined name.
  mutating func visitDefinedName(_ name: OfficeDefinedName)
  /// Begins traversal of a grid worksheet.
  mutating func willVisitWorksheet(_ worksheet: OfficeWorksheet)
  /// Ends traversal of a grid worksheet.
  mutating func didVisitWorksheet(_ worksheet: OfficeWorksheet)
  /// Begins traversal of one sparse worksheet row.
  mutating func willVisitWorksheetRow(_ row: OfficeWorksheetRow)
  /// Ends traversal of one sparse worksheet row.
  mutating func didVisitWorksheetRow(_ row: OfficeWorksheetRow)
  /// Visits one populated worksheet cell.
  mutating func visitCell(_ cell: OfficeCell)
  /// Visits one inert worksheet OLE object or ActiveX control.
  mutating func visitWorksheetObject(_ object: OfficeWorksheetObject)
  /// Visits one anchored worksheet DrawingML object.
  mutating func visitWorksheetDrawingElement(_ element: OfficeWorksheetDrawingElement)
  /// Visits one worksheet comment or note.
  mutating func visitWorksheetComment(_ comment: OfficeWorksheetComment)

  /// Begins traversal of a WordprocessingML document.
  mutating func willVisitWordDocument(_ document: OfficeWordDocument)
  /// Ends traversal of a WordprocessingML document.
  mutating func didVisitWordDocument(_ document: OfficeWordDocument)
  /// Begins traversal of a header or footer story.
  mutating func willVisitWordStory(_ story: OfficeWordStory)
  /// Ends traversal of a header or footer story.
  mutating func didVisitWordStory(_ story: OfficeWordStory)
  /// Begins traversal of one Word paragraph.
  mutating func willVisitWordParagraph(_ paragraph: OfficeWordParagraph)
  /// Ends traversal of one Word paragraph.
  mutating func didVisitWordParagraph(_ paragraph: OfficeWordParagraph)
  /// Visits one Word text run.
  mutating func visitWordRun(_ run: OfficeWordRun)
  /// Visits one ordered inline drawing or equation.
  mutating func visitWordInlineContent(_ content: OfficeWordInlineContent)
  /// Visits one Word hyperlink.
  mutating func visitWordHyperlink(_ hyperlink: OfficeWordHyperlink)
  /// Visits one Word section declaration.
  mutating func visitWordSection(_ section: OfficeWordSection)
  /// Visits one assembled Word field.
  mutating func visitWordField(_ field: OfficeWordField)
  /// Visits one assembled Word bookmark.
  mutating func visitWordBookmark(_ bookmark: OfficeWordBookmark)
  /// Visits one structured document tag.
  mutating func visitWordContentControl(_ contentControl: OfficeWordContentControl)
  /// Begins traversal of one Word table.
  mutating func willVisitWordTable(_ table: OfficeWordTable)
  /// Ends traversal of one Word table.
  mutating func didVisitWordTable(_ table: OfficeWordTable)
  /// Begins traversal of one Word table row.
  mutating func willVisitWordTableRow(_ row: OfficeWordTableRow)
  /// Ends traversal of one Word table row.
  mutating func didVisitWordTableRow(_ row: OfficeWordTableRow)
  /// Begins traversal of one Word table cell.
  mutating func willVisitWordTableCell(_ cell: OfficeWordTableCell)
  /// Ends traversal of one Word table cell.
  mutating func didVisitWordTableCell(_ cell: OfficeWordTableCell)
  /// Visits one inline or floating DrawingML object.
  mutating func visitWordDrawing(_ drawing: OfficeWordDrawing)
  /// Visits one Office Math equation.
  mutating func visitWordEquation(_ equation: OfficeWordEquation)
  /// Visits one Word comment.
  mutating func visitWordComment(_ comment: OfficeWordComment)
  /// Visits one footnote or endnote.
  mutating func visitWordNote(_ note: OfficeWordNote)
  /// Visits one inert embedded object or package.
  mutating func visitWordEmbeddedObject(_ object: OfficeWordEmbeddedObject)
  /// Visits one alternative-format import.
  mutating func visitWordAlternativeFormatImport(_ import: OfficeWordAlternativeFormatImport)
  /// Visits one effective legacy VML shape.
  mutating func visitWordLegacyShape(_ shape: OfficeWordVMLShape)

  /// Visits one relationship-backed internal or external attachment.
  mutating func visitAttachment(_ attachment: OfficeAttachment)
}

extension OfficeDocumentVisitor {
  public mutating func willVisitDocument(_: OfficeDocument) {}
  public mutating func didVisitDocument(_: OfficeDocument) {}
  public mutating func willVisitPresentation(_: OfficePresentation) {}
  public mutating func didVisitPresentation(_: OfficePresentation) {}
  public mutating func willVisitSlide(_: OfficeSlide) {}
  public mutating func didVisitSlide(_: OfficeSlide) {}
  public mutating func willVisitSlideLayer(_: OfficeSlideLayer) {}
  public mutating func didVisitSlideLayer(_: OfficeSlideLayer) {}
  public mutating func willVisitSlideNotes(_: OfficeSlideNotes) {}
  public mutating func didVisitSlideNotes(_: OfficeSlideNotes) {}
  public mutating func willVisitSlideElement(_: OfficeSlideElement) {}
  public mutating func didVisitSlideElement(_: OfficeSlideElement) {}
  public mutating func visitSlideTextParagraph(_: OfficeSlideTextParagraph) {}
  public mutating func visitSlideTextRun(_: OfficeSlideTextRun) {}
  public mutating func visitPresentationComment(_: OfficePresentationComment) {}
  public mutating func willVisitWorkbook(_: OfficeWorkbook) {}
  public mutating func didVisitWorkbook(_: OfficeWorkbook) {}
  public mutating func visitWorkbookSheet(_: OfficeSheetReference) {}
  public mutating func visitDefinedName(_: OfficeDefinedName) {}
  public mutating func willVisitWorksheet(_: OfficeWorksheet) {}
  public mutating func didVisitWorksheet(_: OfficeWorksheet) {}
  public mutating func willVisitWorksheetRow(_: OfficeWorksheetRow) {}
  public mutating func didVisitWorksheetRow(_: OfficeWorksheetRow) {}
  public mutating func visitCell(_: OfficeCell) {}
  public mutating func visitWorksheetObject(_: OfficeWorksheetObject) {}
  public mutating func visitWorksheetDrawingElement(_: OfficeWorksheetDrawingElement) {}
  public mutating func visitWorksheetComment(_: OfficeWorksheetComment) {}
  public mutating func willVisitWordDocument(_: OfficeWordDocument) {}
  public mutating func didVisitWordDocument(_: OfficeWordDocument) {}
  public mutating func willVisitWordStory(_: OfficeWordStory) {}
  public mutating func didVisitWordStory(_: OfficeWordStory) {}
  public mutating func willVisitWordParagraph(_: OfficeWordParagraph) {}
  public mutating func didVisitWordParagraph(_: OfficeWordParagraph) {}
  public mutating func visitWordRun(_: OfficeWordRun) {}
  public mutating func visitWordInlineContent(_: OfficeWordInlineContent) {}
  public mutating func visitWordHyperlink(_: OfficeWordHyperlink) {}
  public mutating func visitWordSection(_: OfficeWordSection) {}
  public mutating func visitWordField(_: OfficeWordField) {}
  public mutating func visitWordBookmark(_: OfficeWordBookmark) {}
  public mutating func visitWordContentControl(_: OfficeWordContentControl) {}
  public mutating func willVisitWordTable(_: OfficeWordTable) {}
  public mutating func didVisitWordTable(_: OfficeWordTable) {}
  public mutating func willVisitWordTableRow(_: OfficeWordTableRow) {}
  public mutating func didVisitWordTableRow(_: OfficeWordTableRow) {}
  public mutating func willVisitWordTableCell(_: OfficeWordTableCell) {}
  public mutating func didVisitWordTableCell(_: OfficeWordTableCell) {}
  public mutating func visitWordDrawing(_: OfficeWordDrawing) {}
  public mutating func visitWordEquation(_: OfficeWordEquation) {}
  public mutating func visitWordComment(_: OfficeWordComment) {}
  public mutating func visitWordNote(_: OfficeWordNote) {}
  public mutating func visitWordEmbeddedObject(_: OfficeWordEmbeddedObject) {}
  public mutating func visitWordAlternativeFormatImport(_: OfficeWordAlternativeFormatImport) {}
  public mutating func visitWordLegacyShape(_: OfficeWordVMLShape) {}
  public mutating func visitAttachment(_: OfficeAttachment) {}
}

extension OfficeDocument {
  /// Parses semantic content and visits it in deterministic authored order.
  ///
  /// The traversal is synchronous because package and XML reads are local. PowerPoint slides and
  /// Excel worksheets are parsed one at a time rather than retained as a whole-document tree.
  public func traverse<Visitor: OfficeDocumentVisitor>(
    using visitor: inout Visitor
  ) throws {
    switch kind {
    case .presentation:
      let presentation = try OfficePresentation(document: self)
      visitor.willVisitDocument(self)
      try traverse(presentation, using: &visitor)
      visitor.didVisitDocument(self)
    case .spreadsheet:
      let workbook = try OfficeWorkbook(document: self)
      visitor.willVisitDocument(self)
      try traverse(workbook, using: &visitor)
      visitor.didVisitDocument(self)
    case .wordProcessing:
      let wordDocument = try OfficeWordDocument(document: self)
      visitor.willVisitDocument(self)
      traverse(wordDocument, using: &visitor)
      visitor.didVisitDocument(self)
    }
  }

  private func traverse<Visitor: OfficeDocumentVisitor>(
    _ presentation: OfficePresentation,
    using visitor: inout Visitor
  ) throws {
    visitor.willVisitPresentation(presentation)
    for index in presentation.slides.indices {
      let slide = try presentation.slide(at: index)
      visitor.willVisitSlide(slide)
      for element in slide.elements { traverse(element, using: &visitor) }
      if let layoutLayer = slide.layoutLayer { traverse(layoutLayer, using: &visitor) }
      if let masterLayer = slide.masterLayer { traverse(masterLayer, using: &visitor) }
      if let notes = slide.notes { traverse(notes, using: &visitor) }
      for comment in slide.comments { visitor.visitPresentationComment(comment) }
      for attachment in slide.relatedAttachments { visitor.visitAttachment(attachment) }
      visitor.didVisitSlide(slide)
    }
    visitor.didVisitPresentation(presentation)
  }

  private func traverse<Visitor: OfficeDocumentVisitor>(
    _ element: OfficeSlideElement,
    using visitor: inout Visitor
  ) {
    visitor.willVisitSlideElement(element)
    if let textBody = element.textBody {
      for paragraph in textBody.paragraphs {
        visitor.visitSlideTextParagraph(paragraph)
        for run in paragraph.runs { visitor.visitSlideTextRun(run) }
      }
    }
    for attachment in element.attachments { visitor.visitAttachment(attachment) }
    for child in element.children { traverse(child, using: &visitor) }
    visitor.didVisitSlideElement(element)
  }

  private func traverse<Visitor: OfficeDocumentVisitor>(
    _ layer: OfficeSlideLayer,
    using visitor: inout Visitor
  ) {
    visitor.willVisitSlideLayer(layer)
    for element in layer.elements { traverse(element, using: &visitor) }
    for attachment in layer.relatedAttachments { visitor.visitAttachment(attachment) }
    visitor.didVisitSlideLayer(layer)
  }

  private func traverse<Visitor: OfficeDocumentVisitor>(
    _ notes: OfficeSlideNotes,
    using visitor: inout Visitor
  ) {
    visitor.willVisitSlideNotes(notes)
    for element in notes.elements { traverse(element, using: &visitor) }
    if let masterLayer = notes.masterLayer { traverse(masterLayer, using: &visitor) }
    for attachment in notes.relatedAttachments { visitor.visitAttachment(attachment) }
    visitor.didVisitSlideNotes(notes)
  }

  private func traverse<Visitor: OfficeDocumentVisitor>(
    _ workbook: OfficeWorkbook,
    using visitor: inout Visitor
  ) throws {
    visitor.willVisitWorkbook(workbook)
    for sheet in workbook.sheets { visitor.visitWorkbookSheet(sheet) }
    for name in workbook.definedNames { visitor.visitDefinedName(name) }
    for attachment in workbook.attachments { visitor.visitAttachment(attachment) }
    for index in workbook.worksheets.indices {
      let worksheet = try workbook.worksheet(at: index)
      visitor.willVisitWorksheet(worksheet)
      for row in worksheet.rows {
        visitor.willVisitWorksheetRow(row)
        for cell in row.cells { visitor.visitCell(cell) }
        visitor.didVisitWorksheetRow(row)
      }
      for drawing in worksheet.drawings {
        for element in drawing.elements {
          visitor.visitWorksheetDrawingElement(element)
          for attachment in element.attachments { visitor.visitAttachment(attachment) }
        }
      }
      for comment in worksheet.comments { visitor.visitWorksheetComment(comment) }
      for object in worksheet.objects { visitor.visitWorksheetObject(object) }
      for attachment in worksheet.attachments { visitor.visitAttachment(attachment) }
      visitor.didVisitWorksheet(worksheet)
    }
    visitor.didVisitWorkbook(workbook)
  }

  private func traverse<Visitor: OfficeDocumentVisitor>(
    _ document: OfficeWordDocument,
    using visitor: inout Visitor
  ) {
    visitor.willVisitWordDocument(document)
    traverse(document.body, using: &visitor)
    for story in document.headers + document.footers {
      visitor.willVisitWordStory(story)
      traverse(story.content, using: &visitor)
      for attachment in story.attachments { visitor.visitAttachment(attachment) }
      visitor.didVisitWordStory(story)
    }
    if let footnotes = document.footnotes {
      for note in footnotes.notes {
        visitor.visitWordNote(note)
        traverse(note.content, using: &visitor)
      }
      for attachment in footnotes.attachments { visitor.visitAttachment(attachment) }
    }
    if let endnotes = document.endnotes {
      for note in endnotes.notes {
        visitor.visitWordNote(note)
        traverse(note.content, using: &visitor)
      }
      for attachment in endnotes.attachments { visitor.visitAttachment(attachment) }
    }
    for comment in document.comments {
      visitor.visitWordComment(comment)
      traverse(comment.content, using: &visitor)
      for attachment in comment.attachments { visitor.visitAttachment(attachment) }
    }
    for attachment in document.attachments { visitor.visitAttachment(attachment) }
    visitor.didVisitWordDocument(document)
  }

  private func traverse<Visitor: OfficeDocumentVisitor>(
    _ body: OfficeWordBody,
    using visitor: inout Visitor
  ) {
    for block in body.blocks { traverse(block, body: body, using: &visitor) }
    for section in body.sections { visitor.visitWordSection(section) }
    for field in body.fields { visitor.visitWordField(field) }
    for bookmark in body.bookmarks { visitor.visitWordBookmark(bookmark) }
    for contentControl in body.contentControls {
      visitor.visitWordContentControl(contentControl)
    }
    for drawing in body.drawings where drawing.inlinePosition == nil {
      visit(drawing, using: &visitor)
    }
    for equation in body.equations where equation.inlinePosition == nil {
      visitor.visitWordEquation(equation)
    }
    for object in body.embeddedObjects {
      visitor.visitWordEmbeddedObject(object)
      visitor.visitAttachment(object.attachment)
    }
    for item in body.alternativeFormatImports {
      visitor.visitWordAlternativeFormatImport(item)
      visitor.visitAttachment(item.attachment)
    }
    for shape in body.legacyShapes {
      visitor.visitWordLegacyShape(shape)
      for attachment in shape.attachments { visitor.visitAttachment(attachment) }
    }
  }

  private func traverse<Visitor: OfficeDocumentVisitor>(
    _ block: OfficeWordBlock,
    body: OfficeWordBody,
    using visitor: inout Visitor
  ) {
    switch block {
    case .paragraph(let paragraph):
      visitor.willVisitWordParagraph(paragraph)
      let inlineContent = body.inlineContent(in: paragraph.location)
      for (runIndex, run) in paragraph.runs.enumerated() {
        for content in inlineContent
        where content.position?.runIndex == runIndex
          && content.position?.characterOffset == 0
        {
          visit(content, using: &visitor)
        }
        visitor.visitWordRun(run)
        for content in inlineContent
        where content.position?.runIndex == runIndex
          && content.position?.characterOffset != 0
        {
          visit(content, using: &visitor)
        }
      }
      for content in inlineContent where (content.position?.runIndex ?? 0) >= paragraph.runs.count {
        visit(content, using: &visitor)
      }
      for hyperlink in paragraph.hyperlinks { visitor.visitWordHyperlink(hyperlink) }
      visitor.didVisitWordParagraph(paragraph)
    case .table(let table):
      visitor.willVisitWordTable(table)
      for row in table.rows {
        visitor.willVisitWordTableRow(row)
        for cell in row.cells {
          visitor.willVisitWordTableCell(cell)
          for child in cell.blocks { traverse(child, body: body, using: &visitor) }
          visitor.didVisitWordTableCell(cell)
        }
        visitor.didVisitWordTableRow(row)
      }
      visitor.didVisitWordTable(table)
    }
  }

  private func visit<Visitor: OfficeDocumentVisitor>(
    _ content: OfficeWordInlineContent,
    using visitor: inout Visitor
  ) {
    visitor.visitWordInlineContent(content)
    switch content {
    case .drawing(let drawing): visit(drawing, using: &visitor)
    case .equation(let equation): visitor.visitWordEquation(equation)
    }
  }

  private func visit<Visitor: OfficeDocumentVisitor>(
    _ drawing: OfficeWordDrawing,
    using visitor: inout Visitor
  ) {
    visitor.visitWordDrawing(drawing)
    for attachment in drawing.attachments { visitor.visitAttachment(attachment) }
  }
}
