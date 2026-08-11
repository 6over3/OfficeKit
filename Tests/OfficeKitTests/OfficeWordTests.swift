import Foundation
import Testing

@testable import OfficeKit

@Test func wordAlternativeFormatImportsRemainLazyAndURLBacked() throws {
  let url = try makeAltChunkDocument()
  defer { try? FileManager.default.removeItem(at: url) }
  let document = try OfficeWordDocument(contentsOf: url)
  let item = try #require(document.body.alternativeFormatImports.first)

  #expect(document.body.paragraphs.map(\.text) == ["Before", "After"])
  #expect(item.index == 0)
  #expect(item.relationshipID == OfficeRelationshipID(rawValue: "rIdAltChunk"))
  #expect(item.attachment.part?.name.rawValue == "/word/imported.html")
  #expect(item.attachment.declaredContentType?.rawValue == "text/html")
  let payloadURL = try item.attachment.url()
  #expect(payloadURL.isFileURL)
  #expect(try String(contentsOf: payloadURL, encoding: .utf8) == "<p>Imported content</p>")
}

@Test func wordMarkupCompatibilitySelectsSupportedDrawingChoice() throws {
  let url = try FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/mcdoc.docx")
  let document = try OfficeWordDocument(contentsOf: url)

  #expect(document.body.drawings.count == 1)
  #expect(document.body.legacyShapes.isEmpty)
}

@Test func wordCommentsCanAnchorAtTableStructureBoundaries() throws {
  let url = try FixtureCatalog.url(
    for:
      "Open-XML-SDK/TestDataStorage/v2FxTestFiles/wordprocessing/table/TableComments/rowWithComment.docx"
  )
  let document = try OfficeWordDocument(contentsOf: url)
  let comment = try #require(document.comments.first)

  #expect(document.comments.count == 1)
  #expect(comment.identifier == 0)
  #expect(comment.anchor != nil)
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/OpenXmlDomTest/DocumentTraverseTest.cs,
// TraverseWordDocument.
@Test(arguments: [
  (filename: "AdjustRightInd.docx", paragraphCount: 2, styleCount: 4),
  (filename: "AutoSpaceDE.docx", paragraphCount: 2, styleCount: 8),
  (filename: "Empty.docx", paragraphCount: 3, styleCount: 4),
])
func traverseWordDocument(
  filename: String,
  paragraphCount: Int,
  styleCount: Int
) throws {
  let document = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/wordprocessing/paragraph/" + filename
    )
  )
  let section = try #require(document.body.sections.first)
  let stylesAttachment = try #require(
    document.attachments.first {
      $0.relationship.type.isEquivalent(to: .styles)
    })

  #expect(document.document.kind == .wordProcessing)
  #expect(document.body.sourcePart.name.rawValue == "/word/document.xml")
  #expect(document.body.paragraphs.count == paragraphCount)
  #expect(document.body.blocks.count == paragraphCount)
  #expect(document.styles.count == styleCount)
  #expect(document.styles.first?.identifier == "Normal")
  #expect(document.styles.first?.type == "paragraph")
  #expect(document.styles.first?.isDefault == true)
  #expect(document.styles.last?.identifier == (styleCount == 8 ? "FooterChar" : "NoList"))
  #expect(document.styles.allSatisfy { $0.sourcePart.name.rawValue == "/word/styles.xml" })
  #expect(try stylesAttachment.url().isFileURL)

  for (index, paragraph) in document.body.paragraphs.enumerated() {
    #expect(paragraph.index == index)
    #expect(paragraph.text == paragraph.runs.map(\.text).joined())
    #expect(paragraph.sourcePart == document.document.mainPart)
    #expect(paragraph.spatialInfo.frame == nil)
    guard case .unresolved = paragraph.spatialInfo.resolution else {
      Issue.record("Ordinary Word paragraph geometry must remain layout-dependent.")
      continue
    }
  }

  #expect(document.body.sections.count == 1)
  #expect(section.pageWidth?.points == 612)
  #expect(section.pageHeight?.points == 792)
  #expect(section.spatialInfo.frame == OfficeRect(x: 0, y: 0, width: 612, height: 792))
  #expect(section.spatialInfo.resolution == .exact)
}

@Test func wordParagraphPropertiesRetainFixtureSemantics() throws {
  let adjustDocument = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/wordprocessing/paragraph/AdjustRightInd.docx"
    )
  )
  let adjustedParagraph = try #require(adjustDocument.body.paragraphs.first)
  #expect(adjustedParagraph.text == " ‘AdjustRightOnOff’ is off.")
  #expect(adjustedParagraph.properties.automaticallySpacesEastAsianAndLatinText == false)
  #expect(adjustedParagraph.properties.automaticallySpacesEastAsianTextAndNumbers == false)
  #expect(adjustedParagraph.properties.adjustsRightIndent == false)

  let spacingDocument = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/wordprocessing/paragraph/AutoSpaceDE.docx"
    )
  )
  #expect(spacingDocument.body.paragraphs[0].text == "AutoSpaceDE is on (default).")
  #expect(spacingDocument.body.paragraphs[0].properties.wrapsAtCharacter == false)
  #expect(spacingDocument.body.paragraphs[1].text == "AutoSpaceDE is off.")
  #expect(
    spacingDocument.body.paragraphs[1].properties.automaticallySpacesEastAsianAndLatinText
      == false
  )
  #expect(spacingDocument.body.paragraphs[1].properties.textAlignment == "center")
}

@Test func wordStylesDefaultsLanguageAndInheritanceResolveWithoutLosingDirectFormatting() throws {
  let document = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Complex01.docx")
  )
  let defaults = try #require(document.styleDefaults)
  let headingStyle = try #require(document.style(identifiedBy: "Heading1"))
  let heading = try #require(
    document.body.paragraphs.first {
      $0.properties.styleIdentifier == "Heading1"
    })
  let highlightedRun = try #require(
    document.body.paragraphs.lazy.flatMap(\.runs).first {
      $0.properties.highlight == "yellow"
    })
  let resolvedParagraph = document.resolvedParagraphProperties(for: heading)
  let resolvedRun = document.resolvedRunProperties(
    for: try #require(heading.runs.first),
    in: heading
  )

  #expect(defaults.runProperties.fontSize?.points == 11)
  #expect(defaults.runProperties.language == "en-US")
  #expect(defaults.runProperties.eastAsianLanguage == "en-US")
  #expect(defaults.runProperties.bidirectionalLanguage == "ar-SA")
  #expect(defaults.runProperties.asciiThemeFont == "minorHAnsi")
  #expect(defaults.paragraphProperties.spacingAfter?.points == 8)
  #expect(defaults.paragraphProperties.lineSpacing?.points == 12.95)
  #expect(defaults.paragraphProperties.lineSpacingRule == "auto")
  #expect(headingStyle.basedOnIdentifier == "Normal")
  #expect(headingStyle.runProperties.fontSize?.points == 16)
  #expect(resolvedParagraph.keepsWithNext == true)
  #expect(resolvedParagraph.spacingBefore?.points == 12)
  #expect(resolvedRun.fontSize?.points == 16)
  #expect(resolvedRun.language == "en-US")
  #expect(highlightedRun.properties.highlight == "yellow")
}

@Test func wordHyperlinksResolveExternalURLsWithoutFetching() throws {
  let document = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Hyperlink.docx")
  )
  let paragraph = try #require(document.body.paragraphs.last)
  let hyperlink = try #require(paragraph.hyperlinks.first)
  let attachment = try #require(hyperlink.attachment)

  #expect(paragraph.text == "EricWhite.com")
  #expect(paragraph.runs.first?.styleIdentifier == "Hyperlink")
  #expect(hyperlink.runRange == 0..<1)
  #expect(hyperlink.relationshipID == OfficeRelationshipID(rawValue: "rId4"))
  #expect(attachment.relationship.type.isEquivalent(to: .hyperlink))
  #expect(attachment.part == nil)
  #expect(try attachment.url().absoluteString == "http://www.ericwhite.com")
}

// Upstream fixture generated by DocumentFormat.OpenXml.Tests/GeneratedClass003.cs.
@Test func wordCommentsResolveTextMetadataAndAuthoredRunAnchors() throws {
  let document = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Comments.docx")
  )
  let comment = try #require(document.comments.first)
  let anchor = try #require(comment.anchor)
  let commentsAttachment = try #require(
    document.attachments.first {
      $0.relationship.type.isEquivalent(to: .comments)
    })

  #expect(document.comments.count == 1)
  #expect(document.body.commentAnchors.count == 1)
  #expect(comment.identifier == 1)
  #expect(comment.author == "Eric White")
  #expect(comment.initials == "EW")
  #expect(comment.dateText == "2014-10-28T20:22:00Z")
  #expect(comment.date != nil)
  #expect(comment.paragraphs == ["This is a comment."])
  #expect(comment.text == "This is a comment.")
  #expect(comment.content.paragraphs.map(\.text) == ["This is a comment."])
  #expect(comment.content.sourcePart == comment.sourcePart)
  #expect(comment.sourcePart.name.rawValue == "/word/comments.xml")
  #expect(anchor.identifier == 1)
  #expect(anchor.start == OfficeWordTextPosition(paragraph: .body(blockIndex: 0), runIndex: 1))
  #expect(anchor.end == OfficeWordTextPosition(paragraph: .body(blockIndex: 0), runIndex: 2))
  #expect(
    anchor.referenceRun
      == OfficeWordTextPosition(paragraph: .body(blockIndex: 0), runIndex: 2)
  )
  #expect(anchor.spatialInfo.frame == nil)
  guard case .unresolved = anchor.spatialInfo.resolution else {
    Issue.record("Word comment marker coordinates must remain layout-dependent.")
    return
  }
  #expect(commentsAttachment.part?.name.rawValue == "/word/comments.xml")
  #expect(try commentsAttachment.url().isFileURL)
}

// Upstream fixture generated by DocumentFormat.OpenXml.Tests/GeneratedClass001.cs.
@Test func wordStoriesNotesAndSectionReferencesResolveThroughRelationships() throws {
  let document = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Complex01.docx")
  )
  let footnotes = try #require(document.footnotes)
  let endnotes = try #require(document.endnotes)
  let footnote = try #require(footnotes.note(identifiedBy: 1))
  let endnote = try #require(endnotes.note(identifiedBy: 1))
  let footnotesAttachment = try #require(
    document.attachments.first {
      $0.relationship.type == .footnotes
    })
  let endnotesAttachment = try #require(
    document.attachments.first {
      $0.relationship.type == .endnotes
    })
  let runs = document.body.paragraphs.flatMap(\.runs)

  #expect(document.headers.count == 9)
  #expect(document.footers.count == 9)
  #expect(document.headers.map(\.part.name.rawValue).contains("/word/header2.xml"))
  #expect(document.footers.map(\.part.name.rawValue).contains("/word/footer9.xml"))
  #expect(document.headers.allSatisfy { $0.kind == .header })
  #expect(document.footers.allSatisfy { $0.kind == .footer })
  #expect(document.headers.allSatisfy { $0.content.sourcePart == $0.part })
  #expect(document.body.sections.count == 4)
  #expect(document.body.sections.flatMap(\.headerReferences).count == 9)
  #expect(document.body.sections.flatMap(\.footerReferences).count == 9)
  #expect(
    document.body.sections.flatMap(\.headerReferences).map(\.variant).contains(.firstPage)
  )
  #expect(
    document.body.sections.flatMap(\.footerReferences).map(\.variant).contains(.evenPages)
  )

  #expect(footnotes.notes.map(\.identifier) == [-1, 0, 1])
  #expect(endnotes.notes.map(\.identifier) == [-1, 0, 1])
  #expect(footnotes.notes[0].type == "separator")
  #expect(endnotes.notes[1].type == "continuationSeparator")
  #expect(footnote.text == " This is a footnote.")
  #expect(endnote.text == " This is an endnote.")
  #expect(footnote.content.paragraphs.count == 1)
  #expect(footnote.content.paragraphs[0].properties.styleIdentifier == "FootnoteText")
  #expect(footnote.content.paragraphs[0].runs.count == 2)
  #expect(endnote.content.paragraphs.count == 1)
  #expect(footnote.content.sourcePart == footnotes.sourcePart)
  #expect(runs.compactMap(\.footnoteIdentifier) == [1])
  #expect(runs.compactMap(\.endnoteIdentifier) == [1])
  #expect(try footnotesAttachment.url().isFileURL)
  #expect(try endnotesAttachment.url().isFileURL)
}

@Test func wordNumberingResolvesConcreteListLevelsAndExactIndents() throws {
  let document = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Complex01.docx")
  )
  let numbering = try #require(document.numbering)
  let paragraph = try #require(
    document.body.paragraphs.first {
      $0.properties.numberingIdentifier == 1 && $0.properties.numberingLevel == 0
    })
  let list = try #require(document.listInfo(for: paragraph))
  let bulletLevel = try #require(numbering.level(numberingIdentifier: 3, level: 0))
  let attachment = try #require(
    document.attachments.first {
      $0.relationship.type == .numbering
    })

  #expect(numbering.abstractDefinitions.count == 4)
  #expect(numbering.instances.count == 4)
  #expect(numbering.instances.map(\.identifier) == [1, 2, 3, 4])
  #expect(numbering.instances.map(\.abstractIdentifier) == [3, 0, 1, 2])
  #expect(list.numberingIdentifier == 1)
  #expect(list.levelIndex == 0)
  #expect(list.start == 1)
  #expect(list.format == "decimal")
  #expect(list.levelText == "%1)")
  #expect(list.justification == "left")
  #expect(list.leftIndent?.points == 18)
  #expect(list.hangingIndent?.points == 18)
  #expect(bulletLevel.format == "bullet")
  #expect(bulletLevel.text == "")
  #expect(attachment.part?.name.rawValue == "/word/numbering.xml")
  #expect(try attachment.url().isFileURL)
}

@Test func wordFieldsBookmarksAndTrackedRevisionViewsPreserveAuthoredContent() throws {
  let document = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Complex01.docx")
  )
  let dateField = try #require(
    document.body.fields.first {
      $0.instruction.contains("DATE")
    })
  let bookmark = try #require(document.body.bookmarks.first { $0.name == "ABookmark" })
  let revisedParagraph = try #require(
    document.body.paragraphs.first { paragraph in
      paragraph.runs.contains { $0.revision?.kind == .deletion }
    })
  let deletedRun = try #require(
    revisedParagraph.runs.first {
      $0.revision?.kind == .deletion
    })

  #expect(document.body.fields.count == 6)
  #expect(dateField.kind == .complex)
  #expect(dateField.instruction == " DATE \\@ \"dddd, MMMM d, yyyy\" ")
  #expect(dateField.separator != nil)
  #expect(dateField.end != nil)
  #expect(document.body.bookmarks.count == 7)
  #expect(bookmark.identifier == 2)
  #expect(
    bookmark.start
      == .text(
        OfficeWordTextPosition(
          paragraph: .body(blockIndex: 54),
          runIndex: 0
        )))
  #expect(
    bookmark.end
      == .text(
        OfficeWordTextPosition(
          paragraph: .body(blockIndex: 54),
          runIndex: 1
        )))
  #expect(deletedRun.text == " in the embed code for the video you want to add")
  #expect(deletedRun.revision?.identifier == 8)
  #expect(deletedRun.revision?.author == "Eric White")
  #expect(deletedRun.revision?.dateText == "2014-10-28T20:44:00Z")
  #expect(!revisedParagraph.text.contains("embed code for the video"))
  #expect(revisedParagraph.text(view: .final) == revisedParagraph.text)
  #expect(revisedParagraph.text(view: .original).contains("embed code for the video"))
  #expect(revisedParagraph.text(view: .all).contains("embed code for the video"))
}

@Test func wordContentControlsPreserveScopePropertiesItemsAndText() throws {
  let document = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Complex01.docx")
  )
  let richText = try #require(
    document.body.contentControls.first {
      $0.tag == "RichTextContentControl"
    })
  let plainText = try #require(
    document.body.contentControls.first {
      $0.tag == "Plain text content control"
    })
  let checkbox = try #require(
    document.body.contentControls.first {
      $0.tag == "CheckBoxContentControl"
    })
  let comboBox = try #require(
    document.body.contentControls.first {
      $0.kind == "comboBox"
    })
  let table = try #require(document.body.contentControls.first { $0.tag == "MyTable" })

  // Four controls live in unselected markup-compatibility branches.
  #expect(document.body.contentControls.count == 28)
  #expect(richText.alias == "RichTextContentControl")
  #expect(richText.scope == .block)
  #expect(richText.text.contains("Video provides a powerful way"))
  #expect(plainText.kind == "text")
  #expect(plainText.scope == .inline)
  #expect(plainText.text == "make your document look professionally")
  #expect(checkbox.kind == "checkbox")
  #expect(checkbox.isChecked == false)
  #expect(checkbox.text == "☐")
  #expect(comboBox.items.map(\.displayText) == ["One", "Two", "Three"])
  #expect(comboBox.items.map(\.value) == ["One", "Two", "Three"])
  #expect(table.scope == .block)
  #expect(table.text.contains("Lorem ipsum dolor sit amet"))
  #expect(table.text.contains("600"))
  #expect(
    document.body.contentControls.allSatisfy {
      $0.sourcePart.name.rawValue == "/word/document.xml"
    })
}

@Test func wordTablesPreserveNestedBlockOrderSpansAndVerticalMerges() throws {
  let document = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/wordprocessing/complexDocx/"
        + "complex tables.docx"
    )
  )
  let table = try #require(document.body.tables.first)
  let cells = table.rows.flatMap(\.cells)
  let nestedContainer = try #require(cells.first { !$0.tables.isEmpty })
  let nestedTable = try #require(nestedContainer.tables.first)

  #expect(document.body.tables.count == 1)
  #expect(table.rows.count == 11)
  #expect(table.columnWidths.count == 11)
  #expect(table.columnWidths.first?.points == 63.7)
  #expect(table.preferredWidth == OfficeWordWidth(type: "auto", value: 0))
  #expect(table.indentation?.length?.points == 5.4)
  #expect(cells.filter { $0.gridSpan > 1 }.map(\.gridSpan) == [11, 3, 2])
  #expect(cells.filter { $0.verticalMerge == "restart" }.count == 2)
  #expect(cells.filter { $0.verticalMerge == "continue" }.count == 2)
  #expect(nestedTable.rows.count == 2)
  #expect(cells.first?.preferredWidth?.length?.points == 653.4)
  #expect(cells.first?.borders.top?.style == "single")
  #expect(cells.first?.borders.top?.sizeInEighthPoints == 18)
  #expect(cells.first?.shading?.pattern == "pct25")
  #expect(cells.contains { $0.shading?.fill == "FFCC00" })
  #expect(
    nestedContainer.blocks.contains { block in
      guard case .table = block else { return false }
      return true
    })
  #expect(table.text.contains("LEADERSHIP SKILLS"))
  #expect(table.text.contains("High Tech Industry"))
}

@Test func wordDrawingMLPreservesAnchorsGeometryAndLazyRelatedResources() throws {
  let document = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Complex01.docx")
  )
  let first = try #require(document.body.drawings.first)
  let chartDrawing = try #require(document.body.drawings.first { $0.kind == .chart })
  let diagramDrawing = try #require(document.body.drawings.first { $0.kind == .diagram })
  let pictureDrawing = try #require(document.body.drawings.first { $0.picture != nil })
  let firstPosition = try #require(first.inlinePosition)
  let inlineContent = document.body.inlineContent(in: firstPosition.paragraph)

  #expect(document.body.drawings.count == 17)
  #expect(document.body.drawings.filter { $0.anchor.placement == .floating }.count == 11)
  #expect(document.body.drawings.filter { $0.anchor.placement == .inline }.count == 6)
  #expect(document.body.drawings.allSatisfy { $0.inlinePosition != nil })
  #expect(
    inlineContent.contains { content in
      guard case .drawing(let drawing) = content else { return false }
      return drawing.sourceOrder == first.sourceOrder
    })
  #expect(first.kind == .shape)
  #expect(first.identifier == 1)
  #expect(first.name == "Elbow Connector 6")
  #expect(first.anchor.width.emu == 2_149_522)
  #expect(first.anchor.height.emu == 1_207_827)
  #expect(first.anchor.horizontalPosition?.relativeFrom == "column")
  #expect(first.anchor.horizontalPosition?.offset?.emu == 2_422_478)
  #expect(first.anchor.verticalPosition?.relativeFrom == "paragraph")
  #expect(first.anchor.verticalPosition?.offset?.emu == 211_539)
  #expect(first.anchor.wrap == "wrapNone")
  #expect(first.anchor.allowsOverlap == true)
  #expect(first.spatialInfo.frame == nil)
  guard case .unresolved = first.spatialInfo.resolution else {
    Issue.record("Floating Word geometry must remain relative to the pagination result.")
    return
  }

  guard case .chart(let chartReference) = chartDrawing.graphicContent else {
    Issue.record("Expected a typed chart relationship.")
    return
  }
  #expect(chartReference.part.name.rawValue == "/word/charts/chart1.xml")
  #expect(try chartReference.attachment.url().isFileURL)

  guard case .diagram(let diagram) = diagramDrawing.graphicContent else {
    Issue.record("Expected typed SmartArt relationship roles.")
    return
  }
  #expect(diagram.data?.part?.name.rawValue == "/word/diagrams/data1.xml")
  #expect(diagram.layout?.part?.name.rawValue == "/word/diagrams/layout1.xml")
  #expect(diagram.quickStyle?.part?.name.rawValue == "/word/diagrams/quickStyle1.xml")
  #expect(diagram.colors?.part?.name.rawValue == "/word/diagrams/colors1.xml")
  #expect(try diagram.data?.url().isFileURL == true)
  #expect(try pictureDrawing.picture?.primaryImage?.url().isFileURL == true)
}

@Test func wordEquationsEmbeddedPackagesAndEffectiveVMLRemainInspectable() throws {
  let document = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Complex01.docx")
  )
  let equation = try #require(document.body.equations.first)
  let embeddedObject = try #require(document.body.embeddedObjects.first)
  let legacyShape = try #require(document.body.legacyShapes.first)
  let equationPosition = try #require(equation.inlinePosition)

  #expect(document.body.equations.count == 1)
  #expect(equation.isDisplay)
  #expect(equation.text.hasPrefix("x+a"))
  #expect(equation.text.contains("k=0"))
  #expect(equation.text.hasSuffix("n-k"))
  #expect(
    document.body.inlineContent(in: equationPosition.paragraph).contains { content in
      guard case .equation(let candidate) = content else { return false }
      return candidate.sourceOrder == equation.sourceOrder
    })
  #expect(document.body.embeddedObjects.count == 1)
  #expect(embeddedObject.relationshipID.rawValue == "rId19")
  #expect(embeddedObject.type == "Embed")
  #expect(embeddedObject.programIdentifier == "Excel.Sheet.12")
  #expect(embeddedObject.drawingAspect == "Content")
  #expect(
    embeddedObject.attachment.part?.name.rawValue
      == "/word/embeddings/Microsoft_Excel_Worksheet2.xlsx")
  #expect(try embeddedObject.attachment.url().isFileURL)
  #expect(document.body.legacyShapes.count == 1)
  #expect(legacyShape.identifier == "_x0000_i1025")
  #expect(legacyShape.spatialInfo.frame?.size.width == 361.9)
  #expect(legacyShape.spatialInfo.frame?.size.height == 145.9)
  #expect(document.body.alternativeFormatImports.isEmpty)
}

// Upstream fixture generated by DocumentFormat.OpenXml.Tests/GeneratedClass002.cs.
@Test func strictWordDocumentRetainsSemanticsAndCanonicalizesRelationships() throws {
  let document = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Strict01.docx")
  )
  let transitional = try OfficeWordDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Complex01.docx")
  )
  let relationships = try document.document.package.relationships(
    from: .part(document.document.mainPart.name)
  )
  let strictChart = try #require(
    relationships.first {
      $0.type.isEquivalent(to: .chart)
    })
  let chartDrawing = try #require(document.body.drawings.first { $0.kind == .chart })
  let imageDrawing = try #require(document.body.drawings.first { $0.picture != nil })
  let firstSection = try #require(document.body.sections.first)
  let settings = try #require(document.settings)

  #expect(document.document.conformance == .strict)
  #expect(!document.body.paragraphs.isEmpty)
  #expect(!document.body.tables.isEmpty)
  #expect(document.headers.count == 9)
  #expect(document.footers.count == 9)
  #expect(document.footnotes?.note(identifiedBy: 1)?.text == " This is a footnote.")
  #expect(document.endnotes?.note(identifiedBy: 1)?.text == " This is an endnote.")
  #expect(document.numbering?.instances.count == 4)
  #expect(document.body.drawings.count == 18)
  #expect(firstSection.pageWidth?.points == 612)
  #expect(firstSection.pageHeight?.points == 792)
  #expect(firstSection.margins?.top?.points == 72)
  #expect(firstSection.margins?.header?.points == 36)
  #expect(firstSection.columns?.spacing?.points == 36)
  #expect(firstSection.spatialInfo.resolution == .exact)
  #expect(settings.view == "web")
  #expect(settings.zoomPercentage == 120)
  #expect(settings.defaultTabStop?.points == 36)
  #expect(settings.characterSpacingControl == "doNotCompress")
  #expect(settings.themeLanguage == "en-US")
  #expect(settings.compatibilityMode == 15)
  #expect(settings.decimalSymbol == ".")
  #expect(settings.listSeparator == ",")
  #expect(settings.sourcePart.name.rawValue == "/word/settings.xml")
  #expect(
    document.styleDefaults?.runProperties.fontSize
      == transitional.styleDefaults?.runProperties.fontSize)
  #expect(
    document.styleDefaults?.paragraphProperties.spacingAfter
      == transitional.styleDefaults?.paragraphProperties.spacingAfter)
  #expect(document.body.sections.map(\.pageWidth) == transitional.body.sections.map(\.pageWidth))
  #expect(document.body.sections.map(\.pageHeight) == transitional.body.sections.map(\.pageHeight))
  #expect(strictChart.type.rawValue.hasPrefix("http://purl.oclc.org/ooxml/"))
  #expect(strictChart.type != .chart)
  #expect(strictChart.type.isEquivalent(to: .chart))
  #expect(try chartDrawing.attachments.first?.url().isFileURL == true)
  #expect(try imageDrawing.picture?.primaryImage?.url().isFileURL == true)
}

private func makeAltChunkDocument() throws -> URL {
  try makeSyntheticOfficePackage(
    entries: [
      "[Content_Types].xml": """
      <?xml version="1.0" encoding="UTF-8"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Default Extension="html" ContentType="text/html"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>
      """,
      "_rels/.rels": """
      <?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
      """,
      "word/document.xml": """
      <?xml version="1.0" encoding="UTF-8"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p><w:r><w:t>Before</w:t></w:r></w:p>
          <w:altChunk r:id="rIdAltChunk"/>
          <w:p><w:r><w:t>After</w:t></w:r></w:p>
        </w:body>
      </w:document>
      """,
      "word/_rels/document.xml.rels": """
      <?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdAltChunk" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="imported.html"/>
      </Relationships>
      """,
      "word/imported.html": "<p>Imported content</p>",
    ],
    pathExtension: "docx"
  )
}
