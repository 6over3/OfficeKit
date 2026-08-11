import Foundation

package final class WordDocumentParser {
  private static let wordNamespace =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  private static let relationshipNamespace =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  private static let word2010Namespace =
    "http://schemas.microsoft.com/office/word/2010/wordml"

  private enum TextTarget {
    case text
    case instruction
  }

  private struct ParagraphBuilder {
    let index: Int
    let identifier: String?
    let startDepth: Int
    var properties = WordParagraphPropertiesBuilder()
    var runs: [OfficeWordRun] = []
    var hyperlinks: [OfficeWordHyperlink] = []
  }

  private struct RunBuilder {
    let startDepth: Int
    let revision: OfficeWordRevision?
    var text = ""
    var instruction = ""
    var properties = WordRunPropertiesBuilder()
    var footnoteIdentifier: Int64?
    var endnoteIdentifier: Int64?

    var value: OfficeWordRun {
      OfficeWordRun(
        text: text,
        fieldInstruction: instruction.isEmpty ? nil : instruction,
        properties: properties.value,
        footnoteIdentifier: footnoteIdentifier,
        endnoteIdentifier: endnoteIdentifier,
        revision: revision
      )
    }
  }

  private struct HyperlinkBuilder {
    let startDepth: Int
    let startRunIndex: Int
    let anchor: String?
    let tooltip: String?
    let relationshipID: OfficeRelationshipID?
  }

  private struct CellBuilder {
    let index: Int
    let startDepth: Int
    var propertiesDepth: Int?
    var bordersDepth: Int?
    var gridSpan = 1
    var verticalMerge: String?
    var preferredWidth: OfficeWordWidth?
    var verticalAlignment: String?
    var borders = BordersBuilder()
    var shading: OfficeWordShading?
    var blocks: [OfficeWordBlock] = []
  }

  private struct RowBuilder {
    let index: Int
    let startDepth: Int
    var propertiesDepth: Int?
    var height: OfficeLength?
    var heightRule: String?
    var preventsPageSplit: Bool?
    var repeatsAsHeader: Bool?
    var cells: [OfficeWordTableCell] = []
  }

  private struct TableBuilder {
    let index: Int
    let startDepth: Int
    var propertiesDepth: Int?
    var gridDepth: Int?
    var bordersDepth: Int?
    var columnWidths: [OfficeLength] = []
    var styleIdentifier: String?
    var preferredWidth: OfficeWordWidth?
    var indentation: OfficeWordWidth?
    var alignment: String?
    var layout: String?
    var borders = BordersBuilder()
    var shading: OfficeWordShading?
    var rows: [OfficeWordTableRow] = []
  }

  private struct BordersBuilder {
    var top: OfficeWordBorder?
    var leading: OfficeWordBorder?
    var bottom: OfficeWordBorder?
    var trailing: OfficeWordBorder?
    var insideHorizontal: OfficeWordBorder?
    var insideVertical: OfficeWordBorder?

    var value: OfficeWordBorders {
      OfficeWordBorders(
        top: top,
        leading: leading,
        bottom: bottom,
        trailing: trailing,
        insideHorizontal: insideHorizontal,
        insideVertical: insideVertical
      )
    }

    mutating func set(_ border: OfficeWordBorder, for localName: String) {
      switch localName {
      case "top": top = border
      case "left", "start": leading = border
      case "bottom": bottom = border
      case "right", "end": trailing = border
      case "insideH": insideHorizontal = border
      case "insideV": insideVertical = border
      default: break
      }
    }
  }

  private struct SectionBuilder {
    let startDepth: Int
    var type: String?
    var pageWidth: OfficeLength?
    var pageHeight: OfficeLength?
    var orientation: String?
    var margins: OfficeWordPageMargins?
    var columns: OfficeWordColumns?
    var hasTitlePage: Bool?
    var verticalAlignment: String?
    var pageNumberStart: Int?
    var headerReferences: [OfficeWordHeaderFooterReference] = []
    var footerReferences: [OfficeWordHeaderFooterReference] = []
  }

  private struct CommentAnchorBuilder {
    var start: OfficeWordTextPosition?
    var end: OfficeWordTextPosition?
    var referenceRun: OfficeWordTextPosition?
  }

  private struct RevisionContext {
    let startDepth: Int
    let revision: OfficeWordRevision
  }

  private struct FieldBuilder {
    let order: Int
    let kind: OfficeWordFieldKind
    let startDepth: Int?
    let start: OfficeWordTextPosition
    var instruction: String
    var separator: OfficeWordTextPosition?
  }

  private struct OrderedField {
    let order: Int
    let field: OfficeWordField
  }

  private struct BookmarkBuilder {
    let name: String?
    let start: OfficeWordBookmarkPosition
    var end: OfficeWordBookmarkPosition?
  }

  private struct ContentControlBuilder {
    let order: Int
    let startDepth: Int
    let scope: OfficeWordContentControlScope
    var propertiesDepth: Int?
    var contentDepth: Int?
    var identifier: Int64?
    var alias: String?
    var tag: String?
    var kind = "richText"
    var isChecked: Bool?
    var items: [OfficeWordContentControlItem] = []
    var binding: OfficeWordContentControlBinding?
    var text = ""
  }

  private struct OrderedContentControl {
    let order: Int
    let value: OfficeWordContentControl
  }

  private let part: OfficePart
  private let package: OfficePackage
  private let relationshipsByID: [OfficeRelationshipID: OfficeRelationship]
  private let drawingParser: WordDrawingParser
  private let auxiliaryContentParser: WordAuxiliaryContentParser
  private var depth = 0
  private var sourceOrder = 0
  private var paragraphPropertiesDepth: Int?
  private var runPropertiesDepth: Int?
  private var textDepth: Int?
  private var textTarget: TextTarget?
  private var paragraph: ParagraphBuilder?
  private var run: RunBuilder?
  private var hyperlink: HyperlinkBuilder?
  private var tableStack: [TableBuilder] = []
  private var rowStack: [RowBuilder] = []
  private var cellStack: [CellBuilder] = []
  private var section: SectionBuilder?
  private var blocks: [OfficeWordBlock] = []
  private var sections: [OfficeWordSection] = []
  private var commentAnchorOrder: [Int64] = []
  private var commentAnchorsByID: [Int64: CommentAnchorBuilder] = [:]
  private var revisions: [RevisionContext] = []
  private var fieldStack: [FieldBuilder] = []
  private var completedFields: [OrderedField] = []
  private var nextFieldOrder = 0
  private var bookmarkOrder: [Int64] = []
  private var bookmarksByID: [Int64: BookmarkBuilder] = [:]
  private var contentControlStack: [ContentControlBuilder] = []
  private var completedContentControls: [OrderedContentControl] = []
  private var nextContentControlOrder = 0

  package init(
    part: OfficePart,
    package: OfficePackage,
    relationships: [OfficeRelationship]
  ) {
    self.part = part
    self.package = package
    self.relationshipsByID = Dictionary(
      relationships.map { ($0.id, $0) },
      uniquingKeysWith: { first, _ in first }
    )
    self.drawingParser = WordDrawingParser(
      part: part,
      package: package,
      relationships: relationships
    )
    self.auxiliaryContentParser = WordAuxiliaryContentParser(
      part: part,
      package: package,
      relationships: relationships
    )
  }

  package static func parse(
    part: OfficePart,
    package: OfficePackage,
    relationships: [OfficeRelationship]
  ) throws -> OfficeWordBody {
    let parser = WordDocumentParser(part: part, package: package, relationships: relationships)
    try package.parseXML(in: part, compatibility: .commonOffice, parser.consume)
    return parser.body
  }

  package var body: OfficeWordBody {
    OfficeWordBody(
      blocks: blocks,
      sections: sections,
      commentAnchors: commentAnchors,
      fields: fields,
      bookmarks: bookmarks,
      contentControls: contentControls,
      drawings: drawingParser.drawings,
      equations: auxiliaryContentParser.equations,
      alternativeFormatImports: auxiliaryContentParser.alternativeFormatImports,
      embeddedObjects: auxiliaryContentParser.embeddedObjects,
      legacyShapes: auxiliaryContentParser.legacyShapes,
      sourcePart: part
    )
  }

  package func consume(_ event: OfficeXMLEvent) throws {
    sourceOrder += 1
    let inlinePosition = currentInlinePosition()
    try drawingParser.consume(
      event,
      sourceOrder: sourceOrder,
      inlinePosition: inlinePosition
    )
    try auxiliaryContentParser.consume(
      event,
      sourceOrder: sourceOrder,
      inlinePosition: inlinePosition
    )
    switch event {
    case .startElement(let name, let attributes, _, _):
      depth += 1
      if name.namespaceURI == Self.word2010Namespace {
        consumeWord2010Start(name.localName, attributes: attributes)
        return
      }
      guard name.namespaceURI == Self.wordNamespace else { return }
      try consumeStart(name.localName, attributes: attributes)
    case .text(let text, _):
      guard textDepth != nil else { return }
      switch textTarget {
      case .text:
        run?.text.append(text)
        appendToContentControls(text)
      case .instruction:
        run?.instruction.append(text)
        if let index = fieldStack.indices.last(where: { fieldStack[$0].kind == .complex }) {
          fieldStack[index].instruction.append(text)
        }
      case nil: break
      }
    case .endElement(let name, _):
      if name.namespaceURI == Self.wordNamespace {
        try consumeEnd(name.localName)
      }
      depth -= 1
    case .startDocument, .endDocument:
      break
    }
  }

  private func consumeStart(
    _ localName: String,
    attributes: [OfficeXMLAttribute]
  ) throws {
    switch localName {
    case "tbl":
      let index = cellStack.last?.blocks.count ?? blocks.count
      tableStack.append(TableBuilder(index: index, startDepth: depth))
    case "tblPr":
      if tableStack.last?.startDepth == depth - 1 {
        tableStack[tableStack.count - 1].propertiesDepth = depth
      }
    case "tblGrid":
      if tableStack.last?.startDepth == depth - 1 {
        tableStack[tableStack.count - 1].gridDepth = depth
      }
    case "tblBorders":
      if tableStack.last?.propertiesDepth != nil {
        tableStack[tableStack.count - 1].bordersDepth = depth
      }
    case "tr":
      guard !tableStack.isEmpty else { return }
      rowStack.append(
        RowBuilder(index: tableStack[tableStack.count - 1].rows.count, startDepth: depth))
    case "trPr":
      if rowStack.last?.startDepth == depth - 1 {
        rowStack[rowStack.count - 1].propertiesDepth = depth
      }
    case "tc":
      guard !rowStack.isEmpty else { return }
      cellStack.append(
        CellBuilder(index: rowStack[rowStack.count - 1].cells.count, startDepth: depth))
    case "tcPr":
      if cellStack.last?.startDepth == depth - 1 {
        cellStack[cellStack.count - 1].propertiesDepth = depth
      }
    case "tcBorders":
      if cellStack.last?.propertiesDepth != nil {
        cellStack[cellStack.count - 1].bordersDepth = depth
      }
    case "p":
      guard paragraph == nil else { return }
      let index =
        cellStack.last?.blocks.reduce(into: 0) { count, block in
          if case .paragraph = block { count += 1 }
        } ?? blocks.count
      paragraph = ParagraphBuilder(
        index: index,
        identifier: wordAttribute("paraId", in: attributes),
        startDepth: depth
      )
    case "pPr":
      guard paragraph != nil else { return }
      paragraphPropertiesDepth = depth
    case "r":
      guard paragraph != nil, run == nil else { return }
      run = RunBuilder(startDepth: depth, revision: revisions.last?.revision)
    case "rPr":
      guard run != nil else { return }
      runPropertiesDepth = depth
    case "t", "delText":
      beginText(.text)
    case "instrText", "delInstrText":
      beginText(.instruction)
    case "tab":
      run?.text.append("\t")
      appendToContentControls("\t")
    case "br", "cr":
      run?.text.append("\n")
      appendToContentControls("\n")
    case "noBreakHyphen":
      run?.text.append("\u{2011}")
      appendToContentControls("\u{2011}")
    case "softHyphen":
      run?.text.append("\u{00AD}")
      appendToContentControls("\u{00AD}")
    case "sdt":
      beginContentControl()
    case "sdtPr":
      if !contentControlStack.isEmpty {
        contentControlStack[contentControlStack.count - 1].propertiesDepth = depth
      }
    case "sdtContent":
      if !contentControlStack.isEmpty {
        contentControlStack[contentControlStack.count - 1].contentDepth = depth
      }
    case "hyperlink":
      beginHyperlink(attributes)
    case "ins":
      beginRevision(.insertion, attributes: attributes)
    case "del":
      beginRevision(.deletion, attributes: attributes)
    case "moveFrom":
      beginRevision(.moveFrom, attributes: attributes)
    case "moveTo":
      beginRevision(.moveTo, attributes: attributes)
    case "fldSimple":
      try beginSimpleField(attributes)
    case "fldChar":
      try consumeFieldCharacter(attributes)
    case "bookmarkStart":
      try beginBookmark(attributes)
    case "bookmarkEnd":
      try endBookmark(attributes)
    case "commentRangeStart":
      try recordCommentMarker(.start, attributes: attributes)
    case "commentRangeEnd":
      try recordCommentMarker(.end, attributes: attributes)
    case "commentReference":
      try recordCommentMarker(.reference, attributes: attributes)
    case "sectPr":
      guard section == nil else { return }
      section = SectionBuilder(startDepth: depth)
    default:
      consumeContentControlProperty(localName, attributes: attributes)
      try consumeProperty(localName, attributes: attributes)
    }
  }

  private func consumeEnd(_ localName: String) throws {
    if textDepth == depth,
      localName == "t" || localName == "delText" || localName == "instrText"
        || localName == "delInstrText"
    {
      textDepth = nil
      textTarget = nil
    }
    if runPropertiesDepth == depth, localName == "rPr" { runPropertiesDepth = nil }
    if paragraphPropertiesDepth == depth, localName == "pPr" { paragraphPropertiesDepth = nil }
    if localName == "tblBorders", tableStack.last?.bordersDepth == depth {
      tableStack[tableStack.count - 1].bordersDepth = nil
    }
    if localName == "tblGrid", tableStack.last?.gridDepth == depth {
      tableStack[tableStack.count - 1].gridDepth = nil
    }
    if localName == "tblPr", tableStack.last?.propertiesDepth == depth {
      tableStack[tableStack.count - 1].propertiesDepth = nil
    }
    if localName == "trPr", rowStack.last?.propertiesDepth == depth {
      rowStack[rowStack.count - 1].propertiesDepth = nil
    }
    if localName == "tcBorders", cellStack.last?.bordersDepth == depth {
      cellStack[cellStack.count - 1].bordersDepth = nil
    }
    if localName == "tcPr", cellStack.last?.propertiesDepth == depth {
      cellStack[cellStack.count - 1].propertiesDepth = nil
    }
    if localName == "sdtPr", contentControlStack.last?.propertiesDepth == depth {
      contentControlStack[contentControlStack.count - 1].propertiesDepth = nil
    }
    if localName == "sdtContent", contentControlStack.last?.contentDepth == depth {
      contentControlStack[contentControlStack.count - 1].contentDepth = nil
    }

    if localName == "r", run?.startDepth == depth, let completed = run {
      paragraph?.runs.append(completed.value)
      run = nil
      return
    }
    if localName == "hyperlink", hyperlink?.startDepth == depth, let completed = hyperlink {
      finishHyperlink(completed)
      hyperlink = nil
      return
    }
    if localName == "fldSimple",
      let index = fieldStack.indices.last(where: { fieldStack[$0].startDepth == depth })
    {
      finishField(at: index, end: currentTextPosition())
    }
    if localName == "p", paragraph?.startDepth == depth, let completed = paragraph {
      let value = makeParagraph(completed)
      if !cellStack.isEmpty {
        cellStack[cellStack.count - 1].blocks.append(.paragraph(value))
      } else if tableStack.isEmpty {
        blocks.append(.paragraph(value))
      }
      paragraph = nil
      return
    }
    if localName == "tc", cellStack.last?.startDepth == depth {
      let completed = cellStack.removeLast()
      guard !rowStack.isEmpty else { return }
      rowStack[rowStack.count - 1].cells.append(
        OfficeWordTableCell(
          index: completed.index,
          gridSpan: completed.gridSpan,
          verticalMerge: completed.verticalMerge,
          preferredWidth: completed.preferredWidth,
          verticalAlignment: completed.verticalAlignment,
          borders: completed.borders.value,
          shading: completed.shading,
          blocks: completed.blocks,
          text: completed.blocks.map(blockText).joined(separator: "\n")
        )
      )
      return
    }
    if localName == "tr", rowStack.last?.startDepth == depth {
      let completed = rowStack.removeLast()
      guard !tableStack.isEmpty else { return }
      tableStack[tableStack.count - 1].rows.append(
        OfficeWordTableRow(
          index: completed.index,
          cells: completed.cells,
          height: completed.height,
          heightRule: completed.heightRule,
          preventsPageSplit: completed.preventsPageSplit,
          repeatsAsHeader: completed.repeatsAsHeader
        )
      )
      return
    }
    if localName == "tbl", tableStack.last?.startDepth == depth {
      let completed = makeTable(tableStack.removeLast())
      if !cellStack.isEmpty {
        cellStack[cellStack.count - 1].blocks.append(.table(completed))
      } else {
        blocks.append(.table(completed))
      }
      return
    }
    if localName == "sectPr", section?.startDepth == depth, let completed = section {
      sections.append(makeSection(completed))
      section = nil
    }
    if localName == "sdt", contentControlStack.last?.startDepth == depth {
      finishContentControl()
    }
    if let index = revisions.indices.last(where: { revisions[$0].startDepth == depth }),
      revisionElement(localName) != nil
    {
      revisions.remove(at: index)
    }
  }

  private func consumeProperty(
    _ localName: String,
    attributes: [OfficeXMLAttribute]
  ) throws {
    if paragraphPropertiesDepth != nil {
      paragraph?.properties.consume(localName, attributes: attributes)
    }
    if runPropertiesDepth != nil {
      run?.properties.consume(localName, attributes: attributes)
    }
    if run != nil {
      switch localName {
      case "footnoteReference":
        run?.footnoteIdentifier = wordAttribute("id", in: attributes).flatMap(Int64.init)
      case "endnoteReference":
        run?.endnoteIdentifier = wordAttribute("id", in: attributes).flatMap(Int64.init)
      default: break
      }
    }
    if !tableStack.isEmpty {
      let index = tableStack.count - 1
      if tableStack[index].gridDepth != nil, localName == "gridCol",
        let width = wordAttribute("w", in: attributes).flatMap(wordTwipLength)
      {
        tableStack[index].columnWidths.append(width)
      }
      if tableStack[index].propertiesDepth != nil {
        switch localName {
        case "tblStyle":
          tableStack[index].styleIdentifier = wordAttribute("val", in: attributes)
        case "tblW":
          tableStack[index].preferredWidth = wordWidth(attributes)
        case "tblInd":
          tableStack[index].indentation = wordWidth(attributes)
        case "jc":
          tableStack[index].alignment = wordAttribute("val", in: attributes)
        case "tblLayout":
          tableStack[index].layout = wordAttribute("type", in: attributes)
        case "shd":
          tableStack[index].shading = wordShading(attributes)
        default:
          if tableStack[index].bordersDepth != nil,
            let border = wordBorder(localName, attributes: attributes)
          {
            tableStack[index].borders.set(border, for: localName)
          }
        }
      }
    }
    if !rowStack.isEmpty, rowStack[rowStack.count - 1].propertiesDepth != nil {
      let index = rowStack.count - 1
      switch localName {
      case "trHeight":
        rowStack[index].height = wordAttribute("val", in: attributes).flatMap(wordTwipLength)
        rowStack[index].heightRule = wordAttribute("hRule", in: attributes)
      case "cantSplit":
        rowStack[index].preventsPageSplit = wordOnOffValue(attributes)
      case "tblHeader":
        rowStack[index].repeatsAsHeader = wordOnOffValue(attributes)
      default:
        break
      }
    }
    if !cellStack.isEmpty {
      let index = cellStack.count - 1
      switch localName {
      case "gridSpan":
        if let span = wordAttribute("val", in: attributes).flatMap(Int.init), span > 0 {
          cellStack[index].gridSpan = span
        }
      case "vMerge":
        cellStack[index].verticalMerge = wordAttribute("val", in: attributes) ?? "continue"
      case "tcW" where cellStack[index].propertiesDepth != nil:
        cellStack[index].preferredWidth = wordWidth(attributes)
      case "vAlign" where cellStack[index].propertiesDepth != nil:
        cellStack[index].verticalAlignment = wordAttribute("val", in: attributes)
      case "shd" where cellStack[index].propertiesDepth != nil:
        cellStack[index].shading = wordShading(attributes)
      default: break
      }
      if cellStack[index].bordersDepth != nil,
        let border = wordBorder(localName, attributes: attributes)
      {
        cellStack[index].borders.set(border, for: localName)
      }
    }
    if section != nil {
      switch localName {
      case "type": section?.type = wordAttribute("val", in: attributes)
      case "pgSz":
        section?.pageWidth = wordAttribute("w", in: attributes).flatMap(wordTwipLength)
        section?.pageHeight = wordAttribute("h", in: attributes).flatMap(wordTwipLength)
        section?.orientation = wordAttribute("orient", in: attributes)
      case "pgMar":
        section?.margins = OfficeWordPageMargins(
          top: wordAttribute("top", in: attributes).flatMap(wordTwipLength),
          right: wordAttribute("right", in: attributes).flatMap(wordTwipLength),
          bottom: wordAttribute("bottom", in: attributes).flatMap(wordTwipLength),
          left: wordAttribute("left", in: attributes).flatMap(wordTwipLength),
          header: wordAttribute("header", in: attributes).flatMap(wordTwipLength),
          footer: wordAttribute("footer", in: attributes).flatMap(wordTwipLength),
          gutter: wordAttribute("gutter", in: attributes).flatMap(wordTwipLength)
        )
      case "cols":
        section?.columns = OfficeWordColumns(
          count: wordAttribute("num", in: attributes).flatMap(Int.init),
          spacing: wordAttribute("space", in: attributes).flatMap(wordTwipLength),
          hasSeparator: wordOnOffAttribute("sep", in: attributes),
          usesEqualWidths: wordOnOffAttribute("equalWidth", in: attributes),
          columns: []
        )
      case "col":
        if let columns = section?.columns {
          section?.columns = OfficeWordColumns(
            count: columns.count,
            spacing: columns.spacing,
            hasSeparator: columns.hasSeparator,
            usesEqualWidths: columns.usesEqualWidths,
            columns: columns.columns + [
              OfficeWordColumn(
                width: wordAttribute("w", in: attributes).flatMap(wordTwipLength),
                spacing: wordAttribute("space", in: attributes).flatMap(wordTwipLength)
              )
            ]
          )
        }
      case "titlePg": section?.hasTitlePage = wordOnOffValue(attributes)
      case "vAlign": section?.verticalAlignment = wordAttribute("val", in: attributes)
      case "pgNumType":
        section?.pageNumberStart = wordAttribute("start", in: attributes).flatMap(Int.init)
      case "headerReference":
        section?.headerReferences.append(
          try headerFooterReference(
            attributes,
            expectedType: .header,
            element: "w:headerReference"
          )
        )
      case "footerReference":
        section?.footerReferences.append(
          try headerFooterReference(
            attributes,
            expectedType: .footer,
            element: "w:footerReference"
          )
        )
      default: break
      }
    }
  }

  private func beginText(_ target: TextTarget) {
    guard run != nil else { return }
    textDepth = depth
    textTarget = target
  }

  private func beginRevision(
    _ kind: OfficeWordRevisionKind,
    attributes: [OfficeXMLAttribute]
  ) {
    let dateText = wordAttribute("date", in: attributes)
    let revision = OfficeWordRevision(
      kind: kind,
      identifier: wordAttribute("id", in: attributes).flatMap(Int64.init),
      author: wordAttribute("author", in: attributes),
      dateText: dateText,
      date: dateText.flatMap(ISO8601DateFormatter().date(from:))
    )
    revisions.append(RevisionContext(startDepth: depth, revision: revision))
  }

  private func beginSimpleField(_ attributes: [OfficeXMLAttribute]) throws {
    guard let position = currentTextPosition() else {
      throw OfficeKitError.invalidXML(
        part: part.name.rawValue,
        message: "w:fldSimple appears outside a paragraph."
      )
    }
    fieldStack.append(
      FieldBuilder(
        order: nextFieldOrder,
        kind: .simple,
        startDepth: depth,
        start: position,
        instruction: wordAttribute("instr", in: attributes) ?? ""
      )
    )
    nextFieldOrder += 1
  }

  private func consumeFieldCharacter(_ attributes: [OfficeXMLAttribute]) throws {
    guard let type = wordAttribute("fldCharType", in: attributes),
      let position = currentTextPosition() else {
      throw OfficeKitError.invalidXML(
        part: part.name.rawValue,
        message: "w:fldChar requires w:fldCharType inside a paragraph."
      )
    }
    switch type {
    case "begin":
      fieldStack.append(
        FieldBuilder(
          order: nextFieldOrder,
          kind: .complex,
          startDepth: nil,
          start: position,
          instruction: ""
        )
      )
      nextFieldOrder += 1
    case "separate":
      if let index = fieldStack.indices.last(where: { fieldStack[$0].kind == .complex }) {
        fieldStack[index].separator = position
      }
    case "end":
      guard let index = fieldStack.indices.last(where: { fieldStack[$0].kind == .complex }) else {
        return
      }
      finishField(at: index, end: position)
    default:
      break
    }
  }

  private func finishField(at index: Int, end: OfficeWordTextPosition?) {
    let completed = fieldStack.remove(at: index)
    completedFields.append(
      OrderedField(
        order: completed.order,
        field: OfficeWordField(
          kind: completed.kind,
          instruction: completed.instruction,
          start: completed.start,
          separator: completed.separator,
          end: end,
          sourcePart: part
        )
      )
    )
  }

  private var fields: [OfficeWordField] {
    completedFields.sorted { $0.order < $1.order }.map(\.field)
  }

  private func beginBookmark(_ attributes: [OfficeXMLAttribute]) throws {
    guard let identifierText = wordAttribute("id", in: attributes),
      let identifier = Int64(identifierText) else {
      throw OfficeKitError.invalidXML(
        part: part.name.rawValue,
        message: "w:bookmarkStart requires a numeric w:id attribute."
      )
    }
    if bookmarksByID[identifier] == nil { bookmarkOrder.append(identifier) }
    bookmarksByID[identifier] = BookmarkBuilder(
      name: wordAttribute("name", in: attributes),
      start: currentBookmarkPosition()
    )
  }

  private func endBookmark(_ attributes: [OfficeXMLAttribute]) throws {
    guard let identifierText = wordAttribute("id", in: attributes),
      let identifier = Int64(identifierText) else {
      throw OfficeKitError.invalidXML(
        part: part.name.rawValue,
        message: "w:bookmarkEnd requires a numeric w:id attribute."
      )
    }
    bookmarksByID[identifier]?.end = currentBookmarkPosition()
  }

  private func currentBookmarkPosition() -> OfficeWordBookmarkPosition {
    if let position = currentTextPosition() { return .text(position) }
    return .bodyBlock(blocks.count)
  }

  private var bookmarks: [OfficeWordBookmark] {
    bookmarkOrder.compactMap { identifier in
      guard let bookmark = bookmarksByID[identifier] else { return nil }
      return OfficeWordBookmark(
        identifier: identifier,
        name: bookmark.name,
        start: bookmark.start,
        end: bookmark.end,
        sourcePart: part
      )
    }
  }

  private func beginContentControl() {
    let scope: OfficeWordContentControlScope
    if paragraph != nil {
      scope = .inline
    } else if !cellStack.isEmpty {
      scope = .cell
    } else if !tableStack.isEmpty {
      scope = .row
    } else {
      scope = .block
    }
    contentControlStack.append(
      ContentControlBuilder(
        order: nextContentControlOrder,
        startDepth: depth,
        scope: scope
      )
    )
    nextContentControlOrder += 1
  }

  private func consumeContentControlProperty(
    _ localName: String,
    attributes: [OfficeXMLAttribute]
  ) {
    guard !contentControlStack.isEmpty,
      contentControlStack[contentControlStack.count - 1].propertiesDepth != nil else { return }
    let index = contentControlStack.count - 1
    switch localName {
    case "alias": contentControlStack[index].alias = wordAttribute("val", in: attributes)
    case "tag": contentControlStack[index].tag = wordAttribute("val", in: attributes)
    case "id":
      contentControlStack[index].identifier = wordAttribute("val", in: attributes).flatMap(
        Int64.init)
    case "text", "picture", "comboBox", "dropDownList", "date", "group",
      "repeatingSection", "repeatingSectionItem":
      contentControlStack[index].kind = localName
    case "listItem":
      contentControlStack[index].items.append(
        OfficeWordContentControlItem(
          displayText: wordAttribute("displayText", in: attributes),
          value: wordAttribute("value", in: attributes)
        )
      )
    case "dataBinding":
      contentControlStack[index].binding = OfficeWordContentControlBinding(
        storeItemIdentifier: wordAttribute("storeItemID", in: attributes),
        xpath: wordAttribute("xpath", in: attributes),
        prefixMappings: wordAttribute("prefixMappings", in: attributes)
      )
    default:
      break
    }
  }

  private func consumeWord2010Start(
    _ localName: String,
    attributes: [OfficeXMLAttribute]
  ) {
    guard !contentControlStack.isEmpty,
      contentControlStack[contentControlStack.count - 1].propertiesDepth != nil else { return }
    let index = contentControlStack.count - 1
    switch localName {
    case "checkbox": contentControlStack[index].kind = "checkbox"
    case "checked": contentControlStack[index].isChecked = wordOnOffValue(attributes)
    default: break
    }
  }

  private func appendToContentControls(_ text: String) {
    for index in contentControlStack.indices where contentControlStack[index].contentDepth != nil {
      contentControlStack[index].text.append(text)
    }
  }

  private func finishContentControl() {
    let completed = contentControlStack.removeLast()
    completedContentControls.append(
      OrderedContentControl(
        order: completed.order,
        value: OfficeWordContentControl(
          identifier: completed.identifier,
          alias: completed.alias,
          tag: completed.tag,
          kind: completed.kind,
          isChecked: completed.isChecked,
          items: completed.items,
          binding: completed.binding,
          scope: completed.scope,
          text: completed.text,
          sourcePart: part
        )
      )
    )
  }

  private var contentControls: [OfficeWordContentControl] {
    completedContentControls.sorted { $0.order < $1.order }.map(\.value)
  }

  private enum CommentMarkerKind {
    case start
    case end
    case reference
  }

  private func recordCommentMarker(
    _ kind: CommentMarkerKind,
    attributes: [OfficeXMLAttribute]
  ) throws {
    guard let identifierText = wordAttribute("id", in: attributes),
      let identifier = Int64(identifierText) else {
      throw OfficeKitError.invalidXML(
        part: part.name.rawValue,
        message: "A Word comment marker has an invalid identifier."
      )
    }
    if commentAnchorsByID[identifier] == nil {
      commentAnchorOrder.append(identifier)
      commentAnchorsByID[identifier] = CommentAnchorBuilder()
    }
    guard let position = currentTextPosition() else { return }
    switch kind {
    case .start: commentAnchorsByID[identifier]?.start = position
    case .end: commentAnchorsByID[identifier]?.end = position
    case .reference: commentAnchorsByID[identifier]?.referenceRun = position
    }
  }

  private func currentTextPosition() -> OfficeWordTextPosition? {
    guard let paragraph else { return nil }
    let location: OfficeWordParagraphLocation
    if tableStack.count == 1, let table = tableStack.last, let row = rowStack.last,
      let cell = cellStack.last
    {
      location = .table(
        blockIndex: table.index,
        rowIndex: row.index,
        cellIndex: cell.index,
        paragraphIndex: paragraph.index
      )
    } else if !tableStack.isEmpty,
      tableStack.count == rowStack.count,
      tableStack.count == cellStack.count
    {
      let path = tableStack.indices.map { index in
        OfficeWordTableCellLocation(
          tableBlockIndex: tableStack[index].index,
          rowIndex: rowStack[index].index,
          cellIndex: cellStack[index].index
        )
      }
      location = .nestedTable(path: path, paragraphIndex: paragraph.index)
    } else {
      location = .body(blockIndex: paragraph.index)
    }
    return OfficeWordTextPosition(paragraph: location, runIndex: paragraph.runs.count)
  }

  private func currentInlinePosition() -> OfficeWordInlinePosition? {
    guard let position = currentTextPosition() else { return nil }
    return OfficeWordInlinePosition(
      paragraph: position.paragraph,
      runIndex: position.runIndex,
      characterOffset: run?.text.count ?? 0
    )
  }

  private var commentAnchors: [OfficeWordCommentAnchor] {
    commentAnchorOrder.compactMap { identifier in
      guard let anchor = commentAnchorsByID[identifier] else { return nil }
      return OfficeWordCommentAnchor(
        identifier: identifier,
        start: anchor.start,
        end: anchor.end,
        referenceRun: anchor.referenceRun,
        spatialInfo: OfficeSpatialInfo(
          coordinateSpace: .character,
          geometrySourcePart: part.name,
          frame: nil,
          resolution: .unresolved(
            reason: "Word comment marker placement depends on text layout and pagination."
          )
        ),
        sourcePart: part
      )
    }
  }

  private func beginHyperlink(_ attributes: [OfficeXMLAttribute]) {
    guard let paragraph, hyperlink == nil else { return }
    let rawRelationshipID = wordAttribute(
      "id",
      namespaceURI: Self.relationshipNamespace,
      in: attributes
    )
    hyperlink = HyperlinkBuilder(
      startDepth: depth,
      startRunIndex: paragraph.runs.count,
      anchor: wordAttribute("anchor", in: attributes),
      tooltip: wordAttribute("tooltip", in: attributes),
      relationshipID: rawRelationshipID.map(OfficeRelationshipID.init(rawValue:))
    )
  }

  private func finishHyperlink(_ builder: HyperlinkBuilder) {
    guard let paragraph else { return }
    let relationship = builder.relationshipID.flatMap { relationshipsByID[$0] }
    self.paragraph?.hyperlinks.append(
      OfficeWordHyperlink(
        runRange: builder.startRunIndex..<paragraph.runs.count,
        anchor: builder.anchor,
        tooltip: builder.tooltip,
        relationshipID: builder.relationshipID,
        attachment: relationship.map(package.attachment(referencedBy:))
      )
    )
  }

  private func makeParagraph(_ builder: ParagraphBuilder) -> OfficeWordParagraph {
    OfficeWordParagraph(
      index: builder.index,
      location: currentTextPosition()?.paragraph ?? .body(blockIndex: builder.index),
      identifier: builder.identifier,
      properties: builder.properties.value,
      runs: builder.runs,
      hyperlinks: builder.hyperlinks,
      text: builder.runs.filter { OfficeWordRevisionView.final.includes($0.revision) }
        .map(\.text).joined(),
      spatialInfo: OfficeSpatialInfo(
        coordinateSpace: .paragraph,
        geometrySourcePart: part.name,
        frame: nil,
        resolution: .unresolved(
          reason: "Word paragraph placement depends on pagination, fonts, and layout settings."
        )
      ),
      sourcePart: part
    )
  }

  private func makeTable(_ builder: TableBuilder) -> OfficeWordTable {
    OfficeWordTable(
      index: builder.index,
      rows: builder.rows,
      columnWidths: builder.columnWidths,
      styleIdentifier: builder.styleIdentifier,
      preferredWidth: builder.preferredWidth,
      indentation: builder.indentation,
      alignment: builder.alignment,
      layout: builder.layout,
      borders: builder.borders.value,
      shading: builder.shading,
      text: builder.rows.flatMap(\.cells).map(\.text).joined(separator: "\n"),
      spatialInfo: OfficeSpatialInfo(
        coordinateSpace: .page,
        geometrySourcePart: part.name,
        frame: nil,
        resolution: .unresolved(
          reason: "Word table placement depends on pagination, styles, fonts, and surrounding flow."
        )
      ),
      sourcePart: part
    )
  }

  private func blockText(_ block: OfficeWordBlock) -> String {
    switch block {
    case .paragraph(let paragraph): paragraph.text
    case .table(let table): table.text
    }
  }

  private func wordWidth(_ attributes: [OfficeXMLAttribute]) -> OfficeWordWidth? {
    guard let type = wordAttribute("type", in: attributes),
      let value = wordAttribute("w", in: attributes).flatMap(Int64.init) else { return nil }
    return OfficeWordWidth(type: type, value: value)
  }

  private func wordShading(_ attributes: [OfficeXMLAttribute]) -> OfficeWordShading {
    OfficeWordShading(
      pattern: wordAttribute("val", in: attributes),
      color: wordAttribute("color", in: attributes),
      fill: wordAttribute("fill", in: attributes)
    )
  }

  private func wordBorder(
    _ localName: String,
    attributes: [OfficeXMLAttribute]
  ) -> OfficeWordBorder? {
    let borderNames: Set<String> = [
      "top", "left", "start", "bottom", "right", "end", "insideH", "insideV",
    ]
    guard borderNames.contains(localName), let style = wordAttribute("val", in: attributes) else {
      return nil
    }
    return OfficeWordBorder(
      style: style,
      color: wordAttribute("color", in: attributes),
      sizeInEighthPoints: wordAttribute("sz", in: attributes).flatMap(Int.init),
      spacingInPoints: wordAttribute("space", in: attributes).flatMap(Int.init)
    )
  }

  private func makeSection(_ builder: SectionBuilder) -> OfficeWordSection {
    let frame: OfficeRect?
    if let width = builder.pageWidth, let height = builder.pageHeight {
      frame = OfficeRect(x: 0, y: 0, width: width.points, height: height.points)
    } else {
      frame = nil
    }
    return OfficeWordSection(
      type: builder.type,
      pageWidth: builder.pageWidth,
      pageHeight: builder.pageHeight,
      orientation: builder.orientation,
      margins: builder.margins,
      columns: builder.columns,
      hasTitlePage: builder.hasTitlePage,
      verticalAlignment: builder.verticalAlignment,
      pageNumberStart: builder.pageNumberStart,
      headerReferences: builder.headerReferences,
      footerReferences: builder.footerReferences,
      spatialInfo: OfficeSpatialInfo(
        coordinateSpace: .page,
        geometrySourcePart: part.name,
        frame: frame,
        resolution: frame == nil
          ? .unresolved(reason: "The Word section does not declare both page dimensions.")
          : .exact
      ),
      sourcePart: part
    )
  }

  private func headerFooterReference(
    _ attributes: [OfficeXMLAttribute],
    expectedType: OfficeRelationshipType,
    element: String
  ) throws -> OfficeWordHeaderFooterReference {
    guard
      let rawRelationshipID = wordAttribute(
        "id",
        namespaceURI: Self.relationshipNamespace,
        in: attributes
      ) else {
      throw OfficeKitError.invalidXML(
        part: part.name.rawValue,
        message: "\(element) requires an r:id attribute."
      )
    }
    let relationshipID = OfficeRelationshipID(rawValue: rawRelationshipID)
    guard let relationship = relationshipsByID[relationshipID],
      relationship.type.isEquivalent(to: expectedType) else {
      throw OfficeKitError.invalidPackage(
        "\(element) references a missing or incorrectly typed relationship \(rawRelationshipID)."
      )
    }
    guard let storyPart = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }
    return OfficeWordHeaderFooterReference(
      variant: headerFooterVariant(wordAttribute("type", in: attributes)),
      relationshipID: relationshipID,
      part: storyPart
    )
  }
}

package func wordAttribute(
  _ localName: String,
  namespaceURI: String? = nil,
  in attributes: [OfficeXMLAttribute]
) -> String? {
  attributes.first {
    $0.name.localName == localName
      && (namespaceURI == nil || $0.name.namespaceURI == namespaceURI)
  }?.value
}

package func wordOnOffValue(_ attributes: [OfficeXMLAttribute]) -> Bool? {
  guard let rawValue = wordAttribute("val", in: attributes) else { return true }
  if let value = OfficeValueDecoder.boolean(rawValue) { return value }
  if rawValue == "on" { return true }
  if rawValue == "off" { return false }
  return nil
}

package func wordOnOffAttribute(
  _ localName: String,
  in attributes: [OfficeXMLAttribute]
) -> Bool? {
  guard let rawValue = wordAttribute(localName, in: attributes) else { return nil }
  if let value = OfficeValueDecoder.boolean(rawValue) { return value }
  if rawValue == "on" { return true }
  if rawValue == "off" { return false }
  return nil
}

package func wordTwipLength(_ text: String) -> OfficeLength? {
  if let twips = Int64(text) {
    let (emu, overflow) = twips.multipliedReportingOverflow(by: 635)
    guard !overflow else { return nil }
    return OfficeLength(emu: emu)
  }

  return wordUniversalMeasure(text)
}

private func headerFooterVariant(_ text: String?) -> OfficeWordHeaderFooterVariant {
  switch text {
  case nil, "default": .default
  case "first": .firstPage
  case "even": .evenPages
  case .some(let value): .unknown(value)
  }
}

private func revisionElement(_ localName: String) -> OfficeWordRevisionKind? {
  switch localName {
  case "ins": .insertion
  case "del": .deletion
  case "moveFrom": .moveFrom
  case "moveTo": .moveTo
  default: nil
  }
}
