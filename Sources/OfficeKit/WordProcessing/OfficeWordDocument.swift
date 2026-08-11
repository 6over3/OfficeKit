import Foundation

/// A read-only WordprocessingML document.
///
/// Initialization streams the main document and its style definitions. Relationship-backed
/// resources remain lazy and can be accessed as URLs without first loading their bytes.
public struct OfficeWordDocument: Sendable {
  /// The opened Office document.
  public let document: OfficeDocument

  /// The authored block content and section declarations in the main document body.
  public let body: OfficeWordBody

  /// Named paragraph, character, table, and numbering styles in declaration order.
  public let styles: [OfficeWordStyle]

  /// Default paragraph and run formatting declared by the styles part.
  public let styleDefaults: OfficeWordStyleDefaults?

  /// Authored comments in declaration order.
  public let comments: [OfficeWordComment]

  /// Header stories directly related to the main document.
  public let headers: [OfficeWordStory]

  /// Footer stories directly related to the main document.
  public let footers: [OfficeWordStory]

  /// The document footnote collection, when present.
  public let footnotes: OfficeWordNoteCollection?

  /// The document endnote collection, when present.
  public let endnotes: OfficeWordNoteCollection?

  /// Numbering definitions used by list paragraphs, when present.
  public let numbering: OfficeWordNumbering?

  /// Document-wide WordprocessingML settings, when the package declares them.
  public let settings: OfficeWordSettings?

  /// Resources directly related to the main document part.
  public let attachments: [OfficeAttachment]

  /// Opens and parses a WordprocessingML document.
  public init(contentsOf url: URL, limits: OfficeParsingLimits = .standard) throws {
    try self.init(document: OfficeDocument(contentsOf: url, limits: limits))
  }

  /// Parses an already opened WordprocessingML document.
  public init(document: OfficeDocument) throws {
    guard document.kind == .wordProcessing else {
      throw OfficeKitError.invalidPackage("The Office document is not a word-processing document.")
    }
    let relationships = try document.package.relationships(from: .part(document.mainPart.name))
    let body = try WordDocumentParser.parse(
      part: document.mainPart,
      package: document.package,
      relationships: relationships
    )
    self.document = document
    self.body = body
    let styleSheet = try WordStyleParser.parse(
      documentPart: document.mainPart,
      package: document.package,
      relationships: relationships
    )
    self.styles = styleSheet.styles
    self.styleDefaults = styleSheet.defaults
    self.comments = try WordCommentParser.parse(
      documentPart: document.mainPart,
      package: document.package,
      relationships: relationships,
      anchors: body.commentAnchors
    )
    self.headers = try WordStoryParser.parse(
      kind: .header,
      package: document.package,
      relationships: relationships
    )
    self.footers = try WordStoryParser.parse(
      kind: .footer,
      package: document.package,
      relationships: relationships
    )
    self.footnotes = try WordNoteParser.parse(
      kind: .footnote,
      package: document.package,
      relationships: relationships
    )
    self.endnotes = try WordNoteParser.parse(
      kind: .endnote,
      package: document.package,
      relationships: relationships
    )
    self.numbering = try WordNumberingParser.parse(
      package: document.package,
      relationships: relationships
    )
    self.settings = try WordSettingsParser.parse(
      package: document.package,
      relationships: relationships
    )
    self.attachments = relationships.map(document.package.attachment(referencedBy:))
  }

  /// Resolves a paragraph's direct or style-inherited list definition.
  public func listInfo(for paragraph: OfficeWordParagraph) -> OfficeWordListInfo? {
    guard let numbering else { return nil }
    var numberingIdentifier = paragraph.properties.numberingIdentifier
    var numberingLevel = paragraph.properties.numberingLevel
    if numberingIdentifier == nil, let styleIdentifier = paragraph.properties.styleIdentifier {
      var nextIdentifier: String? = styleIdentifier
      var visited: Set<String> = []
      while let identifier = nextIdentifier, visited.insert(identifier).inserted,
        let style = styles.first(where: { $0.identifier == identifier })
      {
        if numberingIdentifier == nil { numberingIdentifier = style.numberingIdentifier }
        if numberingLevel == nil { numberingLevel = style.numberingLevel }
        if numberingIdentifier != nil { break }
        nextIdentifier = style.basedOnIdentifier
      }
    }
    guard let numberingIdentifier else { return nil }
    let levelIndex = numberingLevel ?? 0
    guard
      let level = numbering.level(
        numberingIdentifier: numberingIdentifier,
        level: levelIndex
      ) else { return nil }
    return OfficeWordListInfo(
      numberingIdentifier: numberingIdentifier,
      levelIndex: levelIndex,
      start: level.start,
      format: level.format,
      levelText: level.text,
      justification: level.justification,
      leftIndent: level.leftIndent,
      hangingIndent: level.hangingIndent
    )
  }

  /// Returns a style by its producer-assigned identifier.
  public func style(identifiedBy identifier: String) -> OfficeWordStyle? {
    styles.first { $0.identifier == identifier }
  }

  /// Resolves document defaults, the paragraph style inheritance chain, and direct formatting.
  public func resolvedParagraphProperties(
    for paragraph: OfficeWordParagraph
  ) -> OfficeWordParagraphProperties {
    var result = styleDefaults?.paragraphProperties ?? WordParagraphPropertiesBuilder().value
    let identifier =
      paragraph.properties.styleIdentifier
      ?? styles.first { $0.type == "paragraph" && $0.isDefault }?.identifier
    for style in styleChain(endingAt: identifier) {
      result = style.paragraphProperties.merging(over: result)
    }
    return paragraph.properties.merging(over: result)
  }

  /// Resolves defaults, paragraph/character style inheritance, and direct run formatting.
  public func resolvedRunProperties(
    for run: OfficeWordRun,
    in paragraph: OfficeWordParagraph
  ) -> OfficeWordRunProperties {
    var result = styleDefaults?.runProperties ?? WordRunPropertiesBuilder().value
    let paragraphStyleIdentifier =
      paragraph.properties.styleIdentifier
      ?? styles.first { $0.type == "paragraph" && $0.isDefault }?.identifier
    for style in styleChain(endingAt: paragraphStyleIdentifier) {
      result = style.runProperties.merging(over: result)
    }
    for style in styleChain(endingAt: run.properties.styleIdentifier) {
      result = style.runProperties.merging(over: result)
    }
    return run.properties.merging(over: result)
  }

  private func styleChain(endingAt identifier: String?) -> [OfficeWordStyle] {
    var chain: [OfficeWordStyle] = []
    var nextIdentifier = identifier
    var visited: Set<String> = []
    while let identifier = nextIdentifier, visited.insert(identifier).inserted,
      let style = style(identifiedBy: identifier)
    {
      chain.append(style)
      nextIdentifier = style.basedOnIdentifier
    }
    return chain.reversed()
  }
}

/// The block-level content of a Word document body.
public struct OfficeWordBody: Sendable {
  /// Paragraphs and tables in authored order.
  public let blocks: [OfficeWordBlock]

  /// Section declarations in authored order.
  public let sections: [OfficeWordSection]

  /// Comment range and reference markers found in the main document flow.
  public let commentAnchors: [OfficeWordCommentAnchor]

  /// Complex and simple fields found in the main document flow.
  public let fields: [OfficeWordField]

  /// Bookmark ranges found in the main document flow.
  public let bookmarks: [OfficeWordBookmark]

  /// Structured document tags/content controls in authored order.
  public let contentControls: [OfficeWordContentControl]

  /// Inline and floating DrawingML objects in authored order.
  public let drawings: [OfficeWordDrawing]

  /// Office Math equations in authored order.
  public let equations: [OfficeWordEquation]

  /// Alternative-format imports referenced by `w:altChunk`.
  public let alternativeFormatImports: [OfficeWordAlternativeFormatImport]

  /// Inert OLE and embedded-package objects in authored order.
  public let embeddedObjects: [OfficeWordEmbeddedObject]

  /// Legacy VML shapes that remain in the effective document markup.
  public let legacyShapes: [OfficeWordVMLShape]

  /// The main document part that declared this body.
  public let sourcePart: OfficePart

  /// Direct child paragraphs in authored order.
  public var paragraphs: [OfficeWordParagraph] {
    blocks.compactMap { block in
      guard case .paragraph(let paragraph) = block else { return nil }
      return paragraph
    }
  }

  /// Direct child tables in authored order.
  public var tables: [OfficeWordTable] {
    blocks.compactMap { block in
      guard case .table(let table) = block else { return nil }
      return table
    }
  }

  /// Drawings and equations anchored in `paragraph`, sorted by run, character, and XML order.
  public func inlineContent(
    in paragraph: OfficeWordParagraphLocation
  ) -> [OfficeWordInlineContent] {
    let drawingContent = drawings.compactMap { drawing -> OfficeWordInlineContent? in
      guard drawing.inlinePosition?.paragraph == paragraph else { return nil }
      return .drawing(drawing)
    }
    let equationContent = equations.compactMap { equation -> OfficeWordInlineContent? in
      guard equation.inlinePosition?.paragraph == paragraph else { return nil }
      return .equation(equation)
    }
    return (drawingContent + equationContent).sorted { lhs, rhs in
      guard let left = lhs.position, let right = rhs.position else {
        return lhs.sourceOrder < rhs.sourceOrder
      }
      if left.runIndex != right.runIndex { return left.runIndex < right.runIndex }
      if left.characterOffset != right.characterOffset {
        return left.characterOffset < right.characterOffset
      }
      return lhs.sourceOrder < rhs.sourceOrder
    }
  }
}

/// Non-text inline content retained in Word source order.
public enum OfficeWordInlineContent: Sendable {
  /// A DrawingML object embedded in text flow.
  case drawing(OfficeWordDrawing)
  /// An Office Math equation embedded in text flow.
  case equation(OfficeWordEquation)

  /// Exact paragraph/run/character position.
  public var position: OfficeWordInlinePosition? {
    switch self {
    case .drawing(let drawing): drawing.inlinePosition
    case .equation(let equation): equation.inlinePosition
    }
  }

  /// Monotonic XML event order within the source part.
  public var sourceOrder: Int {
    switch self {
    case .drawing(let drawing): drawing.sourceOrder
    case .equation(let equation): equation.sourceOrder
    }
  }
}

/// One block-level item in a Word document body.
public indirect enum OfficeWordBlock: Sendable {
  /// A paragraph.
  case paragraph(OfficeWordParagraph)

  /// A table.
  case table(OfficeWordTable)
}

/// The structural location of a paragraph in the main Word document flow.
public enum OfficeWordParagraphLocation: Sendable, Hashable, Codable {
  /// A paragraph that is a direct body block.
  case body(blockIndex: Int)

  /// A paragraph inside a table cell.
  case table(
    blockIndex: Int,
    rowIndex: Int,
    cellIndex: Int,
    paragraphIndex: Int
  )

  /// A paragraph inside a table nested in another table cell.
  case nestedTable(path: [OfficeWordTableCellLocation], paragraphIndex: Int)
}

/// One table/cell step in the path to nested Word content.
public struct OfficeWordTableCellLocation: Sendable, Hashable, Codable {
  /// The table's block index in the body or its containing cell.
  public let tableBlockIndex: Int

  /// The row index within the table.
  public let rowIndex: Int

  /// The cell index within the row.
  public let cellIndex: Int
}

/// A run-relative position in an authored Word paragraph.
public struct OfficeWordTextPosition: Sendable, Hashable, Codable {
  /// The paragraph containing the boundary.
  public let paragraph: OfficeWordParagraphLocation

  /// The zero-based run index or exclusive run boundary.
  ///
  /// Comment starts and references use the following or containing run index. Comment ends use
  /// the exclusive boundary, which can equal `runs.count`.
  public let runIndex: Int
}

/// Authored range and reference markers for one Word comment.
public struct OfficeWordCommentAnchor: Sendable, Hashable, Codable {
  /// The numeric comment identifier shared with the comments part.
  public let identifier: Int64

  /// The inclusive start boundary, when the document declares one.
  public let start: OfficeWordTextPosition?

  /// The exclusive end boundary, when the document declares one.
  public let end: OfficeWordTextPosition?

  /// The run containing the visible comment-reference mark, when declared.
  public let referenceRun: OfficeWordTextPosition?

  /// Final page placement depends on Word's text layout and pagination.
  public let spatialInfo: OfficeSpatialInfo

  /// The main document part that declared these markers.
  public let sourcePart: OfficePart
}

/// One authored Word comment.
public struct OfficeWordComment: Sendable {
  /// The numeric identifier referenced by document comment markers.
  public let identifier: Int64

  /// The author string exactly as declared.
  public let author: String?

  /// The author's initials exactly as declared.
  public let initials: String?

  /// The authored timestamp lexical value.
  public let dateText: String?

  /// The parsed authored timestamp, when valid ISO 8601.
  public let date: Date?

  /// Plain text for each comment paragraph in authored order.
  public let paragraphs: [String]

  /// Plain comment text assembled with paragraph separators.
  public let text: String

  /// Fully parsed comment blocks, formatting, fields, drawings, and embedded content.
  public let content: OfficeWordBody

  /// The matched range/reference markers in the main document, when present.
  public let anchor: OfficeWordCommentAnchor?

  /// The comments part that declared this comment.
  public let sourcePart: OfficePart

  /// Resources directly related to the comments part.
  public let attachments: [OfficeAttachment]
}

/// Paragraph properties that do not require style inheritance to interpret.
public struct OfficeWordParagraphProperties: Sendable, Hashable, Codable {
  /// The paragraph style identifier, when directly assigned.
  public let styleIdentifier: String?

  /// The direct `w:jc` alignment token.
  public let alignment: String?

  /// The direct text-to-line vertical alignment token.
  public let textAlignment: String?

  /// Whether Word automatically adjusts spacing between Latin and East Asian text.
  public let automaticallySpacesEastAsianAndLatinText: Bool?

  /// Whether Word automatically adjusts spacing between East Asian text and numbers.
  public let automaticallySpacesEastAsianTextAndNumbers: Bool?

  /// Whether Word automatically adjusts the right indent when a document grid is active.
  public let adjustsRightIndent: Bool?

  /// Whether Word permits character-level wrapping instead of word-level wrapping.
  public let wrapsAtCharacter: Bool?

  /// Whether the paragraph stays with the following paragraph.
  public let keepsWithNext: Bool?

  /// Whether all lines of the paragraph stay on one page.
  public let keepsLinesTogether: Bool?

  /// Whether the paragraph begins on a new page.
  public let startsOnNewPage: Bool?

  /// The numbering definition identifier, when directly assigned.
  public let numberingIdentifier: Int?

  /// The zero-based numbering level, when directly assigned.
  public let numberingLevel: Int?

  /// Space before the paragraph.
  public let spacingBefore: OfficeLength?

  /// Space after the paragraph.
  public let spacingAfter: OfficeLength?

  /// Direct line spacing.
  public let lineSpacing: OfficeLength?

  /// The direct line-spacing rule token.
  public let lineSpacingRule: String?

  /// The leading paragraph indent.
  public let leadingIndent: OfficeLength?

  /// The trailing paragraph indent.
  public let trailingIndent: OfficeLength?

  /// The first-line indent.
  public let firstLineIndent: OfficeLength?

  /// The hanging indent.
  public let hangingIndent: OfficeLength?

  /// Direct tab stops in declaration order.
  public let tabStops: [OfficeWordTabStop]
}

/// One paragraph tab stop.
public struct OfficeWordTabStop: Sendable, Hashable, Codable {
  /// The tab alignment token, such as `left`, `center`, or `right`.
  public let alignment: String?

  /// The optional leader token.
  public let leader: String?

  /// The authored position from the leading margin.
  public let position: OfficeLength?
}

/// One authored Word paragraph.
public struct OfficeWordParagraph: Sendable {
  /// The zero-based paragraph order among direct body blocks or within its table cell.
  public let index: Int

  /// Exact structural path to this paragraph within its story body.
  public let location: OfficeWordParagraphLocation

  /// The producer-assigned paragraph identifier, when present.
  public let identifier: String?

  /// Direct paragraph properties.
  public let properties: OfficeWordParagraphProperties

  /// Runs in authored order.
  public let runs: [OfficeWordRun]

  /// Hyperlinks that enclose ranges of `runs`.
  public let hyperlinks: [OfficeWordHyperlink]

  /// Plain text assembled from runs, tabs, and line breaks.
  public let text: String

  /// Word does not author a final page-space rectangle for ordinary paragraphs.
  public let spatialInfo: OfficeSpatialInfo

  /// The part that declared this paragraph.
  public let sourcePart: OfficePart

  /// Assembles paragraph text for a tracked-revision view.
  public func text(view: OfficeWordRevisionView) -> String {
    runs.filter { view.includes($0.revision) }.map(\.text).joined()
  }
}

/// One authored run of Word text.
public struct OfficeWordRun: Sendable, Hashable, Codable {
  /// Plain run text, including authored tabs and line breaks.
  public let text: String

  /// Field instruction text declared by `w:instrText`, when present.
  public let fieldInstruction: String?

  /// Direct run formatting authored on this run.
  public let properties: OfficeWordRunProperties

  /// The directly assigned character-style identifier.
  public var styleIdentifier: String? { properties.styleIdentifier }

  /// The directly assigned bold state.
  public var isBold: Bool? { properties.isBold }

  /// The directly assigned italic state.
  public var isItalic: Bool? { properties.isItalic }

  /// The directly assigned underline token.
  public var underline: String? { properties.underline }

  /// The referenced footnote identifier, when this run contains `w:footnoteReference`.
  public let footnoteIdentifier: Int64?

  /// The referenced endnote identifier, when this run contains `w:endnoteReference`.
  public let endnoteIdentifier: Int64?

  /// The tracked revision containing this run, when present.
  public let revision: OfficeWordRevision?
}

/// An exact authored position for non-text inline content in a Word paragraph.
public struct OfficeWordInlinePosition: Sendable, Hashable, Codable {
  /// The paragraph containing the content.
  public let paragraph: OfficeWordParagraphLocation

  /// The zero-based run index active at the content's start.
  public let runIndex: Int

  /// Character offset in that run before the content, using Swift `String` character indexing.
  public let characterOffset: Int
}

/// Direct or style-authored formatting for Word text runs.
public struct OfficeWordRunProperties: Sendable, Hashable, Codable {
  /// Referenced character style identifier.
  public let styleIdentifier: String?
  /// Direct bold state.
  public let isBold: Bool?
  /// Direct italic state.
  public let isItalic: Bool?
  /// Underline style token.
  public let underline: String?
  /// Direct single strike-through state.
  public let isStruckThrough: Bool?
  /// Direct double strike-through state.
  public let isDoubleStruckThrough: Bool?
  /// Whether all-capital display is requested.
  public let usesAllCaps: Bool?
  /// Whether small-capital display is requested.
  public let usesSmallCaps: Bool?
  /// Authored RGB or automatic color token.
  public let color: String?
  /// Highlight color token.
  public let highlight: String?
  /// Latin text size in exact half-point-derived units.
  public let fontSize: OfficeLength?
  /// Complex-script text size in exact half-point-derived units.
  public let complexScriptFontSize: OfficeLength?
  /// Direct ASCII typeface.
  public let asciiFont: String?
  /// Direct high-ANSI typeface.
  public let highAnsiFont: String?
  /// Direct East Asian typeface.
  public let eastAsianFont: String?
  /// Direct complex-script typeface.
  public let complexScriptFont: String?
  /// ASCII theme-font role.
  public let asciiThemeFont: String?
  /// High-ANSI theme-font role.
  public let highAnsiThemeFont: String?
  /// East Asian theme-font role.
  public let eastAsianThemeFont: String?
  /// Complex-script theme-font role.
  public let complexScriptThemeFont: String?
  /// Primary BCP 47 language tag.
  public let language: String?
  /// East Asian BCP 47 language tag.
  public let eastAsianLanguage: String?
  /// Bidirectional BCP 47 language tag.
  public let bidirectionalLanguage: String?
  /// Vertical alignment token such as `superscript`.
  public let verticalAlignment: String?
}

/// Whether a Word DrawingML object participates in text flow or floats relative to it.
public enum OfficeWordDrawingPlacement: String, Sendable, Hashable, Codable {
  /// The object occupies an inline character position.
  case inline

  /// The object floats relative to a page, margin, column, paragraph, character, or line.
  case floating
}

/// An exact one-axis position authored for a floating Word drawing.
public struct OfficeWordRelativePosition: Sendable, Hashable, Codable {
  /// The reference coordinate space token, such as `page`, `margin`, or `paragraph`.
  public let relativeFrom: String

  /// The exact offset from the reference origin in EMU, when authored.
  public let offset: OfficeLength?

  /// The authored alignment token, such as `center`, `inside`, or `right`.
  public let alignment: String?
}

/// Exact distances between a Word drawing and surrounding text.
public struct OfficeWordTextDistances: Sendable, Hashable, Codable {
  /// Distance above the object.
  public let top: OfficeLength

  /// Distance below the object.
  public let bottom: OfficeLength

  /// Distance to the object's left.
  public let left: OfficeLength

  /// Distance to the object's right.
  public let right: OfficeLength
}

/// The exact anchor metadata for one Word DrawingML object.
public struct OfficeWordDrawingAnchor: Sendable, Hashable, Codable {
  /// Inline or floating placement.
  public let placement: OfficeWordDrawingPlacement

  /// The exact authored width.
  public let width: OfficeLength

  /// The exact authored height.
  public let height: OfficeLength

  /// Horizontal floating position, when applicable.
  public let horizontalPosition: OfficeWordRelativePosition?

  /// Vertical floating position, when applicable.
  public let verticalPosition: OfficeWordRelativePosition?

  /// The optional exact simple-position point in EMU.
  public let simplePosition: OfficePointEMU?

  /// Distances reserved between the object and surrounding text.
  public let textDistances: OfficeWordTextDistances

  /// The wrapping element token, such as `wrapSquare`, `wrapTight`, or `wrapNone`.
  public let wrap: String?

  /// The authored wrap-side token, when present.
  public let wrapText: String?

  /// The producer's relative stacking height.
  public let relativeHeight: UInt32?

  /// Whether the object is behind document text.
  public let isBehindDocument: Bool?

  /// Whether overlap with other floating objects is allowed.
  public let allowsOverlap: Bool?

  /// Whether the object is constrained to its table cell.
  public let laysOutInCell: Bool?
}

/// A two-dimensional point retained exactly in EMU.
public struct OfficePointEMU: Sendable, Hashable, Codable {
  /// The horizontal coordinate.
  public let x: OfficeLength

  /// The vertical coordinate.
  public let y: OfficeLength
}

/// The broad semantic kind of a Word DrawingML object.
public enum OfficeWordDrawingKind: String, Sendable, Hashable, Codable {
  /// A raster or vector picture.
  case picture
  /// A chart reference.
  case chart
  /// A SmartArt or other diagram reference.
  case diagram
  /// An individual DrawingML shape.
  case shape
  /// A group of DrawingML objects.
  case group
  /// Generic graphic content without a more specific classification.
  case graphic
  /// DrawingML content that OfficeKit does not currently classify.
  case unknown
}

/// One inline or floating DrawingML object declared by a Word story.
public struct OfficeWordDrawing: Sendable {
  /// Zero-based authored order within the source part.
  public let index: Int

  /// Monotonic XML event order within the source part.
  public let sourceOrder: Int

  /// Exact paragraph/run position for drawings embedded in flow content.
  public let inlinePosition: OfficeWordInlinePosition?

  /// Broad semantic object kind.
  public let kind: OfficeWordDrawingKind

  /// Producer-assigned non-visual identifier, when declared.
  public let identifier: UInt32?

  /// Producer-assigned object name.
  public let name: String?

  /// Accessibility-oriented description.
  public let alternativeText: String?

  /// Accessibility-oriented title.
  public let title: String?

  /// Exact Word anchor metadata.
  public let anchor: OfficeWordDrawingAnchor

  /// Picture-specific lazy image references, when present.
  public let picture: OfficePicture?

  /// Typed chart, diagram, or relationship-backed graphic payload, when recognized.
  public let graphicContent: OfficeGraphicContent?

  /// Relationships explicitly referenced by this drawing object.
  public let attachments: [OfficeAttachment]

  /// Exact extents and honestly qualified placement information.
  public let spatialInfo: OfficeSpatialInfo

  /// The XML part that declared the drawing.
  public let sourcePart: OfficePart
}

/// One Office Math equation reduced to its authored text sequence.
public struct OfficeWordEquation: Sendable, Hashable, Codable {
  /// Zero-based authored order within the source part.
  public let index: Int

  /// Monotonic XML event order within the source part.
  public let sourceOrder: Int

  /// Exact paragraph/run position for inline equations.
  public let inlinePosition: OfficeWordInlinePosition?

  /// Whether the equation is wrapped in a display-math paragraph.
  public let isDisplay: Bool

  /// Text from `m:t` elements in authored order.
  public let text: String

  /// The XML part that declared the equation.
  public let sourcePart: OfficePart
}

/// One relationship-backed `w:altChunk` import.
public struct OfficeWordAlternativeFormatImport: Sendable {
  /// Zero-based authored order within the source part.
  public let index: Int

  /// The relationship identifier declared by the element.
  public let relationshipID: OfficeRelationshipID

  /// The imported resource, kept lazy and URL-accessible.
  public let attachment: OfficeAttachment

  /// The XML part that declared the import.
  public let sourcePart: OfficePart
}

/// One inert OLE object or embedded package referenced from Word markup.
public struct OfficeWordEmbeddedObject: Sendable {
  /// Zero-based authored order within the source part.
  public let index: Int

  /// The relationship identifier declared by the OLE element.
  public let relationshipID: OfficeRelationshipID

  /// Whether the object is embedded or linked, exactly as authored.
  public let type: String?

  /// The producer application identifier, such as `Excel.Sheet.12`.
  public let programIdentifier: String?

  /// The requested presentation aspect.
  public let drawingAspect: String?

  /// The associated VML shape identifier.
  public let shapeIdentifier: String?

  /// The producer-assigned object identifier.
  public let objectIdentifier: String?

  /// The inert embedded or linked resource.
  public let attachment: OfficeAttachment

  /// The XML part that declared the object.
  public let sourcePart: OfficePart
}

/// One legacy VML shape retained after markup-compatibility processing.
public struct OfficeWordVMLShape: Sendable {
  /// Zero-based authored order within the source part.
  public let index: Int

  /// The producer-assigned shape identifier.
  public let identifier: String?

  /// The referenced VML shape-type identifier.
  public let typeIdentifier: String?

  /// The exact authored VML style string.
  public let style: String?

  /// Accessibility-oriented alternative text.
  public let alternativeText: String?

  /// The authored fill color token.
  public let fillColor: String?

  /// The authored stroke color token.
  public let strokeColor: String?

  /// Relationships referenced within the shape, including image data.
  public let attachments: [OfficeAttachment]

  /// Geometry parsed from absolute VML style lengths when available.
  public let spatialInfo: OfficeSpatialInfo

  /// The XML part that declared the shape.
  public let sourcePart: OfficePart
}

/// A tracked-change kind that can contain authored runs.
public enum OfficeWordRevisionKind: String, Sendable, Hashable, Codable {
  /// Inserted content.
  case insertion

  /// Deleted content.
  case deletion

  /// Content moved away from this location.
  case moveFrom

  /// Content moved into this location.
  case moveTo
}

/// Metadata for one tracked revision around a run.
public struct OfficeWordRevision: Sendable, Hashable, Codable {
  /// The revision operation.
  public let kind: OfficeWordRevisionKind

  /// The producer-assigned revision identifier.
  public let identifier: Int64?

  /// The author exactly as declared.
  public let author: String?

  /// The authored timestamp lexical value.
  public let dateText: String?

  /// The parsed authored timestamp, when valid ISO 8601.
  public let date: Date?
}

/// A policy for projecting tracked revisions into plain text.
public enum OfficeWordRevisionView: String, Sendable, Hashable, Codable {
  /// Shows inserted and moved-to content while hiding deletions and moved-from content.
  case final

  /// Shows deleted and moved-from content while hiding insertions and moved-to content.
  case original

  /// Shows every authored run regardless of revision state.
  case all

  package func includes(_ revision: OfficeWordRevision?) -> Bool {
    guard let revision else { return true }
    return switch (self, revision.kind) {
    case (.all, _), (.final, .insertion), (.final, .moveTo),
      (.original, .deletion), (.original, .moveFrom):
      true
    default: false
    }
  }
}

/// Whether a Word field uses a wrapper or begin/separate/end characters.
public enum OfficeWordFieldKind: String, Sendable, Hashable, Codable {
  /// A `w:fldSimple` field.
  case simple

  /// A field represented by `w:fldChar` runs.
  case complex
}

/// One authored Word field and its positions in document flow.
public struct OfficeWordField: Sendable, Hashable, Codable {
  /// The field representation.
  public let kind: OfficeWordFieldKind

  /// The field instruction assembled from `w:instr` or `w:instrText`.
  public let instruction: String

  /// The opening marker or simple-field content boundary.
  public let start: OfficeWordTextPosition

  /// The complex-field separator marker, when declared.
  public let separator: OfficeWordTextPosition?

  /// The closing marker or simple-field content boundary, when declared.
  public let end: OfficeWordTextPosition?

  /// The part that declared this field.
  public let sourcePart: OfficePart
}

/// A bookmark position in text or between body blocks.
public enum OfficeWordBookmarkPosition: Sendable, Hashable, Codable {
  /// A run-relative position inside a paragraph.
  case text(OfficeWordTextPosition)

  /// A boundary between direct body blocks.
  case bodyBlock(Int)
}

/// One authored Word bookmark range.
public struct OfficeWordBookmark: Sendable, Hashable, Codable {
  /// The numeric bookmark identifier.
  public let identifier: Int64

  /// The producer-assigned bookmark name.
  public let name: String?

  /// The opening position.
  public let start: OfficeWordBookmarkPosition

  /// The closing position, when declared.
  public let end: OfficeWordBookmarkPosition?

  /// The part that declared this bookmark.
  public let sourcePart: OfficePart
}

/// The structural scope occupied by a Word content control.
public enum OfficeWordContentControlScope: String, Sendable, Hashable, Codable {
  /// A control around one or more body blocks.
  case block

  /// A control within a paragraph.
  case inline

  /// A control around a table row.
  case row

  /// A control around cell content.
  case cell
}

/// One item offered by a combo-box or drop-down content control.
public struct OfficeWordContentControlItem: Sendable, Hashable, Codable {
  /// The user-facing item text.
  public let displayText: String?

  /// The stored item value.
  public let value: String?
}

/// Data-binding metadata for a Word content control.
public struct OfficeWordContentControlBinding: Sendable, Hashable, Codable {
  /// The custom XML store item identifier.
  public let storeItemIdentifier: String?

  /// The XPath selecting the bound value.
  public let xpath: String?

  /// Namespace prefix mappings used by `xpath`.
  public let prefixMappings: String?
}

/// One structured document tag/content control.
public struct OfficeWordContentControl: Sendable, Hashable, Codable {
  /// The signed producer-assigned control identifier.
  public let identifier: Int64?

  /// The user-facing alias.
  public let alias: String?

  /// The producer-assigned tag.
  public let tag: String?

  /// The control kind token, such as `text`, `picture`, `checkbox`, or `comboBox`.
  public let kind: String

  /// Whether a checkbox control is checked, when declared.
  public let isChecked: Bool?

  /// Items offered by a combo-box or drop-down control.
  public let items: [OfficeWordContentControlItem]

  /// Custom XML binding metadata, when declared.
  public let binding: OfficeWordContentControlBinding?

  /// Whether the control is block-, inline-, row-, or cell-scoped.
  public let scope: OfficeWordContentControlScope

  /// Plain text contained by the control, including nested controls.
  public let text: String

  /// The part that declared this control.
  public let sourcePart: OfficePart
}

/// A hyperlink around a contiguous range of paragraph runs.
public struct OfficeWordHyperlink: Sendable {
  /// The half-open range in the containing paragraph's `runs` array.
  public let runRange: Range<Int>

  /// An internal bookmark target, when declared.
  public let anchor: String?

  /// Producer-authored hover text, when declared.
  public let tooltip: String?

  /// The relationship identifier, when the link is relationship-backed.
  public let relationshipID: OfficeRelationshipID?

  /// The lazy external or packaged link target, when resolvable.
  public let attachment: OfficeAttachment?
}

/// One Word table in the main document flow.
public struct OfficeWordTable: Sendable {
  /// The zero-based block index in the body or containing cell.
  public let index: Int

  /// Rows in authored order.
  public let rows: [OfficeWordTableRow]

  /// Exact authored table-grid column widths.
  public let columnWidths: [OfficeLength]

  /// Direct table style identifier.
  public let styleIdentifier: String?

  /// Direct preferred table width.
  public let preferredWidth: OfficeWordWidth?

  /// Direct table indentation.
  public let indentation: OfficeWordWidth?

  /// Direct table alignment token.
  public let alignment: String?

  /// Direct layout token, commonly `fixed` or `autofit`.
  public let layout: String?

  /// Direct table borders.
  public let borders: OfficeWordBorders

  /// Direct table shading.
  public let shading: OfficeWordShading?

  /// Plain text assembled from all cells.
  public let text: String

  /// Final table placement depends on Word's layout engine.
  public let spatialInfo: OfficeSpatialInfo

  /// The part that declared this table.
  public let sourcePart: OfficePart
}

/// One authored Word table row.
public struct OfficeWordTableRow: Sendable {
  /// The zero-based row index.
  public let index: Int

  /// Cells in authored order.
  public let cells: [OfficeWordTableCell]

  /// Direct row height, when declared.
  public let height: OfficeLength?

  /// Height interpretation token, such as `atLeast` or `exact`.
  public let heightRule: String?

  /// Whether the row is prohibited from splitting across pages.
  public let preventsPageSplit: Bool?

  /// Whether the row repeats as a table header.
  public let repeatsAsHeader: Bool?
}

/// One authored Word table cell.
public struct OfficeWordTableCell: Sendable {
  /// The zero-based cell index within its row.
  public let index: Int

  /// The number of layout-grid columns spanned by this cell.
  public let gridSpan: Int

  /// The direct vertical-merge token, such as `restart` or `continue`.
  public let verticalMerge: String?

  /// Direct preferred cell width.
  public let preferredWidth: OfficeWordWidth?

  /// Direct vertical-alignment token.
  public let verticalAlignment: String?

  /// Direct cell borders.
  public let borders: OfficeWordBorders

  /// Direct cell shading.
  public let shading: OfficeWordShading?

  /// Paragraph and nested-table content in authored order.
  public let blocks: [OfficeWordBlock]

  /// Direct child paragraphs in authored order.
  public var paragraphs: [OfficeWordParagraph] {
    blocks.compactMap { block in
      guard case .paragraph(let paragraph) = block else { return nil }
      return paragraph
    }
  }

  /// Direct child nested tables in authored order.
  public var tables: [OfficeWordTable] {
    blocks.compactMap { block in
      guard case .table(let table) = block else { return nil }
      return table
    }
  }

  /// Plain text assembled from the cell's paragraphs.
  public let text: String
}

/// A Word table or cell width with its original unit interpretation.
public struct OfficeWordWidth: Sendable, Hashable, Codable {
  /// The source width type, such as `dxa`, `pct`, `auto`, or `nil`.
  public let type: String

  /// The exact source integer value.
  public let value: Int64

  /// The exact length for a twentieths-of-a-point width.
  public var length: OfficeLength? {
    guard type == "dxa" else { return nil }
    return OfficeLength(emu: value * 635)
  }

  /// Percentage width, where `100` means one hundred percent.
  public var percent: Double? {
    guard type == "pct" else { return nil }
    return Double(value) / 50
  }
}

/// One Word border edge with exact eighth-point sizing.
public struct OfficeWordBorder: Sendable, Hashable, Codable {
  /// Border style token, such as `single`, `double`, or `nil`.
  public let style: String

  /// Authored color token.
  public let color: String?

  /// Exact border width in eighths of a point.
  public let sizeInEighthPoints: Int?

  /// Exact border spacing in points.
  public let spacingInPoints: Int?
}

/// Direct border edges for a Word table or cell.
public struct OfficeWordBorders: Sendable, Hashable, Codable {
  /// Top border.
  public let top: OfficeWordBorder?
  /// Leading border in logical reading order.
  public let leading: OfficeWordBorder?
  /// Bottom border.
  public let bottom: OfficeWordBorder?
  /// Trailing border in logical reading order.
  public let trailing: OfficeWordBorder?
  /// Border between adjacent rows.
  public let insideHorizontal: OfficeWordBorder?
  /// Border between adjacent columns.
  public let insideVertical: OfficeWordBorder?

  /// No directly declared borders.
  public static let none = OfficeWordBorders(
    top: nil,
    leading: nil,
    bottom: nil,
    trailing: nil,
    insideHorizontal: nil,
    insideVertical: nil
  )
}

/// Direct Word table or cell shading.
public struct OfficeWordShading: Sendable, Hashable, Codable {
  /// Pattern token, such as `clear` or `pct25`.
  public let pattern: String?

  /// Foreground color token.
  public let color: String?

  /// Background fill color token.
  public let fill: String?
}

/// Exact authored page margins for a Word section.
public struct OfficeWordPageMargins: Sendable, Hashable, Codable {
  /// Top margin, converted exactly from twentieths of a point.
  public let top: OfficeLength?

  /// Right margin, converted exactly from twentieths of a point.
  public let right: OfficeLength?

  /// Bottom margin, converted exactly from twentieths of a point.
  public let bottom: OfficeLength?

  /// Left margin, converted exactly from twentieths of a point.
  public let left: OfficeLength?

  /// Header distance from the page edge.
  public let header: OfficeLength?

  /// Footer distance from the page edge.
  public let footer: OfficeLength?

  /// Additional gutter width.
  public let gutter: OfficeLength?
}

/// One explicitly sized text column in a Word section.
public struct OfficeWordColumn: Sendable, Hashable, Codable {
  /// The authored column width.
  public let width: OfficeLength?

  /// The authored space following the column.
  public let spacing: OfficeLength?
}

/// Authored multi-column layout settings for a Word section.
public struct OfficeWordColumns: Sendable, Hashable, Codable {
  /// The declared column count, when present.
  public let count: Int?

  /// The default spacing between equal-width columns.
  public let spacing: OfficeLength?

  /// Whether a separator line is drawn between columns.
  public let hasSeparator: Bool?

  /// Whether Word computes columns at equal widths.
  public let usesEqualWidths: Bool?

  /// Explicit unequal-width column definitions in authored order.
  public let columns: [OfficeWordColumn]
}

/// An authored Word section declaration.
public struct OfficeWordSection: Sendable, Hashable, Codable {
  /// The section-break token, such as `continuous` or `nextPage`.
  public let type: String?

  /// Exact authored page width.
  public let pageWidth: OfficeLength?

  /// Exact authored page height.
  public let pageHeight: OfficeLength?

  /// The direct page-orientation token.
  public let orientation: String?

  /// Exact authored page margins.
  public let margins: OfficeWordPageMargins?

  /// Authored section column settings.
  public let columns: OfficeWordColumns?

  /// Whether the section uses a distinct first-page header and footer.
  public let hasTitlePage: Bool?

  /// The vertical alignment token for content on section pages.
  public let verticalAlignment: String?

  /// The starting page number, when directly declared.
  public let pageNumberStart: Int?

  /// Header relationships selected by this section.
  public let headerReferences: [OfficeWordHeaderFooterReference]

  /// Footer relationships selected by this section.
  public let footerReferences: [OfficeWordHeaderFooterReference]

  /// Exact page geometry when both dimensions are declared.
  public let spatialInfo: OfficeSpatialInfo

  /// The part that declared this section.
  public let sourcePart: OfficePart
}

/// Document-wide settings that influence Word layout and authored behavior.
public struct OfficeWordSettings: Sendable, Hashable, Codable {
  /// The requested document view token, such as `print` or `web`.
  public let view: String?

  /// The requested zoom percentage, with an optional Strict `%` suffix normalized away.
  public let zoomPercentage: Int?

  /// The default tab interval.
  public let defaultTabStop: OfficeLength?

  /// The character-spacing control token.
  public let characterSpacingControl: String?

  /// Whether revision tracking is enabled.
  public let tracksRevisions: Bool?

  /// Whether fields should update when the document opens.
  public let updatesFieldsOnOpen: Bool?

  /// Whether facing-page margins are mirrored.
  public let mirrorsMargins: Bool?

  /// Whether odd and even pages use distinct headers and footers.
  public let hasEvenAndOddHeaders: Bool?

  /// The theme language for Latin text.
  public let themeLanguage: String?

  /// The theme language for East Asian text.
  public let eastAsianThemeLanguage: String?

  /// The theme language for bidirectional text.
  public let bidirectionalThemeLanguage: String?

  /// The document's formula decimal symbol.
  public let decimalSymbol: String?

  /// The document's formula list separator.
  public let listSeparator: String?

  /// The requested Word compatibility-mode version.
  public let compatibilityMode: Int?

  /// The settings part that declared these values.
  public let sourcePart: OfficePart
}

/// A named Word style definition.
public struct OfficeWordStyle: Sendable, Hashable, Codable {
  /// The producer-assigned style identifier.
  public let identifier: String

  /// The style kind token, such as `paragraph`, `character`, or `table`.
  public let type: String

  /// The user-facing style name, when declared.
  public let name: String?

  /// The inherited base style identifier, when declared.
  public let basedOnIdentifier: String?

  /// The style selected for a following paragraph, when declared.
  public let nextIdentifier: String?

  /// Whether this is the default style for its kind.
  public let isDefault: Bool

  /// Whether the style is marked as a primary style in the Word UI.
  public let isPrimary: Bool

  /// The numbering definition inherited by paragraphs using this style, when declared.
  public let numberingIdentifier: Int?

  /// The numbering level inherited by paragraphs using this style, when declared.
  public let numberingLevel: Int?

  /// Paragraph formatting declared by the style.
  public let paragraphProperties: OfficeWordParagraphProperties

  /// Run formatting declared by the style.
  public let runProperties: OfficeWordRunProperties

  /// The styles part that declared this record.
  public let sourcePart: OfficePart
}

/// Default paragraph and run formatting declared by `w:docDefaults`.
public struct OfficeWordStyleDefaults: Sendable, Hashable, Codable {
  /// Default paragraph formatting.
  public let paragraphProperties: OfficeWordParagraphProperties

  /// Default run formatting.
  public let runProperties: OfficeWordRunProperties

  /// The styles part that declared these defaults.
  public let sourcePart: OfficePart
}

extension OfficeWordParagraphProperties {
  fileprivate func merging(over base: Self) -> Self {
    Self(
      styleIdentifier: styleIdentifier ?? base.styleIdentifier,
      alignment: alignment ?? base.alignment,
      textAlignment: textAlignment ?? base.textAlignment,
      automaticallySpacesEastAsianAndLatinText: automaticallySpacesEastAsianAndLatinText
        ?? base.automaticallySpacesEastAsianAndLatinText,
      automaticallySpacesEastAsianTextAndNumbers: automaticallySpacesEastAsianTextAndNumbers
        ?? base.automaticallySpacesEastAsianTextAndNumbers,
      adjustsRightIndent: adjustsRightIndent ?? base.adjustsRightIndent,
      wrapsAtCharacter: wrapsAtCharacter ?? base.wrapsAtCharacter,
      keepsWithNext: keepsWithNext ?? base.keepsWithNext,
      keepsLinesTogether: keepsLinesTogether ?? base.keepsLinesTogether,
      startsOnNewPage: startsOnNewPage ?? base.startsOnNewPage,
      numberingIdentifier: numberingIdentifier ?? base.numberingIdentifier,
      numberingLevel: numberingLevel ?? base.numberingLevel,
      spacingBefore: spacingBefore ?? base.spacingBefore,
      spacingAfter: spacingAfter ?? base.spacingAfter,
      lineSpacing: lineSpacing ?? base.lineSpacing,
      lineSpacingRule: lineSpacingRule ?? base.lineSpacingRule,
      leadingIndent: leadingIndent ?? base.leadingIndent,
      trailingIndent: trailingIndent ?? base.trailingIndent,
      firstLineIndent: firstLineIndent ?? base.firstLineIndent,
      hangingIndent: hangingIndent ?? base.hangingIndent,
      tabStops: tabStops.isEmpty ? base.tabStops : tabStops
    )
  }
}

extension OfficeWordRunProperties {
  fileprivate func merging(over base: Self) -> Self {
    Self(
      styleIdentifier: styleIdentifier ?? base.styleIdentifier,
      isBold: isBold ?? base.isBold,
      isItalic: isItalic ?? base.isItalic,
      underline: underline ?? base.underline,
      isStruckThrough: isStruckThrough ?? base.isStruckThrough,
      isDoubleStruckThrough: isDoubleStruckThrough ?? base.isDoubleStruckThrough,
      usesAllCaps: usesAllCaps ?? base.usesAllCaps,
      usesSmallCaps: usesSmallCaps ?? base.usesSmallCaps,
      color: color ?? base.color,
      highlight: highlight ?? base.highlight,
      fontSize: fontSize ?? base.fontSize,
      complexScriptFontSize: complexScriptFontSize ?? base.complexScriptFontSize,
      asciiFont: asciiFont ?? base.asciiFont,
      highAnsiFont: highAnsiFont ?? base.highAnsiFont,
      eastAsianFont: eastAsianFont ?? base.eastAsianFont,
      complexScriptFont: complexScriptFont ?? base.complexScriptFont,
      asciiThemeFont: asciiThemeFont ?? base.asciiThemeFont,
      highAnsiThemeFont: highAnsiThemeFont ?? base.highAnsiThemeFont,
      eastAsianThemeFont: eastAsianThemeFont ?? base.eastAsianThemeFont,
      complexScriptThemeFont: complexScriptThemeFont ?? base.complexScriptThemeFont,
      language: language ?? base.language,
      eastAsianLanguage: eastAsianLanguage ?? base.eastAsianLanguage,
      bidirectionalLanguage: bidirectionalLanguage ?? base.bidirectionalLanguage,
      verticalAlignment: verticalAlignment ?? base.verticalAlignment
    )
  }
}

/// The page variant selected by a Word header or footer reference.
public enum OfficeWordHeaderFooterVariant: Sendable, Hashable, Codable {
  /// The normal header or footer.
  case `default`

  /// The header or footer used on the first page of a section.
  case firstPage

  /// The header or footer used on even-numbered pages.
  case evenPages

  /// A producer extension retained without interpretation.
  case unknown(String)
}

/// A section's relationship to one header or footer story.
public struct OfficeWordHeaderFooterReference: Sendable, Hashable, Codable {
  /// The selected page variant.
  public let variant: OfficeWordHeaderFooterVariant

  /// The relationship identifier declared by the section.
  public let relationshipID: OfficeRelationshipID

  /// The resolved story part.
  public let part: OfficePart
}

/// The role of an auxiliary Word story.
public enum OfficeWordStoryKind: String, Sendable, Hashable, Codable {
  /// A page header.
  case header

  /// A page footer.
  case footer
}

/// Parsed content from one Word header or footer part.
public struct OfficeWordStory: Sendable {
  /// Whether this story is a header or footer.
  public let kind: OfficeWordStoryKind

  /// The story XML part.
  public let part: OfficePart

  /// Ordered paragraphs and tables in the story.
  public let content: OfficeWordBody

  /// Resources directly related to the story part.
  public let attachments: [OfficeAttachment]

  /// Plain text assembled from direct paragraphs and tables.
  public var text: String {
    content.blocks.map { block in
      switch block {
      case .paragraph(let paragraph): paragraph.text
      case .table(let table): table.text
      }
    }.joined(separator: "\n")
  }
}

/// Whether a note belongs to the footnote or endnote collection.
public enum OfficeWordNoteKind: String, Sendable, Hashable, Codable {
  /// A footnote.
  case footnote

  /// An endnote.
  case endnote
}

/// One authored Word footnote or endnote.
public struct OfficeWordNote: Sendable {
  /// Whether this is a footnote or endnote.
  public let kind: OfficeWordNoteKind

  /// The signed note identifier. Negative values identify separator stories.
  public let identifier: Int64

  /// A special-note token such as `separator` or `continuationSeparator`.
  public let type: String?

  /// Plain text for each note paragraph in authored order.
  public let paragraphs: [String]

  /// Plain note text assembled with paragraph separators.
  public let text: String

  /// Fully parsed note blocks, formatting, fields, drawings, equations, and embedded content.
  public let content: OfficeWordBody

  /// The note collection part that declared this note.
  public let sourcePart: OfficePart
}

/// A parsed footnote or endnote collection and its lazy resources.
public struct OfficeWordNoteCollection: Sendable {
  /// Whether this collection contains footnotes or endnotes.
  public let kind: OfficeWordNoteKind

  /// Notes in declaration order, including separator stories.
  public let notes: [OfficeWordNote]

  /// The footnotes or endnotes XML part.
  public let sourcePart: OfficePart

  /// Resources directly related to the note collection.
  public let attachments: [OfficeAttachment]

  /// Returns the note with `identifier`, when present.
  public func note(identifiedBy identifier: Int64) -> OfficeWordNote? {
    notes.first { $0.identifier == identifier }
  }
}

/// One level in an abstract Word numbering definition.
public struct OfficeWordNumberingLevel: Sendable, Hashable, Codable {
  /// The zero-based list level.
  public let index: Int

  /// The first number used by this level, when declared.
  public let start: Int?

  /// The numbering format token, such as `decimal`, `lowerLetter`, or `bullet`.
  public let format: String?

  /// The level text template, such as `%1.` or a bullet glyph.
  public let text: String?

  /// The marker justification token.
  public let justification: String?

  /// A paragraph style linked to this level.
  public let paragraphStyleIdentifier: String?

  /// Exact left indentation converted from twentieths of a point.
  public let leftIndent: OfficeLength?

  /// Exact hanging indentation converted from twentieths of a point.
  public let hangingIndent: OfficeLength?
}

/// One reusable multilevel numbering template.
public struct OfficeWordAbstractNumbering: Sendable, Hashable, Codable {
  /// The abstract numbering identifier.
  public let identifier: Int

  /// The multilevel behavior token.
  public let multiLevelType: String?

  /// Levels in declaration order.
  public let levels: [OfficeWordNumberingLevel]
}

/// An instance-level replacement for one abstract numbering level.
public struct OfficeWordNumberingLevelOverride: Sendable, Hashable, Codable {
  /// The overridden zero-based level.
  public let levelIndex: Int

  /// A replacement starting number.
  public let start: Int?

  /// A complete replacement level, when declared.
  public let level: OfficeWordNumberingLevel?
}

/// One concrete numbering instance selected by paragraph `numId` values.
public struct OfficeWordNumberingInstance: Sendable, Hashable, Codable {
  /// The concrete numbering identifier.
  public let identifier: Int

  /// The referenced abstract numbering template.
  public let abstractIdentifier: Int

  /// Per-level overrides in declaration order.
  public let levelOverrides: [OfficeWordNumberingLevelOverride]
}

/// Parsed Word numbering definitions.
public struct OfficeWordNumbering: Sendable, Hashable, Codable {
  /// Reusable abstract numbering templates.
  public let abstractDefinitions: [OfficeWordAbstractNumbering]

  /// Concrete numbering instances selected by paragraphs.
  public let instances: [OfficeWordNumberingInstance]

  /// The numbering part.
  public let sourcePart: OfficePart

  /// Resolves the effective level for a concrete numbering identifier.
  public func level(numberingIdentifier: Int, level levelIndex: Int) -> OfficeWordNumberingLevel? {
    guard let instance = instances.first(where: { $0.identifier == numberingIdentifier }) else {
      return nil
    }
    if let override = instance.levelOverrides.first(where: { $0.levelIndex == levelIndex }) {
      if let level = override.level { return level }
      if let start = override.start,
        let base = abstractDefinitions.first(where: {
          $0.identifier == instance.abstractIdentifier
        })?.levels.first(where: { $0.index == levelIndex })
      {
        return OfficeWordNumberingLevel(
          index: base.index,
          start: start,
          format: base.format,
          text: base.text,
          justification: base.justification,
          paragraphStyleIdentifier: base.paragraphStyleIdentifier,
          leftIndent: base.leftIndent,
          hangingIndent: base.hangingIndent
        )
      }
    }
    return abstractDefinitions.first(where: {
      $0.identifier == instance.abstractIdentifier
    })?.levels.first(where: { $0.index == levelIndex })
  }
}

/// A paragraph's resolved list semantics.
public struct OfficeWordListInfo: Sendable, Hashable, Codable {
  /// The concrete numbering identifier selected by the paragraph or its style.
  public let numberingIdentifier: Int

  /// The resolved zero-based list level.
  public let levelIndex: Int

  /// The first number used by this level, when declared.
  public let start: Int?

  /// The numbering format token.
  public let format: String?

  /// The marker text template.
  public let levelText: String?

  /// The marker justification token.
  public let justification: String?

  /// Exact left indentation.
  public let leftIndent: OfficeLength?

  /// Exact hanging indentation.
  public let hangingIndent: OfficeLength?
}
