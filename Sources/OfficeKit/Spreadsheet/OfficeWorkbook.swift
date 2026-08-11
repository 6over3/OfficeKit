import Foundation

/// A read-only, lazily parsed Excel workbook.
///
/// Initialization indexes the workbook's ordered worksheet references. Worksheet XML and shared
/// strings are parsed only when `worksheet(at:)` is requested.
public struct OfficeWorkbook: Sendable {
  private let storage: SpreadsheetStorage

  /// The opened Office document.
  public let document: OfficeDocument

  /// Worksheets in workbook order.
  public let worksheets: [OfficeWorksheetReference]

  /// All workbook sheets in authored order, including chart, dialog, and macro sheets.
  public let sheets: [OfficeSheetReference]

  /// The serial-date epoch selected by the workbook.
  public let dateSystem: OfficeSpreadsheetDateSystem

  /// Workbook application properties.
  public let properties: OfficeWorkbookProperties

  /// Workbook-scoped and worksheet-scoped defined names.
  public let definedNames: [OfficeDefinedName]

  /// Calculation behavior declared by the workbook.
  public let calculation: OfficeWorkbookCalculation

  /// Every relationship-backed resource owned by the workbook part.
  public let attachments: [OfficeAttachment]

  /// Opens and indexes an Excel workbook.
  public init(contentsOf url: URL, limits: OfficeParsingLimits = .standard) throws {
    try self.init(document: OfficeDocument(contentsOf: url, limits: limits))
  }

  /// Indexes an already opened SpreadsheetML document.
  public init(document: OfficeDocument) throws {
    guard document.kind == .spreadsheet else {
      throw OfficeKitError.invalidPackage("The Office document is not a spreadsheet.")
    }
    let values = try SpreadsheetWorkbookParser.parse(
      package: document.package,
      workbookPart: document.mainPart
    )
    let storage = SpreadsheetStorage(package: document.package, workbookPart: document.mainPart)
    self.document = document
    self.storage = storage
    self.sheets = values.sheets
    self.worksheets = values.worksheets
    self.dateSystem = values.dateSystem
    self.properties = values.properties
    self.definedNames = values.definedNames
    self.calculation = values.calculation
    self.attachments = values.attachments
  }

  /// Parses one worksheet's sparse rows, cells, formulas, values, and relationships.
  public func worksheet(at index: Int) throws -> OfficeWorksheet {
    guard worksheets.indices.contains(index) else {
      throw OfficeKitError.invalidPackage("Worksheet index \(index) is out of bounds.")
    }
    return try SpreadsheetWorksheetParser.parse(
      reference: worksheets[index],
      package: document.package,
      sharedStrings: storage.sharedStrings(),
      styles: storage.styles(),
      dateSystem: dateSystem
    )
  }

  /// Streams sparse worksheet rows to `body` without retaining completed rows.
  ///
  /// Shared strings and styles are cached at workbook scope, but worksheet memory is bounded by
  /// the current row and XML parser buffers. Throwing from `body` stops parsing immediately and
  /// propagates the same error.
  public func streamRows(
    inWorksheetAt index: Int,
    _ body: @escaping (OfficeWorksheetRow) throws -> Void
  ) throws {
    guard worksheets.indices.contains(index) else {
      throw OfficeKitError.invalidPackage("Worksheet index \(index) is out of bounds.")
    }
    try SpreadsheetWorksheetParser.streamRows(
      reference: worksheets[index],
      package: document.package,
      sharedStrings: storage.sharedStrings(),
      styles: storage.styles(),
      dateSystem: dateSystem,
      body
    )
  }
}

// SAFETY: the package and part are immutable Sendable values, and every access to either mutable
// cache is serialized by `lock`.
private final class SpreadsheetStorage: @unchecked Sendable {
  private let package: OfficePackage
  private let workbookPart: OfficePart
  private let lock = NSLock()
  private var cachedSharedStrings: Result<[OfficeSpreadsheetRichText], OfficeKitError>?
  private var cachedStyles: Result<[OfficeCellStyle], OfficeKitError>?

  init(package: OfficePackage, workbookPart: OfficePart) {
    self.package = package
    self.workbookPart = workbookPart
  }

  func sharedStrings() throws -> [OfficeSpreadsheetRichText] {
    try lock.withLock {
      if let cachedSharedStrings { return try cachedSharedStrings.get() }
      do {
        let strings = try SpreadsheetSharedStringParser.parse(
          package: package,
          workbookPart: workbookPart
        )
        cachedSharedStrings = .success(strings)
        return strings
      } catch let error as OfficeKitError {
        cachedSharedStrings = .failure(error)
        throw error
      } catch {
        let wrapped = OfficeKitError.invalidPackage("Unexpected shared-string parsing failure.")
        cachedSharedStrings = .failure(wrapped)
        throw wrapped
      }
    }
  }

  func styles() throws -> [OfficeCellStyle] {
    try lock.withLock {
      if let cachedStyles { return try cachedStyles.get() }
      do {
        let styles = try SpreadsheetStyleParser.parse(
          package: package,
          workbookPart: workbookPart
        )
        cachedStyles = .success(styles)
        return styles
      } catch let error as OfficeKitError {
        cachedStyles = .failure(error)
        throw error
      } catch {
        let wrapped = OfficeKitError.invalidPackage("Unexpected cell-style parsing failure.")
        cachedStyles = .failure(wrapped)
        throw wrapped
      }
    }
  }
}

private enum SpreadsheetWorkbookParser {
  private static let spreadsheetNamespace =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  private static let relationshipNamespace =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  struct Values {
    let sheets: [OfficeSheetReference]
    let worksheets: [OfficeWorksheetReference]
    let dateSystem: OfficeSpreadsheetDateSystem
    let properties: OfficeWorkbookProperties
    let definedNames: [OfficeDefinedName]
    let calculation: OfficeWorkbookCalculation
    let attachments: [OfficeAttachment]
  }

  static func parse(
    package: OfficePackage,
    workbookPart: OfficePart
  ) throws -> Values {
    struct RawSheet {
      let name: String
      let identifier: UInt32
      let visibility: OfficeWorksheetVisibility
      let relationshipID: OfficeRelationshipID
    }

    struct RawDefinedName {
      let name: String
      let localSheetIndex: Int?
      let isHidden: Bool
      let description: String?
      let comment: String?
      var formula = ""
    }

    var rawSheets: [RawSheet] = []
    var dateSystem = OfficeSpreadsheetDateSystem.nineteenHundred
    var codeName: String?
    var filtersPrivacy: Bool?
    var updateLinks: String?
    var calculationIdentifier: UInt32?
    var calculationMode: String?
    var calculatesFullyOnLoad: Bool?
    var forcesFullCalculation: Bool?
    var iterates: Bool?
    var iterationCount: UInt32?
    var iterationDelta: Double?
    var definedNames: [OfficeDefinedName] = []
    var currentDefinedName: RawDefinedName?
    var definedNameDepth: Int?
    var depth = 0
    try package.parseXML(in: workbookPart, compatibility: .commonOffice) { event in
      switch event {
      case .startElement(let name, let attributes, _, _):
        depth += 1
        guard name.namespaceURI == spreadsheetNamespace else { return }
        switch name.localName {
        case "workbookPr":
          let uses1904Epoch =
            attribute("date1904", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false
          dateSystem = uses1904Epoch ? .nineteenOhFour : .nineteenHundred
          codeName = attribute("codeName", in: attributes)
          filtersPrivacy = attribute("filterPrivacy", in: attributes)
            .flatMap(OfficeValueDecoder.boolean)
          updateLinks = attribute("updateLinks", in: attributes)
        case "calcPr":
          calculationIdentifier = attribute("calcId", in: attributes).flatMap(UInt32.init)
          calculationMode = attribute("calcMode", in: attributes)
          calculatesFullyOnLoad = attribute("fullCalcOnLoad", in: attributes)
            .flatMap(OfficeValueDecoder.boolean)
          forcesFullCalculation = attribute("forceFullCalc", in: attributes)
            .flatMap(OfficeValueDecoder.boolean)
          iterates = attribute("iterate", in: attributes).flatMap(OfficeValueDecoder.boolean)
          iterationCount = attribute("iterateCount", in: attributes).flatMap(UInt32.init)
          iterationDelta = attribute("iterateDelta", in: attributes).flatMap(Double.init)
        case "definedName":
          guard let name = attribute("name", in: attributes) else {
            throw OfficeKitError.invalidPackage("Workbook contains a defined name without a name.")
          }
          currentDefinedName = RawDefinedName(
            name: name,
            localSheetIndex: attribute("localSheetId", in: attributes).flatMap(Int.init),
            isHidden: attribute("hidden", in: attributes)
              .flatMap(OfficeValueDecoder.boolean) ?? false,
            description: attribute("description", in: attributes),
            comment: attribute("comment", in: attributes)
          )
          definedNameDepth = depth
        case "sheet":
          guard let sheetName = attribute("name", in: attributes),
            let identifier = attribute("sheetId", in: attributes).flatMap(UInt32.init),
            let relationship = attribute(
              "id",
              namespace: relationshipNamespace,
              in: attributes
            ) else {
            throw OfficeKitError.invalidPackage("Workbook contains an incomplete sheet reference.")
          }
          let rawVisibility = attribute("state", in: attributes) ?? "visible"
          guard let visibility = OfficeWorksheetVisibility(rawValue: rawVisibility) else {
            throw OfficeKitError.invalidPackage(
              "Worksheet \(sheetName) has unknown visibility \(rawVisibility)."
            )
          }
          rawSheets.append(
            RawSheet(
              name: sheetName,
              identifier: identifier,
              visibility: visibility,
              relationshipID: OfficeRelationshipID(rawValue: relationship)
            )
          )
        default: break
        }
      case .text(let text, _):
        if definedNameDepth != nil { currentDefinedName?.formula.append(text) }
      case .endElement(let name, _):
        if name.namespaceURI == spreadsheetNamespace, name.localName == "definedName",
          definedNameDepth == depth, let completed = currentDefinedName
        {
          definedNames.append(
            OfficeDefinedName(
              name: completed.name,
              formula: completed.formula,
              localSheetIndex: completed.localSheetIndex,
              isHidden: completed.isHidden,
              description: completed.description,
              comment: completed.comment
            ))
          currentDefinedName = nil
          definedNameDepth = nil
        }
        depth -= 1
      case .startDocument, .endDocument:
        break
      }
    }

    let sheets = try rawSheets.enumerated().map { index, raw in
      guard
        let relationship = try package.relationship(
          identifiedBy: raw.relationshipID,
          from: .part(workbookPart.name)
        ) else {
        throw OfficeKitError.invalidPackage(
          "Worksheet \(raw.name) references missing relationship \(raw.relationshipID.rawValue)."
        )
      }
      let kind: OfficeSheetKind
      if relationship.type.isEquivalent(to: .worksheet) {
        kind = .worksheet
      } else if relationship.type.isEquivalent(to: .chartSheet) {
        kind = .chart
      } else if relationship.type.isEquivalent(to: .dialogSheet) {
        kind = .dialog
      } else if relationship.type.isEquivalent(to: .macroSheet)
        || relationship.type.isEquivalent(to: .internationalMacroSheet)
      {
        kind = .macro
      } else {
        kind = .unknown
      }
      return OfficeSheetReference(
        index: index,
        name: raw.name,
        identifier: raw.identifier,
        visibility: raw.visibility,
        relationshipID: raw.relationshipID,
        kind: kind,
        part: package.part(referencedBy: relationship),
        attachment: package.attachment(referencedBy: relationship)
      )
    }
    let worksheets = sheets.compactMap { sheet -> OfficeWorksheetReference? in
      guard sheet.kind == .worksheet, let part = sheet.part else { return nil }
      return OfficeWorksheetReference(
        index: sheet.index,
        name: sheet.name,
        identifier: sheet.identifier,
        visibility: sheet.visibility,
        relationshipID: sheet.relationshipID,
        part: part
      )
    }
    let relationships = try package.relationships(from: .part(workbookPart.name))
    return Values(
      sheets: sheets,
      worksheets: worksheets,
      dateSystem: dateSystem,
      properties: OfficeWorkbookProperties(
        dateSystem: dateSystem,
        codeName: codeName,
        filtersPrivacy: filtersPrivacy,
        updateLinks: updateLinks
      ),
      definedNames: definedNames,
      calculation: OfficeWorkbookCalculation(
        calculationIdentifier: calculationIdentifier,
        mode: calculationMode,
        calculatesFullyOnLoad: calculatesFullyOnLoad,
        forcesFullCalculation: forcesFullCalculation,
        iterates: iterates,
        iterationCount: iterationCount,
        iterationDelta: iterationDelta
      ),
      attachments: relationships.map(package.attachment(referencedBy:))
    )
  }
}

private struct SpreadsheetRichTextRunBuilder {
  let startDepth: Int
  var text = ""
  var fontName: String?
  var sizeInPoints: Double?
  var isBold: Bool?
  var isItalic: Bool?
  var underline: String?
  var isStruckThrough: Bool?
  var color: OfficeSpreadsheetColor?
  var verticalAlignment: String?
  var family: UInt32?
  var characterSet: UInt32?
  var scheme: String?

  mutating func consume(
    localName: String,
    attributes: [OfficeXMLAttribute]
  ) {
    switch localName {
    case "rFont": fontName = attribute("val", in: attributes)
    case "sz": sizeInPoints = attribute("val", in: attributes).flatMap(Double.init)
    case "b": isBold = richTextBoolean(attributes)
    case "i": isItalic = richTextBoolean(attributes)
    case "u":
      underline =
        richTextBoolean(attributes)
        ? attribute("val", in: attributes) ?? "single" : nil
    case "strike": isStruckThrough = richTextBoolean(attributes)
    case "color": color = richTextColor(attributes)
    case "vertAlign": verticalAlignment = attribute("val", in: attributes)
    case "family": family = attribute("val", in: attributes).flatMap(UInt32.init)
    case "charset": characterSet = attribute("val", in: attributes).flatMap(UInt32.init)
    case "scheme": scheme = attribute("val", in: attributes)
    default: break
    }
  }

  var value: OfficeSpreadsheetTextRun {
    OfficeSpreadsheetTextRun(
      text: text,
      properties: OfficeSpreadsheetTextRunProperties(
        fontName: fontName,
        sizeInPoints: sizeInPoints,
        isBold: isBold,
        isItalic: isItalic,
        underline: underline,
        isStruckThrough: isStruckThrough,
        color: color,
        verticalAlignment: verticalAlignment,
        family: family,
        characterSet: characterSet,
        scheme: scheme
      )
    )
  }
}

private enum SpreadsheetSharedStringParser {
  private static let spreadsheetNamespace =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  static func parse(
    package: OfficePackage,
    workbookPart: OfficePart
  ) throws -> [OfficeSpreadsheetRichText] {
    guard
      let relationship = try package.relationships(
        from: .part(workbookPart.name),
        ofType: .sharedStrings
      ).first else { return [] }
    guard let part = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }

    var strings: [OfficeSpreadsheetRichText] = []
    var currentText: String?
    var currentRuns: [OfficeSpreadsheetTextRun] = []
    var run: SpreadsheetRichTextRunBuilder?
    var plainRunText: String?
    var textDepth: Int?
    var phoneticDepth: Int?
    var depth = 0
    try package.parseXML(in: part, compatibility: .commonOffice) { event in
      switch event {
      case .startElement(let name, let attributes, _, _):
        depth += 1
        guard name.namespaceURI == spreadsheetNamespace else { return }
        if name.localName == "si" {
          currentText = ""
          currentRuns = []
        } else if name.localName == "r", currentText != nil, phoneticDepth == nil {
          run = SpreadsheetRichTextRunBuilder(startDepth: depth)
        } else if name.localName == "rPh" {
          phoneticDepth = depth
        } else if name.localName == "t", currentText != nil, phoneticDepth == nil {
          textDepth = depth
          if run == nil { plainRunText = "" }
        } else if run != nil {
          run?.consume(localName: name.localName, attributes: attributes)
        }
      case .text(let text, _):
        if textDepth != nil {
          currentText?.append(text)
          if run != nil { run?.text.append(text) } else { plainRunText?.append(text) }
        }
      case .endElement(let name, _):
        if textDepth == depth, name.namespaceURI == spreadsheetNamespace,
          name.localName == "t"
        {
          textDepth = nil
          if let completedPlainRun = plainRunText {
            currentRuns.append(
              OfficeSpreadsheetTextRun(
                text: completedPlainRun,
                properties: .none
              ))
            plainRunText = nil
          }
        }
        if name.namespaceURI == spreadsheetNamespace, name.localName == "r",
          run?.startDepth == depth, let completed = run
        {
          currentRuns.append(completed.value)
          run = nil
        }
        if name.namespaceURI == spreadsheetNamespace, name.localName == "rPh",
          phoneticDepth == depth
        {
          phoneticDepth = nil
        }
        if name.namespaceURI == spreadsheetNamespace, name.localName == "si",
          let completed = currentText
        {
          guard strings.count < package.parsingLimits.maximumSharedStringCount else {
            throw OfficeKitError.limitExceeded(
              limit: .sharedStringCount,
              actual: UInt64(strings.count + 1),
              maximum: UInt64(package.parsingLimits.maximumSharedStringCount)
            )
          }
          strings.append(OfficeSpreadsheetRichText(text: completed, runs: currentRuns))
          currentText = nil
          currentRuns = []
        }
        depth -= 1
      case .startDocument, .endDocument:
        break
      }
    }
    return strings
  }

}

private final class SpreadsheetWorksheetParser {
  private static let spreadsheetNamespace =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  private static let relationshipNamespace =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  private struct RawCell {
    let reference: OfficeCellReference
    let styleIndex: UInt32?
    let sourceType: String?
    var rawValue: String?
    var inlineString = ""
    var inlineRuns: [OfficeSpreadsheetTextRun] = []
    var inlineRun: SpreadsheetRichTextRunBuilder?
    var inlinePlainRunText: String?
    var inlinePhoneticDepth: Int?
    var formulaText: String?
    var formulaType: String?
    var formulaSharedIndex: UInt32?
    var formulaReference: String?
    var valueDepth: Int?
    var formulaDepth: Int?
    var inlineTextDepth: Int?
  }

  private struct RawRow {
    let index: Int
    let height: Double?
    let isHidden: Bool
    let styleIndex: UInt32?
    let outlineLevel: UInt8?
    let isCollapsed: Bool
    var cells: [OfficeCell] = []
  }

  private struct ViewBuilder {
    let startDepth: Int
    let workbookViewIndex: UInt32?
    let isTabSelected: Bool
    let showGridLines: Bool?
    let showRowAndColumnHeaders: Bool?
    let topLeftCell: OfficeCellReference?
    let zoomScale: UInt32?
    var pane: OfficeWorksheetPane?
    var selections: [OfficeWorksheetSelection] = []
  }

  private enum HeaderFooterTextTarget {
    case oddHeader, oddFooter, evenHeader, evenFooter, firstHeader, firstFooter
  }

  private struct HeaderFooterBuilder {
    let startDepth: Int
    let differentOddAndEven: Bool?
    let differentFirst: Bool?
    let scalesWithDocument: Bool?
    let alignsWithPageMargins: Bool?
    var oddHeader: String?
    var oddFooter: String?
    var evenHeader: String?
    var evenFooter: String?
    var firstHeader: String?
    var firstFooter: String?
  }

  private struct ConditionalFormattingBuilder {
    let startDepth: Int
    let ranges: [OfficeCellRange]
    var rules: [OfficeConditionalFormattingRule] = []
  }

  private struct ConditionalRuleBuilder {
    let startDepth: Int
    let type: String
    let priority: UInt32?
    let differentialStyleIdentifier: UInt32?
    let comparisonOperator: String?
    let text: String?
    let timePeriod: String?
    let rank: UInt32?
    let selectsBottomValues: Bool?
    let rankIsPercentage: Bool?
    let stopsIfTrue: Bool?
    var formulas: [String] = []
    var dataBar: OfficeConditionalDataBar?
    var colorScale: OfficeConditionalColorScale?
    var iconSet: OfficeConditionalIconSet?

    var value: OfficeConditionalFormattingRule {
      OfficeConditionalFormattingRule(
        type: type,
        priority: priority,
        differentialStyleIdentifier: differentialStyleIdentifier,
        comparisonOperator: comparisonOperator,
        text: text,
        timePeriod: timePeriod,
        rank: rank,
        selectsBottomValues: selectsBottomValues,
        rankIsPercentage: rankIsPercentage,
        stopsIfTrue: stopsIfTrue,
        formulas: formulas,
        dataBar: dataBar,
        colorScale: colorScale,
        iconSet: iconSet
      )
    }
  }

  private enum ObjectMarkerTarget {
    case from
    case to
  }

  private struct ObjectMarkerBuilder {
    var columnIndex: Int?
    var columnOffset: Int64?
    var rowIndex: Int?
    var rowOffset: Int64?

    var value: OfficeWorksheetCellMarker? {
      guard let columnIndex, let columnOffset, let rowIndex, let rowOffset else { return nil }
      return OfficeWorksheetCellMarker(
        columnIndex: columnIndex,
        columnOffset: OfficeLength(emu: columnOffset),
        rowIndex: rowIndex,
        rowOffset: OfficeLength(emu: rowOffset)
      )
    }
  }

  private struct WorksheetObjectBuilder {
    let startDepth: Int
    let kind: OfficeWorksheetObjectKind
    let shapeIdentifier: UInt32?
    let name: String?
    let programIdentifier: String?
    let relationshipID: OfficeRelationshipID?
    var previewRelationshipID: OfficeRelationshipID?
    var usesDefaultSize: Bool?
    var movesWithCells: Bool?
    var sizesWithCells: Bool?
    var from = ObjectMarkerBuilder()
    var to = ObjectMarkerBuilder()
  }

  private let reference: OfficeWorksheetReference
  private let package: OfficePackage
  private let sharedStrings: [OfficeSpreadsheetRichText]
  private let styles: [OfficeCellStyle]
  private let dateSystem: OfficeSpreadsheetDateSystem
  private let relationshipsByID: [OfficeRelationshipID: OfficeRelationship]
  private let retainsRows: Bool
  private let collectsWorksheetMetadata: Bool
  private let rowHandler: ((OfficeWorksheetRow) throws -> Void)?
  private var depth = 0
  private var dimensionReference: String?
  private var columns: [OfficeWorksheetColumn] = []
  private var views: [OfficeWorksheetView] = []
  private var currentView: ViewBuilder?
  private var rows: [OfficeWorksheetRow] = []
  private var mergedRanges: [OfficeCellRange] = []
  private var autoFilterRange: OfficeCellRange?
  private var pageMargins: OfficeWorksheetPageMargins?
  private var pageSetup: OfficeWorksheetPageSetup?
  private var headerFooter: OfficeWorksheetHeaderFooter?
  private var headerFooterBuilder: HeaderFooterBuilder?
  private var headerFooterTextTarget: HeaderFooterTextTarget?
  private var headerFooterTextDepth: Int?
  private var conditionalFormatting: [OfficeConditionalFormatting] = []
  private var conditionalFormattingBuilder: ConditionalFormattingBuilder?
  private var conditionalRuleBuilder: ConditionalRuleBuilder?
  private var conditionalFormula = ""
  private var conditionalFormulaDepth: Int?
  private var conditionalVisualDepth: Int?
  private var conditionalThresholds: [OfficeConditionalFormatThreshold] = []
  private var conditionalColors: [OfficeSpreadsheetColor] = []
  private var objects: [OfficeWorksheetObject] = []
  private var unresolvedObjects: [WorksheetObjectBuilder] = []
  private var objectBuilder: WorksheetObjectBuilder?
  private var objectMarkerTarget: ObjectMarkerTarget?
  private var objectMarkerName: String?
  private var objectMarkerDepth: Int?
  private var objectMarkerText = ""
  private var hyperlinks: [OfficeWorksheetHyperlink] = []
  private var currentRow: RawRow?
  private var currentCell: RawCell?
  private var nextRowIndex = 0
  private var nextColumnIndex = 0
  private var retainedCellCount = 0

  private init(
    reference: OfficeWorksheetReference,
    package: OfficePackage,
    sharedStrings: [OfficeSpreadsheetRichText],
    styles: [OfficeCellStyle],
    dateSystem: OfficeSpreadsheetDateSystem,
    relationshipsByID: [OfficeRelationshipID: OfficeRelationship],
    retainsRows: Bool,
    collectsWorksheetMetadata: Bool,
    rowHandler: ((OfficeWorksheetRow) throws -> Void)?
  ) {
    self.reference = reference
    self.package = package
    self.sharedStrings = sharedStrings
    self.styles = styles
    self.dateSystem = dateSystem
    self.relationshipsByID = relationshipsByID
    self.retainsRows = retainsRows
    self.collectsWorksheetMetadata = collectsWorksheetMetadata
    self.rowHandler = rowHandler
  }

  static func parse(
    reference: OfficeWorksheetReference,
    package: OfficePackage,
    sharedStrings: [OfficeSpreadsheetRichText],
    styles: [OfficeCellStyle],
    dateSystem: OfficeSpreadsheetDateSystem
  ) throws -> OfficeWorksheet {
    let relationships = try package.relationships(from: .part(reference.part.name))
    let relationshipsByID = Dictionary(
      relationships.map { ($0.id, $0) },
      uniquingKeysWith: { first, _ in first }
    )
    let parser = SpreadsheetWorksheetParser(
      reference: reference,
      package: package,
      sharedStrings: sharedStrings,
      styles: styles,
      dateSystem: dateSystem,
      relationshipsByID: relationshipsByID,
      retainsRows: true,
      collectsWorksheetMetadata: true,
      rowHandler: nil
    )
    try package.parseXML(in: reference.part, compatibility: .commonOffice, parser.consume)
    try parser.resolveWorksheetObjects()
    let relatedAttachments = relationships.map {
      package.attachment(referencedBy: $0)
    }
    let drawings = try relationships.compactMap { relationship -> OfficeWorksheetDrawing? in
      guard relationship.type.isEquivalent(to: .drawing) else { return nil }
      guard let drawingPart = package.part(referencedBy: relationship) else {
        throw OfficeKitError.missingPart(relationship.rawTarget)
      }
      return try SpreadsheetDrawingParser.parse(part: drawingPart, package: package)
    }
    let tables = try relationships.compactMap {
      relationship -> OfficeSpreadsheetTableReference? in
      guard relationship.type.isEquivalent(to: .table) else { return nil }
      guard let tablePart = package.part(referencedBy: relationship) else {
        throw OfficeKitError.missingPart(relationship.rawTarget)
      }
      return OfficeSpreadsheetTableReference(
        package: package,
        attachment: package.attachment(referencedBy: relationship),
        part: tablePart
      )
    }
    let commentShapes = try relationships.compactMap {
      relationship -> [OfficeCellReference: OfficeWorksheetCommentShape]? in
      guard relationship.type.isEquivalent(to: .vmlDrawing) else { return nil }
      guard let vmlPart = package.part(referencedBy: relationship) else {
        throw OfficeKitError.missingPart(relationship.rawTarget)
      }
      return try SpreadsheetVMLCommentParser.parse(part: vmlPart, package: package)
    }.reduce(into: [:]) { result, shapes in
      result.merge(shapes, uniquingKeysWith: { first, _ in first })
    }
    let parsedComments = try relationships.compactMap {
      relationship -> [OfficeWorksheetComment]? in
      guard relationship.type.isEquivalent(to: .comments) else { return nil }
      guard let commentPart = package.part(referencedBy: relationship) else {
        throw OfficeKitError.missingPart(relationship.rawTarget)
      }
      return try SpreadsheetCommentParser.parse(part: commentPart, package: package)
    }.flatMap { $0 }
    let comments = parsedComments.map { comment in
      OfficeWorksheetComment(
        reference: comment.reference,
        authorIndex: comment.authorIndex,
        author: comment.author,
        shapeIdentifier: comment.shapeIdentifier,
        text: comment.text,
        sourcePart: comment.sourcePart,
        shape: commentShapes[comment.reference]
      )
    }
    let attachments = uniqueSpreadsheetAttachments(
      relatedAttachments + drawings.flatMap(\.attachments) + parser.objects.flatMap(\.attachments)
    )
    return OfficeWorksheet(
      reference: reference,
      dimensionReference: parser.dimensionReference,
      columns: parser.columns,
      views: parser.views,
      rows: parser.rows,
      mergedRanges: parser.mergedRanges,
      autoFilterRange: parser.autoFilterRange,
      pageMargins: parser.pageMargins,
      pageSetup: parser.pageSetup,
      headerFooter: parser.headerFooter,
      conditionalFormatting: parser.conditionalFormatting,
      objects: parser.objects,
      hyperlinks: parser.hyperlinks,
      drawings: drawings,
      tables: tables,
      comments: comments,
      attachments: attachments
    )
  }

  static func streamRows(
    reference: OfficeWorksheetReference,
    package: OfficePackage,
    sharedStrings: [OfficeSpreadsheetRichText],
    styles: [OfficeCellStyle],
    dateSystem: OfficeSpreadsheetDateSystem,
    _ body: @escaping (OfficeWorksheetRow) throws -> Void
  ) throws {
    let parser = SpreadsheetWorksheetParser(
      reference: reference,
      package: package,
      sharedStrings: sharedStrings,
      styles: styles,
      dateSystem: dateSystem,
      relationshipsByID: [:],
      retainsRows: false,
      collectsWorksheetMetadata: false,
      rowHandler: body
    )
    try package.parseXML(in: reference.part, compatibility: .commonOffice, parser.consume)
  }

  private func consume(_ event: OfficeXMLEvent) throws {
    switch event {
    case .startElement(let name, let attributes, _, _):
      depth += 1
      if collectsWorksheetMetadata {
        consumeWorksheetObjectStart(name: name, attributes: attributes)
      }
      guard name.namespaceURI == Self.spreadsheetNamespace else { return }
      switch name.localName {
      case "dimension":
        dimensionReference = attribute("ref", in: attributes)
      case "col" where collectsWorksheetMetadata:
        guard let first = attribute("min", in: attributes).flatMap(Int.init), first > 0,
          let last = attribute("max", in: attributes).flatMap(Int.init),
          last >= first, last <= 16_384 else {
          throw OfficeKitError.invalidPackage("Worksheet contains an invalid column range.")
        }
        columns.append(
          OfficeWorksheetColumn(
            firstIndex: first - 1,
            lastIndex: last - 1,
            width: attribute("width", in: attributes).flatMap(Double.init),
            isHidden: attribute("hidden", in: attributes)
              .flatMap(OfficeValueDecoder.boolean) ?? false,
            isBestFit: attribute("bestFit", in: attributes)
              .flatMap(OfficeValueDecoder.boolean) ?? false,
            styleIndex: attribute("style", in: attributes).flatMap(UInt32.init),
            outlineLevel: attribute("outlineLevel", in: attributes).flatMap(UInt8.init),
            isCollapsed: attribute("collapsed", in: attributes)
              .flatMap(OfficeValueDecoder.boolean) ?? false
          ))
      case "sheetView" where collectsWorksheetMetadata:
        currentView = ViewBuilder(
          startDepth: depth,
          workbookViewIndex: attribute("workbookViewId", in: attributes).flatMap(UInt32.init),
          isTabSelected: attribute("tabSelected", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false,
          showGridLines: attribute("showGridLines", in: attributes)
            .flatMap(OfficeValueDecoder.boolean),
          showRowAndColumnHeaders: attribute("showRowColHeaders", in: attributes)
            .flatMap(OfficeValueDecoder.boolean),
          topLeftCell: attribute("topLeftCell", in: attributes)
            .flatMap(OfficeCellReference.init(rawValue:)),
          zoomScale: attribute("zoomScale", in: attributes).flatMap(UInt32.init)
        )
      case "pane" where currentView != nil:
        currentView?.pane = OfficeWorksheetPane(
          horizontalSplit: attribute("xSplit", in: attributes).flatMap(Double.init),
          verticalSplit: attribute("ySplit", in: attributes).flatMap(Double.init),
          topLeftCell: attribute("topLeftCell", in: attributes)
            .flatMap(OfficeCellReference.init(rawValue:)),
          activePane: attribute("activePane", in: attributes),
          state: attribute("state", in: attributes)
        )
      case "selection" where currentView != nil:
        currentView?.selections.append(
          OfficeWorksheetSelection(
            pane: attribute("pane", in: attributes),
            activeCell: attribute("activeCell", in: attributes)
              .flatMap(OfficeCellReference.init(rawValue:)),
            ranges: attribute("sqref", in: attributes)?.split(separator: " ").map(String.init) ?? []
          ))
      case "row":
        let rowNumber = attribute("r", in: attributes).flatMap(Int.init)
        let index = rowNumber.map { $0 - 1 } ?? nextRowIndex
        guard index >= 0, index < 1_048_576 else {
          throw OfficeKitError.invalidPackage("Worksheet contains an invalid row index.")
        }
        currentRow = RawRow(
          index: index,
          height: attribute("ht", in: attributes).flatMap(Double.init),
          isHidden: attribute("hidden", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false,
          styleIndex: attribute("s", in: attributes).flatMap(UInt32.init),
          outlineLevel: attribute("outlineLevel", in: attributes).flatMap(UInt8.init),
          isCollapsed: attribute("collapsed", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false
        )
        nextColumnIndex = 0
      case "c":
        guard let currentRow else {
          throw OfficeKitError.invalidPackage("Worksheet cell appears outside a row.")
        }
        let rawReference =
          attribute("r", in: attributes)
          ?? spreadsheetA1Reference(columnIndex: nextColumnIndex, rowIndex: currentRow.index)
        guard let cellReference = OfficeCellReference(rawValue: rawReference),
          cellReference.rowIndex == currentRow.index else {
          throw OfficeKitError.invalidPackage(
            "Worksheet contains invalid cell reference \(rawReference).")
        }
        currentCell = RawCell(
          reference: cellReference,
          styleIndex: attribute("s", in: attributes).flatMap(UInt32.init),
          sourceType: attribute("t", in: attributes)
        )
      case "v":
        currentCell?.rawValue = ""
        currentCell?.valueDepth = depth
      case "f":
        currentCell?.formulaText = ""
        currentCell?.formulaType = attribute("t", in: attributes)
        currentCell?.formulaSharedIndex = attribute("si", in: attributes).flatMap(UInt32.init)
        currentCell?.formulaReference = attribute("ref", in: attributes)
        currentCell?.formulaDepth = depth
      case "r"
      where currentCell?.sourceType == "inlineStr"
        && currentCell?.inlinePhoneticDepth == nil:
        currentCell?.inlineRun = SpreadsheetRichTextRunBuilder(startDepth: depth)
      case "rPh" where currentCell?.sourceType == "inlineStr":
        currentCell?.inlinePhoneticDepth = depth
      case "t":
        if currentCell?.sourceType == "inlineStr", currentCell?.inlinePhoneticDepth == nil {
          currentCell?.inlineTextDepth = depth
          if currentCell?.inlineRun == nil { currentCell?.inlinePlainRunText = "" }
        }
      case _ where currentCell?.inlineRun != nil:
        currentCell?.inlineRun?.consume(localName: name.localName, attributes: attributes)
      case "mergeCell":
        guard collectsWorksheetMetadata else { break }
        guard let rawRange = attribute("ref", in: attributes),
          let range = OfficeCellRange(rawValue: rawRange) else {
          throw OfficeKitError.invalidPackage("Worksheet contains an invalid merged-cell range.")
        }
        mergedRanges.append(range)
      case "hyperlink":
        if collectsWorksheetMetadata { try readHyperlink(attributes: attributes) }
      case "autoFilter" where collectsWorksheetMetadata:
        if let rawRange = attribute("ref", in: attributes) {
          guard let range = OfficeCellRange(rawValue: rawRange) else {
            throw OfficeKitError.invalidPackage("Worksheet contains an invalid autofilter range.")
          }
          autoFilterRange = range
        }
      case "pageMargins" where collectsWorksheetMetadata:
        pageMargins = OfficeWorksheetPageMargins(
          left: attribute("left", in: attributes).flatMap(Double.init),
          right: attribute("right", in: attributes).flatMap(Double.init),
          top: attribute("top", in: attributes).flatMap(Double.init),
          bottom: attribute("bottom", in: attributes).flatMap(Double.init),
          header: attribute("header", in: attributes).flatMap(Double.init),
          footer: attribute("footer", in: attributes).flatMap(Double.init)
        )
      case "pageSetup" where collectsWorksheetMetadata:
        pageSetup = OfficeWorksheetPageSetup(
          paperSize: attribute("paperSize", in: attributes).flatMap(UInt32.init),
          scale: attribute("scale", in: attributes).flatMap(UInt32.init),
          firstPageNumber: attribute("firstPageNumber", in: attributes).flatMap(UInt32.init),
          fitToWidth: attribute("fitToWidth", in: attributes).flatMap(UInt32.init),
          fitToHeight: attribute("fitToHeight", in: attributes).flatMap(UInt32.init),
          orientation: attribute("orientation", in: attributes),
          pageOrder: attribute("pageOrder", in: attributes),
          blackAndWhite: attribute("blackAndWhite", in: attributes)
            .flatMap(OfficeValueDecoder.boolean),
          draft: attribute("draft", in: attributes).flatMap(OfficeValueDecoder.boolean),
          relationshipID: attribute("id", namespace: Self.relationshipNamespace, in: attributes)
            .map(OfficeRelationshipID.init(rawValue:))
        )
      case "headerFooter" where collectsWorksheetMetadata:
        headerFooterBuilder = HeaderFooterBuilder(
          startDepth: depth,
          differentOddAndEven: attribute("differentOddEven", in: attributes)
            .flatMap(OfficeValueDecoder.boolean),
          differentFirst: attribute("differentFirst", in: attributes)
            .flatMap(OfficeValueDecoder.boolean),
          scalesWithDocument: attribute("scaleWithDoc", in: attributes)
            .flatMap(OfficeValueDecoder.boolean),
          alignsWithPageMargins: attribute("alignWithMargins", in: attributes)
            .flatMap(OfficeValueDecoder.boolean)
        )
      case "conditionalFormatting" where collectsWorksheetMetadata:
        guard let references = attribute("sqref", in: attributes) else {
          throw OfficeKitError.invalidPackage(
            "Worksheet contains conditional formatting without a target range."
          )
        }
        let ranges = references.split(whereSeparator: { $0.isWhitespace }).map(String.init)
        let parsedRanges = ranges.compactMap(OfficeCellRange.init(rawValue:))
        guard !parsedRanges.isEmpty, parsedRanges.count == ranges.count else {
          throw OfficeKitError.invalidPackage(
            "Worksheet contains an invalid conditional-format range."
          )
        }
        conditionalFormattingBuilder = ConditionalFormattingBuilder(
          startDepth: depth,
          ranges: parsedRanges
        )
      case "cfRule" where conditionalFormattingBuilder != nil:
        guard let type = attribute("type", in: attributes) else {
          throw OfficeKitError.invalidPackage("Conditional-format rule has no type.")
        }
        conditionalRuleBuilder = ConditionalRuleBuilder(
          startDepth: depth,
          type: type,
          priority: attribute("priority", in: attributes).flatMap(UInt32.init),
          differentialStyleIdentifier: attribute("dxfId", in: attributes).flatMap(UInt32.init),
          comparisonOperator: attribute("operator", in: attributes),
          text: attribute("text", in: attributes),
          timePeriod: attribute("timePeriod", in: attributes),
          rank: attribute("rank", in: attributes).flatMap(UInt32.init),
          selectsBottomValues: attribute("bottom", in: attributes)
            .flatMap(OfficeValueDecoder.boolean),
          rankIsPercentage: attribute("percent", in: attributes)
            .flatMap(OfficeValueDecoder.boolean),
          stopsIfTrue: attribute("stopIfTrue", in: attributes)
            .flatMap(OfficeValueDecoder.boolean)
        )
      case "formula" where conditionalRuleBuilder != nil:
        conditionalFormula = ""
        conditionalFormulaDepth = depth
      case "dataBar", "colorScale", "iconSet":
        guard conditionalRuleBuilder != nil else { break }
        conditionalVisualDepth = depth
        conditionalThresholds = []
        conditionalColors = []
        if name.localName == "dataBar" {
          conditionalRuleBuilder?.dataBar = OfficeConditionalDataBar(
            minimumLength: attribute("minLength", in: attributes).flatMap(UInt32.init),
            maximumLength: attribute("maxLength", in: attributes).flatMap(UInt32.init),
            showsValue: attribute("showValue", in: attributes)
              .flatMap(OfficeValueDecoder.boolean),
            thresholds: [],
            color: nil
          )
        } else if name.localName == "iconSet" {
          conditionalRuleBuilder?.iconSet = OfficeConditionalIconSet(
            name: attribute("iconSet", in: attributes),
            isReversed: attribute("reverse", in: attributes)
              .flatMap(OfficeValueDecoder.boolean),
            showsValue: attribute("showValue", in: attributes)
              .flatMap(OfficeValueDecoder.boolean),
            usesPercentages: attribute("percent", in: attributes)
              .flatMap(OfficeValueDecoder.boolean),
            thresholds: []
          )
        }
      case "cfvo" where conditionalVisualDepth != nil:
        guard let kind = attribute("type", in: attributes) else {
          throw OfficeKitError.invalidPackage("Conditional-format threshold has no type.")
        }
        conditionalThresholds.append(
          OfficeConditionalFormatThreshold(
            kind: kind,
            value: attribute("val", in: attributes),
            isGreaterThanOrEqual: attribute("gte", in: attributes)
              .flatMap(OfficeValueDecoder.boolean)
          ))
      case "color" where conditionalVisualDepth != nil:
        conditionalColors.append(richTextColor(attributes))
      case "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter":
        guard headerFooterBuilder != nil else { break }
        headerFooterTextTarget = headerFooterTarget(name.localName)
        headerFooterTextDepth = depth
      default:
        break
      }
    case .text(let text, _):
      if objectMarkerDepth != nil { objectMarkerText.append(text) }
      if currentCell?.valueDepth != nil {
        currentCell?.rawValue?.append(text)
      }
      if currentCell?.formulaDepth != nil {
        currentCell?.formulaText?.append(text)
      }
      if currentCell?.inlineTextDepth != nil {
        currentCell?.inlineString.append(text)
        if currentCell?.inlineRun != nil {
          currentCell?.inlineRun?.text.append(text)
        } else {
          currentCell?.inlinePlainRunText?.append(text)
        }
      }
      if conditionalFormulaDepth != nil { conditionalFormula.append(text) }
      appendHeaderFooterText(text)
    case .endElement(let name, _):
      if collectsWorksheetMetadata {
        try consumeWorksheetObjectEnd(name: name)
      }
      guard name.namespaceURI == Self.spreadsheetNamespace else {
        depth -= 1
        return
      }
      if currentCell?.valueDepth == depth, name.localName == "v" {
        currentCell?.valueDepth = nil
      }
      if currentCell?.formulaDepth == depth, name.localName == "f" {
        currentCell?.formulaDepth = nil
      }
      if currentCell?.inlineTextDepth == depth, name.localName == "t" {
        currentCell?.inlineTextDepth = nil
        if let plainRun = currentCell?.inlinePlainRunText {
          currentCell?.inlineRuns.append(
            OfficeSpreadsheetTextRun(
              text: plainRun,
              properties: .none
            ))
          currentCell?.inlinePlainRunText = nil
        }
      }
      if conditionalFormulaDepth == depth, name.localName == "formula" {
        conditionalRuleBuilder?.formulas.append(conditionalFormula)
        conditionalFormula = ""
        conditionalFormulaDepth = nil
      }
      if conditionalVisualDepth == depth {
        finishConditionalVisual(localName: name.localName)
      }
      if name.localName == "cfRule", conditionalRuleBuilder?.startDepth == depth,
        let rule = conditionalRuleBuilder
      {
        conditionalFormattingBuilder?.rules.append(rule.value)
        conditionalRuleBuilder = nil
      }
      if name.localName == "conditionalFormatting",
        conditionalFormattingBuilder?.startDepth == depth,
        let formatting = conditionalFormattingBuilder
      {
        conditionalFormatting.append(
          OfficeConditionalFormatting(
            ranges: formatting.ranges,
            rules: formatting.rules
          ))
        conditionalFormattingBuilder = nil
      }
      if name.localName == "r", currentCell?.inlineRun?.startDepth == depth,
        let run = currentCell?.inlineRun
      {
        currentCell?.inlineRuns.append(run.value)
        currentCell?.inlineRun = nil
      }
      if name.localName == "rPh", currentCell?.inlinePhoneticDepth == depth {
        currentCell?.inlinePhoneticDepth = nil
      }
      if name.localName == "c", let rawCell = currentCell {
        guard currentRow != nil else {
          throw OfficeKitError.invalidPackage("Worksheet cell closes outside a row.")
        }
        if retainsRows {
          guard retainedCellCount < package.parsingLimits.maximumRetainedWorksheetCells else {
            throw OfficeKitError.limitExceeded(
              limit: .retainedWorksheetCells,
              actual: UInt64(retainedCellCount + 1),
              maximum: UInt64(package.parsingLimits.maximumRetainedWorksheetCells)
            )
          }
          retainedCellCount += 1
        }
        currentRow?.cells.append(try makeCell(from: rawCell))
        nextColumnIndex = rawCell.reference.columnIndex + 1
        currentCell = nil
      }
      if name.localName == "row", let rawRow = currentRow {
        let completed = OfficeWorksheetRow(
          index: rawRow.index,
          height: rawRow.height,
          isHidden: rawRow.isHidden,
          styleIndex: rawRow.styleIndex,
          outlineLevel: rawRow.outlineLevel,
          isCollapsed: rawRow.isCollapsed,
          cells: rawRow.cells
        )
        if retainsRows {
          guard rows.count < package.parsingLimits.maximumRetainedWorksheetRows else {
            throw OfficeKitError.limitExceeded(
              limit: .retainedWorksheetRows,
              actual: UInt64(rows.count + 1),
              maximum: UInt64(package.parsingLimits.maximumRetainedWorksheetRows)
            )
          }
          rows.append(completed)
        }
        try rowHandler?(completed)
        nextRowIndex = rawRow.index + 1
        currentRow = nil
      }
      if name.localName == "sheetView", currentView?.startDepth == depth,
        let completed = currentView
      {
        views.append(
          OfficeWorksheetView(
            workbookViewIndex: completed.workbookViewIndex,
            isTabSelected: completed.isTabSelected,
            showGridLines: completed.showGridLines,
            showRowAndColumnHeaders: completed.showRowAndColumnHeaders,
            topLeftCell: completed.topLeftCell,
            zoomScale: completed.zoomScale,
            pane: completed.pane,
            selections: completed.selections
          ))
        currentView = nil
      }
      if headerFooterTextDepth == depth {
        headerFooterTextDepth = nil
        headerFooterTextTarget = nil
      }
      if name.localName == "headerFooter", headerFooterBuilder?.startDepth == depth,
        let completed = headerFooterBuilder
      {
        headerFooter = OfficeWorksheetHeaderFooter(
          differentOddAndEven: completed.differentOddAndEven,
          differentFirst: completed.differentFirst,
          scalesWithDocument: completed.scalesWithDocument,
          alignsWithPageMargins: completed.alignsWithPageMargins,
          oddHeader: completed.oddHeader,
          oddFooter: completed.oddFooter,
          evenHeader: completed.evenHeader,
          evenFooter: completed.evenFooter,
          firstHeader: completed.firstHeader,
          firstFooter: completed.firstFooter
        )
        headerFooterBuilder = nil
      }
      depth -= 1
    case .startDocument, .endDocument:
      break
    }
  }

  private func headerFooterTarget(_ localName: String) -> HeaderFooterTextTarget? {
    switch localName {
    case "oddHeader": .oddHeader
    case "oddFooter": .oddFooter
    case "evenHeader": .evenHeader
    case "evenFooter": .evenFooter
    case "firstHeader": .firstHeader
    case "firstFooter": .firstFooter
    default: nil
    }
  }

  private func appendHeaderFooterText(_ text: String) {
    guard headerFooterTextDepth != nil, let target = headerFooterTextTarget,
      var builder = headerFooterBuilder else { return }
    switch target {
    case .oddHeader: builder.oddHeader = (builder.oddHeader ?? "") + text
    case .oddFooter: builder.oddFooter = (builder.oddFooter ?? "") + text
    case .evenHeader: builder.evenHeader = (builder.evenHeader ?? "") + text
    case .evenFooter: builder.evenFooter = (builder.evenFooter ?? "") + text
    case .firstHeader: builder.firstHeader = (builder.firstHeader ?? "") + text
    case .firstFooter: builder.firstFooter = (builder.firstFooter ?? "") + text
    }
    headerFooterBuilder = builder
  }

  private func finishConditionalVisual(localName: String) {
    switch localName {
    case "dataBar":
      guard let dataBar = conditionalRuleBuilder?.dataBar else { break }
      conditionalRuleBuilder?.dataBar = OfficeConditionalDataBar(
        minimumLength: dataBar.minimumLength,
        maximumLength: dataBar.maximumLength,
        showsValue: dataBar.showsValue,
        thresholds: conditionalThresholds,
        color: conditionalColors.first
      )
    case "colorScale":
      conditionalRuleBuilder?.colorScale = OfficeConditionalColorScale(
        thresholds: conditionalThresholds,
        colors: conditionalColors
      )
    case "iconSet":
      guard let iconSet = conditionalRuleBuilder?.iconSet else { break }
      conditionalRuleBuilder?.iconSet = OfficeConditionalIconSet(
        name: iconSet.name,
        isReversed: iconSet.isReversed,
        showsValue: iconSet.showsValue,
        usesPercentages: iconSet.usesPercentages,
        thresholds: conditionalThresholds
      )
    default: break
    }
    conditionalVisualDepth = nil
    conditionalThresholds = []
    conditionalColors = []
  }

  private func consumeWorksheetObjectStart(
    name: OfficeXMLName,
    attributes: [OfficeXMLAttribute]
  ) {
    if name.namespaceURI == Self.spreadsheetNamespace {
      if name.localName == "oleObject" || name.localName == "control" {
        let kind: OfficeWorksheetObjectKind = name.localName == "oleObject" ? .ole : .control
        objectBuilder = WorksheetObjectBuilder(
          startDepth: depth,
          kind: kind,
          shapeIdentifier: attribute("shapeId", in: attributes).flatMap(UInt32.init),
          name: attribute("name", in: attributes),
          programIdentifier: attribute("progId", in: attributes),
          relationshipID: attribute(
            "id",
            namespace: Self.relationshipNamespace,
            in: attributes
          ).map(OfficeRelationshipID.init(rawValue:))
        )
      } else if objectBuilder != nil,
        name.localName == "objectPr" || name.localName == "controlPr"
      {
        objectBuilder?.previewRelationshipID = attribute(
          "id",
          namespace: Self.relationshipNamespace,
          in: attributes
        ).map(OfficeRelationshipID.init(rawValue:))
        objectBuilder?.usesDefaultSize = attribute("defaultSize", in: attributes)
          .flatMap(OfficeValueDecoder.boolean)
      } else if objectBuilder != nil, name.localName == "anchor" {
        objectBuilder?.movesWithCells = attribute("moveWithCells", in: attributes)
          .flatMap(OfficeValueDecoder.boolean)
        objectBuilder?.sizesWithCells = attribute("sizeWithCells", in: attributes)
          .flatMap(OfficeValueDecoder.boolean)
      } else if objectBuilder != nil, name.localName == "from" {
        objectMarkerTarget = .from
      } else if objectBuilder != nil, name.localName == "to" {
        objectMarkerTarget = .to
      }
    }
    guard objectBuilder != nil, objectMarkerTarget != nil,
      name.localName == "col" || name.localName == "colOff"
        || name.localName == "row" || name.localName == "rowOff" else { return }
    objectMarkerName = name.localName
    objectMarkerDepth = depth
    objectMarkerText = ""
  }

  private func consumeWorksheetObjectEnd(name: OfficeXMLName) throws {
    if objectMarkerDepth == depth, let markerName = objectMarkerName {
      guard let value = Int64(objectMarkerText.trimmingCharacters(in: .whitespacesAndNewlines)) else {
        throw OfficeKitError.invalidPackage("Worksheet object contains an invalid anchor value.")
      }
      try setObjectMarkerValue(value, named: markerName)
      objectMarkerName = nil
      objectMarkerDepth = nil
      objectMarkerText = ""
    }
    if name.namespaceURI == Self.spreadsheetNamespace,
      name.localName == "from" || name.localName == "to"
    {
      objectMarkerTarget = nil
    }
    guard name.namespaceURI == Self.spreadsheetNamespace,
      let completed = objectBuilder, completed.startDepth == depth,
      name.localName == "oleObject" || name.localName == "control" else { return }
    unresolvedObjects.append(completed)
    objectBuilder = nil
  }

  private func resolveWorksheetObjects() throws {
    objects = try unresolvedObjects.map(makeWorksheetObject)
    unresolvedObjects.removeAll(keepingCapacity: false)
  }

  private func setObjectMarkerValue(_ value: Int64, named name: String) throws {
    guard value >= 0, let target = objectMarkerTarget, var builder = objectBuilder else {
      throw OfficeKitError.invalidPackage("Worksheet object contains a negative anchor value.")
    }
    func update(_ marker: inout ObjectMarkerBuilder) throws {
      switch name {
      case "col":
        guard value < 16_384 else {
          throw OfficeKitError.invalidPackage("Worksheet object column exceeds the Excel grid.")
        }
        marker.columnIndex = Int(value)
      case "colOff": marker.columnOffset = value
      case "row":
        guard value < 1_048_576 else {
          throw OfficeKitError.invalidPackage("Worksheet object row exceeds the Excel grid.")
        }
        marker.rowIndex = Int(value)
      case "rowOff": marker.rowOffset = value
      default: break
      }
    }
    switch target {
    case .from: try update(&builder.from)
    case .to: try update(&builder.to)
    }
    objectBuilder = builder
  }

  private func makeWorksheetObject(
    _ builder: WorksheetObjectBuilder
  ) throws -> OfficeWorksheetObject {
    guard let relationshipID = builder.relationshipID,
      let relationship = relationshipsByID[relationshipID] else {
      throw OfficeKitError.invalidPackage("Worksheet object has no resolvable relationship.")
    }
    let hasExpectedType: Bool
    switch builder.kind {
    case .ole:
      hasExpectedType =
        relationship.type.isEquivalent(to: .oleObject)
        || relationship.type.isEquivalent(to: .embeddedPackage)
    case .control:
      hasExpectedType = relationship.type.isEquivalent(to: .control)
    }
    guard hasExpectedType else {
      throw OfficeKitError.invalidPackage("Worksheet object relationship has the wrong type.")
    }
    let attachment = package.attachment(referencedBy: relationship)
    let previewImage = builder.previewRelationshipID.flatMap { relationshipsByID[$0] }
      .map(package.attachment(referencedBy:))
    var attachments = [attachment]
    if let previewImage { attachments.append(previewImage) }
    if let part = attachment.part {
      attachments += try package.relationships(from: .part(part.name)).map {
        package.attachment(referencedBy: $0)
      }
    }
    let anchor = builder.from.value.flatMap { from in
      builder.to.value.map { to in
        OfficeWorksheetDrawingAnchor.twoCell(from: from, to: to, editBehavior: nil)
      }
    }
    return OfficeWorksheetObject(
      kind: builder.kind,
      shapeIdentifier: builder.shapeIdentifier,
      name: builder.name,
      programIdentifier: builder.programIdentifier,
      usesDefaultSize: builder.usesDefaultSize,
      movesWithCells: builder.movesWithCells,
      sizesWithCells: builder.sizesWithCells,
      anchor: anchor,
      attachment: attachment,
      previewImage: previewImage,
      attachments: uniqueSpreadsheetAttachments(attachments),
      spatialInfo: OfficeSpatialInfo(
        coordinateSpace: .worksheet,
        geometrySourcePart: reference.part.name,
        frame: nil,
        resolution: .unresolved(
          reason: "Worksheet object placement requires resolved row heights and column widths."
        )
      ),
      sourcePart: reference.part
    )
  }

  private func makeCell(from raw: RawCell) throws -> OfficeCell {
    let value: OfficeCellValue?
    let richText: OfficeSpreadsheetRichText?
    switch raw.sourceType {
    case "b":
      value = raw.rawValue.flatMap(OfficeValueDecoder.boolean).map(OfficeCellValue.boolean)
      richText = nil
    case "d":
      value = raw.rawValue.map(OfficeCellValue.date)
      richText = nil
    case "e":
      value = raw.rawValue.map(OfficeCellValue.error)
      richText = nil
    case "inlineStr":
      value = .string(raw.inlineString)
      richText = OfficeSpreadsheetRichText(
        text: raw.inlineString,
        runs: raw.inlineRuns
      )
    case "s":
      guard let text = raw.rawValue,
        let index = Int(text),
        sharedStrings.indices.contains(index) else {
        throw OfficeKitError.invalidPackage(
          "Cell \(raw.reference.rawValue) has an invalid shared-string index."
        )
      }
      value = .string(sharedStrings[index].text)
      richText = sharedStrings[index]
    case "str":
      value = raw.rawValue.map(OfficeCellValue.string)
      richText = nil
    case "n", nil:
      if let text = raw.rawValue {
        value =
          Double(text).map(OfficeCellValue.number)
          ?? .unknown(type: raw.sourceType ?? "n", rawValue: text)
      } else {
        value = nil
      }
      richText = nil
    case let sourceType?:
      value = raw.rawValue.map { .unknown(type: sourceType, rawValue: $0) }
      richText = nil
    }
    let formula = raw.formulaText.map {
      OfficeCellFormula(
        text: $0,
        type: raw.formulaType,
        sharedIndex: raw.formulaSharedIndex,
        reference: raw.formulaReference
      )
    }
    let style: OfficeCellStyle?
    if let styleIndex = raw.styleIndex {
      guard styles.indices.contains(Int(styleIndex)) else {
        throw OfficeKitError.invalidPackage(
          "Cell \(raw.reference.rawValue) references missing style \(styleIndex)."
        )
      }
      style = styles[Int(styleIndex)]
    } else {
      style = styles.first
    }
    let dateValue: Date?
    if case .number(let serial) = value, style?.numberFormat.isDate == true {
      dateValue = dateSystem.date(fromSerial: serial)
    } else {
      dateValue = nil
    }
    return OfficeCell(
      reference: raw.reference,
      styleIndex: raw.styleIndex,
      style: style,
      sourceType: raw.sourceType,
      rawValue: raw.rawValue,
      value: value,
      richText: richText,
      formula: formula,
      dateValue: dateValue
    )
  }

  private func readHyperlink(attributes: [OfficeXMLAttribute]) throws {
    guard let rawRange = attribute("ref", in: attributes),
      let range = OfficeCellRange(rawValue: rawRange) else {
      throw OfficeKitError.invalidPackage("Worksheet contains an invalid hyperlink range.")
    }
    let relationshipID = attribute(
      "id",
      namespace: Self.relationshipNamespace,
      in: attributes
    ).map(OfficeRelationshipID.init(rawValue:))
    let attachment: OfficeAttachment?
    if let relationshipID {
      guard let relationship = relationshipsByID[relationshipID],
        relationship.type.isEquivalent(to: .hyperlink) else {
        throw OfficeKitError.invalidPackage(
          "Worksheet hyperlink references missing relationship \(relationshipID.rawValue)."
        )
      }
      attachment = package.attachment(referencedBy: relationship)
    } else {
      attachment = nil
    }
    hyperlinks.append(
      OfficeWorksheetHyperlink(
        reference: range,
        location: attribute("location", in: attributes),
        display: attribute("display", in: attributes),
        tooltip: attribute("tooltip", in: attributes),
        attachment: attachment
      )
    )
  }

}

private func attribute(
  _ localName: String,
  namespace: String? = nil,
  in attributes: [OfficeXMLAttribute]
) -> String? {
  attributes.first {
    $0.name.localName == localName
      && (namespace == nil || $0.name.namespaceURI == namespace)
  }?.value
}

private func richTextBoolean(_ attributes: [OfficeXMLAttribute]) -> Bool {
  attribute("val", in: attributes).flatMap(OfficeValueDecoder.boolean) ?? true
}

private func richTextColor(
  _ attributes: [OfficeXMLAttribute]
) -> OfficeSpreadsheetColor {
  OfficeSpreadsheetColor(
    argb: attribute("rgb", in: attributes),
    indexed: attribute("indexed", in: attributes).flatMap(UInt32.init),
    theme: attribute("theme", in: attributes).flatMap(UInt32.init),
    tint: attribute("tint", in: attributes).flatMap(Double.init),
    isAutomatic: attribute("auto", in: attributes).flatMap(OfficeValueDecoder.boolean)
  )
}

private func uniqueSpreadsheetAttachments(
  _ attachments: [OfficeAttachment]
) -> [OfficeAttachment] {
  var relationships = Set<OfficeRelationship>()
  return attachments.filter { relationships.insert($0.relationship).inserted }
}

package func spreadsheetA1Reference(columnIndex: Int, rowIndex: Int) -> String {
  guard columnIndex >= 0, rowIndex >= 0 else { return "" }
  var column = columnIndex + 1
  var letters = ""
  while column > 0 {
    let remainder = (column - 1) % 26
    letters.insert(Character(UnicodeScalar(UInt8(65 + remainder))), at: letters.startIndex)
    column = (column - 1) / 26
  }
  return letters + String(rowIndex + 1)
}
