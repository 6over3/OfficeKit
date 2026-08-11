import Foundation

/// The serial-date epoch selected by an Excel workbook.
public enum OfficeSpreadsheetDateSystem: String, Sendable, Hashable, Codable {
  /// The Windows-compatible 1900 system, including Excel's fictitious serial day 60.
  case nineteenHundred

  /// The legacy Mac 1904 system.
  case nineteenOhFour

  /// Converts a serial value to an absolute UTC date when the serial represents a real day.
  ///
  /// Serial 60 in the 1900 system returns `nil` because 1900-02-29 did not exist.
  public func date(fromSerial serial: Double) -> Date? {
    guard serial.isFinite else { return nil }
    let secondsPerDay = 86_400.0
    switch self {
    case .nineteenHundred:
      guard serial < 60 || serial >= 61 else { return nil }
      let unixEpochSerial = serial < 60 ? 25_568.0 : 25_569.0
      return Date(timeIntervalSince1970: (serial - unixEpochSerial) * secondsPerDay)
    case .nineteenOhFour:
      return Date(timeIntervalSince1970: (serial - 24_107.0) * secondsPerDay)
    }
  }
}

/// An authored workbook-scoped name and its source formula or range expression.
public struct OfficeDefinedName: Sendable, Hashable, Codable {
  /// The case-insensitive workbook name.
  public let name: String

  /// Formula or range text exactly as stored, without evaluation.
  public let formula: String

  /// Zero-based worksheet scope, or `nil` for workbook scope.
  public let localSheetIndex: Int?

  /// Whether the name is hidden from normal workbook UI.
  public let isHidden: Bool

  /// Producer-authored description, when present.
  public let description: String?

  /// Producer-authored comment, when present.
  public let comment: String?
}

/// Calculation behavior declared by an Excel workbook.
public struct OfficeWorkbookCalculation: Sendable, Hashable, Codable {
  /// The producer calculation-engine identifier.
  public let calculationIdentifier: UInt32?

  /// Calculation mode such as `auto`, `autoNoTable`, or `manual`.
  public let mode: String?

  /// Whether a complete recalculation is requested when opening.
  public let calculatesFullyOnLoad: Bool?

  /// Whether a forced full calculation is requested.
  public let forcesFullCalculation: Bool?

  /// Whether iterative calculation is enabled.
  public let iterates: Bool?

  /// Maximum iterative-calculation count.
  public let iterationCount: UInt32?

  /// Maximum iterative-calculation change threshold.
  public let iterationDelta: Double?
}

/// Workbook application properties that affect parsing semantics.
public struct OfficeWorkbookProperties: Sendable, Hashable, Codable {
  /// The selected spreadsheet date epoch.
  public let dateSystem: OfficeSpreadsheetDateSystem

  /// Workbook VBA code name, when authored.
  public let codeName: String?

  /// Whether personally identifying information should be filtered when saving.
  public let filtersPrivacy: Bool?

  /// The external-link update policy token.
  public let updateLinks: String?
}

/// The number format selected by a cell format record.
public struct OfficeNumberFormat: Sendable, Hashable, Codable {
  /// The workbook or built-in number-format identifier.
  public let identifier: UInt32

  /// The custom or known built-in format code, when available.
  public let code: String?

  /// Whether the format code represents a calendar date or time.
  public let isDate: Bool
}

/// A color reference authored in SpreadsheetML.
public struct OfficeSpreadsheetColor: Sendable, Hashable, Codable {
  /// Eight-digit ARGB hexadecimal text, when directly authored.
  public let argb: String?

  /// Indexed palette position, when authored.
  public let indexed: UInt32?

  /// Workbook theme color index, when authored.
  public let theme: UInt32?

  /// Tint or shade adjustment from `-1` through `1`.
  public let tint: Double?

  /// Whether the producer requests an automatic color.
  public let isAutomatic: Bool?
}

/// One font record from a workbook style table.
public struct OfficeSpreadsheetFont: Sendable, Hashable, Codable {
  /// Typeface name, when declared.
  public let name: String?
  /// Font size measured in points.
  public let sizeInPoints: Double?
  /// Whether bold formatting is enabled.
  public let isBold: Bool
  /// Whether italic formatting is enabled.
  public let isItalic: Bool
  /// Underline style token, when enabled.
  public let underline: String?
  /// Whether strike-through formatting is enabled.
  public let isStruckThrough: Bool
  /// Authored font color.
  public let color: OfficeSpreadsheetColor?
  /// OOXML font-family classification.
  public let family: UInt32?
  /// Character-set identifier.
  public let characterSet: UInt32?
  /// Theme font scheme token.
  public let scheme: String?
}

/// One fill record from a workbook style table.
public struct OfficeSpreadsheetFill: Sendable, Hashable, Codable {
  /// Pattern fill token, when present.
  public let patternType: String?
  /// Pattern foreground color.
  public let foregroundColor: OfficeSpreadsheetColor?
  /// Pattern background color.
  public let backgroundColor: OfficeSpreadsheetColor?
}

/// One edge of a SpreadsheetML border.
public struct OfficeSpreadsheetBorderEdge: Sendable, Hashable, Codable {
  /// Border line-style token.
  public let style: String?
  /// Border color.
  public let color: OfficeSpreadsheetColor?
}

/// One border record from a workbook style table.
public struct OfficeSpreadsheetBorder: Sendable, Hashable, Codable {
  /// Leading edge in logical reading order.
  public let leading: OfficeSpreadsheetBorderEdge
  /// Trailing edge in logical reading order.
  public let trailing: OfficeSpreadsheetBorderEdge
  /// Top edge.
  public let top: OfficeSpreadsheetBorderEdge
  /// Bottom edge.
  public let bottom: OfficeSpreadsheetBorderEdge
  /// Diagonal edge style.
  public let diagonal: OfficeSpreadsheetBorderEdge
  /// Whether the diagonal rises from bottom-leading to top-trailing.
  public let diagonalUp: Bool
  /// Whether the diagonal falls from top-leading to bottom-trailing.
  public let diagonalDown: Bool
}

/// Direct cell alignment stored on a cell-format record.
public struct OfficeSpreadsheetAlignment: Sendable, Hashable, Codable {
  /// Horizontal alignment token.
  public let horizontal: String?
  /// Vertical alignment token.
  public let vertical: String?
  /// Authored OOXML text-rotation value.
  public let textRotation: Int?
  /// Whether cell text wraps.
  public let wrapsText: Bool?
  /// Whether text shrinks to fit the cell.
  public let shrinksToFit: Bool?
  /// Alignment indentation level.
  public let indent: Int?
  /// Relative indentation adjustment.
  public let relativeIndent: Int?
  /// OOXML reading-order value.
  public let readingOrder: UInt32?
  /// Whether the final line is justified.
  public let justifiesLastLine: Bool?
}

/// Direct cell protection stored on a cell-format record.
public struct OfficeSpreadsheetProtection: Sendable, Hashable, Codable {
  /// Whether editing is locked when sheet protection is active.
  public let isLocked: Bool?
  /// Whether formulas are hidden when sheet protection is active.
  public let isHidden: Bool?
}

/// Formatting authored on one SpreadsheetML rich-text run.
public struct OfficeSpreadsheetTextRunProperties: Sendable, Hashable, Codable {
  /// Typeface name.
  public let fontName: String?
  /// Font size measured in points.
  public let sizeInPoints: Double?
  /// Direct bold state.
  public let isBold: Bool?
  /// Direct italic state.
  public let isItalic: Bool?
  /// Underline style token.
  public let underline: String?
  /// Direct strike-through state.
  public let isStruckThrough: Bool?
  /// Direct text color.
  public let color: OfficeSpreadsheetColor?
  /// Vertical alignment token such as `superscript`.
  public let verticalAlignment: String?
  /// OOXML font-family classification.
  public let family: UInt32?
  /// Character-set identifier.
  public let characterSet: UInt32?
  /// Theme font scheme token.
  public let scheme: String?

  /// An unformatted run-property value.
  public static let none = Self(
    fontName: nil,
    sizeInPoints: nil,
    isBold: nil,
    isItalic: nil,
    underline: nil,
    isStruckThrough: nil,
    color: nil,
    verticalAlignment: nil,
    family: nil,
    characterSet: nil,
    scheme: nil
  )
}

/// One run in a shared or inline SpreadsheetML rich string.
public struct OfficeSpreadsheetTextRun: Sendable, Hashable, Codable {
  /// Visible text in source order.
  public let text: String
  /// Formatting authored for this run.
  public let properties: OfficeSpreadsheetTextRunProperties
}

/// Lossless visible text and authored runs for a SpreadsheetML string.
public struct OfficeSpreadsheetRichText: Sendable, Hashable, Codable {
  /// Plain visible text assembled from all runs.
  public let text: String
  /// Authored runs in source order.
  public let runs: [OfficeSpreadsheetTextRun]
}

/// The resolved style record selected by a worksheet cell.
public struct OfficeCellStyle: Sendable, Hashable, Codable {
  /// The zero-based cell-format index.
  public let index: UInt32

  /// The resolved number format.
  public let numberFormat: OfficeNumberFormat

  /// The zero-based font-table index.
  public let fontIndex: UInt32

  /// The zero-based fill-table index.
  public let fillIndex: UInt32

  /// The zero-based border-table index.
  public let borderIndex: UInt32

  /// Resolved font record.
  public let font: OfficeSpreadsheetFont

  /// Resolved fill record.
  public let fill: OfficeSpreadsheetFill

  /// Resolved border record.
  public let border: OfficeSpreadsheetBorder

  /// Direct alignment overrides.
  public let alignment: OfficeSpreadsheetAlignment?

  /// Direct protection overrides.
  public let protection: OfficeSpreadsheetProtection?
}

/// The visibility state of a worksheet tab.
public enum OfficeWorksheetVisibility: String, Sendable, Hashable, Codable {
  /// The worksheet is visible in the workbook UI.
  case visible

  /// The worksheet is hidden but can be made visible through normal application UI.
  case hidden

  /// The worksheet is hidden and requires a programmatic or advanced action to reveal.
  case veryHidden
}

/// The semantic kind of a sheet listed by an Excel workbook.
public enum OfficeSheetKind: String, Sendable, Hashable, Codable {
  /// A grid worksheet that can be parsed into sparse rows and cells.
  case worksheet
  /// A chart-only sheet.
  case chart
  /// A legacy dialog sheet.
  case dialog
  /// A legacy Excel 4 macro sheet, retained as inert content.
  case macro
  /// A sheet relationship unknown to this OfficeKit version.
  case unknown
}

/// An ordered workbook sheet reference, including non-grid sheet kinds.
public struct OfficeSheetReference: Sendable {
  /// The zero-based position among all workbook sheets.
  public let index: Int
  /// The sheet name shown to the user.
  public let name: String
  /// The workbook-local numeric sheet identifier.
  public let identifier: UInt32
  /// The sheet tab visibility.
  public let visibility: OfficeWorksheetVisibility
  /// The relationship identifier used by the workbook.
  public let relationshipID: OfficeRelationshipID
  /// The semantic kind derived from the relationship type.
  public let kind: OfficeSheetKind
  /// The resolved sheet XML part, when its internal target exists.
  public let part: OfficePart?
  /// The relationship-backed sheet resource, available even for unknown or missing targets.
  public let attachment: OfficeAttachment
}

/// An A1 cell address with zero-based row and column indices.
public struct OfficeCellReference: RawRepresentable, Sendable, Hashable, Codable,
  CustomStringConvertible
{
  /// The source A1 spelling.
  public let rawValue: String

  /// The zero-based column index.
  public let columnIndex: Int

  /// The zero-based row index.
  public let rowIndex: Int

  /// Parses an A1 cell address.
  public init?(rawValue: String) {
    let bytes = Array(rawValue.utf8)
    var offset = 0
    var column = 0
    while offset < bytes.count {
      let byte = bytes[offset]
      guard byte >= 65, byte <= 90 else { break }
      let (multiplied, multiplyOverflow) = column.multipliedReportingOverflow(by: 26)
      let (next, addOverflow) = multiplied.addingReportingOverflow(Int(byte - 64))
      guard !multiplyOverflow, !addOverflow else { return nil }
      column = next
      offset += 1
    }
    guard column > 0, offset < bytes.count else { return nil }
    let rowText = String(decoding: bytes[offset...], as: UTF8.self)
    guard column <= 16_384, let rowNumber = Int(rowText),
      rowNumber > 0, rowNumber <= 1_048_576 else { return nil }

    self.rawValue = rawValue
    self.columnIndex = column - 1
    self.rowIndex = rowNumber - 1
  }

  /// The source A1 spelling.
  public var description: String { rawValue }
}

/// A formula authored for one worksheet cell.
public struct OfficeCellFormula: Sendable, Hashable, Codable {
  /// Formula text without the leading equals sign.
  public let text: String

  /// The OOXML formula kind, such as `shared`, `array`, or `dataTable`.
  public let type: String?

  /// The shared-formula index, when declared.
  public let sharedIndex: UInt32?

  /// The source range associated with an array or shared formula.
  public let reference: String?
}

/// A rectangular A1 cell range.
public struct OfficeCellRange: RawRepresentable, Sendable, Hashable, Codable,
  CustomStringConvertible
{
  /// The source A1 range spelling.
  public let rawValue: String

  /// The top-left range endpoint.
  public let start: OfficeCellReference

  /// The bottom-right range endpoint.
  public let end: OfficeCellReference

  /// Parses a single-cell or colon-separated A1 range.
  public init?(rawValue: String) {
    let endpoints = rawValue.split(separator: ":", omittingEmptySubsequences: false)
    guard endpoints.count == 1 || endpoints.count == 2,
      let start = OfficeCellReference(rawValue: String(endpoints[0])),
      let end = OfficeCellReference(rawValue: String(endpoints[endpoints.count - 1])),
      start.rowIndex <= end.rowIndex,
      start.columnIndex <= end.columnIndex else { return nil }
    self.rawValue = rawValue
    self.start = start
    self.end = end
  }

  /// The source A1 range spelling.
  public var description: String { rawValue }
}

/// A hyperlink applied to a worksheet cell or range.
public struct OfficeWorksheetHyperlink: Sendable {
  /// The linked worksheet cells.
  public let reference: OfficeCellRange

  /// An internal workbook target such as `Sheet2!A1`, when declared.
  public let location: String?

  /// Producer-authored display text, when declared.
  public let display: String?

  /// Producer-authored hover text, when declared.
  public let tooltip: String?

  /// The lazy external relationship, when the hyperlink targets a URL.
  public let attachment: OfficeAttachment?
}

/// A typed cached or literal SpreadsheetML cell value.
public enum OfficeCellValue: Sendable, Hashable, Codable {
  /// A boolean value.
  case boolean(Bool)

  /// An IEEE 754 numeric value.
  case number(Double)

  /// A shared, inline, literal, or formula-result string.
  case string(String)

  /// An Excel error token such as `#DIV/0!`.
  case error(String)

  /// An ISO 8601 date lexical value from a cell explicitly typed as a date.
  case date(String)

  /// A producer extension or malformed value retained without interpretation.
  case unknown(type: String, rawValue: String)
}

/// One populated cell from a worksheet's sparse cell stream.
public struct OfficeCell: Sendable, Hashable, Codable {
  /// The parsed A1 address.
  public let reference: OfficeCellReference

  /// The zero-based index into the workbook cell-format table, when declared.
  public let styleIndex: UInt32?

  /// The resolved cell format, when its style index is valid.
  public let style: OfficeCellStyle?

  /// The exact SpreadsheetML cell type token, when declared.
  public let sourceType: String?

  /// The exact cached or literal value text from `x:v`, when present.
  public let rawValue: String?

  /// The resolved typed value, including shared-string lookup.
  public let value: OfficeCellValue?

  /// Rich-text runs for shared or inline strings, when applicable.
  public let richText: OfficeSpreadsheetRichText?

  /// The authored formula and its sharing metadata, when present.
  public let formula: OfficeCellFormula?

  /// The cached numeric value interpreted through a date number format and workbook epoch.
  public let dateValue: Date?
}

/// One sparse worksheet row.
public struct OfficeWorksheetRow: Sendable, Hashable, Codable {
  /// The zero-based worksheet row index.
  public let index: Int

  /// The authored row height in points, when explicitly declared.
  public let height: Double?

  /// Whether the row is hidden.
  public let isHidden: Bool

  /// The row's default cell-format index, when authored.
  public let styleIndex: UInt32?

  /// The outline level from zero through seven.
  public let outlineLevel: UInt8?

  /// Whether the outlined row is collapsed.
  public let isCollapsed: Bool

  /// Populated cells in source order; omitted cells are not synthesized.
  public let cells: [OfficeCell]
}

/// A contiguous worksheet column range sharing authored properties.
public struct OfficeWorksheetColumn: Sendable, Hashable, Codable {
  /// Zero-based first column covered by this declaration.
  public let firstIndex: Int

  /// Zero-based last column covered by this declaration.
  public let lastIndex: Int

  /// Excel's authored character-based column width.
  public let width: Double?

  /// Whether these columns are hidden.
  public let isHidden: Bool

  /// Whether the producer marked the width as best-fit.
  public let isBestFit: Bool

  /// The default cell-format index for these columns.
  public let styleIndex: UInt32?

  /// The outline level from zero through seven.
  public let outlineLevel: UInt8?

  /// Whether the outlined columns are collapsed.
  public let isCollapsed: Bool
}

/// Frozen or split pane metadata from a worksheet view.
public struct OfficeWorksheetPane: Sendable, Hashable, Codable {
  /// Horizontal split position or frozen-column count.
  public let horizontalSplit: Double?
  /// Vertical split position or frozen-row count.
  public let verticalSplit: Double?
  /// Top-left visible cell in the lower-right pane.
  public let topLeftCell: OfficeCellReference?
  /// Active pane token.
  public let activePane: String?
  /// Pane state such as `frozen` or `split`.
  public let state: String?
}

/// Current selection metadata from a worksheet view.
public struct OfficeWorksheetSelection: Sendable, Hashable, Codable {
  /// Pane containing the selection.
  public let pane: String?
  /// Active cell within the selection.
  public let activeCell: OfficeCellReference?
  /// Authored selection ranges without normalization.
  public let ranges: [String]
}

/// One worksheet view as authored for a workbook window.
public struct OfficeWorksheetView: Sendable, Hashable, Codable {
  /// Zero-based workbook-view index associated with this view.
  public let workbookViewIndex: UInt32?
  /// Whether the sheet tab is selected.
  public let isTabSelected: Bool
  /// Whether grid lines are shown.
  public let showGridLines: Bool?
  /// Whether row and column headings are shown.
  public let showRowAndColumnHeaders: Bool?
  /// Top-left visible cell.
  public let topLeftCell: OfficeCellReference?
  /// View zoom percentage.
  public let zoomScale: UInt32?
  /// Frozen or split pane metadata.
  public let pane: OfficeWorksheetPane?
  /// Selections in source order.
  public let selections: [OfficeWorksheetSelection]
}

/// Physical page margins, measured in inches.
public struct OfficeWorksheetPageMargins: Sendable, Hashable, Codable {
  /// Left margin in inches.
  public let left: Double?
  /// Right margin in inches.
  public let right: Double?
  /// Top margin in inches.
  public let top: Double?
  /// Bottom margin in inches.
  public let bottom: Double?
  /// Header margin in inches.
  public let header: Double?
  /// Footer margin in inches.
  public let footer: Double?
}

/// Authored worksheet print-page setup.
public struct OfficeWorksheetPageSetup: Sendable, Hashable, Codable {
  /// OOXML paper-size identifier.
  public let paperSize: UInt32?
  /// Print scaling percentage.
  public let scale: UInt32?
  /// Authored first page number.
  public let firstPageNumber: UInt32?
  /// Target number of pages wide.
  public let fitToWidth: UInt32?
  /// Target number of pages tall.
  public let fitToHeight: UInt32?
  /// Page orientation token.
  public let orientation: String?
  /// Print page-order token.
  public let pageOrder: String?
  /// Whether printing is black and white.
  public let blackAndWhite: Bool?
  /// Whether draft-quality printing is requested.
  public let draft: Bool?
  /// Relationship to an external printer-settings part.
  public let relationshipID: OfficeRelationshipID?
}

/// Authored worksheet header and footer control strings.
public struct OfficeWorksheetHeaderFooter: Sendable, Hashable, Codable {
  /// Whether odd and even pages use different content.
  public let differentOddAndEven: Bool?
  /// Whether the first page uses different content.
  public let differentFirst: Bool?
  /// Whether header/footer content scales with the document.
  public let scalesWithDocument: Bool?
  /// Whether content aligns to page margins.
  public let alignsWithPageMargins: Bool?
  /// Odd-page header control string.
  public let oddHeader: String?
  /// Odd-page footer control string.
  public let oddFooter: String?
  /// Even-page header control string.
  public let evenHeader: String?
  /// Even-page footer control string.
  public let evenFooter: String?
  /// First-page header control string.
  public let firstHeader: String?
  /// First-page footer control string.
  public let firstFooter: String?
}

/// A threshold used by a conditional-format data bar, color scale, or icon set.
public struct OfficeConditionalFormatThreshold: Sendable, Hashable, Codable {
  /// The threshold kind, such as `min`, `max`, `num`, `percent`, or `formula`.
  public let kind: String
  /// The authored threshold value or formula.
  public let value: String?
  /// Whether the comparison is greater-than-or-equal rather than strictly greater-than.
  public let isGreaterThanOrEqual: Bool?
}

/// A conditional-format data-bar visualization.
public struct OfficeConditionalDataBar: Sendable, Hashable, Codable {
  /// Minimum bar length as an authored percentage.
  public let minimumLength: UInt32?
  /// Maximum bar length as an authored percentage.
  public let maximumLength: UInt32?
  /// Whether the underlying cell value remains visible.
  public let showsValue: Bool?
  /// Ordered lower and upper thresholds.
  public let thresholds: [OfficeConditionalFormatThreshold]
  /// The positive bar color.
  public let color: OfficeSpreadsheetColor?
}

/// A conditional-format color-scale visualization.
public struct OfficeConditionalColorScale: Sendable, Hashable, Codable {
  /// Ordered scale thresholds.
  public let thresholds: [OfficeConditionalFormatThreshold]
  /// Ordered colors corresponding to the thresholds.
  public let colors: [OfficeSpreadsheetColor]
}

/// A conditional-format icon-set visualization.
public struct OfficeConditionalIconSet: Sendable, Hashable, Codable {
  /// The icon-set token, such as `3TrafficLights1`.
  public let name: String?
  /// Whether icons are displayed in reverse order.
  public let isReversed: Bool?
  /// Whether the underlying cell value remains visible.
  public let showsValue: Bool?
  /// Whether numeric threshold values are percentages.
  public let usesPercentages: Bool?
  /// Ordered icon thresholds.
  public let thresholds: [OfficeConditionalFormatThreshold]
}

/// One authored worksheet conditional-format rule.
public struct OfficeConditionalFormattingRule: Sendable, Hashable, Codable {
  /// Rule type token such as `cellIs`, `dataBar`, `colorScale`, or `top10`.
  public let type: String
  /// Evaluation priority declared by the producer.
  public let priority: UInt32?
  /// Differential-style index applied by the rule.
  public let differentialStyleIdentifier: UInt32?
  /// Comparison operator for value-based rules.
  public let comparisonOperator: String?
  /// Text operand for text-matching rules.
  public let text: String?
  /// Time-period token for date-based rules.
  public let timePeriod: String?
  /// Rank for top/bottom rules.
  public let rank: UInt32?
  /// Whether a top/bottom rule selects the bottom values.
  public let selectsBottomValues: Bool?
  /// Whether rank is interpreted as a percentage.
  public let rankIsPercentage: Bool?
  /// Whether later rules stop after this rule matches.
  public let stopsIfTrue: Bool?
  /// Formula operands in authored order.
  public let formulas: [String]
  /// Data-bar details when this is a data-bar rule.
  public let dataBar: OfficeConditionalDataBar?
  /// Color-scale details when this is a color-scale rule.
  public let colorScale: OfficeConditionalColorScale?
  /// Icon-set details when this is an icon-set rule.
  public let iconSet: OfficeConditionalIconSet?
}

/// Conditional-format rules applied to one or more worksheet ranges.
public struct OfficeConditionalFormatting: Sendable, Hashable, Codable {
  /// Target ranges in authored order.
  public let ranges: [OfficeCellRange]
  /// Rules in authored order.
  public let rules: [OfficeConditionalFormattingRule]
}

/// The inert embedded-object family declared by a worksheet.
public enum OfficeWorksheetObjectKind: String, Sendable, Hashable, Codable {
  /// An OLE object or embedded Office package.
  case ole
  /// An ActiveX control descriptor.
  case control
}

/// A worksheet OLE object or ActiveX control that OfficeKit never activates.
public struct OfficeWorksheetObject: Sendable {
  /// The embedded-object family.
  public let kind: OfficeWorksheetObjectKind
  /// The VML shape identifier used to associate legacy display markup.
  public let shapeIdentifier: UInt32?
  /// The producer-assigned control name.
  public let name: String?
  /// The OLE program identifier, such as `PowerPoint.Show.12`.
  public let programIdentifier: String?
  /// Whether the producer requests default object dimensions.
  public let usesDefaultSize: Bool?
  /// Whether the object moves with its surrounding cells.
  public let movesWithCells: Bool?
  /// Whether the object resizes with its surrounding cells.
  public let sizesWithCells: Bool?
  /// Exact authored cell-grid anchor, when present.
  public let anchor: OfficeWorksheetDrawingAnchor?
  /// Relationship-backed OLE payload or ActiveX descriptor.
  public let attachment: OfficeAttachment
  /// Relationship-backed preview image, when present.
  public let previewImage: OfficeAttachment?
  /// Primary, preview, and nested descriptor resources without duplicate relationships.
  public let attachments: [OfficeAttachment]
  /// Authored worksheet-relative placement, explicitly unresolved without grid layout.
  public let spatialInfo: OfficeSpatialInfo
  /// The worksheet part that declared this object.
  public let sourcePart: OfficePart
}

/// An ordered workbook reference to a worksheet part.
public struct OfficeWorksheetReference: Sendable, Hashable, Codable {
  /// The zero-based workbook order.
  public let index: Int

  /// The worksheet name shown to the user.
  public let name: String

  /// The workbook-local numeric sheet identifier.
  public let identifier: UInt32

  /// The worksheet tab visibility.
  public let visibility: OfficeWorksheetVisibility

  /// The relationship identifier used by the workbook.
  public let relationshipID: OfficeRelationshipID

  /// The resolved worksheet XML part.
  public let part: OfficePart
}

/// The parsed sparse content of one worksheet.
public struct OfficeWorksheet: Sendable {
  /// The ordered workbook reference that led to this worksheet.
  public let reference: OfficeWorksheetReference

  /// The authored used-range reference, when declared.
  public let dimensionReference: String?

  /// Authored column-range declarations.
  public let columns: [OfficeWorksheetColumn]

  /// Worksheet views, panes, and selections in declaration order.
  public let views: [OfficeWorksheetView]

  /// Populated rows in source order; omitted rows are not synthesized.
  public let rows: [OfficeWorksheetRow]

  /// Authored merged-cell ranges in source order.
  public let mergedRanges: [OfficeCellRange]

  /// The worksheet autofilter range, when declared.
  public let autoFilterRange: OfficeCellRange?

  /// Physical print margins in inches.
  public let pageMargins: OfficeWorksheetPageMargins?

  /// Print page setup metadata.
  public let pageSetup: OfficeWorksheetPageSetup?

  /// Header/footer control strings.
  public let headerFooter: OfficeWorksheetHeaderFooter?

  /// Conditional-format declarations in authored order.
  public let conditionalFormatting: [OfficeConditionalFormatting]

  /// Inert OLE objects and ActiveX controls in authored order.
  public let objects: [OfficeWorksheetObject]

  /// Cell and range hyperlinks in source order.
  public let hyperlinks: [OfficeWorksheetHyperlink]

  /// Parsed worksheet DrawingML parts, including exact anchors and nested attachments.
  public let drawings: [OfficeWorksheetDrawing]

  /// Lazily parsed structured table definitions related from this worksheet.
  public let tables: [OfficeSpreadsheetTableReference]

  /// Legacy note-style comments in source order.
  public let comments: [OfficeWorksheetComment]

  /// Every relationship-backed resource owned by the worksheet.
  public let attachments: [OfficeAttachment]

  /// Finds a populated cell by A1 reference without synthesizing empty cells.
  public func cell(at reference: OfficeCellReference) -> OfficeCell? {
    guard let row = rows.first(where: { $0.index == reference.rowIndex }) else { return nil }
    return row.cells.first { $0.reference == reference }
  }
}
