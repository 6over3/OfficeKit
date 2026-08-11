/// A worksheet table definition that is parsed only when requested.
public struct OfficeSpreadsheetTableReference: Sendable {
  private let package: OfficePackage

  /// The worksheet relationship to the table.
  public let attachment: OfficeAttachment

  /// The table XML part.
  public let part: OfficePart

  package init(package: OfficePackage, attachment: OfficeAttachment, part: OfficePart) {
    self.package = package
    self.attachment = attachment
    self.part = part
  }

  /// Parses the table range, columns, formulas, and style settings.
  public func table() throws -> OfficeSpreadsheetTable {
    try SpreadsheetTableParser.parse(part: part, package: package)
  }
}

/// A structured table defined over a worksheet cell range.
public struct OfficeSpreadsheetTable: Sendable {
  /// The table XML part.
  public let sourcePart: OfficePart

  /// The workbook-local numeric identifier.
  public let identifier: UInt32

  /// The internal table name.
  public let name: String

  /// The name displayed to formulas and users.
  public let displayName: String

  /// The exact worksheet range occupied by the table.
  public let range: OfficeCellRange

  /// The number of header rows.
  public let headerRowCount: UInt32

  /// The number of totals rows.
  public let totalsRowCount: UInt32

  /// Columns in table order.
  public let columns: [OfficeSpreadsheetTableColumn]

  /// The applied table style settings, when declared.
  public let style: OfficeSpreadsheetTableStyle?
}

/// One structured worksheet-table column.
public struct OfficeSpreadsheetTableColumn: Sendable, Hashable, Codable {
  /// The table-local numeric identifier.
  public let identifier: UInt32

  /// The column heading.
  public let name: String

  /// The unique producer-assigned column name, when declared.
  public let uniqueName: String?

  /// The label displayed in the totals row, when declared.
  public let totalsRowLabel: String?

  /// The built-in totals-row function, when declared.
  public let totalsRowFunction: String?

  /// The calculated-column formula, when declared.
  public let calculatedColumnFormula: String?

  /// The totals-row formula, when declared.
  public let totalsRowFormula: String?
}

/// Display options for a structured worksheet table.
public struct OfficeSpreadsheetTableStyle: Sendable, Hashable, Codable {
  /// The workbook table-style name.
  public let name: String?

  /// Whether the first column receives special styling.
  public let showsFirstColumn: Bool

  /// Whether the last column receives special styling.
  public let showsLastColumn: Bool

  /// Whether alternating row stripes are enabled.
  public let showsRowStripes: Bool

  /// Whether alternating column stripes are enabled.
  public let showsColumnStripes: Bool
}
