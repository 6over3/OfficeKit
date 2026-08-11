/// Resource limits applied while opening and reading an Office package.
public struct OfficeParsingLimits: Sendable, Equatable {
  /// The maximum number of ZIP entries accepted in one package.
  public let maximumEntryCount: Int

  /// The maximum uncompressed size of any single ZIP entry.
  public let maximumEntrySize: UInt64

  /// The maximum sum of the declared uncompressed sizes of all ZIP entries.
  public let maximumTotalUncompressedSize: UInt64

  /// The maximum permitted ratio between uncompressed and compressed entry sizes.
  public let maximumCompressionRatio: Double

  /// The maximum size of package metadata loaded into memory for XML parsing.
  public let maximumMetadataPartSize: UInt64

  /// The default limit for a caller-requested in-memory part read.
  public let defaultDataReadLimit: UInt64

  /// The largest XML part read into memory instead of streamed through a temporary file.
  public let maximumInMemoryXMLPartSize: UInt64

  /// The maximum number of relationships accepted from one relationship part.
  public let maximumRelationshipsPerPart: Int

  /// Limits applied to every XML part parsed through this package.
  public let xmlParsingLimits: OfficeXMLParsingLimits

  /// The maximum number of rich strings retained from one workbook shared-string table.
  public let maximumSharedStringCount: Int

  /// The maximum number of rows retained by `OfficeWorkbook.worksheet(at:)`.
  ///
  /// Row streaming through `OfficeWorkbook.streamRows(inWorksheetAt:_:)` does not retain rows and
  /// is not constrained by this collection limit.
  public let maximumRetainedWorksheetRows: Int

  /// The maximum number of cells retained by `OfficeWorkbook.worksheet(at:)`.
  ///
  /// Row streaming does not retain completed cells and is not constrained by this total.
  public let maximumRetainedWorksheetCells: Int

  /// Creates a set of package resource limits.
  public init(
    maximumEntryCount: Int = 50_000,
    maximumEntrySize: UInt64 = 512 * 1_024 * 1_024,
    maximumTotalUncompressedSize: UInt64 = 8 * 1_024 * 1_024 * 1_024,
    maximumCompressionRatio: Double = 1_000,
    maximumMetadataPartSize: UInt64 = 16 * 1_024 * 1_024,
    defaultDataReadLimit: UInt64 = 64 * 1_024 * 1_024,
    maximumInMemoryXMLPartSize: UInt64 = 4 * 1_024 * 1_024,
    maximumRelationshipsPerPart: Int = 100_000,
    xmlParsingLimits: OfficeXMLParsingLimits = .standard,
    maximumSharedStringCount: Int = 2_000_000,
    maximumRetainedWorksheetRows: Int = 1_048_576,
    maximumRetainedWorksheetCells: Int = 10_000_000
  ) {
    self.maximumEntryCount = max(0, maximumEntryCount)
    self.maximumEntrySize = maximumEntrySize
    self.maximumTotalUncompressedSize = maximumTotalUncompressedSize
    self.maximumCompressionRatio =
      maximumCompressionRatio.isFinite
      ? max(0, maximumCompressionRatio) : 0
    self.maximumMetadataPartSize = maximumMetadataPartSize
    self.defaultDataReadLimit = defaultDataReadLimit
    self.maximumInMemoryXMLPartSize = maximumInMemoryXMLPartSize
    self.maximumRelationshipsPerPart = max(0, maximumRelationshipsPerPart)
    self.xmlParsingLimits = xmlParsingLimits
    self.maximumSharedStringCount = max(0, maximumSharedStringCount)
    self.maximumRetainedWorksheetRows = max(0, maximumRetainedWorksheetRows)
    self.maximumRetainedWorksheetCells = max(0, maximumRetainedWorksheetCells)
  }

  /// Limits suitable for normal locally opened Office documents.
  public static let standard = OfficeParsingLimits()
}
