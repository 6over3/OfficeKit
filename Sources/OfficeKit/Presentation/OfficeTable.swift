/// The semantic payload stored in a PowerPoint graphic frame.
public enum OfficeGraphicContent: Sendable {
  /// An inline DrawingML table.
  case table(OfficeTable)

  /// A lazily parsed chart part.
  case chart(OfficeChartReference)

  /// A SmartArt diagram whose defining parts remain lazy attachments.
  case diagram(OfficeDiagramReference)

  /// A 3D model and its raster fallback image.
  case model3D(OfficeModel3DReference)

  /// A relationship-backed or extension graphic identified by its source URI.
  case related(uri: String)
}

/// A DrawingML table embedded directly in a presentation shape tree.
public struct OfficeTable: Sendable {
  /// Exact authored grid-column widths in EMU.
  public let columnWidths: [OfficeLength]

  /// Rows in source order.
  public let rows: [OfficeTableRow]

  /// The table style identifier, when declared.
  public let styleIdentifier: String?
}

/// One DrawingML table row.
public struct OfficeTableRow: Sendable {
  /// The zero-based source row index.
  public let index: Int

  /// The exact authored row height in EMU.
  ///
  /// A zero height commonly means the producer expects text layout to choose the rendered height.
  public let height: OfficeLength

  /// Cells in source order.
  public let cells: [OfficeTableCell]
}

/// One DrawingML table cell with its logical grid position and spatial fidelity.
public struct OfficeTableCell: Sendable {
  /// The zero-based row index.
  public let rowIndex: Int

  /// The zero-based logical grid-column index.
  public let columnIndex: Int

  /// The number of grid columns covered by the cell.
  public let columnSpan: Int

  /// The number of grid rows covered by the cell.
  public let rowSpan: Int

  /// Whether this is a continuation cell in a horizontal merge.
  public let isHorizontallyMerged: Bool

  /// Whether this is a continuation cell in a vertical merge.
  public let isVerticallyMerged: Bool

  /// Plain cell text in paragraph order.
  public let text: String

  /// The cell frame in slide coordinates when row and column dimensions are resolvable.
  public let spatialInfo: OfficeSpatialInfo
}
