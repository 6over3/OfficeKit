/// An exact VML comment-box location within the worksheet grid.
public struct OfficeWorksheetCommentMarker: Sendable, Hashable, Codable {
  /// The zero-based column index.
  public let columnIndex: Int

  /// The horizontal position in 1/1024 of the column width.
  public let columnOffset: Int

  /// The zero-based row index.
  public let rowIndex: Int

  /// The vertical position in 1/256 of the row height.
  public let rowOffset: Int
}

/// The exact grid anchor for a legacy VML comment box.
public struct OfficeWorksheetCommentAnchor: Sendable, Hashable, Codable {
  /// The top-left grid marker.
  public let from: OfficeWorksheetCommentMarker

  /// The bottom-right grid marker.
  public let to: OfficeWorksheetCommentMarker
}

/// The legacy VML shape used to display a worksheet comment.
public struct OfficeWorksheetCommentShape: Sendable, Hashable, Codable {
  /// The producer-assigned VML shape identifier.
  public let identifier: String?

  /// The exact Excel grid anchor, when declared.
  public let anchor: OfficeWorksheetCommentAnchor?

  /// Whether the comment box is initially visible.
  public let isVisible: Bool

  /// Whether the box moves when its cells move.
  public let movesWithCells: Bool

  /// Whether the box resizes when its cells resize.
  public let sizesWithCells: Bool

  /// The VML z-index, when declared.
  public let zIndex: Int?

  /// Exact point-based VML box geometry when all four style values are present.
  public let spatialInfo: OfficeSpatialInfo

  /// The VML part that declared the shape.
  public let sourcePart: OfficePart
}

/// One legacy note-style comment attached to a worksheet cell.
public struct OfficeWorksheetComment: Sendable, Hashable, Codable {
  /// The commented cell.
  public let reference: OfficeCellReference

  /// The zero-based author index declared by the comment.
  public let authorIndex: UInt32

  /// The resolved author name, when the index is valid.
  public let author: String?

  /// The legacy VML shape identifier, when declared.
  public let shapeIdentifier: UInt32?

  /// Plain comment text assembled from rich-text runs.
  public let text: String

  /// The comment-list part that declared this comment.
  public let sourcePart: OfficePart

  /// The matched VML display shape and its spatial information, when available.
  public let shape: OfficeWorksheetCommentShape?
}
