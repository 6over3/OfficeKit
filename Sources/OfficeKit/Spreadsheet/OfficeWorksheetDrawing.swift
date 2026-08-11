/// An exact location within the worksheet cell grid.
public struct OfficeWorksheetCellMarker: Sendable, Hashable, Codable {
  /// The zero-based column index.
  public let columnIndex: Int

  /// The horizontal offset from the column's leading edge in EMU.
  public let columnOffset: OfficeLength

  /// The zero-based row index.
  public let rowIndex: Int

  /// The vertical offset from the row's top edge in EMU.
  public let rowOffset: OfficeLength
}

/// The way an Excel drawing object is positioned relative to the worksheet grid.
public enum OfficeWorksheetDrawingAnchor: Sendable, Hashable, Codable {
  /// An object bounded by two exact cell-grid markers.
  case twoCell(
    from: OfficeWorksheetCellMarker,
    to: OfficeWorksheetCellMarker,
    editBehavior: String?
  )

  /// An object starting at one exact marker with a fixed EMU extent.
  case oneCell(
    from: OfficeWorksheetCellMarker,
    width: OfficeLength,
    height: OfficeLength
  )

  /// An object with a worksheet-absolute EMU position and extent.
  case absolute(x: OfficeLength, y: OfficeLength, width: OfficeLength, height: OfficeLength)
}

/// The semantic kind of an object in a worksheet drawing part.
public enum OfficeWorksheetDrawingElementKind: String, Sendable, Hashable, Codable {
  /// A geometric shape or text box.
  case shape

  /// A picture backed by one or more image relationships.
  case picture

  /// A connector between drawing objects.
  case connector

  /// A chart, diagram, or other graphic frame.
  case graphicFrame

  /// A group containing drawing objects.
  case group
}

/// One anchored object from an Excel worksheet DrawingML part.
public struct OfficeWorksheetDrawingElement: Sendable {
  /// The drawing-object kind.
  public let kind: OfficeWorksheetDrawingElementKind

  /// The producer-assigned non-visual identifier, when declared.
  public let identifier: UInt32?

  /// The producer-assigned object name.
  public let name: String?

  /// Accessibility-oriented alternative text.
  public let alternativeText: String?

  /// Plain DrawingML text in source order.
  public let text: String

  /// The exact worksheet anchor.
  public let anchor: OfficeWorksheetDrawingAnchor

  /// A typed chart, diagram, or extension payload, when recognized.
  public let graphicContent: OfficeGraphicContent?

  /// Picture-specific image relationships.
  public let picture: OfficePicture?

  /// Resolved or honestly unresolved worksheet-space geometry.
  public let spatialInfo: OfficeSpatialInfo

  /// Relationships explicitly referenced by this drawing object.
  public let attachments: [OfficeAttachment]
}

/// One worksheet DrawingML part and its anchored objects.
public struct OfficeWorksheetDrawing: Sendable {
  /// The drawing XML part.
  public let part: OfficePart

  /// Objects in their authored back-to-front order.
  public let elements: [OfficeWorksheetDrawingElement]

  /// Every payload relationship owned by the drawing part.
  public let attachments: [OfficeAttachment]
}
