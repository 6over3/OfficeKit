/// A presentation comment author declared in the presentation-level author list.
public struct OfficePresentationCommentAuthor: Sendable, Hashable, Codable {
  /// The numeric identifier referenced by slide comments.
  public let identifier: UInt32

  /// The author's display name.
  public let name: String

  /// The author's initials, when declared.
  public let initials: String?

  /// The producer's author-color index, when declared.
  public let colorIndex: UInt32?
}

/// An exact PowerPoint comment position stored in DrawingML coordinate units.
public struct OfficePresentationCommentPosition: Sendable, Hashable, Codable {
  /// The exact horizontal coordinate in EMU.
  public let x: OfficeLength

  /// The exact vertical coordinate in EMU.
  public let y: OfficeLength

  /// The position converted to points.
  public var point: OfficePoint {
    OfficePoint(x: x.points, y: y.points)
  }
}

/// A legacy PresentationML comment attached to a slide.
public struct OfficePresentationComment: Sendable, Hashable, Codable {
  /// The comment's index within its author namespace.
  public let index: UInt32

  /// The referenced author identifier.
  public let authorID: UInt32

  /// The resolved presentation-level author, when declared.
  public let author: OfficePresentationCommentAuthor?

  /// The timestamp exactly as serialized in the package.
  public let dateTime: String?

  /// The comment text.
  public let text: String

  /// The authored comment marker position.
  public let position: OfficePresentationCommentPosition

  /// The comment-list part that declared this value.
  public let sourcePart: OfficePartName
}
