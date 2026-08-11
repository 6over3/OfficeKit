/// An exact OOXML percentage stored in 1/1000 percent units.
public struct OfficePercentage: RawRepresentable, Sendable, Hashable, Codable,
  CustomStringConvertible
{
  /// The exact source value, where `100000` represents 100 percent.
  public let rawValue: Int64

  /// Creates a percentage from its exact OOXML integer representation.
  public init(rawValue: Int64) {
    self.rawValue = rawValue
  }

  /// The percentage represented as a unit fraction, where `1` is 100 percent.
  public var fraction: Double { Double(rawValue) / 100_000 }

  /// The percentage represented as a conventional value from zero to 100.
  public var percent: Double { Double(rawValue) / 1_000 }

  /// A locale-independent percentage description.
  public var description: String { "\(percent)%" }
}

/// Exact crop offsets applied to an image's source rectangle.
public struct OfficeCropRectangle: Sendable, Hashable, Codable {
  /// The amount removed from the left edge.
  public let left: OfficePercentage

  /// The amount removed from the top edge.
  public let top: OfficePercentage

  /// The amount removed from the right edge.
  public let right: OfficePercentage

  /// The amount removed from the bottom edge.
  public let bottom: OfficePercentage
}

/// Picture-specific content from a PresentationML picture element.
public struct OfficePicture: Sendable {
  /// The preferred image relationship, including vector images when directly referenced.
  public let primaryImage: OfficeAttachment?

  /// All image relationships referenced by the picture, including fallbacks.
  public let images: [OfficeAttachment]

  /// The authored source crop rectangle. `nil` means no `a:srcRect` was declared.
  public let cropRectangle: OfficeCropRectangle?
}
