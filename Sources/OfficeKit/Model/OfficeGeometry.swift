import Foundation

/// A signed length stored exactly in English Metric Units (EMU).
public struct OfficeLength: Sendable, Hashable, Codable, Comparable, CustomStringConvertible {
  /// The number of EMU in one typographic point.
  public static let emuPerPoint: Int64 = 12_700

  /// The number of EMU in one inch.
  public static let emuPerInch: Int64 = 914_400

  /// The exact source length in EMU.
  public let emu: Int64

  /// Creates an exact EMU length.
  public init(emu: Int64) {
    self.emu = emu
  }

  /// Creates a length from points, rounded to the nearest EMU.
  public init(points: Double) {
    self.emu = Int64((points * Double(Self.emuPerPoint)).rounded())
  }

  /// The length in typographic points.
  public var points: Double { Double(emu) / Double(Self.emuPerPoint) }

  /// The length in inches.
  public var inches: Double { Double(emu) / Double(Self.emuPerInch) }

  /// Converts the length to pixels for a specified dots-per-inch value.
  public func pixels(atDPI dpi: Double) -> Double {
    inches * dpi
  }

  /// The length rendered as an EMU count.
  public var description: String { "\(emu) EMU" }

  /// Orders lengths by their exact EMU values.
  public static func < (lhs: Self, rhs: Self) -> Bool {
    lhs.emu < rhs.emu
  }
}

/// A two-dimensional point measured in typographic points.
public struct OfficePoint: Sendable, Hashable, Codable {
  /// The horizontal coordinate.
  public let x: Double

  /// The vertical coordinate.
  public let y: Double

  /// Creates a point.
  public init(x: Double, y: Double) {
    self.x = x
    self.y = y
  }

  /// The zero point.
  public static let zero = OfficePoint(x: 0, y: 0)
}

/// A two-dimensional size measured in typographic points.
public struct OfficeSize: Sendable, Hashable, Codable {
  /// The horizontal extent.
  public let width: Double

  /// The vertical extent.
  public let height: Double

  /// Creates a size.
  public init(width: Double, height: Double) {
    self.width = width
    self.height = height
  }

  /// The zero size.
  public static let zero = OfficeSize(width: 0, height: 0)
}

/// An axis-aligned rectangle measured in typographic points.
public struct OfficeRect: Sendable, Hashable, Codable {
  /// The rectangle origin.
  public let origin: OfficePoint

  /// The rectangle size.
  public let size: OfficeSize

  /// Creates a rectangle from an origin and size.
  public init(origin: OfficePoint, size: OfficeSize) {
    self.origin = origin
    self.size = size
  }

  /// Creates a rectangle from scalar coordinates.
  public init(x: Double, y: Double, width: Double, height: Double) {
    self.init(
      origin: OfficePoint(x: x, y: y),
      size: OfficeSize(width: width, height: height)
    )
  }

  /// The smallest horizontal coordinate.
  public var minX: Double { min(origin.x, origin.x + size.width) }

  /// The largest horizontal coordinate.
  public var maxX: Double { max(origin.x, origin.x + size.width) }

  /// The smallest vertical coordinate.
  public var minY: Double { min(origin.y, origin.y + size.height) }

  /// The largest vertical coordinate.
  public var maxY: Double { max(origin.y, origin.y + size.height) }

  /// The zero rectangle.
  public static let zero = OfficeRect(origin: .zero, size: .zero)

  /// Returns the axis-aligned bounds of this rectangle after applying `transform`.
  public func applying(_ transform: OfficeAffineTransform) -> OfficeRect {
    let points = [
      OfficePoint(x: minX, y: minY),
      OfficePoint(x: maxX, y: minY),
      OfficePoint(x: minX, y: maxY),
      OfficePoint(x: maxX, y: maxY),
    ].map(transform.applying(to:))
    let xs = points.map(\.x)
    let ys = points.map(\.y)
    let minimumX = xs.min() ?? 0
    let maximumX = xs.max() ?? 0
    let minimumY = ys.min() ?? 0
    let maximumY = ys.max() ?? 0
    return OfficeRect(
      x: minimumX,
      y: minimumY,
      width: maximumX - minimumX,
      height: maximumY - minimumY
    )
  }
}

/// A two-dimensional affine transform independent of CoreGraphics.
public struct OfficeAffineTransform: Sendable, Hashable, Codable {
  /// The x-to-x scale or rotation coefficient.
  public let a: Double

  /// The x-to-y rotation or shear coefficient.
  public let b: Double

  /// The y-to-x rotation or shear coefficient.
  public let c: Double

  /// The y-to-y scale or rotation coefficient.
  public let d: Double

  /// The horizontal translation in points.
  public let tx: Double

  /// The vertical translation in points.
  public let ty: Double

  /// Creates an affine transform using the conventional six coefficients.
  public init(a: Double, b: Double, c: Double, d: Double, tx: Double, ty: Double) {
    self.a = a
    self.b = b
    self.c = c
    self.d = d
    self.tx = tx
    self.ty = ty
  }

  /// The identity transform.
  public static let identity = OfficeAffineTransform(a: 1, b: 0, c: 0, d: 1, tx: 0, ty: 0)

  /// Creates a translation.
  public static func translation(x: Double, y: Double) -> Self {
    Self(a: 1, b: 0, c: 0, d: 1, tx: x, ty: y)
  }

  /// Creates independent horizontal and vertical scaling.
  public static func scale(x: Double, y: Double) -> Self {
    Self(a: x, b: 0, c: 0, d: y, tx: 0, ty: 0)
  }

  /// Creates a counterclockwise rotation in radians.
  public static func rotation(radians: Double) -> Self {
    let cosine = cos(radians)
    let sine = sin(radians)
    return Self(a: cosine, b: sine, c: -sine, d: cosine, tx: 0, ty: 0)
  }

  /// Applies the transform to a point.
  public func applying(to point: OfficePoint) -> OfficePoint {
    OfficePoint(
      x: a * point.x + c * point.y + tx,
      y: b * point.x + d * point.y + ty
    )
  }

  /// Returns a transform that applies this transform and then `next`.
  public func followed(by next: Self) -> Self {
    Self(
      a: next.a * a + next.c * b,
      b: next.b * a + next.d * b,
      c: next.a * c + next.c * d,
      d: next.b * c + next.d * d,
      tx: next.a * tx + next.c * ty + next.tx,
      ty: next.b * tx + next.d * ty + next.ty
    )
  }
}

/// The document coordinate space in which authored geometry is expressed.
public enum OfficeCoordinateSpace: String, Sendable, Hashable, Codable {
  /// A PowerPoint slide canvas.
  case slide
  /// A PowerPoint slide layout.
  case slideLayout
  /// A PowerPoint slide master.
  case slideMaster
  /// A PowerPoint speaker-notes page.
  case notesPage
  /// A PowerPoint notes master.
  case notesMaster
  /// An Excel worksheet drawing surface.
  case worksheet
  /// A Word page.
  case page
  /// A Word page-margin rectangle.
  case margin
  /// A Word column.
  case column
  /// A Word paragraph.
  case paragraph
  /// A Word character anchor.
  case character
  /// A nested DrawingML group coordinate system.
  case group
  /// A table cell.
  case tableCell
  /// A producer extension whose coordinate system is not understood.
  case unknown
}

/// The confidence with which OfficeKit resolved authored geometry.
public enum OfficeSpatialResolution: Sendable, Hashable, Codable {
  /// The public geometry is an exact unit conversion from source values.
  case exact
  /// The geometry was derived from exact anchors and parent transforms.
  case derived
  /// Final geometry depends on layout information OfficeKit does not have.
  case unresolved(reason: String)
}

/// Exact DrawingML transform values retained in their source units.
public struct OfficeDrawingTransform: Sendable, Hashable, Codable {
  /// The horizontal offset in EMU.
  public let x: OfficeLength

  /// The vertical offset in EMU.
  public let y: OfficeLength

  /// The horizontal extent in EMU.
  public let width: OfficeLength

  /// The vertical extent in EMU.
  public let height: OfficeLength

  /// A group's child-coordinate horizontal origin in EMU.
  public let childX: OfficeLength?

  /// A group's child-coordinate vertical origin in EMU.
  public let childY: OfficeLength?

  /// A group's child-coordinate horizontal extent in EMU.
  public let childWidth: OfficeLength?

  /// A group's child-coordinate vertical extent in EMU.
  public let childHeight: OfficeLength?

  /// The exact clockwise rotation in 1/60,000 degree units.
  public let rotationUnits: Int64

  /// Whether the source requests a horizontal flip.
  public let isFlippedHorizontally: Bool

  /// Whether the source requests a vertical flip.
  public let isFlippedVertically: Bool

  /// Creates an exact DrawingML transform.
  public init(
    x: OfficeLength,
    y: OfficeLength,
    width: OfficeLength,
    height: OfficeLength,
    childX: OfficeLength? = nil,
    childY: OfficeLength? = nil,
    childWidth: OfficeLength? = nil,
    childHeight: OfficeLength? = nil,
    rotationUnits: Int64 = 0,
    isFlippedHorizontally: Bool = false,
    isFlippedVertically: Bool = false
  ) {
    self.x = x
    self.y = y
    self.width = width
    self.height = height
    self.childX = childX
    self.childY = childY
    self.childWidth = childWidth
    self.childHeight = childHeight
    self.rotationUnits = rotationUnits
    self.isFlippedHorizontally = isFlippedHorizontally
    self.isFlippedVertically = isFlippedVertically
  }

  /// The authored clockwise rotation in radians.
  public var rotationRadians: Double {
    Double(rotationUnits) / 60_000 * .pi / 180
  }
}

/// Spatial information retained for a semantic Office element.
public struct OfficeSpatialInfo: Sendable, Hashable, Codable {
  /// The element's coordinate space.
  public let coordinateSpace: OfficeCoordinateSpace

  /// The exact source DrawingML transform, when the element uses one.
  public let sourceTransform: OfficeDrawingTransform?

  /// The package part that declared `sourceTransform`, when known.
  public let geometrySourcePart: OfficePartName?

  /// The authored or derived axis-aligned frame in points, when resolvable.
  public let frame: OfficeRect?

  /// The transform from the element's local coordinates to its parent coordinates.
  public let transformToParent: OfficeAffineTransform?

  /// The composed transform from local coordinates to the document surface, when known.
  public let transformToDocument: OfficeAffineTransform?

  /// The authored clockwise rotation in radians.
  public let rotation: Double

  /// Whether the source requests a horizontal flip.
  public let isFlippedHorizontally: Bool

  /// Whether the source requests a vertical flip.
  public let isFlippedVertically: Bool

  /// The zero-based order within its containing drawing tree.
  public let zIndex: Int?

  /// The fidelity of the resolved frame and transform.
  public let resolution: OfficeSpatialResolution

  /// Creates spatial information.
  public init(
    coordinateSpace: OfficeCoordinateSpace,
    sourceTransform: OfficeDrawingTransform? = nil,
    geometrySourcePart: OfficePartName? = nil,
    frame: OfficeRect?,
    transformToParent: OfficeAffineTransform? = nil,
    transformToDocument: OfficeAffineTransform? = nil,
    rotation: Double = 0,
    isFlippedHorizontally: Bool = false,
    isFlippedVertically: Bool = false,
    zIndex: Int? = nil,
    resolution: OfficeSpatialResolution
  ) {
    self.coordinateSpace = coordinateSpace
    self.sourceTransform = sourceTransform
    self.geometrySourcePart = geometrySourcePart
    self.frame = frame
    self.transformToParent = transformToParent
    self.transformToDocument = transformToDocument
    self.rotation = rotation
    self.isFlippedHorizontally = isFlippedHorizontally
    self.isFlippedVertically = isFlippedVertically
    self.zIndex = zIndex
    self.resolution = resolution
  }
}
