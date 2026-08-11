import Foundation

/// An ordered reference from a presentation to one slide part.
public struct OfficeSlideReference: Sendable, Hashable, Codable {
  /// The zero-based presentation order.
  public let index: Int

  /// The numeric slide identifier declared by PresentationML.
  public let identifier: UInt32

  /// The relationship identifier used by the presentation part.
  public let relationshipID: OfficeRelationshipID

  /// The resolved slide XML part.
  public let part: OfficePart

  package init(
    index: Int,
    identifier: UInt32,
    relationshipID: OfficeRelationshipID,
    part: OfficePart
  ) {
    self.index = index
    self.identifier = identifier
    self.relationshipID = relationshipID
    self.part = part
  }
}

/// A read-only PowerPoint presentation index.
///
/// Initialization parses the presentation part and its small comment-author list, when present.
/// Slide XML, comments, notes, media, charts, and embedded resources remain lazy until the
/// corresponding slide is requested.
public struct OfficePresentation: Sendable {
  /// The opened Office document.
  public let document: OfficeDocument

  /// The authored slide canvas size in points.
  public let slideSize: OfficeSize

  /// The authored notes-page size in points, when declared.
  public let notesSize: OfficeSize?

  /// Slides in presentation order, resolved through relationship identifiers.
  public let slides: [OfficeSlideReference]

  /// Legacy presentation comment authors, indexed at presentation open time.
  public let commentAuthors: [OfficePresentationCommentAuthor]

  /// Opens and indexes a PowerPoint presentation.
  public init(contentsOf url: URL, limits: OfficeParsingLimits = .standard) throws {
    try self.init(document: OfficeDocument(contentsOf: url, limits: limits))
  }

  /// Indexes an already opened PowerPoint document.
  public init(document: OfficeDocument) throws {
    guard document.kind == .presentation else {
      throw OfficeKitError.invalidPackage("The Office document is not a presentation.")
    }

    let values = try PresentationIndexParser.parse(
      package: document.package,
      presentationPart: document.mainPart
    )
    let commentAuthors = try PresentationCommentParser.authors(
      presentationPart: document.mainPart,
      package: document.package
    )
    self.document = document
    self.slideSize = values.slideSize
    self.notesSize = values.notesSize
    self.slides = values.slides
    self.commentAuthors = commentAuthors
  }

  /// Parses one slide's shape tree, text, relationship-backed attachments, and authored geometry.
  public func slide(at index: Int) throws -> OfficeSlide {
    guard slides.indices.contains(index) else {
      throw OfficeKitError.invalidPackage("Slide index \(index) is out of bounds.")
    }
    return try PresentationSlideParser.parse(
      reference: slides[index],
      package: document.package,
      commentAuthors: commentAuthors
    )
  }
}

private enum PresentationIndexParser {
  private static let presentationNamespace =
    "http://schemas.openxmlformats.org/presentationml/2006/main"
  private static let relationshipNamespace =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  struct Values {
    let slideSize: OfficeSize
    let notesSize: OfficeSize?
    let slides: [OfficeSlideReference]
  }

  struct RawSlide {
    let identifier: UInt32
    let relationshipID: OfficeRelationshipID
  }

  static func parse(package: OfficePackage, presentationPart: OfficePart) throws -> Values {
    var rawSlideSize: (width: Int64, height: Int64)?
    var rawNotesSize: (width: Int64, height: Int64)?
    var rawSlides: [RawSlide] = []

    try package.parseXML(in: presentationPart, compatibility: .commonOffice) { event in
      guard case .startElement(let name, let attributes, _, _) = event,
        name.namespaceURI == presentationNamespace else { return }
      switch name.localName {
      case "sldSz":
        rawSlideSize = try size(from: attributes, element: "p:sldSz")
      case "notesSz":
        rawNotesSize = try size(from: attributes, element: "p:notesSz")
      case "sldId":
        guard let rawIdentifier = attribute("id", in: attributes),
          let identifier = UInt32(rawIdentifier),
          let rawRelationshipID = attribute(
            "id",
            namespaceURI: relationshipNamespace,
            in: attributes
          ) else {
          throw OfficeKitError.invalidXML(
            part: presentationPart.name.rawValue,
            message: "p:sldId requires numeric id and r:id attributes."
          )
        }
        rawSlides.append(
          RawSlide(
            identifier: identifier,
            relationshipID: OfficeRelationshipID(rawValue: rawRelationshipID)
          )
        )
      default:
        break
      }
    }

    guard let rawSlideSize else {
      throw OfficeKitError.invalidXML(
        part: presentationPart.name.rawValue,
        message: "The presentation does not declare p:sldSz."
      )
    }

    var slides: [OfficeSlideReference] = []
    slides.reserveCapacity(rawSlides.count)
    for (index, rawSlide) in rawSlides.enumerated() {
      guard
        let relationship = try package.relationship(
          identifiedBy: rawSlide.relationshipID,
          from: .part(presentationPart.name)
        ) else {
        throw OfficeKitError.invalidPackage(
          "Slide \(rawSlide.identifier) references missing relationship "
            + "\(rawSlide.relationshipID.rawValue)."
        )
      }
      guard relationship.type.isEquivalent(to: .slide) else {
        throw OfficeKitError.invalidPackage(
          "Relationship \(relationship.id.rawValue) for slide \(rawSlide.identifier) "
            + "does not target a slide."
        )
      }
      guard let part = package.part(referencedBy: relationship) else {
        throw OfficeKitError.missingPart(relationship.rawTarget)
      }
      slides.append(
        OfficeSlideReference(
          index: index,
          identifier: rawSlide.identifier,
          relationshipID: rawSlide.relationshipID,
          part: part
        )
      )
    }

    return Values(
      slideSize: pointSize(from: rawSlideSize),
      notesSize: rawNotesSize.map(pointSize(from:)),
      slides: slides
    )
  }

  private static func size(
    from attributes: [OfficeXMLAttribute],
    element: String
  ) throws -> (width: Int64, height: Int64) {
    guard let rawWidth = attribute("cx", in: attributes),
      let width = Int64(rawWidth),
      let rawHeight = attribute("cy", in: attributes),
      let height = Int64(rawHeight),
      width >= 0, height >= 0 else {
      throw OfficeKitError.invalidPackage("\(element) requires nonnegative cx and cy values.")
    }
    return (width, height)
  }

  private static func pointSize(from size: (width: Int64, height: Int64)) -> OfficeSize {
    OfficeSize(
      width: OfficeLength(emu: size.width).points,
      height: OfficeLength(emu: size.height).points
    )
  }

  private static func attribute(
    _ localName: String,
    namespaceURI: String? = nil,
    in attributes: [OfficeXMLAttribute]
  ) -> String? {
    attributes.first {
      $0.name.localName == localName && $0.name.namespaceURI == namespaceURI
    }?.value
  }
}
