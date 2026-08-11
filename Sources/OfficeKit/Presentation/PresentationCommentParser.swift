package enum PresentationCommentParser {
  private static let presentationNamespace =
    "http://schemas.openxmlformats.org/presentationml/2006/main"

  package static func authors(
    presentationPart: OfficePart,
    package: OfficePackage
  ) throws -> [OfficePresentationCommentAuthor] {
    guard
      let relationship = try package.relationships(
        from: .part(presentationPart.name),
        ofType: .commentAuthors
      ).first else { return [] }
    guard let part = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }

    var authors: [OfficePresentationCommentAuthor] = []
    try package.parseXML(in: part, compatibility: .commonOffice) { event in
      guard case .startElement(let name, let attributes, _, _) = event,
        name.namespaceURI == presentationNamespace,
        name.localName == "cmAuthor" else { return }
      guard let identifier = attribute("id", in: attributes).flatMap(UInt32.init),
        let name = attribute("name", in: attributes) else {
        throw OfficeKitError.invalidXML(
          part: part.name.rawValue,
          message: "p:cmAuthor requires id and name attributes."
        )
      }
      authors.append(
        OfficePresentationCommentAuthor(
          identifier: identifier,
          name: name,
          initials: attribute("initials", in: attributes),
          colorIndex: attribute("clrIdx", in: attributes).flatMap(UInt32.init)
        )
      )
    }
    return authors
  }

  package static func comments(
    slidePart: OfficePart,
    authorsByID: [UInt32: OfficePresentationCommentAuthor],
    package: OfficePackage
  ) throws -> [OfficePresentationComment] {
    guard
      let relationship = try package.relationships(
        from: .part(slidePart.name),
        ofType: .comments
      ).first else { return [] }
    guard let part = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }

    let session = CommentParsingSession(
      sourcePart: part.name,
      authorsByID: authorsByID
    )
    try package.parseXML(in: part, compatibility: .commonOffice, session.consume)
    return session.comments
  }

  private static func attribute(
    _ localName: String,
    in attributes: [OfficeXMLAttribute]
  ) -> String? {
    attributes.first { $0.name.namespaceURI == nil && $0.name.localName == localName }?.value
  }
}

private final class CommentParsingSession {
  private static let presentationNamespace =
    "http://schemas.openxmlformats.org/presentationml/2006/main"

  struct Builder {
    let depth: Int
    let index: UInt32
    let authorID: UInt32
    let dateTime: String?
    var x: Int64?
    var y: Int64?
    var text = ""
    var textDepth: Int?
  }

  let sourcePart: OfficePartName
  let authorsByID: [UInt32: OfficePresentationCommentAuthor]
  var comments: [OfficePresentationComment] = []
  var depth = 0
  var builder: Builder?

  init(
    sourcePart: OfficePartName,
    authorsByID: [UInt32: OfficePresentationCommentAuthor]
  ) {
    self.sourcePart = sourcePart
    self.authorsByID = authorsByID
  }

  func consume(_ event: OfficeXMLEvent) throws {
    switch event {
    case .startElement(let name, let attributes, _, _):
      depth += 1
      if name.namespaceURI == Self.presentationNamespace, name.localName == "cm" {
        guard let index = attribute("idx", in: attributes).flatMap(UInt32.init),
          let authorID = attribute("authorId", in: attributes).flatMap(UInt32.init) else {
          throw OfficeKitError.invalidXML(
            part: sourcePart.rawValue,
            message: "p:cm requires numeric idx and authorId attributes."
          )
        }
        builder = Builder(
          depth: depth,
          index: index,
          authorID: authorID,
          dateTime: attribute("dt", in: attributes)
        )
      } else if name.namespaceURI == Self.presentationNamespace, name.localName == "pos" {
        builder?.x = attribute("x", in: attributes).flatMap(Int64.init)
        builder?.y = attribute("y", in: attributes).flatMap(Int64.init)
      } else if name.namespaceURI == Self.presentationNamespace, name.localName == "text" {
        builder?.textDepth = depth
      }

    case .text(let text, _):
      guard builder?.textDepth != nil else { return }
      builder?.text.append(text)

    case .endElement(let name, _):
      if builder?.textDepth == depth, name.namespaceURI == Self.presentationNamespace,
        name.localName == "text"
      {
        builder?.textDepth = nil
      }
      if let completed = builder, completed.depth == depth,
        name.namespaceURI == Self.presentationNamespace, name.localName == "cm"
      {
        guard let x = completed.x, let y = completed.y else {
          throw OfficeKitError.invalidXML(
            part: sourcePart.rawValue,
            message: "p:cm requires a p:pos with numeric x and y attributes."
          )
        }
        comments.append(
          OfficePresentationComment(
            index: completed.index,
            authorID: completed.authorID,
            author: authorsByID[completed.authorID],
            dateTime: completed.dateTime,
            text: completed.text,
            position: OfficePresentationCommentPosition(
              x: OfficeLength(emu: x),
              y: OfficeLength(emu: y)
            ),
            sourcePart: sourcePart
          )
        )
        builder = nil
      }
      depth -= 1

    case .startDocument, .endDocument:
      break
    }
  }

  private func attribute(
    _ localName: String,
    in attributes: [OfficeXMLAttribute]
  ) -> String? {
    attributes.first { $0.name.namespaceURI == nil && $0.name.localName == localName }?.value
  }
}
