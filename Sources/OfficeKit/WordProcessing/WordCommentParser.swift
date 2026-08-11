import Foundation

package enum WordCommentParser {
  private static let wordNamespace =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

  private struct Builder {
    let startDepth: Int
    let identifier: Int64
    let author: String?
    let initials: String?
    let dateText: String?
    let parser: WordDocumentParser
  }

  package static func parse(
    documentPart: OfficePart,
    package: OfficePackage,
    relationships: [OfficeRelationship],
    anchors: [OfficeWordCommentAnchor]
  ) throws -> [OfficeWordComment] {
    guard let relationship = relationships.first(where: { $0.type.isEquivalent(to: .comments) }) else {
      return []
    }
    guard let part = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }
    let anchorsByID = Dictionary(
      anchors.map { ($0.identifier, $0) },
      uniquingKeysWith: { first, _ in first }
    )
    let dateFormatter = ISO8601DateFormatter()
    let commentRelationships = try package.relationships(from: .part(part.name))
    let commentAttachments = commentRelationships.map(package.attachment(referencedBy:))
    var depth = 0
    var builder: Builder?
    var comments: [OfficeWordComment] = []

    try package.parseXML(in: part, compatibility: .commonOffice) { event in
      switch event {
      case .startElement(let name, let attributes, _, _):
        depth += 1
        guard name.namespaceURI == wordNamespace else { return }
        switch name.localName {
        case "comment":
          guard let identifierText = wordAttribute("id", in: attributes),
            let identifier = Int64(identifierText) else {
            throw OfficeKitError.invalidXML(
              part: part.name.rawValue,
              message: "w:comment requires a numeric w:id attribute."
            )
          }
          builder = Builder(
            startDepth: depth,
            identifier: identifier,
            author: wordAttribute("author", in: attributes),
            initials: wordAttribute("initials", in: attributes),
            dateText: wordAttribute("date", in: attributes),
            parser: WordDocumentParser(
              part: part,
              package: package,
              relationships: commentRelationships
            )
          )
        default:
          break
        }
        try builder?.parser.consume(event)
      case .text:
        try builder?.parser.consume(event)
      case .endElement(let name, _):
        try builder?.parser.consume(event)
        if name.namespaceURI == wordNamespace {
          if name.localName == "comment", builder?.startDepth == depth,
            let completed = builder
          {
            let content = completed.parser.body
            let paragraphs = content.paragraphs.map(\.text)
            comments.append(
              OfficeWordComment(
                identifier: completed.identifier,
                author: completed.author,
                initials: completed.initials,
                dateText: completed.dateText,
                date: completed.dateText.flatMap(dateFormatter.date(from:)),
                paragraphs: paragraphs,
                text: paragraphs.joined(separator: "\n"),
                content: content,
                anchor: anchorsByID[completed.identifier],
                sourcePart: part,
                attachments: commentAttachments
              )
            )
            builder = nil
          }
        }
        depth -= 1
      case .startDocument, .endDocument:
        break
      }
    }
    return comments
  }
}
