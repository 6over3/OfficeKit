package enum WordNoteParser {
  private static let wordNamespace =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

  private struct Builder {
    let startDepth: Int
    let identifier: Int64
    let type: String?
    let parser: WordDocumentParser
  }

  package static func parse(
    kind: OfficeWordNoteKind,
    package: OfficePackage,
    relationships: [OfficeRelationship]
  ) throws -> OfficeWordNoteCollection? {
    let relationshipType: OfficeRelationshipType = kind == .footnote ? .footnotes : .endnotes
    guard
      let relationship = relationships.first(where: {
        $0.type.isEquivalent(to: relationshipType)
      }) else { return nil }
    guard let part = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }
    let noteElement = kind == .footnote ? "footnote" : "endnote"
    let noteRelationships = try package.relationships(from: .part(part.name))
    var depth = 0
    var builder: Builder?
    var notes: [OfficeWordNote] = []
    try package.parseXML(in: part, compatibility: .commonOffice) { event in
      switch event {
      case .startElement(let name, let attributes, _, _):
        depth += 1
        guard name.namespaceURI == wordNamespace else { return }
        if name.localName == noteElement {
          guard let identifierText = wordAttribute("id", in: attributes),
            let identifier = Int64(identifierText) else {
            throw OfficeKitError.invalidXML(
              part: part.name.rawValue,
              message: "w:\(noteElement) requires a numeric w:id attribute."
            )
          }
          builder = Builder(
            startDepth: depth,
            identifier: identifier,
            type: wordAttribute("type", in: attributes),
            parser: WordDocumentParser(
              part: part,
              package: package,
              relationships: noteRelationships
            )
          )
        }
        try builder?.parser.consume(event)
      case .text:
        try builder?.parser.consume(event)
      case .endElement(let name, _):
        try builder?.parser.consume(event)
        if name.namespaceURI == wordNamespace {
          if name.localName == noteElement, builder?.startDepth == depth,
            let completed = builder
          {
            let content = completed.parser.body
            let paragraphs = content.paragraphs.map(\.text)
            notes.append(
              OfficeWordNote(
                kind: kind,
                identifier: completed.identifier,
                type: completed.type,
                paragraphs: paragraphs,
                text: paragraphs.joined(separator: "\n"),
                content: content,
                sourcePart: part
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
    return OfficeWordNoteCollection(
      kind: kind,
      notes: notes,
      sourcePart: part,
      attachments: noteRelationships.map(package.attachment(referencedBy:))
    )
  }
}
