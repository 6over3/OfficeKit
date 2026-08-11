package enum SpreadsheetCommentParser {
  private static let spreadsheetNamespace =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  private struct RawComment {
    let reference: OfficeCellReference
    let authorIndex: UInt32
    let shapeIdentifier: UInt32?
    var text = ""
  }

  package static func parse(part: OfficePart, package: OfficePackage) throws
    -> [OfficeWorksheetComment]
  {
    var authors: [String] = []
    var comments: [RawComment] = []
    var currentAuthor: String?
    var currentComment: RawComment?
    var authorDepth: Int?
    var textDepth: Int?
    var depth = 0

    try package.parseXML(in: part, compatibility: .commonOffice) { event in
      switch event {
      case .startElement(let name, let attributes, _, _):
        depth += 1
        guard name.namespaceURI == spreadsheetNamespace else { return }
        if name.localName == "author" {
          currentAuthor = ""
          authorDepth = depth
        } else if name.localName == "comment" {
          guard let rawReference = commentAttribute("ref", in: attributes),
            let reference = OfficeCellReference(rawValue: rawReference),
            let authorIndex = commentAttribute("authorId", in: attributes).flatMap(UInt32.init) else {
            throw OfficeKitError.invalidPackage("Worksheet contains an incomplete comment.")
          }
          currentComment = RawComment(
            reference: reference,
            authorIndex: authorIndex,
            shapeIdentifier: commentAttribute("shapeId", in: attributes).flatMap(UInt32.init)
          )
        } else if name.localName == "t", currentComment != nil {
          textDepth = depth
        }
      case .text(let text, _):
        if authorDepth != nil { currentAuthor?.append(text) }
        if textDepth != nil { currentComment?.text.append(text) }
      case .endElement(let name, _):
        if authorDepth == depth, name.namespaceURI == spreadsheetNamespace,
          name.localName == "author",
          let author = currentAuthor
        {
          authors.append(author)
          currentAuthor = nil
          authorDepth = nil
        }
        if textDepth == depth, name.namespaceURI == spreadsheetNamespace,
          name.localName == "t"
        {
          textDepth = nil
        }
        if name.namespaceURI == spreadsheetNamespace, name.localName == "comment",
          let comment = currentComment
        {
          comments.append(comment)
          currentComment = nil
        }
        depth -= 1
      case .startDocument, .endDocument:
        break
      }
    }

    return comments.map { comment in
      OfficeWorksheetComment(
        reference: comment.reference,
        authorIndex: comment.authorIndex,
        author: Int(comment.authorIndex) < authors.count ? authors[Int(comment.authorIndex)] : nil,
        shapeIdentifier: comment.shapeIdentifier,
        text: comment.text,
        sourcePart: part,
        shape: nil
      )
    }
  }
}

private func commentAttribute(
  _ localName: String,
  in attributes: [OfficeXMLAttribute]
) -> String? {
  attributes.first { $0.name.localName == localName && $0.name.namespaceURI == nil }?.value
}
