package enum WordStyleParser {
  private static let wordNamespace =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

  private struct Builder {
    let startDepth: Int
    let identifier: String
    let type: String
    let isDefault: Bool
    var name: String?
    var basedOnIdentifier: String?
    var nextIdentifier: String?
    var isPrimary = false
    var numberingIdentifier: Int?
    var numberingLevel: Int?
    var paragraphProperties = WordParagraphPropertiesBuilder()
    var runProperties = WordRunPropertiesBuilder()
  }

  package struct Result {
    let styles: [OfficeWordStyle]
    let defaults: OfficeWordStyleDefaults?
  }

  package static func parse(
    documentPart: OfficePart,
    package: OfficePackage,
    relationships: [OfficeRelationship]
  ) throws -> Result {
    guard let relationship = relationships.first(where: { $0.type.isEquivalent(to: .styles) }) else {
      return Result(styles: [], defaults: nil)
    }
    guard let part = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }

    var depth = 0
    var paragraphPropertiesDepth: Int?
    var runPropertiesDepth: Int?
    var defaultsDepth: Int?
    var defaultParagraphPropertiesDepth: Int?
    var defaultRunPropertiesDepth: Int?
    var defaultParagraphProperties = WordParagraphPropertiesBuilder()
    var defaultRunProperties = WordRunPropertiesBuilder()
    var builder: Builder?
    var styles: [OfficeWordStyle] = []
    try package.parseXML(in: part, compatibility: .commonOffice) { event in
      switch event {
      case .startElement(let name, let attributes, _, _):
        depth += 1
        guard name.namespaceURI == wordNamespace else { return }
        if name.localName == "docDefaults" { defaultsDepth = depth }
        if name.localName == "style" {
          guard let identifier = wordAttribute("styleId", in: attributes),
            let type = wordAttribute("type", in: attributes) else {
            throw OfficeKitError.invalidXML(
              part: part.name.rawValue,
              message: "w:style requires w:styleId and w:type attributes."
            )
          }
          builder = Builder(
            startDepth: depth,
            identifier: identifier,
            type: type,
            isDefault: wordOnOffValue(attributes) ?? false
          )
          return
        }
        if defaultsDepth != nil, builder == nil {
          if name.localName == "pPr" { defaultParagraphPropertiesDepth = depth }
          if name.localName == "rPr" { defaultRunPropertiesDepth = depth }
          if defaultParagraphPropertiesDepth != nil {
            defaultParagraphProperties.consume(name.localName, attributes: attributes)
          }
          if defaultRunPropertiesDepth != nil {
            defaultRunProperties.consume(name.localName, attributes: attributes)
          }
        }
        guard builder != nil else { return }
        switch name.localName {
        case "name": builder?.name = wordAttribute("val", in: attributes)
        case "basedOn": builder?.basedOnIdentifier = wordAttribute("val", in: attributes)
        case "next": builder?.nextIdentifier = wordAttribute("val", in: attributes)
        case "qFormat": builder?.isPrimary = wordOnOffValue(attributes) ?? false
        case "pPr":
          if depth == (builder?.startDepth ?? 0) + 1 { paragraphPropertiesDepth = depth }
        case "rPr":
          if depth == (builder?.startDepth ?? 0) + 1 { runPropertiesDepth = depth }
        case "numId":
          if paragraphPropertiesDepth != nil {
            builder?.numberingIdentifier = wordAttribute("val", in: attributes).flatMap(Int.init)
          }
        case "ilvl":
          if paragraphPropertiesDepth != nil {
            builder?.numberingLevel = wordAttribute("val", in: attributes).flatMap(Int.init)
          }
        default: break
        }
        if paragraphPropertiesDepth != nil {
          builder?.paragraphProperties.consume(name.localName, attributes: attributes)
        }
        if runPropertiesDepth != nil {
          builder?.runProperties.consume(name.localName, attributes: attributes)
        }
      case .endElement(let name, _):
        if name.namespaceURI == wordNamespace, name.localName == "pPr",
          paragraphPropertiesDepth == depth
        {
          paragraphPropertiesDepth = nil
        }
        if name.namespaceURI == wordNamespace, name.localName == "rPr",
          runPropertiesDepth == depth
        {
          runPropertiesDepth = nil
        }
        if name.namespaceURI == wordNamespace, name.localName == "pPr",
          defaultParagraphPropertiesDepth == depth
        {
          defaultParagraphPropertiesDepth = nil
        }
        if name.namespaceURI == wordNamespace, name.localName == "rPr",
          defaultRunPropertiesDepth == depth
        {
          defaultRunPropertiesDepth = nil
        }
        if name.namespaceURI == wordNamespace, name.localName == "style",
          builder?.startDepth == depth, let completed = builder
        {
          styles.append(
            OfficeWordStyle(
              identifier: completed.identifier,
              type: completed.type,
              name: completed.name,
              basedOnIdentifier: completed.basedOnIdentifier,
              nextIdentifier: completed.nextIdentifier,
              isDefault: completed.isDefault,
              isPrimary: completed.isPrimary,
              numberingIdentifier: completed.numberingIdentifier,
              numberingLevel: completed.numberingLevel,
              paragraphProperties: completed.paragraphProperties.value,
              runProperties: completed.runProperties.value,
              sourcePart: part
            )
          )
          builder = nil
        }
        if name.namespaceURI == wordNamespace, name.localName == "docDefaults",
          defaultsDepth == depth
        {
          defaultsDepth = nil
        }
        depth -= 1
      case .text, .startDocument, .endDocument:
        break
      }
    }
    return Result(
      styles: styles,
      defaults: OfficeWordStyleDefaults(
        paragraphProperties: defaultParagraphProperties.value,
        runProperties: defaultRunProperties.value,
        sourcePart: part
      )
    )
  }
}
