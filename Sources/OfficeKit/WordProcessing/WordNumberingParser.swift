package enum WordNumberingParser {
  private static let wordNamespace =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

  private struct AbstractBuilder {
    let startDepth: Int
    let identifier: Int
    var multiLevelType: String?
    var levels: [OfficeWordNumberingLevel] = []
  }

  private struct LevelBuilder {
    let startDepth: Int
    let index: Int
    var start: Int?
    var format: String?
    var text: String?
    var justification: String?
    var paragraphStyleIdentifier: String?
    var leftIndent: OfficeLength?
    var hangingIndent: OfficeLength?

    var value: OfficeWordNumberingLevel {
      OfficeWordNumberingLevel(
        index: index,
        start: start,
        format: format,
        text: text,
        justification: justification,
        paragraphStyleIdentifier: paragraphStyleIdentifier,
        leftIndent: leftIndent,
        hangingIndent: hangingIndent
      )
    }
  }

  private struct InstanceBuilder {
    let startDepth: Int
    let identifier: Int
    var abstractIdentifier: Int?
    var overrides: [OfficeWordNumberingLevelOverride] = []
  }

  private struct OverrideBuilder {
    let startDepth: Int
    let levelIndex: Int
    var start: Int?
    var level: OfficeWordNumberingLevel?
  }

  package static func parse(
    package: OfficePackage,
    relationships: [OfficeRelationship]
  ) throws -> OfficeWordNumbering? {
    guard
      let relationship = relationships.first(where: {
        $0.type.isEquivalent(to: .numbering)
      }) else { return nil }
    guard let part = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }

    var depth = 0
    var abstract: AbstractBuilder?
    var level: LevelBuilder?
    var instance: InstanceBuilder?
    var override: OverrideBuilder?
    var abstracts: [OfficeWordAbstractNumbering] = []
    var instances: [OfficeWordNumberingInstance] = []

    try package.parseXML(in: part, compatibility: .commonOffice) { event in
      switch event {
      case .startElement(let name, let attributes, _, _):
        depth += 1
        guard name.namespaceURI == wordNamespace else { return }
        switch name.localName {
        case "abstractNum":
          guard
            let identifier = wordAttribute(
              "abstractNumId",
              in: attributes
            ).flatMap(Int.init) else {
            throw OfficeKitError.invalidXML(
              part: part.name.rawValue,
              message: "w:abstractNum requires a numeric w:abstractNumId attribute."
            )
          }
          abstract = AbstractBuilder(startDepth: depth, identifier: identifier)
        case "multiLevelType":
          abstract?.multiLevelType = wordAttribute("val", in: attributes)
        case "lvl":
          guard abstract != nil || override != nil,
            let index = wordAttribute("ilvl", in: attributes).flatMap(Int.init) else { break }
          level = LevelBuilder(startDepth: depth, index: index)
        case "start": level?.start = wordAttribute("val", in: attributes).flatMap(Int.init)
        case "numFmt": level?.format = wordAttribute("val", in: attributes)
        case "lvlText": level?.text = wordAttribute("val", in: attributes)
        case "lvlJc": level?.justification = wordAttribute("val", in: attributes)
        case "pStyle": level?.paragraphStyleIdentifier = wordAttribute("val", in: attributes)
        case "ind":
          level?.leftIndent = wordAttribute("left", in: attributes).flatMap(wordTwipLength)
          level?.hangingIndent = wordAttribute("hanging", in: attributes).flatMap(wordTwipLength)
        case "num":
          guard let identifier = wordAttribute("numId", in: attributes).flatMap(Int.init) else {
            throw OfficeKitError.invalidXML(
              part: part.name.rawValue,
              message: "w:num requires a numeric w:numId attribute."
            )
          }
          instance = InstanceBuilder(startDepth: depth, identifier: identifier)
        case "abstractNumId":
          instance?.abstractIdentifier = wordAttribute("val", in: attributes).flatMap(Int.init)
        case "lvlOverride":
          guard instance != nil,
            let index = wordAttribute("ilvl", in: attributes).flatMap(Int.init) else { break }
          override = OverrideBuilder(startDepth: depth, levelIndex: index)
        case "startOverride":
          override?.start = wordAttribute("val", in: attributes).flatMap(Int.init)
        default:
          break
        }
      case .endElement(let name, _):
        if name.namespaceURI == wordNamespace {
          if name.localName == "lvl", level?.startDepth == depth, let completed = level {
            if override != nil {
              override?.level = completed.value
            } else {
              abstract?.levels.append(completed.value)
            }
            level = nil
          }
          if name.localName == "lvlOverride", override?.startDepth == depth,
            let completed = override
          {
            instance?.overrides.append(
              OfficeWordNumberingLevelOverride(
                levelIndex: completed.levelIndex,
                start: completed.start,
                level: completed.level
              )
            )
            override = nil
          }
          if name.localName == "abstractNum", abstract?.startDepth == depth,
            let completed = abstract
          {
            abstracts.append(
              OfficeWordAbstractNumbering(
                identifier: completed.identifier,
                multiLevelType: completed.multiLevelType,
                levels: completed.levels
              )
            )
            abstract = nil
          }
          if name.localName == "num", instance?.startDepth == depth, let completed = instance {
            guard let abstractIdentifier = completed.abstractIdentifier else {
              throw OfficeKitError.invalidXML(
                part: part.name.rawValue,
                message: "w:num \(completed.identifier) is missing w:abstractNumId."
              )
            }
            instances.append(
              OfficeWordNumberingInstance(
                identifier: completed.identifier,
                abstractIdentifier: abstractIdentifier,
                levelOverrides: completed.overrides
              )
            )
            instance = nil
          }
        }
        depth -= 1
      case .text, .startDocument, .endDocument:
        break
      }
    }
    return OfficeWordNumbering(
      abstractDefinitions: abstracts,
      instances: instances,
      sourcePart: part
    )
  }
}
