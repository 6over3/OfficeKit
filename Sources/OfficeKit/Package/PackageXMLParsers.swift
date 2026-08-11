import Foundation

private let contentTypesNamespace =
  "http://schemas.openxmlformats.org/package/2006/content-types"
private let packageRelationshipsNamespace =
  "http://schemas.openxmlformats.org/package/2006/relationships"

package struct ContentTypeMap: Sendable {
  let defaults: [String: OfficeContentType]
  let overrides: [OfficePartName: OfficeContentType]

  func contentType(for partName: OfficePartName) -> OfficeContentType? {
    if let override = overrides[partName] { return override }
    let extensionStart = partName.rawValue.lastIndex(of: ".")
    guard let extensionStart else { return nil }
    let fileExtension = partName.rawValue[partName.rawValue.index(after: extensionStart)...]
    return defaults[fileExtension.lowercased()]
  }
}

package enum PackageXMLParser {
  static func contentTypes(from data: Data) throws -> ContentTypeMap {
    let source = "/[Content_Types].xml"
    var defaults: [String: OfficeContentType] = [:]
    var overrides: [OfficePartName: OfficeContentType] = [:]
    var sawRoot = false
    var depth = 0
    try OfficeXMLReader(data: data, source: source).parse { event in
      switch event {
      case .startElement(let name, let attributes, _, _):
        depth += 1
        if depth == 1 {
          guard name.localName == "Types", name.namespaceURI == contentTypesNamespace else {
            throw OfficeKitError.invalidXML(
              part: source,
              message: "The document element must be Types in the OPC content-types namespace."
            )
          }
          sawRoot = true
        } else if depth == 2, name.namespaceURI == contentTypesNamespace {
          switch name.localName {
          case "Default":
            guard let fileExtension = attribute(named: "Extension", in: attributes),
              let contentType = attribute(named: "ContentType", in: attributes),
              !fileExtension.isEmpty, !contentType.isEmpty else {
              throw OfficeKitError.invalidXML(
                part: source, message: "Default requires Extension and ContentType."
              )
            }
            let key = fileExtension.lowercased()
            if let existing = defaults[key], existing.rawValue != contentType {
              throw OfficeKitError.invalidPackage(
                "Conflicting default content types for .\(key)."
              )
            }
            defaults[key] = OfficeContentType(rawValue: contentType)
          case "Override":
            guard let rawPartName = attribute(named: "PartName", in: attributes),
              let contentType = attribute(named: "ContentType", in: attributes),
              !contentType.isEmpty else {
              throw OfficeKitError.invalidXML(
                part: source, message: "Override requires PartName and ContentType."
              )
            }
            let partName = try OfficePartName(rawValue: rawPartName)
            if let existing = overrides[partName], existing.rawValue != contentType {
              throw OfficeKitError.invalidPackage(
                "Conflicting content-type overrides for \(partName.rawValue)."
              )
            }
            overrides[partName] = OfficeContentType(rawValue: contentType)
          default:
            break
          }
        }
      case .endElement:
        depth -= 1
      default:
        break
      }
    }
    guard sawRoot else {
      throw OfficeKitError.invalidXML(
        part: source, message: "Missing Types root element."
      )
    }
    return ContentTypeMap(defaults: defaults, overrides: overrides)
  }

  static func relationships(
    from data: Data,
    source: OfficeRelationshipSource,
    maximumCount: Int
  ) throws -> [OfficeRelationship] {
    let part = "/" + source.relationshipArchivePath
    let maximumCount = max(0, maximumCount)
    var relationships: [OfficeRelationship] = []
    var identifiers: Set<OfficeRelationshipID> = []
    var sawRoot = false
    var depth = 0
    try OfficeXMLReader(data: data, source: part).parse { event in
      switch event {
      case .startElement(let name, let attributes, _, _):
        depth += 1
        if depth == 1 {
          guard name.localName == "Relationships",
            name.namespaceURI == packageRelationshipsNamespace else {
            throw OfficeKitError.invalidXML(
              part: part,
              message:
                "The document element must be Relationships in the OPC relationships namespace."
            )
          }
          sawRoot = true
        } else if depth == 2, name.namespaceURI == packageRelationshipsNamespace,
          name.localName == "Relationship"
        {
          guard relationships.count < maximumCount else {
            throw OfficeKitError.limitExceeded(
              limit: .relationshipsPerPart,
              actual: UInt64(relationships.count + 1),
              maximum: UInt64(maximumCount)
            )
          }
          guard let rawID = attribute(named: "Id", in: attributes), !rawID.isEmpty,
            let rawType = attribute(named: "Type", in: attributes), !rawType.isEmpty,
            let rawTarget = attribute(named: "Target", in: attributes), !rawTarget.isEmpty else {
            throw OfficeKitError.invalidXML(
              part: part, message: "Relationship requires Id, Type, and Target."
            )
          }
          let id = OfficeRelationshipID(rawValue: rawID)
          guard identifiers.insert(id).inserted else {
            throw OfficeKitError.invalidPackage(
              "Duplicate relationship identifier \(rawID) in \(source.description)."
            )
          }
          let target: OfficeRelationshipTarget
          if attribute(named: "TargetMode", in: attributes)?
            .caseInsensitiveCompare("External") == .orderedSame
          {
            target = .external(rawTarget)
          } else {
            let resolved = try RelationshipTargetResolver.resolve(rawTarget, from: source)
            target = .internalPart(resolved.part, fragment: resolved.fragment)
          }
          relationships.append(
            OfficeRelationship(
              source: source,
              id: id,
              type: OfficeRelationshipType(rawValue: rawType),
              target: target,
              rawTarget: rawTarget
            )
          )
        }
      case .endElement:
        depth -= 1
      default:
        break
      }
    }
    guard sawRoot else {
      throw OfficeKitError.invalidXML(part: part, message: "Missing Relationships root element.")
    }
    return relationships
  }

  private static func attribute(
    named localName: String,
    in attributes: [OfficeXMLAttribute]
  ) -> String? {
    attributes.first { $0.name.namespaceURI == nil && $0.name.localName == localName }?.value
  }
}

package enum RelationshipTargetResolver {
  package static func resolve(
    _ rawTarget: String,
    from source: OfficeRelationshipSource
  ) throws -> (part: OfficePartName, fragment: String?) {
    guard !rawTarget.contains("\\"), !rawTarget.hasPrefix("//") else {
      throw OfficeKitError.invalidPartName(rawTarget)
    }

    let targetAndFragment = rawTarget.split(
      separator: "#", maxSplits: 1, omittingEmptySubsequences: false)
    let targetPath = String(targetAndFragment[0])
    let fragment = targetAndFragment.count == 2 ? String(targetAndFragment[1]) : nil
    guard !targetPath.isEmpty, !targetPath.contains("?") else {
      throw OfficeKitError.invalidPartName(rawTarget)
    }

    var components: [Substring]
    if targetPath.hasPrefix("/") {
      components = []
    } else {
      switch source {
      case .package:
        components = []
      case .part(let sourcePart):
        components = sourcePart.archivePath.split(separator: "/")
        if !components.isEmpty { components.removeLast() }
      }
    }

    let relativeComponents = targetPath.split(separator: "/", omittingEmptySubsequences: false)
    for component in relativeComponents {
      if component.isEmpty {
        if targetPath.hasPrefix("/"), component == relativeComponents.first { continue }
        throw OfficeKitError.invalidPartName(rawTarget)
      }
      switch component {
      case ".":
        continue
      case "..":
        guard !components.isEmpty else { throw OfficeKitError.invalidPartName(rawTarget) }
        components.removeLast()
      default:
        if component.contains(":"), components.isEmpty {
          throw OfficeKitError.invalidPartName(rawTarget)
        }
        components.append(component)
      }
    }

    let resolvedPath = "/" + components.joined(separator: "/")
    return (try OfficePartName(rawValue: resolvedPath), fragment?.isEmpty == true ? nil : fragment)
  }
}
