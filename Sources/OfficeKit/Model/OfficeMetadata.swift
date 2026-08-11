import Foundation

/// A typed value from an OOXML custom document property.
public enum OfficeMetadataValue: Sendable, Hashable, Codable {
  /// A string property.
  case string(String)
  /// A signed integer property.
  case integer(Int64)
  /// A floating-point property.
  case double(Double)
  /// A Boolean property.
  case boolean(Bool)
  /// A parsed date together with its exact source spelling.
  case date(Date, lexicalValue: String)
  /// An unrecognized property type retained losslessly as text.
  case unknown(type: String, lexicalValue: String)
}

/// One named custom document property in declaration order.
public struct OfficeCustomProperty: Sendable, Hashable, Codable {
  /// Producer-assigned property identifier.
  public let identifier: Int?

  /// Property name.
  public let name: String

  /// Format identifier used by the producer.
  public let formatIdentifier: String?

  /// Typed property value.
  public let value: OfficeMetadataValue
}

/// Core, extended, and custom metadata read from package property parts.
public struct OfficeMetadata: Sendable, Hashable, Codable {
  /// Document title.
  public let title: String?
  /// Document subject.
  public let subject: String?
  /// Original creator or author.
  public let creator: String?
  /// Producer-authored keyword text.
  public let keywords: String?
  /// Document description or comments.
  public let description: String?
  /// Name of the last modifying user.
  public let lastModifiedBy: String?
  /// Producer-authored revision token.
  public let revision: String?
  /// Document category.
  public let category: String?
  /// Document workflow status.
  public let contentStatus: String?
  /// Document language tag.
  public let language: String?
  /// Parsed creation timestamp.
  public let created: Date?
  /// Exact creation timestamp spelling.
  public let createdLexicalValue: String?
  /// Parsed modification timestamp.
  public let modified: Date?
  /// Exact modification timestamp spelling.
  public let modifiedLexicalValue: String?
  /// Producer application name.
  public let application: String?
  /// Producer application version.
  public let applicationVersion: String?
  /// Organization or company name.
  public let company: String?
  /// Manager name.
  public let manager: String?
  /// Typed custom properties in declaration order.
  public let customProperties: [OfficeCustomProperty]
}

package enum OfficeMetadataParser {
  private struct PropertyBuilder {
    let identifier: Int?
    let name: String
    let formatIdentifier: String?
    let startDepth: Int
    var valueType: String?
    var value = ""
  }

  package static func parse(
    package: OfficePackage,
    rootRelationships: [OfficeRelationship]
  ) throws -> OfficeMetadata {
    var scalarValues: [String: String] = [:]
    var customProperties: [OfficeCustomProperty] = []
    for relationship in rootRelationships {
      guard
        relationship.type.isEquivalent(to: .coreProperties)
          || relationship.type.isEquivalent(to: .extendedProperties)
          || relationship.type.isEquivalent(to: .customProperties) else { continue }
      guard let part = package.part(referencedBy: relationship) else {
        throw OfficeKitError.missingPart(relationship.rawTarget)
      }
      if relationship.type.isEquivalent(to: .customProperties) {
        customProperties = try parseCustomProperties(part: part, package: package)
      } else {
        let parsedValues = try parseScalarProperties(part: part, package: package)
        scalarValues.merge(parsedValues, uniquingKeysWith: { _, last in last })
      }
    }
    let createdText = scalarValues["created"]
    let modifiedText = scalarValues["modified"]
    return OfficeMetadata(
      title: scalarValues["title"],
      subject: scalarValues["subject"],
      creator: scalarValues["creator"],
      keywords: scalarValues["keywords"],
      description: scalarValues["description"],
      lastModifiedBy: scalarValues["lastModifiedBy"],
      revision: scalarValues["revision"],
      category: scalarValues["category"],
      contentStatus: scalarValues["contentStatus"],
      language: scalarValues["language"],
      created: createdText.flatMap(ISO8601DateFormatter().date(from:)),
      createdLexicalValue: createdText,
      modified: modifiedText.flatMap(ISO8601DateFormatter().date(from:)),
      modifiedLexicalValue: modifiedText,
      application: scalarValues["Application"],
      applicationVersion: scalarValues["AppVersion"],
      company: scalarValues["Company"],
      manager: scalarValues["Manager"],
      customProperties: customProperties
    )
  }

  private static func parseScalarProperties(
    part: OfficePart,
    package: OfficePackage
  ) throws -> [String: String] {
    let recognizedNames: Set<String> = [
      "title", "subject", "creator", "keywords", "description", "lastModifiedBy",
      "revision", "category", "contentStatus", "language", "created", "modified",
      "Application", "AppVersion", "Company", "Manager",
    ]
    var depth = 0
    var capture: (name: String, depth: Int, text: String)?
    var values: [String: String] = [:]
    try package.parseXML(in: part) { event in
      switch event {
      case .startElement(let name, _, _, _):
        depth += 1
        if recognizedNames.contains(name.localName) {
          capture = (name.localName, depth, "")
        }
      case .text(let text, _):
        capture?.text.append(text)
      case .endElement(let name, _):
        if capture?.depth == depth, capture?.name == name.localName, let completed = capture {
          values[completed.name] = completed.text
          capture = nil
        }
        depth -= 1
      case .startDocument, .endDocument:
        break
      }
    }
    return values
  }

  private static func parseCustomProperties(
    part: OfficePart,
    package: OfficePackage
  ) throws -> [OfficeCustomProperty] {
    var depth = 0
    var builder: PropertyBuilder?
    var valueDepth: Int?
    var properties: [OfficeCustomProperty] = []
    try package.parseXML(in: part) { event in
      switch event {
      case .startElement(let name, let attributes, _, _):
        depth += 1
        if name.localName == "property",
          let propertyName = metadataAttribute("name", in: attributes)
        {
          builder = PropertyBuilder(
            identifier: metadataAttribute("pid", in: attributes).flatMap(Int.init),
            name: propertyName,
            formatIdentifier: metadataAttribute("fmtid", in: attributes),
            startDepth: depth
          )
        } else if builder != nil, depth == (builder?.startDepth ?? 0) + 1 {
          builder?.valueType = name.localName
          valueDepth = depth
        }
      case .text(let text, _):
        if valueDepth != nil { builder?.value.append(text) }
      case .endElement(let name, _):
        if valueDepth == depth { valueDepth = nil }
        if name.localName == "property", builder?.startDepth == depth, let completed = builder {
          properties.append(
            OfficeCustomProperty(
              identifier: completed.identifier,
              name: completed.name,
              formatIdentifier: completed.formatIdentifier,
              value: metadataValue(type: completed.valueType ?? "unknown", text: completed.value)
            )
          )
          builder = nil
        }
        depth -= 1
      case .startDocument, .endDocument:
        break
      }
    }
    return properties
  }

  private static func metadataValue(type: String, text: String) -> OfficeMetadataValue {
    switch type {
    case "lpstr", "lpwstr", "bstr": return .string(text)
    case "i1", "i2", "i4", "i8", "int", "ui1", "ui2", "ui4", "ui8", "uint":
      return Int64(text).map(OfficeMetadataValue.integer)
        ?? .unknown(type: type, lexicalValue: text)
    case "r4", "r8", "decimal":
      return Double(text).map(OfficeMetadataValue.double)
        ?? .unknown(type: type, lexicalValue: text)
    case "bool":
      return OfficeValueDecoder.boolean(text).map(OfficeMetadataValue.boolean)
        ?? .unknown(type: type, lexicalValue: text)
    case "date", "filetime":
      return ISO8601DateFormatter().date(from: text).map {
        .date($0, lexicalValue: text)
      } ?? .unknown(type: type, lexicalValue: text)
    default:
      return .unknown(type: type, lexicalValue: text)
    }
  }
}

private func metadataAttribute(
  _ localName: String,
  in attributes: [OfficeXMLAttribute]
) -> String? {
  attributes.first { $0.name.localName == localName }?.value
}
