import Foundation
import ZIPFoundation

private let flatOPCNamespace = "http://schemas.microsoft.com/office/2006/xmlPackage"

package enum FlatOPCConverter {
  package static func convert(
    contentsOf sourceURL: URL,
    limits: OfficeParsingLimits
  ) throws -> URL? {
    guard try looksLikeXML(sourceURL) else { return nil }
    let document = try OfficeXMLDocument(
      reading: OfficeXMLReader(
        contentsOf: sourceURL,
        limits: limits.xmlParsingLimits
      )
    )
    guard document.root.name.namespaceURI == flatOPCNamespace,
      document.root.name.localName == "package" else { return nil }

    let payloads = try payloads(in: document.root, limits: limits)
    let temporaryURL = FileManager.default.temporaryDirectory
      .appendingPathComponent("OfficeKit-FlatOPC-\(UUID().uuidString)")
      .appendingPathExtension("zip")
    do {
      let archive = try Archive(url: temporaryURL, accessMode: .create)
      let contentTypes = contentTypesXML(for: payloads)
      try add(contentTypes, at: "[Content_Types].xml", to: archive)
      for payload in payloads {
        try add(payload.data, at: payload.name.archivePath, to: archive)
      }
      return temporaryURL
    } catch {
      try? FileManager.default.removeItem(at: temporaryURL)
      if let error = error as? OfficeKitError { throw error }
      throw OfficeKitError.invalidPackage("Could not prepare the Flat OPC package.")
    }
  }

  private struct Payload {
    let name: OfficePartName
    let contentType: String
    let data: Data
  }

  private static func looksLikeXML(_ url: URL) throws -> Bool {
    let handle = try FileHandle(forReadingFrom: url)
    defer { try? handle.close() }
    let data = try handle.read(upToCount: 1_024) ?? Data()
    let bytes = Array(data.prefix(4))
    if bytes.starts(with: [0xFF, 0xFE]) || bytes.starts(with: [0xFE, 0xFF]) {
      return true
    }
    if bytes.starts(with: [0x00, 0x00, 0xFE, 0xFF])
      || bytes.starts(with: [0xFF, 0xFE, 0x00, 0x00])
      || bytes.starts(with: [0x00, 0x3C, 0x00, 0x3F])
      || bytes.starts(with: [0x3C, 0x00, 0x3F, 0x00])
    {
      return true
    }

    var index = data.starts(with: [0xEF, 0xBB, 0xBF]) ? 3 : 0
    while index < data.count, [0x09, 0x0A, 0x0D, 0x20].contains(data[index]) {
      index += 1
    }
    return index < data.count && data[index] == 0x3C
  }

  private static func payloads(
    in root: OfficeXMLElement,
    limits: OfficeParsingLimits
  ) throws -> [Payload] {
    let partElements = root.childElements.filter {
      $0.name.namespaceURI == flatOPCNamespace && $0.name.localName == "part"
    }
    guard partElements.count < limits.maximumEntryCount else {
      throw OfficeKitError.limitExceeded(
        limit: .entryCount,
        actual: UInt64(partElements.count + 1),
        maximum: UInt64(limits.maximumEntryCount)
      )
    }

    var payloads: [Payload] = []
    payloads.reserveCapacity(partElements.count)
    var names: Set<OfficePartName> = []
    var totalSize: UInt64 = 0
    for partElement in partElements {
      guard let rawName = partElement.attribute(named: "name", namespaceURI: flatOPCNamespace),
        let contentType = partElement.attribute(
          named: "contentType",
          namespaceURI: flatOPCNamespace
        ) else {
        throw OfficeKitError.invalidPackage("A Flat OPC part is missing its name or content type.")
      }
      let name = try OfficePartName(rawValue: rawName)
      guard name.archivePath != "[Content_Types].xml" else {
        throw OfficeKitError.invalidPackage("Flat OPC cannot redefine [Content_Types].xml.")
      }
      guard names.insert(name).inserted else {
        throw OfficeKitError.duplicatePartName(name.rawValue)
      }
      let data = try payload(in: partElement)
      guard UInt64(data.count) <= limits.maximumEntrySize else {
        throw OfficeKitError.limitExceeded(
          limit: .entrySize,
          actual: UInt64(data.count),
          maximum: limits.maximumEntrySize
        )
      }
      let (newTotal, overflow) = totalSize.addingReportingOverflow(UInt64(data.count))
      guard !overflow, newTotal <= limits.maximumTotalUncompressedSize else {
        throw OfficeKitError.limitExceeded(
          limit: .totalUncompressedSize,
          actual: overflow ? .max : newTotal,
          maximum: limits.maximumTotalUncompressedSize
        )
      }
      totalSize = newTotal
      payloads.append(Payload(name: name, contentType: contentType, data: data))
    }
    return payloads
  }

  private static func payload(in part: OfficeXMLElement) throws -> Data {
    let containers = part.childElements.filter { $0.name.namespaceURI == flatOPCNamespace }
    guard containers.count == 1, let container = containers.first else {
      throw OfficeKitError.invalidPackage(
        "A Flat OPC part must contain exactly one XML or binary payload."
      )
    }
    switch container.name.localName {
    case "xmlData":
      guard container.childElements.count == 1 else {
        throw OfficeKitError.invalidPackage(
          "A Flat OPC XML payload must contain exactly one document element."
        )
      }
      return Data(FlatOPCXMLSerializer.serialize(container.children).utf8)
    case "binaryData":
      let lexicalValue = container.textContent.filter { !$0.isWhitespace }
      guard let data = Data(base64Encoded: lexicalValue) else {
        throw OfficeKitError.invalidPackage("A Flat OPC binary payload is not valid Base64.")
      }
      return data
    default:
      throw OfficeKitError.invalidPackage("A Flat OPC part has an unknown payload container.")
    }
  }

  private static func contentTypesXML(for payloads: [Payload]) -> Data {
    var xml = """
      <?xml version="1.0" encoding="UTF-8"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      """
    for payload in payloads {
      xml += "<Override PartName=\"\(escapedAttribute(payload.name.rawValue))\""
      xml += " ContentType=\"\(escapedAttribute(payload.contentType))\"/>"
    }
    xml += "</Types>"
    return Data(xml.utf8)
  }

  private static func add(_ data: Data, at path: String, to archive: Archive) throws {
    try archive.addEntry(
      with: path,
      type: .file,
      uncompressedSize: Int64(data.count)
    ) { position, size in
      data.subdata(in: Int(position)..<Int(position) + size)
    }
  }

  private static func escapedAttribute(_ value: String) -> String {
    value
      .replacingOccurrences(of: "&", with: "&amp;")
      .replacingOccurrences(of: "\"", with: "&quot;")
      .replacingOccurrences(of: "<", with: "&lt;")
  }
}

private enum FlatOPCXMLSerializer {
  static func serialize(_ nodes: [OfficeXMLNode]) -> String {
    var result = ""
    var namespaces: [String: String] = ["xml": "http://www.w3.org/XML/1998/namespace"]
    for node in nodes {
      append(node, namespaces: &namespaces, to: &result)
    }
    return result
  }

  private static func append(
    _ node: OfficeXMLNode,
    namespaces: inout [String: String],
    to result: inout String
  ) {
    switch node {
    case .text(let text, _):
      result += escapedText(text)
    case .element(let element):
      append(element, namespaces: &namespaces, to: &result)
    }
  }

  private static func append(
    _ element: OfficeXMLElement,
    namespaces: inout [String: String],
    to result: inout String
  ) {
    let parentNamespaces = namespaces
    var declarations = element.namespaceDeclarations
    for declaration in declarations {
      if declaration.namespaceURI.isEmpty {
        namespaces.removeValue(forKey: declaration.prefix)
      } else {
        namespaces[declaration.prefix] = declaration.namespaceURI
      }
    }
    addRequiredBinding(
      for: element.name,
      usesDefaultNamespace: true,
      declarations: &declarations,
      namespaces: &namespaces
    )
    for attribute in element.attributes {
      addRequiredBinding(
        for: attribute.name,
        usesDefaultNamespace: false,
        declarations: &declarations,
        namespaces: &namespaces
      )
    }

    result += "<\(element.name.description)"
    for declaration in declarations {
      let name = declaration.prefix.isEmpty ? "xmlns" : "xmlns:\(declaration.prefix)"
      result += " \(name)=\"\(escapedAttribute(declaration.namespaceURI))\""
    }
    for attribute in element.attributes {
      result += " \(attribute.name.description)=\"\(escapedAttribute(attribute.value))\""
    }
    if element.children.isEmpty {
      result += "/>"
    } else {
      result += ">"
      for child in element.children {
        append(child, namespaces: &namespaces, to: &result)
      }
      result += "</\(element.name.description)>"
    }
    namespaces = parentNamespaces
  }

  private static func addRequiredBinding(
    for name: OfficeXMLName,
    usesDefaultNamespace: Bool,
    declarations: inout [OfficeXMLNamespaceDeclaration],
    namespaces: inout [String: String]
  ) {
    let prefix = name.prefix ?? ""
    guard let namespaceURI = name.namespaceURI else {
      if usesDefaultNamespace, name.prefix == nil, namespaces[""] != nil {
        declarations.append(OfficeXMLNamespaceDeclaration(prefix: "", namespaceURI: ""))
        namespaces.removeValue(forKey: "")
      }
      return
    }
    if let existing = namespaces[prefix],
      OfficeXMLNamespace.canonicalize(existing) == namespaceURI
    {
      return
    }
    declarations.removeAll { $0.prefix == prefix }
    declarations.append(
      OfficeXMLNamespaceDeclaration(prefix: prefix, namespaceURI: namespaceURI)
    )
    namespaces[prefix] = namespaceURI
  }

  private static func escapedText(_ value: String) -> String {
    value
      .replacingOccurrences(of: "&", with: "&amp;")
      .replacingOccurrences(of: "<", with: "&lt;")
      .replacingOccurrences(of: ">", with: "&gt;")
  }

  private static func escapedAttribute(_ value: String) -> String {
    escapedText(value)
      .replacingOccurrences(of: "\"", with: "&quot;")
      .replacingOccurrences(of: "\r", with: "&#13;")
      .replacingOccurrences(of: "\n", with: "&#10;")
      .replacingOccurrences(of: "\t", with: "&#9;")
  }
}
