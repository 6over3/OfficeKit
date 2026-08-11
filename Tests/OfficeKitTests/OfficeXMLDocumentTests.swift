import Foundation
import Testing

@testable import OfficeKit

@Test func structuredXMLRetainsUnknownContentAndDocumentOrder() throws {
  let xml = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <x:root xmlns:x="urn:unknown" xmlns:m="urn:metadata" m:flag="yes">before<x:item n="2">inside</x:item>after</x:root>
    """
  let document = try OfficeXMLDocument(reading: OfficeXMLReader(data: Data(xml.utf8)))

  #expect(document.declaration?.encoding == "UTF-8")
  #expect(document.declaration?.isStandalone == true)
  #expect(document.root.name.namespaceURI == "urn:unknown")
  #expect(document.root.attribute(named: "flag", namespaceURI: "urn:metadata") == "yes")
  #expect(document.root.textContent == "beforeinsideafter")
  #expect(document.root.childElements.map(\.name.localName) == ["item"])

  var visited: [(String, [Int])] = []
  document.traverseElements { element, path in
    visited.append((element.name.localName, path))
  }
  #expect(visited.map(\.0) == ["root", "item"])
  #expect(visited[1].1 == [1])
}

@Test func arbitraryXMLPartResolvesEveryRelationshipValuedAttribute() throws {
  let relationshipNamespace =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  let packageURL = try makeSyntheticOfficePackage(
    entries: [
      "[Content_Types].xml": """
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>
      """,
      "custom/data.xml": """
      <future:root xmlns:future="urn:future" xmlns:r="\(relationshipNamespace)" xmlns:not-r="urn:extension/relationships" r:id="rId1" not-r:id="not-a-relationship"><future:item r:embed="rId2"/><future:missing r:href="rId404"/></future:root>
      """,
      "custom/_rels/data.xml.rels": """
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:payload" Target="../media/blob.bin"/><Relationship Id="rId2" Type="urn:external" Target="https://example.com/resource" TargetMode="External"/></Relationships>
      """,
      "media/blob.bin": "payload",
    ],
    pathExtension: "zip"
  )
  defer { try? FileManager.default.removeItem(at: packageURL) }

  let package = try OfficePackage(contentsOf: packageURL)
  let partName = try OfficePartName(rawValue: "/custom/data.xml")
  let part = try #require(package.part(named: partName))
  let parsed = try package.parsedXMLPart(part)

  var streamedXMLParts: Set<OfficePartName> = []
  var streamedEventCount = 0
  try package.streamXMLParts { part, event in
    if case .startDocument = event { streamedXMLParts.insert(part.name) }
    streamedEventCount += 1
  }

  #expect(parsed.document.root.name.localName == "root")
  #expect(part.isXML)
  #expect(package.parts.first { $0.name.rawValue == "/media/blob.bin" }?.isXML == false)
  #expect(streamedXMLParts.count == 2)
  #expect(streamedEventCount > 0)
  #expect(parsed.relationships.count == 2)
  #expect(parsed.attachments.count == 2)
  #expect(
    parsed.relationshipReferences.map { $0.relationshipID.rawValue }
      == ["rId1", "rId2", "rId404"]
  )
  #expect(parsed.relationshipReferences.map { $0.elementPath } == [[], [0], [1]])
  #expect(parsed.relationshipReferences[2].relationship == nil)
  #expect(parsed.relationshipReferences[2].attachment == nil)

  let internalURL = try #require(parsed.relationshipReferences[0].attachment).url()
  #expect(try String(contentsOf: internalURL, encoding: .utf8) == "payload")
  let externalURL = try #require(parsed.relationshipReferences[1].attachment).url()
  #expect(externalURL == URL(string: "https://example.com/resource"))
}

@Test func structuredParserCoversEveryPinnedSchemaElementDeclaration() throws {
  let manifestURL = try FixtureCatalog.url(for: "SchemaCoverage/elements.tsv")
  let rows = try String(contentsOf: manifestURL, encoding: .utf8)
    .split(whereSeparator: \.isNewline)
  let schemaManifestURL = try FixtureCatalog.url(for: "SchemaCoverage/schemas.tsv")
  let schemaRows = try String(contentsOf: schemaManifestURL, encoding: .utf8)
    .split(whereSeparator: \.isNewline)
  var contributingSchemaFiles: Set<Substring> = []
  var expandedNames: Set<OfficeXMLName> = []
  var declaredElementCount = 0
  var declaredEnumerationCount = 0

  #expect(rows.count == 4_119)
  #expect(schemaRows.count == 155)
  for schemaRow in schemaRows {
    let fields = schemaRow.split(separator: "\t", omittingEmptySubsequences: false)
    #expect(fields.count == 4)
    guard fields.count == 4 else { continue }
    declaredElementCount += Int(fields[2]) ?? 0
    declaredEnumerationCount += Int(fields[3]) ?? 0
  }
  #expect(declaredElementCount == 4_119)
  #expect(declaredEnumerationCount == 582)
  for row in rows {
    let fields = row.split(separator: "\t", omittingEmptySubsequences: false)
    #expect(fields.count == 5)
    guard fields.count == 5 else { continue }
    contributingSchemaFiles.insert(fields[0])

    let namespaceURI = String(fields[1])
    let schemaName = fields[2].split(separator: "/").last ?? ""
    let localName = schemaName.split(separator: ":").last ?? ""
    let escapedNamespace =
      namespaceURI
      .replacingOccurrences(of: "&", with: "&amp;")
      .replacingOccurrences(of: "\"", with: "&quot;")
    let xml = "<schema:\(localName) xmlns:schema=\"\(escapedNamespace)\"/>"
    let document = try OfficeXMLDocument(reading: OfficeXMLReader(data: Data(xml.utf8)))

    #expect(document.root.name.localName == String(localName))
    #expect(
      document.root.name.namespaceURI
        == OfficeXMLNamespace.canonicalize(namespaceURI)
    )
    expandedNames.insert(
      OfficeXMLName(namespaceURI: namespaceURI, localName: String(localName))
    )
  }

  #expect(contributingSchemaFiles.count == 139)
  #expect(expandedNames.count == 3_702)
}

@Test func structuredParserCoversEveryPinnedSchemaAttributeDeclaration() throws {
  let manifestURL = try FixtureCatalog.url(for: "SchemaCoverage/attributes.tsv")
  let rows = try String(contentsOf: manifestURL, encoding: .utf8)
    .split(whereSeparator: \.isNewline)
  var expandedNames: Set<OfficeXMLName> = []

  #expect(rows.count == 8_510)
  for row in rows {
    let fields = row.split(separator: "\t", omittingEmptySubsequences: false)
    #expect(fields.count == 8)
    guard fields.count == 8 else { continue }

    let namespaceURI = String(fields[3])
    let localName = String(fields[4])
    let escapedNamespace =
      namespaceURI
      .replacingOccurrences(of: "&", with: "&amp;")
      .replacingOccurrences(of: "\"", with: "&quot;")
    let xml: String
    if namespaceURI.isEmpty {
      xml = "<root \(localName)=\"value\"/>"
    } else if namespaceURI == "http://www.w3.org/XML/1998/namespace" {
      xml = "<root xml:\(localName)=\"value\"/>"
    } else {
      xml = "<root xmlns:attribute=\"\(escapedNamespace)\" attribute:\(localName)=\"value\"/>"
    }
    let document = try OfficeXMLDocument(reading: OfficeXMLReader(data: Data(xml.utf8)))
    let attribute = try #require(document.root.attributes.first)
    let expectedNamespace =
      namespaceURI.isEmpty
      ? nil
      : OfficeXMLNamespace.canonicalize(namespaceURI)

    #expect(attribute.name.localName == localName)
    #expect(attribute.name.namespaceURI == expectedNamespace)
    #expect(attribute.value == "value")
    expandedNames.insert(attribute.name)
  }

  #expect(expandedNames.count == 2_567)
}

@Test func packageLayerCoversEveryPinnedPartDefinition() throws {
  let manifestURL = try FixtureCatalog.url(for: "SchemaCoverage/parts.tsv")
  let rows = try String(contentsOf: manifestURL, encoding: .utf8)
    .split(whereSeparator: \.isNewline)
  var contentTypes: Set<OfficeContentType> = []
  var relationshipTypes: Set<OfficeRelationshipType> = []
  var declaredContentTypes = 0
  var declaredRelationshipTypes = 0
  var rootedParts = 0

  #expect(rows.count == 128)
  for row in rows {
    let fields = row.split(separator: "\t", omittingEmptySubsequences: false)
    #expect(fields.count == 6)
    guard fields.count == 6 else { continue }
    let contentType = String(fields[2])
    let relationshipType = String(fields[3])
    let rootType = String(fields[4])
    if !contentType.isEmpty {
      declaredContentTypes += 1
      let value = OfficeContentType(rawValue: contentType)
      contentTypes.insert(value)
      if !rootType.isEmpty { #expect(value.isXML) }
    }
    if !relationshipType.isEmpty {
      declaredRelationshipTypes += 1
      relationshipTypes.insert(OfficeRelationshipType(rawValue: relationshipType))
    }
    if !rootType.isEmpty { rootedParts += 1 }
  }

  #expect(declaredContentTypes == 111)
  #expect(contentTypes.count == 108)
  #expect(declaredRelationshipTypes == 125)
  #expect(relationshipTypes.count == 118)
  #expect(rootedParts == 96)
}
