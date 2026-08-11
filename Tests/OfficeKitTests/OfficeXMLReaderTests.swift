import Foundation
import Testing
import UniformTypeIdentifiers

@testable import OfficeKit

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Packaging.Tests/OpenXmlPartReaderTests.cs,
// ExtractsInfoFromStream.
@Test(arguments: [
  ("<?xml version='1.0' encoding='UTF-8' standalone='yes'?><root/>", "UTF-8", true),
  ("<?xml version='1.0' encoding='UTF-32' standalone='yes'?><root/>", "UTF-32", true),
  ("<?xml version='1.0' standalone='yes'?><root/>", nil, true),
  ("<?xml version='1.0' standalone='no'?><root/>", nil, false),
  ("<?xml version='1.0'?><root/>", nil, nil),
  ("<?xml version='1.0'?>", nil, nil),
  ("<root/>", nil, nil),
])
func extractsInfoFromStream(xml: String, encoding: String?, standalone: Bool?) {
  let reader = OfficeXMLReader(data: Data(xml.utf8))
  #expect(reader.declaration?.encoding == encoding)
  #expect(reader.declaration?.isStandalone == standalone)
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Packaging.Tests/OpenXmlPartReaderTests.cs,
// CreateElement.
@Test func createElement() throws {
  let xml = """
    <?xml version='1.0' encoding='utf-8' standalone='yes'?>
    <a:root xmlns:a="http://www.tests.com/">
      <a:child>Test</a:child>
      <a:child>Test2</a:child>
    </a:root>
    """
  var names: [OfficeXMLName] = []
  try OfficeXMLReader(data: Data(xml.utf8)).parse { event in
    switch event {
    case .startElement(let name, _, _, _), .endElement(let name, _):
      names.append(name)
    default:
      break
    }
  }

  #expect(names.count == 6)
  #expect(names.allSatisfy { $0.prefix == "a" })
  #expect(names.allSatisfy { $0.namespaceURI == "http://www.tests.com/" })
  #expect(names.map(\.localName) == ["root", "child", "child", "child", "child", "root"])
}

@Test func namespaceDeclarationsAndNamespacedAttributesArePreserved() throws {
  let xml = """
    <p:root xmlns:p="urn:p" xmlns:r="urn:r" r:id="rId1" plain="value"/>
    """
  var start:
    (
      OfficeXMLName,
      [OfficeXMLAttribute],
      [OfficeXMLNamespaceDeclaration]
    )?
  try OfficeXMLReader(data: Data(xml.utf8)).parse { event in
    if case .startElement(let name, let attributes, let declarations, _) = event {
      start = (name, attributes, declarations)
    }
  }

  let element = try #require(start)
  #expect(element.0 == OfficeXMLName(namespaceURI: "urn:p", localName: "root", prefix: "p"))
  #expect(
    Set(element.2)
      == Set([
        OfficeXMLNamespaceDeclaration(prefix: "p", namespaceURI: "urn:p"),
        OfficeXMLNamespaceDeclaration(prefix: "r", namespaceURI: "urn:r"),
      ]))
  #expect(
    element.1.contains {
      $0.name == OfficeXMLName(namespaceURI: "urn:r", localName: "id", prefix: "r")
        && $0.value == "rId1"
    })
  #expect(
    OfficeXMLName(namespaceURI: "urn:r", localName: "id", prefix: "r")
      == OfficeXMLName(namespaceURI: "urn:r", localName: "id", prefix: "relationship")
  )
  #expect(
    element.1.contains {
      $0.name == OfficeXMLName(namespaceURI: nil, localName: "plain") && $0.value == "value"
    })
}

@Test func reservedXMLAttributesRetainTheirNamespace() throws {
  let xml = "<root xml:id='item' xml:space='preserve'/>"
  let document = try OfficeXMLDocument(reading: OfficeXMLReader(data: Data(xml.utf8)))
  let namespace = "http://www.w3.org/XML/1998/namespace"

  #expect(document.root.attribute(named: "id", namespaceURI: namespace) == "item")
  #expect(document.root.attribute(named: "space", namespaceURI: namespace) == "preserve")
}

@Test func undeclaredPrefixesAndDuplicateExpandedAttributesAreRejected() {
  #expect(throws: OfficeKitError.self) {
    try OfficeXMLReader(data: Data("<missing:root/>".utf8)).parse { _ in }
  }
  #expect(throws: OfficeKitError.self) {
    let xml = "<root xmlns:a='urn:value' xmlns:b='urn:value' a:id='1' b:id='2'/>"
    try OfficeXMLReader(data: Data(xml.utf8)).parse { _ in }
  }
}

@Test func strictNamespacesCanonicalizeWithoutDependingOnPrefixes() throws {
  let strict = "http://purl.oclc.org/ooxml/presentationml/main"
  let transitional = "http://schemas.openxmlformats.org/presentationml/2006/main"
  let xml = "<anything:presentation xmlns:anything='\(strict)'/>"
  var root: OfficeXMLName?
  try OfficeXMLReader(data: Data(xml.utf8)).parse { event in
    if case .startElement(let name, _, _, _) = event { root = name }
  }
  #expect(root?.namespaceURI == transitional)
  #expect(root?.localName == "presentation")
  #expect(root?.prefix == "anything")
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// src/DocumentFormat.OpenXml.Framework/Features/OpenXmlNamespaceResolver.cs,
// Strict-to-Transitional namespace map used by OpenXmlNamespaceTests.
@Test(arguments: [
  ("descriptions/base", "http://descriptions.openxmlformats.org/description/base"),
  ("descriptions/full", "http://descriptions.openxmlformats.org/description/full"),
  ("drawingml/compatibility", "http://schemas.openxmlformats.org/drawingml/2006/compatibility"),
  ("drawingml/lockedCanvas", "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas"),
  (
    "officeDocument/bibliography",
    "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
  ),
  (
    "officeDocument/customProperties",
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
  ),
  ("officeDocument/customXml", "http://schemas.openxmlformats.org/officeDocument/2006/customXml"),
  (
    "officeDocument/customXmlDataProps",
    "http://schemas.openxmlformats.org/officeDocument/2006/customXmlDataProps"
  ),
  (
    "officeDocument/docPropsVTypes",
    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
  ),
  (
    "officeDocument/extendedProperties",
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  ),
  ("officeDocument/math", "http://schemas.openxmlformats.org/officeDocument/2006/math"),
  (
    "officeDocument/sharedTypes",
    "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes"
  ),
  ("schemaLibrary/main", "http://schemas.openxmlformats.org/schemaLibrary/2006/main"),
])
func strictNamespaceTableMatchesThePinnedSDK(suffix: String, transitional: String) {
  #expect(OfficeXMLNamespace.canonicalize("http://purl.oclc.org/ooxml/\(suffix)") == transitional)
}

@Test func strictCustomXMLRelationshipUsesTheISOErratumMapping() {
  let strict = OfficeRelationshipType(
    rawValue: "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml"
  )
  #expect(
    strict.canonicalValue
      == "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
  )
}

@Test(arguments: [
  (
    "http://schemas.openxmlformats.org/wordprocessingml/2006/6/main",
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  ),
  (
    "http://schemas.openxmlformats.org/spreadsheetml/2006/7/main",
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  ),
  (
    "http://schemas.openxmlformats.org/presentationml/2006/3/main",
    "http://schemas.openxmlformats.org/presentationml/2006/main"
  ),
  (
    "http://schemas.openxmlformats.org/drawingml/2006/3/main",
    "http://schemas.openxmlformats.org/drawingml/2006/main"
  ),
  (
    "http://schemas.microsoft.com/office/word/2010/11/wordml",
    "http://schemas.microsoft.com/office/word/2012/wordml"
  ),
])
func namespaceVersionAliasesMatchThePinnedSDK(alias: String, canonical: String) {
  #expect(OfficeXMLNamespace.canonicalize(alias) == canonical)
}

@Test func packageXMLStreamingReadsARealPresentationRoot() throws {
  let document = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Presentation.pptx")
  )
  var root: OfficeXMLName?
  try document.package.parseXML(in: document.mainPart) { event in
    guard root == nil else { return }
    if case .startElement(let name, _, _, _) = event { root = name }
  }
  #expect(root?.localName == "presentation")
  #expect(root?.namespaceURI == "http://schemas.openxmlformats.org/presentationml/2006/main")
}

@Test func relationshipIdentifiersResolveInternalMediaAndExternalHyperlinks() throws {
  let presentation = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/mediareference.pptx")
  )
  let slide = try #require(
    presentation.package.part(named: OfficePartName(rawValue: "/ppt/slides/slide1.xml"))
  )
  let imageRelationship = try #require(
    try presentation.package.relationship(
      identifiedBy: OfficeRelationshipID(rawValue: "rId4"),
      from: .part(slide.name)
    )
  )
  let image = try #require(presentation.package.part(referencedBy: imageRelationship))
  #expect(image.name.rawValue == "/ppt/media/image1.png")
  #expect(image.contentType.rawValue == "image/png")

  let attachment = presentation.package.attachment(referencedBy: imageRelationship)
  let firstURL = try attachment.url()
  let secondURL = try attachment.url()
  #expect(firstURL == secondURL)
  #expect(firstURL.isFileURL)
  #expect(FileManager.default.fileExists(atPath: firstURL.path))
  let fileSize = try firstURL.resourceValues(forKeys: [.fileSizeKey]).fileSize
  #expect(fileSize == Int(image.uncompressedSize))
  #expect(attachment.filenameHint == "image1.png")
  #expect(attachment.contentType == .png)
  #expect(attachment.declaredContentType == OfficeContentType(rawValue: "image/png"))

  let word = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Hyperlink.docx")
  )
  let hyperlink = try #require(
    try word.package.relationship(
      identifiedBy: OfficeRelationshipID(rawValue: "rId4"),
      from: .part(word.mainPart.name)
    )
  )
  #expect(hyperlink.target == .external("http://www.ericwhite.com"))
  let externalAttachment = word.package.attachment(referencedBy: hyperlink)
  #expect(try externalAttachment.url().absoluteString == "http://www.ericwhite.com")
}

@Test func xmlReaderRejectsDocumentTypesAndEnforcesResourceLimits() throws {
  let documentType = "<!DOCTYPE root [<!ENTITY x 'expanded'>]><root>&x;</root>"
  #expect(throws: OfficeKitError.self) {
    try OfficeXMLReader(data: Data(documentType.utf8)).parse { _ in }
  }

  let utf16DocumentType = documentType.data(using: .utf16LittleEndian) ?? Data()
  #expect(throws: OfficeKitError.self) {
    try OfficeXMLReader(data: utf16DocumentType).parse { _ in }
  }

  let deeplyNested = "<a><b><c/></b></a>"
  #expect {
    try OfficeXMLReader(
      data: Data(deeplyNested.utf8),
      limits: OfficeXMLParsingLimits(maximumDepth: 2)
    ).parse { _ in }
  } throws: { error in
    (error as? OfficeKitError) == .limitExceeded(limit: .xmlDepth, actual: 3, maximum: 2)
  }

  let tooManyAttributes = "<root a='1' b='2'/>"
  #expect {
    try OfficeXMLReader(
      data: Data(tooManyAttributes.utf8),
      limits: OfficeXMLParsingLimits(maximumAttributesPerElement: 1)
    ).parse { _ in }
  } throws: { error in
    (error as? OfficeKitError)
      == .limitExceeded(
        limit: .xmlAttributesPerElement,
        actual: 2,
        maximum: 1
      )
  }

  let text = "<root>12345</root>"
  #expect {
    try OfficeXMLReader(
      data: Data(text.utf8),
      limits: OfficeXMLParsingLimits(maximumTextSize: 4)
    ).parse { _ in }
  } throws: { error in
    (error as? OfficeKitError) == .limitExceeded(limit: .xmlTextSize, actual: 5, maximum: 4)
  }
}

private struct ConsumerStopped: Error, Equatable {}

@Test func xmlReaderRethrowsConsumerErrorsUnchanged() {
  #expect(throws: ConsumerStopped.self) {
    try OfficeXMLReader(data: Data("<root/>".utf8)).parse { event in
      if case .startElement = event { throw ConsumerStopped() }
    }
  }
}

@Test func packageXMLLimitDefaultsApplyToEveryPartParse() throws {
  let document = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/HelloWorld.docx"),
    limits: OfficeParsingLimits(
      xmlParsingLimits: OfficeXMLParsingLimits(maximumDepth: 2)
    )
  )
  #expect(throws: OfficeKitError.self) {
    try document.package.parseXML(in: document.mainPart) { _ in }
  }
}

@Test func xmlReaderCooperativelyStopsWhenItsTaskIsCancelled() async {
  let source = "<root>" + String(repeating: "<value>1</value>", count: 2_000) + "</root>"
  let reader = OfficeXMLReader(data: Data(source.utf8))
  let task = Task {
    do {
      try await Task.sleep(for: .seconds(60))
    } catch {
      // Continue into the synchronous reader with the task's cancellation bit set.
    }
    try reader.parse { _ in }
  }
  task.cancel()
  do {
    try await task.value
    Issue.record("A cancelled XML parse unexpectedly completed.")
  } catch {
    #expect(error is CancellationError)
  }
}

@Test func fileBackedXMLReaderStreamsAndScansAcrossChunkBoundaries() throws {
  let temporaryDirectory = FileManager.default.temporaryDirectory
    .appendingPathComponent(UUID().uuidString, isDirectory: true)
  try FileManager.default.createDirectory(at: temporaryDirectory, withIntermediateDirectories: true)
  defer { try? FileManager.default.removeItem(at: temporaryDirectory) }

  let validURL = temporaryDirectory.appendingPathComponent("valid.xml")
  try Data("<?xml version='1.0' encoding='UTF-8'?><root><child/></root>".utf8)
    .write(to: validURL)
  let reader = OfficeXMLReader(contentsOf: validURL)
  #expect(reader.declaration?.encoding == "UTF-8")
  var elementCount = 0
  try reader.parse { event in
    if case .startElement = event { elementCount += 1 }
  }
  #expect(elementCount == 2)

  let documentTypeURL = temporaryDirectory.appendingPathComponent("doctype.xml")
  var documentType = Data(repeating: 0x20, count: 65_532)
  documentType.append(Data("<!DOCTYPE root><root/>".utf8))
  try documentType.write(to: documentTypeURL)
  #expect(throws: OfficeKitError.self) {
    try OfficeXMLReader(contentsOf: documentTypeURL).parse { _ in }
  }
}

@Test func packageXMLUsesFileBackedParsingAboveItsMemoryThreshold() throws {
  let limits = OfficeParsingLimits(maximumInMemoryXMLPartSize: 0)
  let document = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Presentation.pptx"),
    limits: limits
  )
  var rootName: String?
  try document.package.parseXML(in: document.mainPart) { event in
    guard rootName == nil else { return }
    if case .startElement(let name, _, _, _) = event { rootName = name.localName }
  }
  #expect(rootName == "presentation")
}

@Test func packageMetadataRejectsDocumentTypesBeyondTheFormerPrefixBoundary() {
  let padding = String(repeating: " ", count: 5_000)
  let contentTypes = Data(
    ("<?xml version='1.0'?>\(padding)"
      + "<!DOCTYPE Types [<!ENTITY payload 'expanded'>]>"
      + "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>"
      + "&payload;</Types>").utf8)
  let relationships = Data(
    ("<?xml version='1.0'?>\(padding)"
      + "<!DOCTYPE Relationships [<!ENTITY payload 'expanded'>]>"
      + "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
      + "&payload;</Relationships>").utf8)

  #expect(throws: OfficeKitError.self) {
    try PackageXMLParser.contentTypes(from: contentTypes)
  }
  #expect(throws: OfficeKitError.self) {
    try PackageXMLParser.relationships(
      from: relationships,
      source: .package,
      maximumCount: 10
    )
  }
}
