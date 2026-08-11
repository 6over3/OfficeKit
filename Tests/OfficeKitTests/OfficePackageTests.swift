import CryptoKit
import Foundation
import Testing

@testable import OfficeKit

@Test func packageMetadataRequiresOPCNamespacesAndDirectChildren() throws {
  let wrongContentTypesNamespace = Data(
    """
    <Types xmlns="urn:not-opc">
      <Default Extension="xml" ContentType="application/xml"/>
    </Types>
    """.utf8)
  #expect(throws: OfficeKitError.self) {
    _ = try PackageXMLParser.contentTypes(from: wrongContentTypesNamespace)
  }

  let nestedContentType = Data(
    """
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Extension xmlns="urn:extension">
        <Default xmlns="http://schemas.openxmlformats.org/package/2006/content-types"
          Extension="xml" ContentType="application/xml"/>
      </Extension>
    </Types>
    """.utf8)
  let contentTypes = try PackageXMLParser.contentTypes(from: nestedContentType)
  #expect(contentTypes.defaults.isEmpty)

  let wrongRelationshipsNamespace = Data(
    """
    <Relationships xmlns="urn:not-opc">
      <Relationship Id="rId1" Type="urn:type" Target="part.xml"/>
    </Relationships>
    """.utf8)
  #expect(throws: OfficeKitError.self) {
    _ = try PackageXMLParser.relationships(
      from: wrongRelationshipsNamespace,
      source: .package,
      maximumCount: 10
    )
  }

  let nestedRelationship = Data(
    """
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Extension xmlns="urn:extension">
        <Relationship xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
          Id="rId1" Type="urn:type" Target="part.xml"/>
      </Extension>
    </Relationships>
    """.utf8)
  let relationships = try PackageXMLParser.relationships(
    from: nestedRelationship,
    source: .package,
    maximumCount: 10
  )
  #expect(relationships.isEmpty)
}

@Test func seededFixtureManifestCoversAndVerifiesEveryUpstreamFile() throws {
  let manifestURL = try FixtureCatalog.url(for: "SHA256SUMS")
  let lines = try String(contentsOf: manifestURL, encoding: .utf8)
    .split(whereSeparator: \.isNewline)
  var manifestedPaths: Set<String> = []

  for line in lines {
    guard line.count > 66 else {
      Issue.record("Malformed fixture checksum line: \(line)")
      continue
    }
    let checksum = String(line.prefix(64))
    let path = String(line.dropFirst(66))
    let data = try Data(contentsOf: FixtureCatalog.url(for: path))
    let actual = SHA256.hash(data: data).map { String(format: "%02x", $0) }.joined()
    #expect(actual == checksum)
    #expect(manifestedPaths.insert(path).inserted)
  }

  let fixturesRoot = try FixtureCatalog.url(for: "Open-XML-SDK")
  let enumerator = try #require(
    FileManager.default.enumerator(at: fixturesRoot, includingPropertiesForKeys: nil))
  let fixturePaths = Set(
    enumerator.compactMap { value -> String? in
      guard let url = value as? URL, !url.hasDirectoryPath else { return nil }
      return "Open-XML-SDK/"
        + url.path.replacingOccurrences(
          of: fixturesRoot.path + "/",
          with: ""
        )
    })
  #expect(manifestedPaths == fixturePaths)
}

@Test func opensMainPartsByPackageRelationship() throws {
  let cases: [(path: String, kind: OfficeDocumentKind, mainPart: String)] = [
    ("Open-XML-SDK/TestFiles/HelloWorld.docx", .wordProcessing, "/word/document.xml"),
    ("Open-XML-SDK/TestFiles/Spreadsheet.xlsx", .spreadsheet, "/xl/workbook.xml"),
    ("Open-XML-SDK/TestFiles/Presentation.pptx", .presentation, "/ppt/presentation.xml"),
  ]

  for testCase in cases {
    let document = try OfficeDocument(contentsOf: FixtureCatalog.url(for: testCase.path))
    #expect(document.kind == testCase.kind)
    #expect(document.conformance == .transitional)
    #expect(document.mainPart.name.rawValue == testCase.mainPart)
    #expect(!document.package.parts.isEmpty)
  }
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/Documents/FlatOpcAndCloningTests.cs,
// DocumentsHaveIdenticalParts.
@Test func flatOPCOpensWithTheSamePartsAndWordSemantics() throws {
  let flatDocument = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/HelloWorldFlatOpc.xml")
  )
  let zippedDocument = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/HelloWorld.docx")
  )

  #expect(flatDocument.kind == .wordProcessing)
  #expect(flatDocument.package.parts.map(\.name) == zippedDocument.package.parts.map(\.name))
  let wordDocument = try OfficeWordDocument(document: flatDocument)
  let paragraphs = wordDocument.body.blocks.compactMap { block -> OfficeWordParagraph? in
    guard case .paragraph(let paragraph) = block else { return nil }
    return paragraph
  }
  #expect(paragraphs.first?.text == "Hello World!")
}

@Test func malformedFlatOPCBinaryDataFailsWithATypedError() throws {
  let xml = """
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/payload.bin" pkg:contentType="application/octet-stream">
        <pkg:binaryData>not-base64!</pkg:binaryData>
      </pkg:part>
    </pkg:package>
    """
  let url = FileManager.default.temporaryDirectory
    .appendingPathComponent(ProcessInfo.processInfo.globallyUniqueString)
    .appendingPathExtension("xml")
  try Data(xml.utf8).write(to: url, options: .atomic)
  defer { try? FileManager.default.removeItem(at: url) }

  #expect(throws: OfficeKitError.self) {
    _ = try OfficePackage(contentsOf: url)
  }
}

@Test func flatOPCBinaryPartsUsePackageOwnedFileURLs() throws {
  let xml = """
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/payload.bin" pkg:contentType="application/octet-stream">
        <pkg:binaryData>cGF5bG9hZA==</pkg:binaryData>
      </pkg:part>
    </pkg:package>
    """
  let sourceURL = FileManager.default.temporaryDirectory
    .appendingPathComponent(ProcessInfo.processInfo.globallyUniqueString)
    .appendingPathExtension("xml")
  try Data(xml.utf8).write(to: sourceURL, options: .atomic)
  defer { try? FileManager.default.removeItem(at: sourceURL) }

  var extractedURL: URL?
  do {
    let package = try OfficePackage(contentsOf: sourceURL)
    let name = try OfficePartName(rawValue: "/payload.bin")
    let part = try #require(package.part(named: name))
    let url = try package.fileURL(for: part)
    extractedURL = url
    #expect(try String(contentsOf: url, encoding: .utf8) == "payload")
  }
  #expect(extractedURL.map { !FileManager.default.fileExists(atPath: $0.path) } == true)
}

@Test(arguments: [
  (
    "TestDataStorage/v2FxTestFiles/presentation/macro enabled.pptm", OfficeDocumentKind.presentation
  ),
  ("TestDataStorage/v2FxTestFiles/spreadsheet/macro.xlsm", OfficeDocumentKind.spreadsheet),
  (
    "TestDataStorage/v2FxTestFiles/wordprocessing/Open XML API testing Macro enabled.docm",
    OfficeDocumentKind.wordProcessing
  ),
])
func macroEnabledFamiliesOpenWithoutExecutingEmbeddedCode(
  relativePath: String,
  expectedKind: OfficeDocumentKind
) throws {
  let url = try FixtureCatalog.url(for: "Open-XML-SDK/\(relativePath)")
  let document = try OfficeDocument(contentsOf: url)

  #expect(document.kind == expectedKind)
  #expect(document.mainPart.contentType.rawValue.localizedCaseInsensitiveContains("macroEnabled"))
  let macroParts = document.package.parts.filter {
    $0.contentType.rawValue.localizedCaseInsensitiveContains("vbaProject")
  }
  for macroPart in macroParts {
    let relationship = try #require(
      try document.package.relationships(from: .part(document.mainPart.name)).first {
        document.package.part(referencedBy: $0) == macroPart
      }
    )
    let attachment = document.package.attachment(referencedBy: relationship)
    #expect(try attachment.url().isFileURL)
  }
}

@Test func documentMetadataAndThumbnailResolveFromRootRelationships() throws {
  let document = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/DocProps.docx")
  )
  let checkedBy = try #require(
    document.metadata.customProperties.first {
      $0.name == "Checked by"
    })
  let language = try #require(
    document.metadata.customProperties.first {
      $0.name == "Language"
    })
  let presentation = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Presentation.pptx")
  )
  let thumbnail = try #require(presentation.thumbnail)

  #expect(document.metadata.creator == "Eric White")
  #expect(document.metadata.application == "Microsoft Office Word")
  #expect(document.metadata.customProperties.count == 2)
  #expect(checkedBy.identifier == 2)
  #expect(checkedBy.value == .string("Eric White"))
  #expect(language.value == .string("English"))
  #expect(presentation.metadata.title == "Title")
  #expect(presentation.metadata.creator == "Eric White")
  #expect(presentation.metadata.createdLexicalValue == "2015-07-04T11:46:55Z")
  #expect(presentation.metadata.created != nil)
  #expect(thumbnail.part != nil)
  #expect(try thumbnail.url().isFileURL)
}

@Test(arguments: seededDocuments)
func everySeededUnencryptedDocumentOpens(testCase: SeededDocument) throws {
  let document = try OfficeDocument(contentsOf: FixtureCatalog.url(for: testCase.path))
  #expect(document.kind == testCase.kind)
  #expect(!document.package.parts.isEmpty)
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Packaging.Tests/OpenXmlPackageTests.cs,
// IsEncryptedOfficeFile_ReturnsTrue_ForEncryptedFilePath.
@Test func isEncryptedOfficeFileReturnsTrueForEncryptedFilePath() throws {
  let url = try FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/encrypted_pptx.pptx")
  #expect(try OfficeDocument.isEncryptedOfficeFile(at: url))
  #expect {
    try OfficeDocument(contentsOf: url)
  } throws: { error in
    (error as? OfficeKitError) == .encryptedDocument(path: url.path)
  }
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Packaging.Tests/OpenXmlPackageTests.cs,
// IsEncryptedOfficeFile_ReturnsFalse_ForUnencryptedFile_FromString.
@Test func isEncryptedOfficeFileReturnsFalseForUnencryptedFileFromString() throws {
  let url = try FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Presentation.pptx")
  #expect(try !OfficeDocument.isEncryptedOfficeFile(at: url))
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Packaging.Tests/OpenXmlPackageTests.cs,
// TestOpenModel3DWrittenByPowerPoint_DotMime.
@Test func testOpenModel3DWrittenByPowerPointDotMime() throws {
  let document = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/3dtestdot.pptx")
  )
  let modelPart = try #require(try firstModel3DPart(in: document.package))
  #expect(modelPart.contentType.rawValue == "model/gltf.binary")
  #expect(modelPart.contentType.canonicalValue == "model/gltf-binary")
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Packaging.Tests/OpenXmlPackageTests.cs,
// TestOpenModel3DWrittenByPowerPoint_DashMime.
@Test func testOpenModel3DWrittenByPowerPointDashMime() throws {
  let document = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/3dtestdash.pptx")
  )
  let modelPart = try #require(try firstModel3DPart(in: document.package))
  #expect(modelPart.contentType.rawValue == "model/gltf-binary")
  #expect(modelPart.contentType.canonicalValue == "model/gltf-binary")
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Packaging.Tests/OpenXmlPackageTests.cs,
// ThrowWithMissingCalcChainPart.
@Test func throwWithMissingCalcChainPart() throws {
  let document = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/missingcalcchainpart.xlsx")
  )
  #expect {
    try document.package.validateRelationshipTargets()
  } throws: { error in
    guard let officeError = error as? OfficeKitError,
      case .danglingRelationship(_, _, let target) = officeError else {
      return false
    }
    return target == "/xl/calcChain.xml"
  }
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Packaging.Tests/OpenXmlPackageTests.cs,
// SucceedWithMissingCalcChainPart.
@Test func succeedWithMissingCalcChainPart() throws {
  let document = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/missingcalcchainpart.xlsx")
  )
  let diagnostics = try document.package.validateRelationshipTargets(policy: .recovering)
  #expect(document.kind == .spreadsheet)
  #expect(diagnostics.count == 1)
  #expect(diagnostics[0].severity == .warning)
  #expect(diagnostics[0].code == .danglingRelationship)
  #expect(diagnostics[0].source.part == "/xl/workbook.xml")
  #expect(diagnostics[0].source.relationshipID == OfficeRelationshipID(rawValue: "rId10"))
  #expect(diagnostics[0].target == "/xl/calcChain.xml")
}

@Test func relationshipTargetsResolveAgainstTheirSourcePart() throws {
  let source = try OfficePartName(rawValue: "/ppt/slides/slide1.xml")
  let resolved = try RelationshipTargetResolver.resolve("../media/image1.png", from: .part(source))
  #expect(resolved.part.rawValue == "/ppt/media/image1.png")
  #expect(resolved.fragment == nil)

  let withFragment = try RelationshipTargetResolver.resolve(
    "notes.xml#comment-1", from: .part(source))
  #expect(withFragment.part.rawValue == "/ppt/slides/notes.xml")
  #expect(withFragment.fragment == "comment-1")
}

@Test func partNamesRejectFilesystemAndTraversalSyntax() {
  let invalidNames = [
    "word/document.xml",
    "/../word/document.xml",
    "/word/./document.xml",
    "/word\\document.xml",
    "/word//document.xml",
    "/word/%2Fdocument.xml",
    "/word/document.xml?query",
  ]
  for name in invalidNames {
    #expect(throws: OfficeKitError.self) {
      try OfficePartName(rawValue: name)
    }
  }
}

@Test func boundedPartReadsAndStreamingCopiesPreserveAttachmentBytes() throws {
  let document = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/HelloWorld.docx")
  )
  let part = document.mainPart
  #expect(throws: OfficeKitError.self) {
    try document.package.data(for: part, limit: 1)
  }

  let data = try document.package.data(for: part)
  #expect(!data.isEmpty)

  let temporaryDirectory = FileManager.default.temporaryDirectory
    .appendingPathComponent(UUID().uuidString, isDirectory: true)
  try FileManager.default.createDirectory(
    at: temporaryDirectory, withIntermediateDirectories: true
  )
  defer { try? FileManager.default.removeItem(at: temporaryDirectory) }
  let destination = temporaryDirectory.appendingPathComponent("document.xml")
  try document.package.copy(part, to: destination)
  #expect(try Data(contentsOf: destination) == data)
  #expect {
    try document.package.copy(part, to: destination)
  } throws: { error in
    (error as? OfficeKitError) == .destinationAlreadyExists(path: destination.path)
  }
}

@Test func concurrentAttachmentURLRequestsShareOneStreamingExtraction() async throws {
  let document = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/mediareference.pptx")
  )
  let slideName = try OfficePartName(rawValue: "/ppt/slides/slide1.xml")
  let relationship = try #require(
    try document.package.relationship(
      identifiedBy: OfficeRelationshipID(rawValue: "rId4"),
      from: .part(slideName)
    )
  )
  let attachment = document.package.attachment(referencedBy: relationship)

  let urls = try await withThrowingTaskGroup(of: URL.self) { group in
    for _ in 0..<16 {
      group.addTask { try attachment.url() }
    }
    var urls: [URL] = []
    for try await url in group { urls.append(url) }
    return urls
  }

  let firstURL = try #require(urls.first)
  #expect(urls.allSatisfy { $0 == firstURL })
  #expect(FileManager.default.fileExists(atPath: firstURL.path))
}

@Test func concurrentRelationshipAndWorksheetCachesRemainDeterministic() async throws {
  let document = try OfficeDocument(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Spreadsheet.xlsx")
  )
  let workbook = try OfficeWorkbook(document: document)
  let expectedRelationshipCount = try document.package.parts.reduce(into: 0) { count, part in
    count += try document.package.relationships(from: .part(part.name)).count
  }
  var expectedCellCount = 0
  try workbook.streamRows(inWorksheetAt: 0) { row in
    expectedCellCount += row.cells.count
  }

  let counts = try await withThrowingTaskGroup(of: Int.self) { group in
    for taskIndex in 0..<16 {
      group.addTask {
        if taskIndex.isMultiple(of: 2) {
          return try document.package.parts.reduce(into: 0) { count, part in
            count += try document.package.relationships(from: .part(part.name)).count
          }
        }
        var cellCount = 0
        try workbook.streamRows(inWorksheetAt: 0) { row in
          cellCount += row.cells.count
        }
        return cellCount
      }
    }
    var counts: [Int] = []
    for try await count in group { counts.append(count) }
    return counts
  }

  #expect(counts.count == 16)
  #expect(
    counts.allSatisfy {
      $0 == expectedRelationshipCount || $0 == expectedCellCount
    })
}

private func firstModel3DPart(in package: OfficePackage) throws -> OfficePart? {
  for part in package.parts {
    for relationship in try package.relationships(from: .part(part.name))
    where relationship.type == .model3D {
      return package.part(referencedBy: relationship)
    }
  }
  return nil
}

struct SeededDocument: Sendable, CustomTestStringConvertible {
  let path: String
  let kind: OfficeDocumentKind

  var testDescription: String { path }
}

let seededDocuments = [
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/O14ISOStrict/PowerPoint/"
      + "[HC]viewPr-PresentationViewProperties-showComments-1.pptx",
    kind: .presentation
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/O14ISOStrict/PowerPoint/"
      + "webPr-WebProperties-ImageTargetResolution-imgsz-1024x786.pptx",
    kind: .presentation
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/block_hyperlink_crash.pptx",
    kind: .presentation
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/Table_Small.pptx",
    kind: .presentation
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/test.pptx",
    kind: .presentation
  ),
  SeededDocument(path: "Open-XML-SDK/TestFiles/3dtestdash.pptx", kind: .presentation),
  SeededDocument(path: "Open-XML-SDK/TestFiles/3dtestdot.pptx", kind: .presentation),
  SeededDocument(path: "Open-XML-SDK/TestFiles/Comments.docx", kind: .wordProcessing),
  SeededDocument(path: "Open-XML-SDK/TestFiles/Complex01.docx", kind: .wordProcessing),
  SeededDocument(path: "Open-XML-SDK/TestFiles/Complex01.xlsx", kind: .spreadsheet),
  SeededDocument(path: "Open-XML-SDK/TestFiles/DocProps.docx", kind: .wordProcessing),
  SeededDocument(path: "Open-XML-SDK/TestFiles/HelloWorld.docx", kind: .wordProcessing),
  SeededDocument(path: "Open-XML-SDK/TestFiles/Hyperlink.docx", kind: .wordProcessing),
  SeededDocument(path: "Open-XML-SDK/TestFiles/Presentation.pptx", kind: .presentation),
  SeededDocument(path: "Open-XML-SDK/TestFiles/Spreadsheet.xlsx", kind: .spreadsheet),
  SeededDocument(path: "Open-XML-SDK/TestFiles/Strict01.docx", kind: .wordProcessing),
  SeededDocument(path: "Open-XML-SDK/TestFiles/basicspreadsheet.xlsx", kind: .spreadsheet),
  SeededDocument(path: "Open-XML-SDK/TestFiles/mcppt.pptx", kind: .presentation),
  SeededDocument(path: "Open-XML-SDK/TestFiles/mediareference.pptx", kind: .presentation),
  SeededDocument(path: "Open-XML-SDK/TestFiles/missingcalcchainpart.xlsx", kind: .spreadsheet),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/wordprocessing/complexDocx/"
      + "complex tables.docx",
    kind: .wordProcessing
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/smallset/"
      + "Text_withExtrusion_200chars+Animation (Fly In, all at once).pptx",
    kind: .presentation
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/smallset/"
      + "Text_withExtrusion_200chars+Animation (Fly In, by letter).pptx",
    kind: .presentation
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/smallset/"
      + "Text_withExtrusion_200chars.pptx",
    kind: .presentation
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/spreadsheet/smallset/SharedWorkbook.xlsx",
    kind: .spreadsheet
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/spreadsheet/smallset/SheetData.xlsx",
    kind: .spreadsheet
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/spreadsheet/smallset/SheetViewsFSB.xlsx",
    kind: .spreadsheet
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/wordprocessing/paragraph/AdjustRightInd.docx",
    kind: .wordProcessing
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/wordprocessing/paragraph/AutoSpaceDE.docx",
    kind: .wordProcessing
  ),
  SeededDocument(
    path: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/wordprocessing/paragraph/Empty.docx",
    kind: .wordProcessing
  ),
]
