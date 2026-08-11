import Foundation
import Testing

@testable import OfficeKit

@Test func deterministicLexicalMutationCorpusNeverTraps() {
  var generator = DeterministicGenerator(state: 0x4F_46_46_49_43_45)
  let alphabet = Array("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789$:+-.eEtruefals#_/".utf8)

  for _ in 0..<5_000 {
    let length = Int(generator.next() % 48)
    let bytes = (0..<length).map { _ in alphabet[Int(generator.next() % UInt64(alphabet.count))] }
    let value = String(decoding: bytes, as: UTF8.self)
    _ = OfficeCellReference(rawValue: value)
    _ = OfficeCellRange(rawValue: value)
    _ = Int(value)
    _ = Double(value)
    _ = Bool(value)
    _ = OfficeValueDecoder.boolean(value)
    _ = OfficeValueDecoder.decimal(value)
  }
}

@Test func deterministicXMLMutationCorpusFailsBoundedlyOrParses() {
  let original = Array("<r xmlns='urn:test'><a x='1'>text</a><b/></r>".utf8)
  var generator = DeterministicGenerator(state: 0x58_4D_4C)
  let limits = OfficeXMLParsingLimits(
    maximumDepth: 32,
    maximumAttributesPerElement: 16,
    maximumEventCount: 1_024,
    maximumTextSize: 4_096
  )

  for iteration in 0..<512 {
    var bytes = original
    let mutationCount = 1 + Int(generator.next() % 4)
    for _ in 0..<mutationCount {
      let index = Int(generator.next() % UInt64(bytes.count))
      bytes[index] ^= UInt8(truncatingIfNeeded: generator.next())
    }
    let reader = OfficeXMLReader(
      data: Data(bytes),
      source: "mutation-\(iteration).xml",
      limits: limits
    )
    do {
      try reader.parse { _ in }
    } catch is OfficeKitError {
      continue
    } catch {
      Issue.record("XML mutation escaped the documented error domain: \(error)")
    }
  }
}

@Test func deterministicOPCMetadataMutationCorpusNeverTraps() {
  let contentTypes = Array(
    ("<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>"
      + "<Default Extension='xml' ContentType='application/xml'/>"
      + "<Override PartName='/word/document.xml' ContentType='application/xml'/>"
      + "</Types>").utf8)
  let relationships = Array(
    ("<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
      + "<Relationship Id='rId1' Type='urn:test' Target='word/document.xml'/>"
      + "</Relationships>").utf8)
  var generator = DeterministicGenerator(state: 0x4F_50_43)

  for _ in 0..<512 {
    var contentTypeMutation = contentTypes
    var relationshipMutation = relationships
    mutate(&contentTypeMutation, using: &generator)
    mutate(&relationshipMutation, using: &generator)
    do {
      _ = try PackageXMLParser.contentTypes(from: Data(contentTypeMutation))
    } catch {
      // Malformed mutations are expected; completing without a trap is the invariant.
    }
    do {
      _ = try PackageXMLParser.relationships(
        from: Data(relationshipMutation),
        source: .package,
        maximumCount: 32
      )
    } catch {
      // Malformed mutations are expected; completing without a trap is the invariant.
    }
  }
}

@Test func deterministicZIPCentralDirectoryMutationCorpusNeverTraps() throws {
  let sourceURL = try FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/HelloWorld.docx")
  let original = try Data(contentsOf: sourceURL)
  let temporaryDirectory = FileManager.default.temporaryDirectory
    .appendingPathComponent("OfficeKitFuzz-\(UUID().uuidString)", isDirectory: true)
  try FileManager.default.createDirectory(
    at: temporaryDirectory,
    withIntermediateDirectories: true
  )
  defer { try? FileManager.default.removeItem(at: temporaryDirectory) }

  let centralDirectoryWindow = min(original.count, 2_048)
  for iteration in 0..<64 {
    var mutation = original
    let distanceFromEnd = 1 + ((iteration * 31) % centralDirectoryWindow)
    let index = mutation.index(mutation.endIndex, offsetBy: -distanceFromEnd)
    mutation[index] ^= UInt8(truncatingIfNeeded: iteration &* 17 &+ 1)
    let url = temporaryDirectory.appendingPathComponent("mutation-\(iteration).docx")
    try mutation.write(to: url, options: .atomic)

    do {
      let document = try OfficeDocument(contentsOf: url)
      var visitor = NoOpVisitor()
      try document.traverse(using: &visitor)
    } catch {
      continue
    }
  }
}

private struct NoOpVisitor: OfficeDocumentVisitor {}

private struct DeterministicGenerator {
  var state: UInt64

  mutating func next() -> UInt64 {
    state = state &* 6_364_136_223_846_793_005 &+ 1_442_695_040_888_963_407
    return state
  }
}

private func mutate(_ bytes: inout [UInt8], using generator: inout DeterministicGenerator) {
  let mutationCount = 1 + Int(generator.next() % 4)
  for _ in 0..<mutationCount {
    let index = Int(generator.next() % UInt64(bytes.count))
    bytes[index] ^= UInt8(truncatingIfNeeded: generator.next())
  }
}
