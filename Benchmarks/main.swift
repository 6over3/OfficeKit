import Darwin
import Foundation
import OfficeKit

private struct BenchmarkVisitor: OfficeDocumentVisitor {
  var cells = 0
  var runs = 0
  var slideElements = 0
  var attachments = 0

  mutating func visitCell(_: OfficeCell) { cells += 1 }
  mutating func visitWordRun(_: OfficeWordRun) { runs += 1 }
  mutating func willVisitSlideElement(_: OfficeSlideElement) { slideElements += 1 }
  mutating func visitAttachment(_: OfficeAttachment) { attachments += 1 }
}

private func seconds(_ duration: Duration) -> Double {
  let components = duration.components
  return Double(components.seconds) + Double(components.attoseconds) / 1e18
}

private func writeJSONLine(_ value: [String: Any], to handle: FileHandle) throws {
  let data = try JSONSerialization.data(withJSONObject: value, options: [.sortedKeys])
  handle.write(data)
  handle.write(Data("\n".utf8))
}

private func peakResidentMemoryBytes() -> UInt64 {
  var usage = rusage()
  guard getrusage(RUSAGE_SELF, &usage) == 0 else { return 0 }
  return UInt64(max(usage.ru_maxrss, 0))
}

struct ExpectedFailure {
  let pathSuffix: String
  let errorSubstring: String
}

func expectedFailures(in file: String) throws -> [ExpectedFailure] {
  try String(contentsOfFile: file, encoding: .utf8)
    .split(whereSeparator: \.isNewline)
    .compactMap { line in
      guard !line.hasPrefix("#") else { return nil }
      let fields = line.split(separator: "\t", maxSplits: 1, omittingEmptySubsequences: false)
      guard fields.count == 2, !fields[0].isEmpty, !fields[1].isEmpty else { return nil }
      return ExpectedFailure(pathSuffix: String(fields[0]), errorSubstring: String(fields[1]))
    }
}

var arguments = Array(CommandLine.arguments.dropFirst())
var maximumSeconds = 5.0
var maximumMemoryBytes: UInt64?
var continuesAfterErrors = false
var streamsRowsOnly = false
var runsDetailedMetrics = false
var scansAllXML = false
var expectedFailureFile: String?
while let option = arguments.first, option.hasPrefix("--") {
  switch option {
  case "--maximum-seconds":
    guard arguments.count >= 2, let parsed = Double(arguments[1]), parsed > 0 else {
      FileHandle.standardError.write(Data("Invalid --maximum-seconds value.\n".utf8))
      exit(2)
    }
    maximumSeconds = parsed
    arguments.removeFirst(2)
  case "--maximum-memory-mib":
    guard arguments.count >= 2, let parsed = UInt64(arguments[1]), parsed > 0,
      parsed <= UInt64.max / (1_024 * 1_024) else {
      FileHandle.standardError.write(Data("Invalid --maximum-memory-mib value.\n".utf8))
      exit(2)
    }
    maximumMemoryBytes = parsed * 1_024 * 1_024
    arguments.removeFirst(2)
  case "--continue-on-error":
    continuesAfterErrors = true
    arguments.removeFirst()
  case "--streaming-rows-only":
    streamsRowsOnly = true
    arguments.removeFirst()
  case "--detailed":
    runsDetailedMetrics = true
    arguments.removeFirst()
  case "--all-xml":
    scansAllXML = true
    arguments.removeFirst()
  case "--expected-errors":
    guard arguments.count >= 2 else {
      FileHandle.standardError.write(Data("Missing --expected-errors path.\n".utf8))
      exit(2)
    }
    expectedFailureFile = arguments[1]
    continuesAfterErrors = true
    arguments.removeFirst(2)
  default:
    FileHandle.standardError.write(Data("Unknown option \(option).\n".utf8))
    exit(2)
  }
}

guard !arguments.isEmpty else {
  FileHandle.standardError.write(
    Data(
      ("Usage: OfficeKitBenchmarks [--maximum-seconds N] [--maximum-memory-mib N] "
        + "[--continue-on-error] "
        + "[--streaming-rows-only] [--detailed] [--all-xml] "
        + "[--expected-errors FILE] FILE_OR_DIRECTORY...\n").utf8))
  exit(2)
}

let supportedExtensions: Set<String> = [
  "pptx", "pptm", "ppsx", "ppsm", "potx", "potm",
  "xlsx", "xlsm", "xltx", "xltm",
  "docx", "docm", "dotx", "dotm",
]

func inputFiles(for paths: [String]) -> [String] {
  paths.flatMap { path -> [String] in
    var isDirectory: ObjCBool = false
    guard FileManager.default.fileExists(atPath: path, isDirectory: &isDirectory),
      isDirectory.boolValue else { return [path] }

    guard let enumerator = FileManager.default.enumerator(atPath: path) else { return [] }
    return enumerator.compactMap { value in
      guard let relativePath = value as? String,
        supportedExtensions.contains(URL(fileURLWithPath: relativePath).pathExtension.lowercased()) else {
        return nil
      }
      return URL(fileURLWithPath: path).appending(path: relativePath).path
    }
  }
  .sorted()
}

let clock = ContinuousClock()
let expectedFailures = try expectedFailureFile.map(expectedFailures(in:)) ?? []
var matchedExpectedFailures: Set<Int> = []
var exceededBudget = false
var exceededMemoryBudget = false
var encounteredError = false

func firstInternalAttachment(in document: OfficeDocument) throws -> OfficeAttachment? {
  var fallback: OfficeAttachment?
  let sources: [OfficeRelationshipSource] =
    [.package]
    + document.package.parts.map { .part($0.name) }
  for source in sources {
    for relationship in try document.package.relationships(from: source) {
      guard let part = document.package.part(referencedBy: relationship),
        part.name != document.mainPart.name else { continue }
      let attachment = document.package.attachment(referencedBy: relationship)
      fallback = fallback ?? attachment
      let contentType = part.contentType.canonicalValue
      if !contentType.hasSuffix("+xml") && !contentType.hasSuffix("/xml") {
        return attachment
      }
    }
  }
  return fallback
}

for path in inputFiles(for: arguments) {
  do {
    let url = URL(fileURLWithPath: path)
    let openStart = clock.now
    let document = try OfficeDocument(contentsOf: url)
    let openSeconds = seconds(openStart.duration(to: clock.now))

    var relationshipCount: Int?
    var relationshipLookupSeconds: Double?
    var firstSlideSeconds: Double?
    var attachmentURLSeconds: Double?
    var attachmentCopySeconds: Double?
    var attachmentBytes: UInt64?
    if runsDetailedMetrics {
      let relationshipStart = clock.now
      var count = try document.package.relationships(from: .package).count
      for part in document.package.parts {
        count += try document.package.relationships(from: .part(part.name)).count
      }
      relationshipLookupSeconds = seconds(relationshipStart.duration(to: clock.now))
      relationshipCount = count

      if document.kind == .presentation {
        let firstSlideStart = clock.now
        let presentation = try OfficePresentation(document: document)
        if !presentation.slides.isEmpty { _ = try presentation.slide(at: 0) }
        firstSlideSeconds = seconds(firstSlideStart.duration(to: clock.now))
      }

      if let attachment = try firstInternalAttachment(in: document) {
        attachmentBytes = attachment.uncompressedSize
        let attachmentURLStart = clock.now
        _ = try attachment.url()
        attachmentURLSeconds = seconds(attachmentURLStart.duration(to: clock.now))

        let destination = FileManager.default.temporaryDirectory
          .appendingPathComponent(ProcessInfo.processInfo.globallyUniqueString)
        defer { try? FileManager.default.removeItem(at: destination) }
        let attachmentCopyStart = clock.now
        try attachment.copy(to: destination)
        attachmentCopySeconds = seconds(attachmentCopyStart.duration(to: clock.now))
      }
    }

    var visitor = BenchmarkVisitor()
    let traversalStart = clock.now
    if streamsRowsOnly, document.kind == .spreadsheet {
      let workbook = try OfficeWorkbook(document: document)
      for index in workbook.worksheets.indices {
        try workbook.streamRows(inWorksheetAt: index) { row in
          visitor.cells += row.cells.count
        }
      }
    } else {
      try document.traverse(using: &visitor)
    }
    let traversalSeconds = seconds(traversalStart.duration(to: clock.now))
    var xmlPartCount = 0
    var xmlEventCount = 0
    var allXMLSeconds: Double?
    if scansAllXML {
      let allXMLStart = clock.now
      try document.package.streamXMLParts { _, event in
        if case .startDocument = event { xmlPartCount += 1 }
        xmlEventCount += 1
      }
      allXMLSeconds = seconds(allXMLStart.duration(to: clock.now))
    }
    let peakMemory = peakResidentMemoryBytes()
    exceededBudget =
      exceededBudget || openSeconds > maximumSeconds
      || traversalSeconds > maximumSeconds
      || (allXMLSeconds ?? 0) > maximumSeconds
    if let maximumMemoryBytes, peakMemory > maximumMemoryBytes {
      exceededMemoryBudget = true
    }

    var record: [String: Any] = [
      "path": path,
      "kind": document.kind.rawValue,
      "parts": document.package.parts.count,
      "declaredUncompressedBytes": document.package.parts.reduce(UInt64(0)) {
        $0 + $1.uncompressedSize
      },
      "peakResidentMemoryBytes": peakMemory,
      "openSeconds": openSeconds,
      "traversalSeconds": traversalSeconds,
      "cells": visitor.cells,
      "runs": visitor.runs,
      "slideElements": visitor.slideElements,
      "attachmentCallbacks": visitor.attachments,
    ]
    if scansAllXML {
      record["xmlParts"] = xmlPartCount
      record["xmlEvents"] = xmlEventCount
      record["allXMLSeconds"] = allXMLSeconds ?? 0
    }
    if runsDetailedMetrics {
      record["relationshipCount"] = relationshipCount ?? 0
      record["relationshipLookupSeconds"] = relationshipLookupSeconds ?? 0
      record["firstSlideSeconds"] = firstSlideSeconds ?? NSNull()
      record["attachmentURLSeconds"] = attachmentURLSeconds ?? NSNull()
      record["attachmentCopySeconds"] = attachmentCopySeconds ?? NSNull()
      record["attachmentBytes"] = attachmentBytes ?? NSNull()
    }
    try writeJSONLine(record, to: .standardOutput)
  } catch {
    let description = String(describing: error)
    let expectedIndex = expectedFailures.indices.first { index in
      path.hasSuffix(expectedFailures[index].pathSuffix)
        && description.contains(expectedFailures[index].errorSubstring)
    }
    if let expectedIndex {
      matchedExpectedFailures.insert(expectedIndex)
    } else {
      encounteredError = true
    }
    let record: [String: Any] = [
      "path": path,
      "error": description,
      "expectedError": expectedIndex != nil,
    ]
    try writeJSONLine(record, to: .standardError)
    guard continuesAfterErrors else { throw error }
  }
}

let missingExpectedFailures = expectedFailures.indices.filter {
  !matchedExpectedFailures.contains($0)
}
for index in missingExpectedFailures {
  try writeJSONLine(
    [
      "missingExpectedError": expectedFailures[index].errorSubstring,
      "pathSuffix": expectedFailures[index].pathSuffix,
    ], to: .standardError)
}
encounteredError = encounteredError || !missingExpectedFailures.isEmpty

if exceededBudget || exceededMemoryBudget || encounteredError {
  let message: String
  if exceededBudget {
    message = "OfficeKit benchmark time budget exceeded.\n"
  } else if exceededMemoryBudget {
    message = "OfficeKit benchmark memory budget exceeded.\n"
  } else {
    message = "OfficeKit corpus scan encountered errors.\n"
  }
  FileHandle.standardError.write(Data(message.utf8))
  exit(1)
}
