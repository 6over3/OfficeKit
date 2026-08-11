import Foundation

/// A resource limit enforced by OfficeKit.
public enum OfficeLimit: String, Sendable, Equatable {
  /// The ZIP entry count.
  case entryCount
  /// The uncompressed size of one ZIP entry.
  case entrySize
  /// The total declared uncompressed ZIP size.
  case totalUncompressedSize
  /// The ratio of uncompressed to compressed entry bytes.
  case compressionRatio
  /// The in-memory size of package metadata.
  case metadataPartSize
  /// The size of a caller-requested in-memory part read.
  case dataReadSize
  /// The relationship count owned by one source.
  case relationshipsPerPart
  /// The nesting depth of an XML part.
  case xmlDepth
  /// The number of attributes on one XML element.
  case xmlAttributesPerElement
  /// The number of events emitted from one XML part.
  case xmlEventCount
  /// The cumulative decoded character-data size in one XML part.
  case xmlTextSize
  /// The number of rich strings retained from one shared-string table.
  case sharedStringCount
  /// The number of worksheet rows retained by a non-streaming parse.
  case retainedWorksheetRows
  /// The number of worksheet cells retained by a non-streaming parse.
  case retainedWorksheetCells
}

/// An error produced while opening or reading an Office document.
public enum OfficeKitError: Error, Sendable, Equatable {
  /// The requested file does not exist.
  case fileNotFound(path: String)
  /// The file is an encrypted Office compound document.
  case encryptedDocument(path: String)
  /// The file is an unsupported legacy compound-binary Office document.
  case unsupportedLegacyBinary(path: String)
  /// The ZIP container could not be opened.
  case unreadableArchive(path: String)
  /// The package lacks its required content-type manifest.
  case missingContentTypes
  /// A part or relationship target contains an invalid OPC name.
  case invalidPartName(String)
  /// More than one ZIP entry maps to the same part name.
  case duplicatePartName(String)
  /// A ZIP entry type is not safe for read-only package access.
  case unsupportedZipEntry(path: String)
  /// Input exceeded an explicitly configured resource limit.
  case limitExceeded(limit: OfficeLimit, actual: UInt64, maximum: UInt64)
  /// The content-type manifest does not cover a package part.
  case missingContentType(part: String)
  /// A requested or referenced part is absent.
  case missingPart(String)
  /// OPC package structure is invalid.
  case invalidPackage(String)
  /// An XML part is malformed or violates parser security policy.
  case invalidXML(part: String, message: String)
  /// An internal relationship resolves to a missing part.
  case danglingRelationship(source: String, id: String, target: String)
  /// A streaming copy would overwrite an existing file.
  case destinationAlreadyExists(path: String)
}

extension OfficeKitError: LocalizedError {
  /// A caller-facing description of the failure.
  public var errorDescription: String? {
    switch self {
    case .fileNotFound(let path):
      "No file exists at \(path)."
    case .encryptedDocument(let path):
      "The Office document at \(path) is encrypted."
    case .unsupportedLegacyBinary(let path):
      "The file at \(path) is a legacy compound-binary Office document."
    case .unreadableArchive(let path):
      "The ZIP package at \(path) could not be opened."
    case .missingContentTypes:
      "The package does not contain [Content_Types].xml."
    case .invalidPartName(let name):
      "The package contains an invalid part name: \(name)."
    case .duplicatePartName(let name):
      "The package contains the part name more than once: \(name)."
    case .unsupportedZipEntry(let path):
      "The package contains an unsupported ZIP entry: \(path)."
    case .limitExceeded(let limit, let actual, let maximum):
      "The package exceeded the \(limit.rawValue) limit (\(actual) > \(maximum))."
    case .missingContentType(let part):
      "The package does not declare a content type for \(part)."
    case .missingPart(let part):
      "The package does not contain the referenced part \(part)."
    case .invalidPackage(let reason):
      "The Office package is invalid: \(reason)"
    case .invalidXML(let part, let message):
      "The XML in \(part) is invalid: \(message)"
    case .danglingRelationship(let source, let id, let target):
      "Relationship \(id) from \(source) targets the missing part \(target)."
    case .destinationAlreadyExists(let path):
      "A file already exists at \(path)."
    }
  }
}
