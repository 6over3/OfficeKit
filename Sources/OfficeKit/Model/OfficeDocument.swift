import Foundation

/// The Office application family represented by a document's main part.
public enum OfficeDocumentKind: String, Sendable, Hashable, Codable {
  /// A PowerPoint PresentationML document.
  case presentation
  /// An Excel SpreadsheetML document.
  case spreadsheet
  /// A Word WordprocessingML document.
  case wordProcessing
}

/// The namespace/relationship conformance family used by an Office document.
public enum OfficeFormatConformance: String, Sendable, Hashable, Codable {
  /// The ECMA-376 Transitional namespace and relationship family.
  case transitional
  /// The ISO/IEC 29500 Strict namespace and relationship family.
  case strict
}

/// A read-only Office Open XML document and its underlying OPC package.
public struct OfficeDocument: Sendable {
  /// The document's PowerPoint, Excel, or Word family.
  public let kind: OfficeDocumentKind

  /// Whether the main package relationship uses Strict or Transitional OOXML.
  public let conformance: OfficeFormatConformance

  /// The part containing the presentation, workbook, or word-processing document root.
  public let mainPart: OfficePart

  /// The complete read-only OPC package.
  public let package: OfficePackage

  /// Core, extended, and custom package metadata.
  public let metadata: OfficeMetadata

  /// The package thumbnail, when declared, with lazy URL access.
  public let thumbnail: OfficeAttachment?

  /// Opens an Office Open XML document and discovers its main part by relationship.
  public init(contentsOf url: URL, limits: OfficeParsingLimits = .standard) throws {
    guard FileManager.default.fileExists(atPath: url.path) else {
      throw OfficeKitError.fileNotFound(path: url.path)
    }
    if Self.isCompoundBinaryFile(at: url) {
      if try Self.isEncryptedOfficeFile(at: url) {
        throw OfficeKitError.encryptedDocument(path: url.path)
      }
      throw OfficeKitError.unsupportedLegacyBinary(path: url.path)
    }

    let package = try OfficePackage(contentsOf: url, limits: limits)
    let rootRelationships = try package.relationships(from: .package)
    guard
      let mainRelationship = rootRelationships.first(where: {
        $0.type.isEquivalent(to: .officeDocument)
      }) else {
      throw OfficeKitError.invalidPackage("Missing package relationship to the main Office part.")
    }
    guard let mainPart = package.part(referencedBy: mainRelationship) else {
      throw OfficeKitError.missingPart(mainRelationship.rawTarget)
    }

    guard let kind = Self.documentKind(for: mainPart.contentType) else {
      throw OfficeKitError.invalidPackage(
        "Unsupported main-part content type \(mainPart.contentType.rawValue)."
      )
    }

    self.kind = kind
    self.conformance = mainRelationship.type == .strictOfficeDocument ? .strict : .transitional
    self.mainPart = mainPart
    self.package = package
    self.metadata = try OfficeMetadataParser.parse(
      package: package,
      rootRelationships: rootRelationships
    )
    self.thumbnail = rootRelationships.first(where: {
      $0.type.isEquivalent(to: .thumbnail)
    }).map(package.attachment(referencedBy:))
  }

  /// Reports whether a file is an encrypted Office compound document.
  ///
  /// This method performs no decryption and returns `false` for ZIP-backed OOXML and unencrypted
  /// legacy compound-binary files.
  public static func isEncryptedOfficeFile(at url: URL) throws -> Bool {
    guard FileManager.default.fileExists(atPath: url.path) else {
      throw OfficeKitError.fileNotFound(path: url.path)
    }
    let data: Data
    do {
      data = try Data(contentsOf: url, options: [.mappedIfSafe])
    } catch {
      throw OfficeKitError.unreadableArchive(path: url.path)
    }
    guard data.starts(with: compoundBinarySignature) else { return false }
    return data.range(of: utf16LittleEndian("EncryptedPackage")) != nil
      && data.range(of: utf16LittleEndian("EncryptionInfo")) != nil
  }

  private static let compoundBinarySignature = Data([
    0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1,
  ])

  private static func isCompoundBinaryFile(at url: URL) -> Bool {
    guard let handle = try? FileHandle(forReadingFrom: url) else { return false }
    defer { try? handle.close() }
    return (try? handle.read(upToCount: compoundBinarySignature.count))?
      .starts(with: compoundBinarySignature) == true
  }

  private static func utf16LittleEndian(_ value: String) -> Data {
    var data = Data()
    data.reserveCapacity(value.utf16.count * 2)
    for codeUnit in value.utf16 {
      data.append(UInt8(truncatingIfNeeded: codeUnit))
      data.append(UInt8(truncatingIfNeeded: codeUnit >> 8))
    }
    return data
  }

  private static func documentKind(for contentType: OfficeContentType) -> OfficeDocumentKind? {
    let value = contentType.rawValue.lowercased()
    if value.contains("presentationml") || value.hasPrefix("application/vnd.ms-powerpoint.") {
      return .presentation
    }
    if value.contains("spreadsheetml") || value.hasPrefix("application/vnd.ms-excel.") {
      return .spreadsheet
    }
    if value.contains("wordprocessingml") || value.hasPrefix("application/vnd.ms-word.") {
      return .wordProcessing
    }
    return nil
  }
}
