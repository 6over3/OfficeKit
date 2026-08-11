import Foundation
import UniformTypeIdentifiers

/// A relationship-backed packaged or external resource.
///
/// Attachment bytes remain lazy. Asking for `url()` streams a packaged resource to a cached
/// temporary file or returns the original external URL without fetching it.
public struct OfficeAttachment: Sendable {
  private let package: OfficePackage

  /// The relationship that references the resource.
  public let relationship: OfficeRelationship

  /// The resolved package part, or `nil` for external and dangling targets.
  public let part: OfficePart?

  package init(
    package: OfficePackage,
    relationship: OfficeRelationship,
    part: OfficePart?
  ) {
    self.package = package
    self.relationship = relationship
    self.part = part
  }

  /// A filename hint derived from the internal part name or external URL.
  public var filenameHint: String? {
    switch relationship.target {
    case .internalPart(let partName, _):
      return partName.rawValue.split(separator: "/").last.map(String.init)
    case .external(let rawURL):
      return URL(string: rawURL)?.lastPathComponent.nonEmpty
    }
  }

  /// The preferred system content type for a packaged attachment, when available.
  public var contentType: UTType? { part?.uniformType }

  /// The exact MIME content type declared by the package.
  ///
  /// This remains available for unknown producer extensions that do not map to a system `UTType`.
  public var declaredContentType: OfficeContentType? { part?.contentType }

  /// The declared uncompressed size for a packaged attachment.
  public var uncompressedSize: UInt64? { part?.uncompressedSize }

  /// Returns a URL that can be handed directly to file- and URL-based APIs.
  ///
  /// Internal resources are streamed to package-owned temporary storage without an intermediate
  /// `Data` allocation. External URLs are returned without network access. A packaged URL remains
  /// valid while this attachment or its originating package remains alive.
  public func url() throws -> URL {
    switch relationship.target {
    case .internalPart:
      guard let part else { throw OfficeKitError.missingPart(relationship.rawTarget) }
      return try package.fileURL(for: part)
    case .external(let rawURL):
      guard let url = URL(string: rawURL) else {
        throw OfficeKitError.invalidPackage("Invalid external attachment URL \(rawURL).")
      }
      return url
    }
  }

  /// Copies a packaged attachment to a permanent destination without loading it into memory.
  ///
  /// External resources are never downloaded and produce an invalid-package error.
  public func copy(to destinationURL: URL) throws {
    guard let part else {
      throw OfficeKitError.invalidPackage("External attachments cannot be copied without fetching.")
    }
    try package.copy(part, to: destinationURL)
  }

  /// Loads packaged attachment bytes under an explicit memory limit.
  ///
  /// Prefer `url()` or `copy(to:)` for large resources. External resources are never fetched.
  public func data(limit: UInt64) throws -> Data {
    guard let part else {
      throw OfficeKitError.invalidPackage("External attachments cannot be loaded without fetching.")
    }
    return try package.data(for: part, limit: limit)
  }
}

extension String {
  fileprivate var nonEmpty: String? { isEmpty ? nil : self }
}
