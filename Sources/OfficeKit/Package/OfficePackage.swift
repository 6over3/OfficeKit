import Foundation

// SAFETY: every access to mutable cache state is serialized by `lock`; loaded relationship
// arrays and errors are immutable Sendable values.
private final class RelationshipCache: @unchecked Sendable {
  private let lock = NSLock()
  private var values: [OfficeRelationshipSource: Result<[OfficeRelationship], OfficeKitError>] = [:]

  func relationships(
    for source: OfficeRelationshipSource,
    load: () throws -> [OfficeRelationship]
  ) throws -> [OfficeRelationship] {
    try lock.withLock {
      if let cached = values[source] { return try cached.get() }
      do {
        let relationships = try load()
        values[source] = .success(relationships)
        return relationships
      } catch let error as OfficeKitError {
        values[source] = .failure(error)
        throw error
      } catch {
        let wrapped = OfficeKitError.invalidPackage("Unexpected relationship loading failure.")
        values[source] = .failure(wrapped)
        throw wrapped
      }
    }
  }
}

/// A read-only Open Packaging Conventions package.
///
/// `OfficePackage` indexes ZIP metadata at initialization but loads part bytes and relationship
/// XML lazily. Its immutable public values are safe to share across concurrency domains.
public struct OfficePackage: Sendable {
  private let storage: ArchiveStorage
  private let partsByName: [OfficePartName: OfficePart]
  private let relationshipCache = RelationshipCache()

  /// All file-like package parts except the special `[Content_Types].xml` stream.
  ///
  /// Relationship parts are included because they are addressable package content. The result is
  /// sorted by canonical part name.
  public let parts: [OfficePart]

  /// Resource limits governing this package and its semantic parsers.
  public let parsingLimits: OfficeParsingLimits

  /// Opens and validates the OPC container at `url` without inflating document XML or media.
  public init(contentsOf url: URL, limits: OfficeParsingLimits = .standard) throws {
    guard FileManager.default.fileExists(atPath: url.path) else {
      throw OfficeKitError.fileNotFound(path: url.path)
    }
    let storage: ArchiveStorage
    do {
      storage = try ArchiveStorage(url: url, limits: limits)
    } catch let archiveError as OfficeKitError {
      guard case .unreadableArchive = archiveError,
        let convertedURL = try FlatOPCConverter.convert(contentsOf: url, limits: limits) else {
        throw archiveError
      }
      do {
        storage = try ArchiveStorage(
          url: convertedURL,
          limits: limits,
          removesArchiveOnDeinit: true
        )
      } catch {
        try? FileManager.default.removeItem(at: convertedURL)
        throw error
      }
    }
    guard storage.contains(path: "[Content_Types].xml") else {
      throw OfficeKitError.missingContentTypes
    }
    let contentTypesData = try storage.readData(
      at: "[Content_Types].xml", maximumSize: limits.maximumMetadataPartSize
    )
    let contentTypes = try PackageXMLParser.contentTypes(from: contentTypesData)

    var partsByName: [OfficePartName: OfficePart] = [:]
    for entry in storage.entries where entry.path != "[Content_Types].xml" {
      let name = try OfficePartName(archivePath: entry.path)
      guard partsByName[name] == nil else {
        throw OfficeKitError.duplicatePartName(name.rawValue)
      }
      guard let contentType = contentTypes.contentType(for: name) else {
        throw OfficeKitError.missingContentType(part: name.rawValue)
      }
      partsByName[name] = OfficePart(
        name: name,
        contentType: contentType,
        uncompressedSize: entry.uncompressedSize
      )
    }

    self.storage = storage
    self.partsByName = partsByName
    self.parts = partsByName.values.sorted { $0.name < $1.name }
    self.parsingLimits = limits
  }

  /// Returns the part with the specified canonical name, or `nil` when it is absent.
  public func part(named name: OfficePartName) -> OfficePart? {
    partsByName[name]
  }

  /// Returns all parts whose canonical MIME type matches `contentType`.
  public func parts(withContentType contentType: OfficeContentType) -> [OfficePart] {
    let canonicalValue = contentType.canonicalValue
    return parts.filter { $0.contentType.canonicalValue == canonicalValue }
  }

  /// Loads a part into memory, enforcing either the supplied limit or the package default.
  ///
  /// Use `copy(_:to:)` for attachments that do not need to be materialized as `Data`.
  public func data(for part: OfficePart, limit: UInt64? = nil) throws -> Data {
    guard partsByName[part.name] == part else {
      throw OfficeKitError.missingPart(part.name.rawValue)
    }
    return try storage.readData(
      at: part.name.archivePath,
      maximumSize: limit ?? storage.limits.defaultDataReadLimit
    )
  }

  /// Copies a part to a new file without first collecting its bytes into one `Data` value.
  ///
  /// This method never overwrites an existing destination.
  public func copy(_ part: OfficePart, to destinationURL: URL) throws {
    guard partsByName[part.name] == part else {
      throw OfficeKitError.missingPart(part.name.rawValue)
    }
    try storage.copyItem(at: part.name.archivePath, to: destinationURL)
  }

  /// Returns a local file URL for a package part without loading its bytes into memory.
  ///
  /// The first call streams the ZIP entry into package-owned temporary storage. Later calls for
  /// the same part return the cached URL. The file remains available while this package or an
  /// `OfficeAttachment` created from it remains alive. Call `copy(_:to:)` when the file must
  /// outlive those values.
  public func fileURL(for part: OfficePart) throws -> URL {
    guard partsByName[part.name] == part else {
      throw OfficeKitError.missingPart(part.name.rawValue)
    }
    return try storage.fileURL(at: part.name.archivePath)
  }

  /// Returns relationships owned by the package or part, parsing them on first access.
  public func relationships(from source: OfficeRelationshipSource) throws -> [OfficeRelationship] {
    try relationshipCache.relationships(for: source) {
      let path = source.relationshipArchivePath
      guard storage.contains(path: path) else { return [] }
      let data = try storage.readData(
        at: path, maximumSize: storage.limits.maximumMetadataPartSize
      )
      return try PackageXMLParser.relationships(
        from: data,
        source: source,
        maximumCount: storage.limits.maximumRelationshipsPerPart
      )
    }
  }

  /// Returns the relationship with `id` owned by `source`, or `nil` when it is absent.
  ///
  /// Semantic readers use this method to resolve `r:id`, `r:embed`, and `r:link` attributes
  /// without interpreting ZIP paths themselves.
  public func relationship(
    identifiedBy id: OfficeRelationshipID,
    from source: OfficeRelationshipSource
  ) throws -> OfficeRelationship? {
    try relationships(from: source).first { $0.id == id }
  }

  /// Returns relationships of `type` owned by `source`, preserving relationship order.
  public func relationships(
    from source: OfficeRelationshipSource,
    ofType type: OfficeRelationshipType
  ) throws -> [OfficeRelationship] {
    try relationships(from: source).filter { $0.type.isEquivalent(to: type) }
  }

  /// Resolves internal relationships of `type` to existing parts.
  ///
  /// External and dangling targets are omitted. The relationships themselves remain available
  /// through `relationships(from:ofType:)` when callers need to distinguish those cases.
  public func relatedParts(
    from source: OfficeRelationshipSource,
    ofType type: OfficeRelationshipType
  ) throws -> [OfficePart] {
    try relationships(from: source, ofType: type).compactMap(part(referencedBy:))
  }

  /// Streams namespace-aware XML events from a package part.
  ///
  /// XML parts up to the package's in-memory threshold are parsed from `Data`; larger parts are
  /// streamed through a package-owned temporary file. Emitted events are never retained. Binary
  /// parts fail with an invalid-package error rather than being passed to the XML parser.
  public func parseXML(
    in part: OfficePart,
    maximumPartSize: UInt64? = nil,
    limits: OfficeXMLParsingLimits? = nil,
    compatibility: OfficeXMLCompatibilityOptions? = nil,
    _ body: @escaping (OfficeXMLEvent) throws -> Void
  ) throws {
    guard part.isXML else {
      throw OfficeKitError.invalidPackage("Part \(part.name.rawValue) is not XML content.")
    }
    let permittedSize = maximumPartSize ?? storage.limits.maximumEntrySize
    let effectiveXMLLimits = limits ?? parsingLimits.xmlParsingLimits
    guard part.uncompressedSize <= permittedSize else {
      throw OfficeKitError.limitExceeded(
        limit: .dataReadSize,
        actual: part.uncompressedSize,
        maximum: permittedSize
      )
    }
    if part.uncompressedSize <= storage.limits.maximumInMemoryXMLPartSize {
      let data = try self.data(for: part, limit: permittedSize)
      try parse(
        OfficeXMLReader(data: data, source: part.name.rawValue, limits: effectiveXMLLimits),
        compatibility: compatibility,
        body
      )
    } else {
      let fileURL = try self.fileURL(for: part)
      try parse(
        OfficeXMLReader(
          contentsOf: fileURL,
          source: part.name.rawValue,
          limits: effectiveXMLLimits
        ),
        compatibility: compatibility,
        body
      )
    }
  }

  /// Streams every XML part in canonical package order.
  ///
  /// Parts are parsed one at a time and emitted events are not retained, so memory remains bounded
  /// by the current XML parser buffers and caller-owned state. Throwing from `body` stops the scan
  /// immediately and propagates the same error.
  public func streamXMLParts(
    compatibility: OfficeXMLCompatibilityOptions? = nil,
    _ body: @escaping (OfficePart, OfficeXMLEvent) throws -> Void
  ) throws {
    for part in parts where part.isXML {
      try parseXML(in: part, compatibility: compatibility) { event in
        try body(part, event)
      }
    }
  }

  private func parse(
    _ reader: OfficeXMLReader,
    compatibility: OfficeXMLCompatibilityOptions?,
    _ body: @escaping (OfficeXMLEvent) throws -> Void
  ) throws {
    if let compatibility {
      try reader.parseCompatible(using: compatibility, body)
    } else {
      try reader.parse(body)
    }
  }

  /// Returns the internal part targeted by `relationship`, or `nil` for external or missing targets.
  public func part(referencedBy relationship: OfficeRelationship) -> OfficePart? {
    guard case .internalPart(let name, _) = relationship.target else { return nil }
    return partsByName[name]
  }

  /// Wraps a relationship as a lazy attachment reference.
  ///
  /// The returned value retains this package, allowing packaged attachment URLs to remain valid
  /// for the attachment's lifetime.
  public func attachment(referencedBy relationship: OfficeRelationship) -> OfficeAttachment {
    OfficeAttachment(
      package: self,
      relationship: relationship,
      part: part(referencedBy: relationship)
    )
  }

  /// Validates that internal relationship targets exist.
  ///
  /// Relationship types in `ignoredMissingTargetTypes` remain visible but do not fail validation.
  /// This is useful for explicitly tolerated producer defects such as a stale calculation-chain
  /// relationship.
  public func validateRelationshipTargets(
    ignoringMissingTargetTypes: Set<OfficeRelationshipType> = []
  ) throws {
    var sources: [OfficeRelationshipSource] = [.package]
    sources.append(contentsOf: parts.map { .part($0.name) })
    for source in sources {
      for relationship in try relationships(from: source) {
        guard case .internalPart(let target, _) = relationship.target,
          partsByName[target] == nil,
          !ignoringMissingTargetTypes.contains(where: relationship.type.isEquivalent(to:)) else {
          continue
        }
        throw OfficeKitError.danglingRelationship(
          source: source.description,
          id: relationship.id.rawValue,
          target: target.rawValue
        )
      }
    }
  }

  /// Validates internal relationship targets under an explicit strict or recovering policy.
  ///
  /// Strict validation throws the first dangling relationship. Recovering validation returns a
  /// stable diagnostic for every dangling relationship and never invents or reconnects a target.
  @discardableResult
  public func validateRelationshipTargets(
    policy: OfficeValidationPolicy
  ) throws -> [OfficeDiagnostic] {
    let diagnostics = try relationshipTargetDiagnostics()
    guard policy == .recovering else {
      if let diagnostic = diagnostics.first {
        throw OfficeKitError.danglingRelationship(
          source: diagnostic.source.part,
          id: diagnostic.source.relationshipID?.rawValue ?? "",
          target: diagnostic.target ?? ""
        )
      }
      return []
    }
    return diagnostics
  }

  /// Returns all dangling internal relationships in deterministic source and relationship order.
  public func relationshipTargetDiagnostics() throws -> [OfficeDiagnostic] {
    var sources: [OfficeRelationshipSource] = [.package]
    sources.append(contentsOf: parts.map { .part($0.name) })
    var diagnostics: [OfficeDiagnostic] = []
    for source in sources {
      for relationship in try relationships(from: source) {
        guard case .internalPart(let target, _) = relationship.target,
          partsByName[target] == nil else { continue }
        diagnostics.append(
          OfficeDiagnostic(
            severity: .warning,
            code: .danglingRelationship,
            message:
              "Relationship \(relationship.id.rawValue) targets missing part \(target.rawValue).",
            source: OfficeSourceReference(
              part: source.description,
              relationshipID: relationship.id,
              qualifiedName: nil
            ),
            target: target.rawValue
          ))
      }
    }
    return diagnostics
  }
}
