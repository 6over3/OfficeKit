import Foundation
import ZIPFoundation

package struct StoredEntry: Sendable {
  let path: String
  let uncompressedSize: UInt64
  let compressedSize: UInt64
}

/// Synchronizes all access to ZIPFoundation's shared file handle.
package final class ArchiveStorage: @unchecked Sendable {
  private let archive: Archive
  private let archiveURLToRemove: URL?
  private let entriesByPath: [String: Entry]
  private let lock = NSLock()
  private var extractionDirectory: URL?
  private var extractedURLs: [String: URL] = [:]

  package let entries: [StoredEntry]
  package let limits: OfficeParsingLimits

  package init(
    url: URL,
    limits: OfficeParsingLimits,
    removesArchiveOnDeinit: Bool = false
  ) throws {
    let archive: Archive
    do {
      archive = try Archive(url: url, accessMode: .read)
    } catch {
      throw OfficeKitError.unreadableArchive(path: url.path)
    }

    var entriesByPath: [String: Entry] = [:]
    var storedEntries: [StoredEntry] = []
    var entryCount = 0
    var totalUncompressedSize: UInt64 = 0

    for entry in archive {
      entryCount += 1
      guard entryCount <= limits.maximumEntryCount else {
        throw OfficeKitError.limitExceeded(
          limit: .entryCount,
          actual: UInt64(entryCount),
          maximum: UInt64(limits.maximumEntryCount)
        )
      }
      guard entry.type != .symlink else {
        throw OfficeKitError.unsupportedZipEntry(path: entry.path)
      }
      guard entriesByPath[entry.path] == nil else {
        throw OfficeKitError.duplicatePartName(entry.path)
      }
      guard entry.uncompressedSize <= limits.maximumEntrySize else {
        throw OfficeKitError.limitExceeded(
          limit: .entrySize,
          actual: entry.uncompressedSize,
          maximum: limits.maximumEntrySize
        )
      }

      let (newTotal, overflow) = totalUncompressedSize.addingReportingOverflow(
        entry.uncompressedSize)
      guard !overflow, newTotal <= limits.maximumTotalUncompressedSize else {
        throw OfficeKitError.limitExceeded(
          limit: .totalUncompressedSize,
          actual: overflow ? .max : newTotal,
          maximum: limits.maximumTotalUncompressedSize
        )
      }
      totalUncompressedSize = newTotal

      if entry.uncompressedSize > 0 {
        guard entry.compressedSize > 0 else {
          throw OfficeKitError.limitExceeded(
            limit: .compressionRatio,
            actual: .max,
            maximum: UInt64(limits.maximumCompressionRatio)
          )
        }
        let ratio = Double(entry.uncompressedSize) / Double(entry.compressedSize)
        guard ratio <= limits.maximumCompressionRatio else {
          throw OfficeKitError.limitExceeded(
            limit: .compressionRatio,
            actual: UInt64(ratio.rounded(.up)),
            maximum: UInt64(limits.maximumCompressionRatio)
          )
        }
      }

      entriesByPath[entry.path] = entry
      if entry.type == .file {
        storedEntries.append(
          StoredEntry(
            path: entry.path,
            uncompressedSize: entry.uncompressedSize,
            compressedSize: entry.compressedSize
          )
        )
      }
    }

    self.archive = archive
    self.archiveURLToRemove = removesArchiveOnDeinit ? url : nil
    self.entriesByPath = entriesByPath
    self.entries = storedEntries
    self.limits = limits
  }

  deinit {
    if let extractionDirectory {
      try? FileManager.default.removeItem(at: extractionDirectory)
    }
    if let archiveURLToRemove {
      try? FileManager.default.removeItem(at: archiveURLToRemove)
    }
  }

  package func contains(path: String) -> Bool {
    entriesByPath[path] != nil
  }

  package func readData(at path: String, maximumSize: UInt64) throws -> Data {
    try lock.withLock {
      guard let entry = entriesByPath[path], entry.type == .file else {
        throw OfficeKitError.missingPart(path)
      }
      guard entry.uncompressedSize <= maximumSize,
        entry.uncompressedSize <= UInt64(Int.max) else {
        throw OfficeKitError.limitExceeded(
          limit: .dataReadSize,
          actual: entry.uncompressedSize,
          maximum: maximumSize
        )
      }

      var data = Data()
      data.reserveCapacity(Int(entry.uncompressedSize))
      do {
        _ = try archive.extract(entry) { chunk in
          guard UInt64(data.count) + UInt64(chunk.count) <= maximumSize else {
            throw OfficeKitError.limitExceeded(
              limit: .dataReadSize,
              actual: UInt64(data.count) + UInt64(chunk.count),
              maximum: maximumSize
            )
          }
          data.append(chunk)
        }
      } catch let error as OfficeKitError {
        throw error
      } catch {
        throw OfficeKitError.invalidPackage("Could not read ZIP entry \(path).")
      }
      return data
    }
  }

  package func copyItem(at path: String, to destinationURL: URL) throws {
    try lock.withLock {
      guard !FileManager.default.fileExists(atPath: destinationURL.path) else {
        throw OfficeKitError.destinationAlreadyExists(path: destinationURL.path)
      }
      guard let entry = entriesByPath[path], entry.type == .file else {
        throw OfficeKitError.missingPart(path)
      }
      do {
        _ = try archive.extract(entry, to: destinationURL)
      } catch let error as OfficeKitError {
        throw error
      } catch {
        throw OfficeKitError.invalidPackage("Could not copy ZIP entry \(path).")
      }
    }
  }

  package func fileURL(at path: String) throws -> URL {
    try lock.withLock {
      if let existingURL = extractedURLs[path] { return existingURL }
      guard let entry = entriesByPath[path], entry.type == .file else {
        throw OfficeKitError.missingPart(path)
      }

      let directory = try extractionDirectoryLocked()
      var destinationURL = directory
      for component in path.split(separator: "/") {
        destinationURL.appendPathComponent(String(component), isDirectory: false)
      }
      let parentURL = destinationURL.deletingLastPathComponent()
      do {
        try FileManager.default.createDirectory(
          at: parentURL,
          withIntermediateDirectories: true,
          attributes: [.posixPermissions: 0o700]
        )
        _ = try archive.extract(entry, to: destinationURL)
      } catch {
        try? FileManager.default.removeItem(at: destinationURL)
        throw OfficeKitError.invalidPackage("Could not expose ZIP entry \(path) as a file URL.")
      }
      extractedURLs[path] = destinationURL
      return destinationURL
    }
  }

  private func extractionDirectoryLocked() throws -> URL {
    if let extractionDirectory { return extractionDirectory }
    let directory = FileManager.default.temporaryDirectory
      .appendingPathComponent("OfficeKit-\(UUID().uuidString)", isDirectory: true)
    do {
      try FileManager.default.createDirectory(
        at: directory,
        withIntermediateDirectories: false,
        attributes: [.posixPermissions: 0o700]
      )
    } catch {
      throw OfficeKitError.invalidPackage("Could not create temporary attachment storage.")
    }
    extractionDirectory = directory
    return directory
  }
}
