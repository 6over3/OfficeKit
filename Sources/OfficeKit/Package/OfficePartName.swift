import Foundation

/// The canonical, package-absolute URI name of an OPC part.
public struct OfficePartName: Sendable, Hashable, Codable, Comparable, CustomStringConvertible {
  /// The canonical name, beginning with `/`.
  public let rawValue: String

  /// Validates and creates a package-absolute part name.
  ///
  /// Part names use URI path syntax rather than filesystem path syntax.
  public init(rawValue: String) throws {
    guard Self.isValid(rawValue) else {
      throw OfficeKitError.invalidPartName(rawValue)
    }
    self.rawValue = rawValue
  }

  /// The canonical package-absolute part name.
  public var description: String { rawValue }

  /// Orders part names by their canonical URI strings.
  public static func < (lhs: Self, rhs: Self) -> Bool {
    lhs.rawValue < rhs.rawValue
  }

  package var archivePath: String {
    String(rawValue.dropFirst())
  }

  package init(archivePath: String) throws {
    guard !archivePath.hasPrefix("/") else {
      throw OfficeKitError.invalidPartName(archivePath)
    }
    try self.init(rawValue: "/" + archivePath)
  }

  private static func isValid(_ value: String) -> Bool {
    guard value.first == "/", value.count > 1, !value.hasSuffix("/") else { return false }
    guard !value.contains("\\"), !value.contains("\0"), !value.contains("?"),
      !value.contains("#"), !value.contains("//") else { return false }

    let lowercased = value.lowercased()
    guard !lowercased.contains("%2f"), !lowercased.contains("%5c"),
      !lowercased.contains("%00") else { return false }

    var index = lowercased.startIndex
    while index < lowercased.endIndex {
      if lowercased[index] == "%" {
        guard let first = lowercased.index(index, offsetBy: 1, limitedBy: lowercased.endIndex),
          let second = lowercased.index(index, offsetBy: 2, limitedBy: lowercased.endIndex),
          first < lowercased.endIndex, second < lowercased.endIndex,
          lowercased[first].isHexDigit, lowercased[second].isHexDigit else { return false }
        index = lowercased.index(after: second)
      } else {
        index = lowercased.index(after: index)
      }
    }

    return value.split(separator: "/", omittingEmptySubsequences: false).dropFirst().allSatisfy {
      segment in
      !segment.isEmpty && segment != "." && segment != ".."
        && segment.lowercased() != "%2e" && segment.lowercased() != "%2e%2e"
    }
  }
}
