import Foundation

enum FixtureCatalog {
  static func url(for relativePath: String) throws -> URL {
    guard let resourceURL = Bundle.module.resourceURL else {
      throw FixtureError.missing(relativePath)
    }
    let fixturesURL = resourceURL.appendingPathComponent("Fixtures", isDirectory: true)
    let url = fixturesURL.appendingPathComponent(relativePath, isDirectory: false)
    let fixturesPath = fixturesURL.standardizedFileURL.path + "/"
    guard url.standardizedFileURL.path.hasPrefix(fixturesPath),
      FileManager.default.fileExists(atPath: url.path) else {
      throw FixtureError.missing(relativePath)
    }
    return url
  }
}

enum FixtureError: Error, Equatable {
  case missing(String)
}
