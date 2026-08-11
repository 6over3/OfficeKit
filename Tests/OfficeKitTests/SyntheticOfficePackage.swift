import Foundation
import ZIPFoundation

func makeSyntheticOfficePackage(
  entries: [String: String],
  pathExtension: String
) throws -> URL {
  let archive = try Archive(accessMode: .create)
  for (path, source) in entries {
    let data = Data(source.utf8)
    try archive.addEntry(
      with: path,
      type: .file,
      uncompressedSize: Int64(data.count)
    ) { position, size in
      data.subdata(in: Int(position)..<Int(position) + size)
    }
  }
  guard let data = archive.data else {
    throw FixtureError.missing("Could not create synthetic Office package.")
  }
  let url = FileManager.default.temporaryDirectory
    .appendingPathComponent(ProcessInfo.processInfo.globallyUniqueString)
    .appendingPathExtension(pathExtension)
  try data.write(to: url, options: .atomic)
  return url
}
