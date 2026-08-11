// swift-tools-version: 6.2

import PackageDescription

let package = Package(
  name: "OfficeKit",
  platforms: [
    .macOS(.v13)
  ],
  products: [
    .library(name: "OfficeKit", targets: ["OfficeKit"])
  ],
  dependencies: [
    .package(url: "https://github.com/weichsel/ZIPFoundation.git", from: "0.9.20"),
    .package(url: "https://github.com/swiftlang/swift-docc-plugin", from: "1.5.0"),
  ],
  targets: [
    .target(
      name: "OfficeKit",
      dependencies: ["ZIPFoundation"],
      resources: [.copy("OfficeKit.docc")]
    ),
    .testTarget(
      name: "OfficeKitTests",
      dependencies: ["OfficeKit", "ZIPFoundation"],
      resources: [.copy("Fixtures")]
    ),
    .executableTarget(
      name: "OfficeKitBenchmarks",
      dependencies: ["OfficeKit"],
      path: "Benchmarks",
      exclude: [
        "generate-million-row-workbook.sh",
        "upstream-corpus-expected-errors.tsv",
      ]
    ),
  ],
  swiftLanguageModes: [.v6]
)
