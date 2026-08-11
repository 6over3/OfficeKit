# OfficeKit

OfficeKit is a read-only Swift package for Word, Excel, and PowerPoint Open XML files. It requires
Swift 6.2 and macOS 13 or later.

## Usage

```swift
import OfficeKit

let document = try OfficeDocument(contentsOf: fileURL)
switch document.kind {
case .wordProcessing:
  let wordDocument = try OfficeWordDocument(document: document)
  print(wordDocument.body.blocks.count)
case .spreadsheet:
  let workbook = try OfficeWorkbook(document: document)
  try workbook.streamRows(inWorksheetAt: 0) { row in
    print(row.cells)
  }
case .presentation:
  let presentation = try OfficePresentation(document: document)
  let slide = try presentation.slide(at: 0)
  let attachmentURLs = try slide.attachments.map { try $0.url() }
  print(attachmentURLs)
}
```

## Development

```sh
swift test
swift build -c release
swift run -c release OfficeKitBenchmarks --detailed --all-xml file.docx file.xlsx file.pptx
```

OfficeKit uses the GNU Affero General Public License v3.0.
