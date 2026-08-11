import Foundation

package enum SpreadsheetVMLCommentParser {
  private static let vmlNamespace = "urn:schemas-microsoft-com:vml"
  private static let excelNamespace = "urn:schemas-microsoft-com:office:excel"

  private enum TextTarget {
    case anchor
    case row
    case column
  }

  private struct Builder {
    let startDepth: Int
    let identifier: String?
    let style: [String: String]
    var isNote = false
    var isVisible = false
    var movesWithCells = false
    var sizesWithCells = false
    var anchorText = ""
    var rowText = ""
    var columnText = ""
  }

  package static func parse(
    part: OfficePart,
    package: OfficePackage
  ) throws -> [OfficeCellReference: OfficeWorksheetCommentShape] {
    var depth = 0
    var builder: Builder?
    var textTarget: TextTarget?
    var textDepth: Int?
    var shapes: [OfficeCellReference: OfficeWorksheetCommentShape] = [:]

    try package.parseXML(in: part) { event in
      switch event {
      case .startElement(let name, let attributes, _, _):
        depth += 1
        if name.namespaceURI == vmlNamespace, name.localName == "shape" {
          builder = Builder(
            startDepth: depth,
            identifier: vmlAttribute("id", in: attributes),
            style: parseVMLStyle(vmlAttribute("style", in: attributes) ?? "")
          )
          return
        }
        guard builder != nil, name.namespaceURI == excelNamespace else { return }
        switch name.localName {
        case "ClientData":
          builder?.isNote = vmlAttribute("ObjectType", in: attributes) == "Note"
        case "MoveWithCells":
          builder?.movesWithCells = true
        case "SizeWithCells":
          builder?.sizesWithCells = true
        case "Visible":
          builder?.isVisible = true
        case "Anchor":
          textTarget = .anchor
          textDepth = depth
        case "Row":
          textTarget = .row
          textDepth = depth
        case "Column":
          textTarget = .column
          textDepth = depth
        default:
          break
        }
      case .text(let text, _):
        switch textTarget {
        case .anchor: builder?.anchorText.append(text)
        case .row: builder?.rowText.append(text)
        case .column: builder?.columnText.append(text)
        case nil: break
        }
      case .endElement(let name, _):
        if textDepth == depth, name.namespaceURI == excelNamespace {
          textTarget = nil
          textDepth = nil
        }
        if builder?.startDepth == depth, name.namespaceURI == vmlNamespace,
          name.localName == "shape",
          let completed = builder
        {
          if completed.isNote,
            let rowIndex = Int(completed.rowText.trimmingCharacters(in: .whitespacesAndNewlines)),
            let columnIndex = Int(
              completed.columnText.trimmingCharacters(in: .whitespacesAndNewlines)
            ),
            let reference = OfficeCellReference(
              rawValue: spreadsheetA1Reference(columnIndex: columnIndex, rowIndex: rowIndex)
            )
          {
            shapes[reference] = makeShape(from: completed, part: part)
          }
          builder = nil
        }
        depth -= 1
      case .startDocument, .endDocument:
        break
      }
    }
    return shapes
  }

  private static func makeShape(
    from builder: Builder,
    part: OfficePart
  ) -> OfficeWorksheetCommentShape {
    let left = pointValue(builder.style["margin-left"])
    let top = pointValue(builder.style["margin-top"])
    let width = pointValue(builder.style["width"])
    let height = pointValue(builder.style["height"])
    let frame: OfficeRect?
    if let left, let top, let width, let height {
      frame = OfficeRect(x: left, y: top, width: width, height: height)
    } else {
      frame = nil
    }
    let visibility = builder.style["visibility"]?.lowercased()
    return OfficeWorksheetCommentShape(
      identifier: builder.identifier,
      anchor: commentAnchor(builder.anchorText),
      isVisible: builder.isVisible || visibility == "visible",
      movesWithCells: builder.movesWithCells,
      sizesWithCells: builder.sizesWithCells,
      zIndex: builder.style["z-index"].flatMap(Int.init),
      spatialInfo: OfficeSpatialInfo(
        coordinateSpace: .worksheet,
        geometrySourcePart: part.name,
        frame: frame,
        zIndex: builder.style["z-index"].flatMap(Int.init),
        resolution: frame == nil
          ? .unresolved(reason: "VML comment shape has incomplete point geometry.")
          : .exact
      ),
      sourcePart: part
    )
  }

  private static func commentAnchor(_ text: String) -> OfficeWorksheetCommentAnchor? {
    let values = text.split(separator: ",", omittingEmptySubsequences: false).compactMap {
      Int($0.trimmingCharacters(in: .whitespacesAndNewlines))
    }
    guard values.count == 8 else { return nil }
    return OfficeWorksheetCommentAnchor(
      from: OfficeWorksheetCommentMarker(
        columnIndex: values[0],
        columnOffset: values[1],
        rowIndex: values[2],
        rowOffset: values[3]
      ),
      to: OfficeWorksheetCommentMarker(
        columnIndex: values[4],
        columnOffset: values[5],
        rowIndex: values[6],
        rowOffset: values[7]
      )
    )
  }

  private static func pointValue(_ text: String?) -> Double? {
    guard let text else { return nil }
    let trimmed = text.trimmingCharacters(in: .whitespacesAndNewlines)
    guard trimmed.lowercased().hasSuffix("pt") else { return nil }
    return Double(trimmed.dropLast(2))
  }

}

private func parseVMLStyle(_ text: String) -> [String: String] {
  var values: [String: String] = [:]
  for declaration in text.split(separator: ";") {
    let components = declaration.split(separator: ":", maxSplits: 1)
    guard components.count == 2 else { continue }
    let name = components[0].trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
    let value = components[1].trimmingCharacters(in: .whitespacesAndNewlines)
    values[name] = value
  }
  return values
}

private func vmlAttribute(
  _ localName: String,
  in attributes: [OfficeXMLAttribute]
) -> String? {
  attributes.first { $0.name.localName == localName }?.value
}
