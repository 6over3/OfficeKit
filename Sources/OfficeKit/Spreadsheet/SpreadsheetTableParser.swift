package enum SpreadsheetTableParser {
  private static let spreadsheetNamespace =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  private enum FormulaTarget {
    case calculated
    case totalsRow
  }

  private struct RawColumn {
    let identifier: UInt32
    let name: String
    let uniqueName: String?
    let totalsRowLabel: String?
    let totalsRowFunction: String?
    var calculatedColumnFormula: String?
    var totalsRowFormula: String?
  }

  package static func parse(part: OfficePart, package: OfficePackage) throws
    -> OfficeSpreadsheetTable
  {
    var identifier: UInt32?
    var name: String?
    var displayName: String?
    var range: OfficeCellRange?
    var headerRowCount: UInt32 = 1
    var totalsRowCount: UInt32 = 0
    var columns: [RawColumn] = []
    var currentColumn: RawColumn?
    var style: OfficeSpreadsheetTableStyle?
    var formulaTarget: FormulaTarget?
    var formulaDepth: Int?
    var depth = 0

    try package.parseXML(in: part, compatibility: .commonOffice) { event in
      switch event {
      case .startElement(let elementName, let attributes, _, _):
        depth += 1
        guard elementName.namespaceURI == spreadsheetNamespace else { return }
        switch elementName.localName {
        case "table":
          identifier = tableAttribute("id", in: attributes).flatMap(UInt32.init)
          name = tableAttribute("name", in: attributes)
          displayName = tableAttribute("displayName", in: attributes)
          range = tableAttribute("ref", in: attributes).flatMap(OfficeCellRange.init(rawValue:))
          headerRowCount =
            tableAttribute("headerRowCount", in: attributes).flatMap(UInt32.init) ?? 1
          if let count = tableAttribute("totalsRowCount", in: attributes).flatMap(UInt32.init) {
            totalsRowCount = count
          } else {
            let isShown =
              tableAttribute("totalsRowShown", in: attributes)
              .flatMap(OfficeValueDecoder.boolean) ?? true
            totalsRowCount = isShown ? 1 : 0
          }
        case "tableColumn":
          guard let columnIdentifier = tableAttribute("id", in: attributes).flatMap(UInt32.init),
            let columnName = tableAttribute("name", in: attributes) else {
            throw OfficeKitError.invalidPackage("Table contains an incomplete column definition.")
          }
          currentColumn = RawColumn(
            identifier: columnIdentifier,
            name: columnName,
            uniqueName: tableAttribute("uniqueName", in: attributes),
            totalsRowLabel: tableAttribute("totalsRowLabel", in: attributes),
            totalsRowFunction: tableAttribute("totalsRowFunction", in: attributes)
          )
        case "calculatedColumnFormula":
          currentColumn?.calculatedColumnFormula = ""
          formulaTarget = .calculated
          formulaDepth = depth
        case "totalsRowFormula":
          currentColumn?.totalsRowFormula = ""
          formulaTarget = .totalsRow
          formulaDepth = depth
        case "tableStyleInfo":
          style = OfficeSpreadsheetTableStyle(
            name: tableAttribute("name", in: attributes),
            showsFirstColumn: tableAttribute("showFirstColumn", in: attributes)
              .flatMap(OfficeValueDecoder.boolean) ?? false,
            showsLastColumn: tableAttribute("showLastColumn", in: attributes)
              .flatMap(OfficeValueDecoder.boolean) ?? false,
            showsRowStripes: tableAttribute("showRowStripes", in: attributes)
              .flatMap(OfficeValueDecoder.boolean) ?? false,
            showsColumnStripes: tableAttribute("showColumnStripes", in: attributes)
              .flatMap(OfficeValueDecoder.boolean) ?? false
          )
        default:
          break
        }
      case .text(let text, _):
        switch formulaTarget {
        case .calculated: currentColumn?.calculatedColumnFormula?.append(text)
        case .totalsRow: currentColumn?.totalsRowFormula?.append(text)
        case nil: break
        }
      case .endElement(let elementName, _):
        if formulaDepth == depth, elementName.namespaceURI == spreadsheetNamespace {
          formulaDepth = nil
          formulaTarget = nil
        }
        if elementName.namespaceURI == spreadsheetNamespace,
          elementName.localName == "tableColumn",
          let column = currentColumn
        {
          columns.append(column)
          currentColumn = nil
        }
        depth -= 1
      case .startDocument, .endDocument:
        break
      }
    }

    guard let identifier, let name, let displayName, let range else {
      throw OfficeKitError.invalidPackage("Table part \(part.name.rawValue) is incomplete.")
    }
    return OfficeSpreadsheetTable(
      sourcePart: part,
      identifier: identifier,
      name: name,
      displayName: displayName,
      range: range,
      headerRowCount: headerRowCount,
      totalsRowCount: totalsRowCount,
      columns: columns.map {
        OfficeSpreadsheetTableColumn(
          identifier: $0.identifier,
          name: $0.name,
          uniqueName: $0.uniqueName,
          totalsRowLabel: $0.totalsRowLabel,
          totalsRowFunction: $0.totalsRowFunction,
          calculatedColumnFormula: $0.calculatedColumnFormula,
          totalsRowFormula: $0.totalsRowFormula
        )
      },
      style: style
    )
  }
}

private func tableAttribute(
  _ localName: String,
  in attributes: [OfficeXMLAttribute]
) -> String? {
  attributes.first { $0.name.localName == localName && $0.name.namespaceURI == nil }?.value
}
