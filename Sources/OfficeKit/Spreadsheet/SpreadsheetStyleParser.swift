package enum SpreadsheetStyleParser {
  private static let spreadsheetNamespace =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  private struct RawFormat {
    let numberFormatID: UInt32
    let fontIndex: UInt32
    let fillIndex: UInt32
    let borderIndex: UInt32
    var alignment: OfficeSpreadsheetAlignment?
    var protection: OfficeSpreadsheetProtection?
  }

  private struct FontBuilder {
    let startDepth: Int
    var name: String?
    var sizeInPoints: Double?
    var isBold = false
    var isItalic = false
    var underline: String?
    var isStruckThrough = false
    var color: OfficeSpreadsheetColor?
    var family: UInt32?
    var characterSet: UInt32?
    var scheme: String?

    var value: OfficeSpreadsheetFont {
      OfficeSpreadsheetFont(
        name: name,
        sizeInPoints: sizeInPoints,
        isBold: isBold,
        isItalic: isItalic,
        underline: underline,
        isStruckThrough: isStruckThrough,
        color: color,
        family: family,
        characterSet: characterSet,
        scheme: scheme
      )
    }
  }

  private struct FillBuilder {
    let startDepth: Int
    var patternType: String?
    var foregroundColor: OfficeSpreadsheetColor?
    var backgroundColor: OfficeSpreadsheetColor?

    var value: OfficeSpreadsheetFill {
      OfficeSpreadsheetFill(
        patternType: patternType,
        foregroundColor: foregroundColor,
        backgroundColor: backgroundColor
      )
    }
  }

  private struct BorderEdgeBuilder {
    var style: String?
    var color: OfficeSpreadsheetColor?

    var value: OfficeSpreadsheetBorderEdge {
      OfficeSpreadsheetBorderEdge(style: style, color: color)
    }
  }

  private struct BorderBuilder {
    let startDepth: Int
    var leading = BorderEdgeBuilder()
    var trailing = BorderEdgeBuilder()
    var top = BorderEdgeBuilder()
    var bottom = BorderEdgeBuilder()
    var diagonal = BorderEdgeBuilder()
    var diagonalUp = false
    var diagonalDown = false

    var value: OfficeSpreadsheetBorder {
      OfficeSpreadsheetBorder(
        leading: leading.value,
        trailing: trailing.value,
        top: top.value,
        bottom: bottom.value,
        diagonal: diagonal.value,
        diagonalUp: diagonalUp,
        diagonalDown: diagonalDown
      )
    }

    mutating func set(_ edge: BorderEdgeBuilder, for localName: String) {
      switch localName {
      case "left", "start": leading = edge
      case "right", "end": trailing = edge
      case "top": top = edge
      case "bottom": bottom = edge
      case "diagonal": diagonal = edge
      default: break
      }
    }
  }

  package static func parse(
    package: OfficePackage,
    workbookPart: OfficePart
  ) throws -> [OfficeCellStyle] {
    guard
      let relationship = try package.relationships(
        from: .part(workbookPart.name),
        ofType: .styles
      ).first else { return [] }
    guard let part = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }

    var customCodes: [UInt32: String] = [:]
    var formats: [RawFormat] = []
    var fonts: [OfficeSpreadsheetFont] = []
    var fills: [OfficeSpreadsheetFill] = []
    var borders: [OfficeSpreadsheetBorder] = []
    var fontsDepth: Int?
    var fillsDepth: Int?
    var bordersDepth: Int?
    var cellFormatsDepth: Int?
    var activeCellFormatDepth: Int?
    var font: FontBuilder?
    var fill: FillBuilder?
    var border: BorderBuilder?
    var borderEdge: (name: String, depth: Int, value: BorderEdgeBuilder)?
    var depth = 0
    try package.parseXML(in: part, compatibility: .commonOffice) { event in
      switch event {
      case .startElement(let name, let attributes, _, _):
        depth += 1
        guard name.namespaceURI == spreadsheetNamespace else { return }
        if name.localName == "fonts" {
          fontsDepth = depth
        } else if name.localName == "font", fontsDepth == depth - 1 {
          font = FontBuilder(startDepth: depth)
        } else if font != nil {
          switch name.localName {
          case "name": font?.name = styleAttribute("val", in: attributes)
          case "sz": font?.sizeInPoints = styleAttribute("val", in: attributes).flatMap(Double.init)
          case "b": font?.isBold = styleBoolean(attributes)
          case "i": font?.isItalic = styleBoolean(attributes)
          case "u":
            font?.underline =
              styleBoolean(attributes) ? styleAttribute("val", in: attributes) ?? "single" : nil
          case "strike": font?.isStruckThrough = styleBoolean(attributes)
          case "color": font?.color = styleColor(attributes)
          case "family": font?.family = styleAttribute("val", in: attributes).flatMap(UInt32.init)
          case "charset":
            font?.characterSet = styleAttribute("val", in: attributes).flatMap(UInt32.init)
          case "scheme": font?.scheme = styleAttribute("val", in: attributes)
          default: break
          }
        } else if name.localName == "fills" {
          fillsDepth = depth
        } else if name.localName == "fill", fillsDepth == depth - 1 {
          fill = FillBuilder(startDepth: depth)
        } else if fill != nil {
          switch name.localName {
          case "patternFill": fill?.patternType = styleAttribute("patternType", in: attributes)
          case "fgColor": fill?.foregroundColor = styleColor(attributes)
          case "bgColor": fill?.backgroundColor = styleColor(attributes)
          default: break
          }
        } else if name.localName == "borders" {
          bordersDepth = depth
        } else if name.localName == "border", bordersDepth == depth - 1 {
          border = BorderBuilder(
            startDepth: depth,
            diagonalUp: styleAttribute("diagonalUp", in: attributes)
              .flatMap(OfficeValueDecoder.boolean) ?? false,
            diagonalDown: styleAttribute("diagonalDown", in: attributes)
              .flatMap(OfficeValueDecoder.boolean) ?? false
          )
        } else if border != nil,
          ["left", "start", "right", "end", "top", "bottom", "diagonal"]
            .contains(name.localName)
        {
          borderEdge = (
            name.localName,
            depth,
            BorderEdgeBuilder(style: styleAttribute("style", in: attributes))
          )
        } else if name.localName == "color", borderEdge != nil {
          borderEdge?.value.color = styleColor(attributes)
        } else if name.localName == "numFmt",
          let identifier = styleAttribute("numFmtId", in: attributes).flatMap(UInt32.init),
          let code = styleAttribute("formatCode", in: attributes)
        {
          customCodes[identifier] = code
        } else if name.localName == "cellXfs" {
          cellFormatsDepth = depth
        } else if name.localName == "xf", cellFormatsDepth == depth - 1 {
          formats.append(
            RawFormat(
              numberFormatID: styleAttribute("numFmtId", in: attributes).flatMap(UInt32.init) ?? 0,
              fontIndex: styleAttribute("fontId", in: attributes).flatMap(UInt32.init) ?? 0,
              fillIndex: styleAttribute("fillId", in: attributes).flatMap(UInt32.init) ?? 0,
              borderIndex: styleAttribute("borderId", in: attributes).flatMap(UInt32.init) ?? 0,
              alignment: nil,
              protection: nil
            )
          )
          activeCellFormatDepth = depth
        } else if name.localName == "alignment", activeCellFormatDepth != nil,
          !formats.isEmpty
        {
          formats[formats.count - 1].alignment = OfficeSpreadsheetAlignment(
            horizontal: styleAttribute("horizontal", in: attributes),
            vertical: styleAttribute("vertical", in: attributes),
            textRotation: styleAttribute("textRotation", in: attributes).flatMap(Int.init),
            wrapsText: styleAttribute("wrapText", in: attributes).flatMap(
              OfficeValueDecoder.boolean),
            shrinksToFit: styleAttribute("shrinkToFit", in: attributes).flatMap(
              OfficeValueDecoder.boolean),
            indent: styleAttribute("indent", in: attributes).flatMap(Int.init),
            relativeIndent: styleAttribute("relativeIndent", in: attributes).flatMap(Int.init),
            readingOrder: styleAttribute("readingOrder", in: attributes).flatMap(UInt32.init),
            justifiesLastLine: styleAttribute("justifyLastLine", in: attributes)
              .flatMap(OfficeValueDecoder.boolean)
          )
        } else if name.localName == "protection", activeCellFormatDepth != nil,
          !formats.isEmpty
        {
          formats[formats.count - 1].protection = OfficeSpreadsheetProtection(
            isLocked: styleAttribute("locked", in: attributes).flatMap(OfficeValueDecoder.boolean),
            isHidden: styleAttribute("hidden", in: attributes).flatMap(OfficeValueDecoder.boolean)
          )
        }
      case .endElement(let name, _):
        if name.namespaceURI == spreadsheetNamespace, name.localName == "font",
          font?.startDepth == depth, let completed = font
        {
          fonts.append(completed.value)
          font = nil
        }
        if name.namespaceURI == spreadsheetNamespace, name.localName == "fonts",
          fontsDepth == depth
        {
          fontsDepth = nil
        }
        if name.namespaceURI == spreadsheetNamespace, name.localName == "fill",
          fill?.startDepth == depth, let completed = fill
        {
          fills.append(completed.value)
          fill = nil
        }
        if name.namespaceURI == spreadsheetNamespace, name.localName == "fills",
          fillsDepth == depth
        {
          fillsDepth = nil
        }
        if name.namespaceURI == spreadsheetNamespace, borderEdge?.name == name.localName,
          borderEdge?.depth == depth, let completed = borderEdge
        {
          border?.set(completed.value, for: completed.name)
          borderEdge = nil
        }
        if name.namespaceURI == spreadsheetNamespace, name.localName == "border",
          border?.startDepth == depth, let completed = border
        {
          borders.append(completed.value)
          border = nil
        }
        if name.namespaceURI == spreadsheetNamespace, name.localName == "borders",
          bordersDepth == depth
        {
          bordersDepth = nil
        }
        if name.namespaceURI == spreadsheetNamespace, name.localName == "xf",
          activeCellFormatDepth == depth
        {
          activeCellFormatDepth = nil
        }
        if cellFormatsDepth == depth, name.namespaceURI == spreadsheetNamespace,
          name.localName == "cellXfs"
        {
          cellFormatsDepth = nil
        }
        depth -= 1
      case .text, .startDocument, .endDocument:
        break
      }
    }

    return try formats.enumerated().map { index, format in
      guard fonts.indices.contains(Int(format.fontIndex)),
        fills.indices.contains(Int(format.fillIndex)),
        borders.indices.contains(Int(format.borderIndex)) else {
        throw OfficeKitError.invalidPackage(
          "Cell style \(index) references a missing font, fill, or border record."
        )
      }
      let code = customCodes[format.numberFormatID] ?? builtInCodes[format.numberFormatID]
      return OfficeCellStyle(
        index: UInt32(index),
        numberFormat: OfficeNumberFormat(
          identifier: format.numberFormatID,
          code: code,
          isDate: isDateFormat(identifier: format.numberFormatID, code: code)
        ),
        fontIndex: format.fontIndex,
        fillIndex: format.fillIndex,
        borderIndex: format.borderIndex,
        font: fonts[Int(format.fontIndex)],
        fill: fills[Int(format.fillIndex)],
        border: borders[Int(format.borderIndex)],
        alignment: format.alignment,
        protection: format.protection
      )
    }
  }

  private static func isDateFormat(identifier: UInt32, code: String?) -> Bool {
    if builtInDateFormatIDs.contains(identifier) { return true }
    guard let code else { return false }

    var semanticCode = ""
    var isQuoted = false
    var skipsNext = false
    var bracket = ""
    var isBracketed = false
    for character in code.lowercased() {
      if skipsNext {
        skipsNext = false
        continue
      }
      if character == "\\" || character == "_" || character == "*" {
        skipsNext = true
        continue
      }
      if character == "\"" {
        isQuoted.toggle()
        continue
      }
      if isQuoted { continue }
      if character == "[" {
        isBracketed = true
        bracket = ""
        continue
      }
      if character == "]", isBracketed {
        if bracket == "h" || bracket == "hh" || bracket == "m" || bracket == "mm"
          || bracket == "s" || bracket == "ss"
        {
          semanticCode.append(contentsOf: bracket)
        }
        isBracketed = false
        continue
      }
      if isBracketed {
        bracket.append(character)
      } else {
        semanticCode.append(character)
      }
    }
    return semanticCode.contains("y") || semanticCode.contains("d")
      || semanticCode.contains("h") || semanticCode.contains("s")
  }

  private static let builtInDateFormatIDs: Set<UInt32> =
    Set(14...22)
    .union(27...36)
    .union(45...47)
    .union(50...58)

  private static let builtInCodes: [UInt32: String] = [
    14: "m/d/yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
  ]
}

private func styleAttribute(
  _ localName: String,
  in attributes: [OfficeXMLAttribute]
) -> String? {
  attributes.first { $0.name.localName == localName && $0.name.namespaceURI == nil }?.value
}

private func styleBoolean(_ attributes: [OfficeXMLAttribute]) -> Bool {
  styleAttribute("val", in: attributes).flatMap(OfficeValueDecoder.boolean) ?? true
}

private func styleColor(
  _ attributes: [OfficeXMLAttribute]
) -> OfficeSpreadsheetColor {
  OfficeSpreadsheetColor(
    argb: styleAttribute("rgb", in: attributes),
    indexed: styleAttribute("indexed", in: attributes).flatMap(UInt32.init),
    theme: styleAttribute("theme", in: attributes).flatMap(UInt32.init),
    tint: styleAttribute("tint", in: attributes).flatMap(Double.init),
    isAutomatic: styleAttribute("auto", in: attributes).flatMap(OfficeValueDecoder.boolean)
  )
}
