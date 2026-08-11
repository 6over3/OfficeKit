package enum PresentationThemeParser {
  private static let drawingNamespace =
    "http://schemas.openxmlformats.org/drawingml/2006/main"

  private struct FontBuilder {
    var latinTypeface: String?
    var eastAsianTypeface: String?
    var complexScriptTypeface: String?
    var supplementalFonts: [OfficeThemeSupplementalFont] = []

    var value: OfficeThemeFontSet {
      OfficeThemeFontSet(
        latinTypeface: latinTypeface,
        eastAsianTypeface: eastAsianTypeface,
        complexScriptTypeface: complexScriptTypeface,
        supplementalFonts: supplementalFonts
      )
    }
  }

  package static func parse(
    masterPart: OfficePart,
    package: OfficePackage
  ) throws -> OfficeTheme? {
    guard
      let relationship = try package.relationships(
        from: .part(masterPart.name),
        ofType: .theme
      ).first else { return nil }
    guard let part = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }

    var name: String?
    var colorSchemeName: String?
    var colors: [OfficeThemeColor] = []
    var colorRole: (name: String, depth: Int)?
    var fontSchemeName: String?
    var majorFonts = FontBuilder()
    var minorFonts = FontBuilder()
    var majorFontDepth: Int?
    var minorFontDepth: Int?
    var depth = 0
    let colorRoles: Set<String> = [
      "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4",
      "accent5", "accent6", "hlink", "folHlink",
    ]

    try package.parseXML(in: part, compatibility: .commonOffice) { event in
      switch event {
      case .startElement(let element, let attributes, _, _):
        depth += 1
        guard element.namespaceURI == drawingNamespace else { return }
        switch element.localName {
        case "theme": name = themeAttribute("name", in: attributes)
        case "clrScheme": colorSchemeName = themeAttribute("name", in: attributes)
        case let role where colorRoles.contains(role): colorRole = (role, depth)
        case "srgbClr":
          if let role = colorRole?.name, let value = themeAttribute("val", in: attributes) {
            colors.append(OfficeThemeColor(name: role, value: .sRGB(value)))
          }
        case "sysClr":
          if let role = colorRole?.name, let value = themeAttribute("val", in: attributes) {
            colors.append(
              OfficeThemeColor(
                name: role,
                value: .system(
                  name: value,
                  lastColor: themeAttribute("lastClr", in: attributes)
                )
              ))
          }
        case "fontScheme": fontSchemeName = themeAttribute("name", in: attributes)
        case "majorFont": majorFontDepth = depth
        case "minorFont": minorFontDepth = depth
        case "latin":
          mutateActiveFont(
            majorDepth: majorFontDepth,
            minorDepth: minorFontDepth,
            major: &majorFonts,
            minor: &minorFonts
          ) { $0.latinTypeface = themeAttribute("typeface", in: attributes) }
        case "ea":
          mutateActiveFont(
            majorDepth: majorFontDepth,
            minorDepth: minorFontDepth,
            major: &majorFonts,
            minor: &minorFonts
          ) { $0.eastAsianTypeface = themeAttribute("typeface", in: attributes) }
        case "cs":
          mutateActiveFont(
            majorDepth: majorFontDepth,
            minorDepth: minorFontDepth,
            major: &majorFonts,
            minor: &minorFonts
          ) { $0.complexScriptTypeface = themeAttribute("typeface", in: attributes) }
        case "font":
          if let script = themeAttribute("script", in: attributes),
            let typeface = themeAttribute("typeface", in: attributes)
          {
            mutateActiveFont(
              majorDepth: majorFontDepth,
              minorDepth: minorFontDepth,
              major: &majorFonts,
              minor: &minorFonts
            ) {
              $0.supplementalFonts.append(
                OfficeThemeSupplementalFont(script: script, typeface: typeface)
              )
            }
          }
        default: break
        }
      case .endElement(let element, _):
        if colorRole?.depth == depth { colorRole = nil }
        if element.namespaceURI == drawingNamespace, element.localName == "majorFont",
          majorFontDepth == depth
        {
          majorFontDepth = nil
        }
        if element.namespaceURI == drawingNamespace, element.localName == "minorFont",
          minorFontDepth == depth
        {
          minorFontDepth = nil
        }
        depth -= 1
      case .text, .startDocument, .endDocument:
        break
      }
    }

    return OfficeTheme(
      name: name,
      colorSchemeName: colorSchemeName,
      colors: colors,
      fontSchemeName: fontSchemeName,
      majorFonts: majorFonts.value,
      minorFonts: minorFonts.value,
      sourcePart: part,
      attachment: package.attachment(referencedBy: relationship)
    )
  }

  private static func mutateActiveFont(
    majorDepth: Int?,
    minorDepth: Int?,
    major: inout FontBuilder,
    minor: inout FontBuilder,
    _ mutation: (inout FontBuilder) -> Void
  ) {
    if majorDepth != nil {
      mutation(&major)
    } else if minorDepth != nil {
      mutation(&minor)
    }
  }
}

private func themeAttribute(
  _ localName: String,
  in attributes: [OfficeXMLAttribute]
) -> String? {
  attributes.first { $0.name.localName == localName && $0.name.namespaceURI == nil }?.value
}
