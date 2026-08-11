package enum WordSettingsParser {
  private static let wordNamespace =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

  package static func parse(
    package: OfficePackage,
    relationships: [OfficeRelationship]
  ) throws -> OfficeWordSettings? {
    guard
      let relationship = relationships.first(where: {
        $0.type.isEquivalent(to: .settings)
      }) else { return nil }
    guard let part = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }

    var view: String?
    var zoomPercentage: Int?
    var defaultTabStop: OfficeLength?
    var characterSpacingControl: String?
    var tracksRevisions: Bool?
    var updatesFieldsOnOpen: Bool?
    var mirrorsMargins: Bool?
    var hasEvenAndOddHeaders: Bool?
    var themeLanguage: String?
    var eastAsianThemeLanguage: String?
    var bidirectionalThemeLanguage: String?
    var decimalSymbol: String?
    var listSeparator: String?
    var compatibilityMode: Int?

    try package.parseXML(in: part, compatibility: .commonOffice) { event in
      guard case .startElement(let name, let attributes, _, _) = event,
        name.namespaceURI == wordNamespace else { return }
      switch name.localName {
      case "view": view = wordAttribute("val", in: attributes)
      case "zoom":
        zoomPercentage = wordAttribute("percent", in: attributes)
          .map { $0.hasSuffix("%") ? String($0.dropLast()) : $0 }
          .flatMap(Int.init)
      case "defaultTabStop":
        defaultTabStop = wordAttribute("val", in: attributes).flatMap(wordTwipLength)
      case "characterSpacingControl":
        characterSpacingControl = wordAttribute("val", in: attributes)
      case "trackRevisions": tracksRevisions = wordOnOffValue(attributes)
      case "updateFields": updatesFieldsOnOpen = wordOnOffValue(attributes)
      case "mirrorMargins": mirrorsMargins = wordOnOffValue(attributes)
      case "evenAndOddHeaders": hasEvenAndOddHeaders = wordOnOffValue(attributes)
      case "themeFontLang":
        themeLanguage = wordAttribute("val", in: attributes)
        eastAsianThemeLanguage = wordAttribute("eastAsia", in: attributes)
        bidirectionalThemeLanguage = wordAttribute("bidi", in: attributes)
      case "decimalSymbol": decimalSymbol = wordAttribute("val", in: attributes)
      case "listSeparator": listSeparator = wordAttribute("val", in: attributes)
      case "compatSetting" where wordAttribute("name", in: attributes) == "compatibilityMode":
        compatibilityMode = wordAttribute("val", in: attributes).flatMap(Int.init)
      default: break
      }
    }

    return OfficeWordSettings(
      view: view,
      zoomPercentage: zoomPercentage,
      defaultTabStop: defaultTabStop,
      characterSpacingControl: characterSpacingControl,
      tracksRevisions: tracksRevisions,
      updatesFieldsOnOpen: updatesFieldsOnOpen,
      mirrorsMargins: mirrorsMargins,
      hasEvenAndOddHeaders: hasEvenAndOddHeaders,
      themeLanguage: themeLanguage,
      eastAsianThemeLanguage: eastAsianThemeLanguage,
      bidirectionalThemeLanguage: bidirectionalThemeLanguage,
      decimalSymbol: decimalSymbol,
      listSeparator: listSeparator,
      compatibilityMode: compatibilityMode,
      sourcePart: part
    )
  }
}
