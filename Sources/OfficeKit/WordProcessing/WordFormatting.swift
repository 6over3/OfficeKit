package struct WordParagraphPropertiesBuilder {
  var styleIdentifier: String?
  var alignment: String?
  var textAlignment: String?
  var automaticallySpacesEastAsianAndLatinText: Bool?
  var automaticallySpacesEastAsianTextAndNumbers: Bool?
  var adjustsRightIndent: Bool?
  var wrapsAtCharacter: Bool?
  var keepsWithNext: Bool?
  var keepsLinesTogether: Bool?
  var startsOnNewPage: Bool?
  var numberingIdentifier: Int?
  var numberingLevel: Int?
  var spacingBefore: OfficeLength?
  var spacingAfter: OfficeLength?
  var lineSpacing: OfficeLength?
  var lineSpacingRule: String?
  var leadingIndent: OfficeLength?
  var trailingIndent: OfficeLength?
  var firstLineIndent: OfficeLength?
  var hangingIndent: OfficeLength?
  var tabStops: [OfficeWordTabStop] = []

  mutating func consume(_ localName: String, attributes: [OfficeXMLAttribute]) {
    switch localName {
    case "pStyle": styleIdentifier = wordAttribute("val", in: attributes)
    case "jc": alignment = wordAttribute("val", in: attributes)
    case "textAlignment": textAlignment = wordAttribute("val", in: attributes)
    case "autoSpaceDE": automaticallySpacesEastAsianAndLatinText = wordOnOffValue(attributes)
    case "autoSpaceDN": automaticallySpacesEastAsianTextAndNumbers = wordOnOffValue(attributes)
    case "adjustRightInd": adjustsRightIndent = wordOnOffValue(attributes)
    case "wordWrap": wrapsAtCharacter = wordOnOffValue(attributes)
    case "keepNext": keepsWithNext = wordOnOffValue(attributes)
    case "keepLines": keepsLinesTogether = wordOnOffValue(attributes)
    case "pageBreakBefore": startsOnNewPage = wordOnOffValue(attributes)
    case "numId": numberingIdentifier = wordAttribute("val", in: attributes).flatMap(Int.init)
    case "ilvl": numberingLevel = wordAttribute("val", in: attributes).flatMap(Int.init)
    case "spacing":
      spacingBefore = wordAttribute("before", in: attributes).flatMap(wordTwipLength)
      spacingAfter = wordAttribute("after", in: attributes).flatMap(wordTwipLength)
      lineSpacing = wordAttribute("line", in: attributes).flatMap(wordTwipLength)
      lineSpacingRule = wordAttribute("lineRule", in: attributes)
    case "ind":
      leadingIndent =
        (wordAttribute("start", in: attributes)
        ?? wordAttribute("left", in: attributes)).flatMap(wordTwipLength)
      trailingIndent =
        (wordAttribute("end", in: attributes)
        ?? wordAttribute("right", in: attributes)).flatMap(wordTwipLength)
      firstLineIndent = wordAttribute("firstLine", in: attributes).flatMap(wordTwipLength)
      hangingIndent = wordAttribute("hanging", in: attributes).flatMap(wordTwipLength)
    case "tab":
      tabStops.append(
        OfficeWordTabStop(
          alignment: wordAttribute("val", in: attributes),
          leader: wordAttribute("leader", in: attributes),
          position: wordAttribute("pos", in: attributes).flatMap(wordTwipLength)
        ))
    default: break
    }
  }

  var value: OfficeWordParagraphProperties {
    OfficeWordParagraphProperties(
      styleIdentifier: styleIdentifier,
      alignment: alignment,
      textAlignment: textAlignment,
      automaticallySpacesEastAsianAndLatinText: automaticallySpacesEastAsianAndLatinText,
      automaticallySpacesEastAsianTextAndNumbers: automaticallySpacesEastAsianTextAndNumbers,
      adjustsRightIndent: adjustsRightIndent,
      wrapsAtCharacter: wrapsAtCharacter,
      keepsWithNext: keepsWithNext,
      keepsLinesTogether: keepsLinesTogether,
      startsOnNewPage: startsOnNewPage,
      numberingIdentifier: numberingIdentifier,
      numberingLevel: numberingLevel,
      spacingBefore: spacingBefore,
      spacingAfter: spacingAfter,
      lineSpacing: lineSpacing,
      lineSpacingRule: lineSpacingRule,
      leadingIndent: leadingIndent,
      trailingIndent: trailingIndent,
      firstLineIndent: firstLineIndent,
      hangingIndent: hangingIndent,
      tabStops: tabStops
    )
  }
}

package struct WordRunPropertiesBuilder {
  var styleIdentifier: String?
  var isBold: Bool?
  var isItalic: Bool?
  var underline: String?
  var isStruckThrough: Bool?
  var isDoubleStruckThrough: Bool?
  var usesAllCaps: Bool?
  var usesSmallCaps: Bool?
  var color: String?
  var highlight: String?
  var fontSize: OfficeLength?
  var complexScriptFontSize: OfficeLength?
  var asciiFont: String?
  var highAnsiFont: String?
  var eastAsianFont: String?
  var complexScriptFont: String?
  var asciiThemeFont: String?
  var highAnsiThemeFont: String?
  var eastAsianThemeFont: String?
  var complexScriptThemeFont: String?
  var language: String?
  var eastAsianLanguage: String?
  var bidirectionalLanguage: String?
  var verticalAlignment: String?

  mutating func consume(_ localName: String, attributes: [OfficeXMLAttribute]) {
    switch localName {
    case "rStyle": styleIdentifier = wordAttribute("val", in: attributes)
    case "b": isBold = wordOnOffValue(attributes)
    case "i": isItalic = wordOnOffValue(attributes)
    case "u": underline = wordAttribute("val", in: attributes) ?? "single"
    case "strike": isStruckThrough = wordOnOffValue(attributes)
    case "dstrike": isDoubleStruckThrough = wordOnOffValue(attributes)
    case "caps": usesAllCaps = wordOnOffValue(attributes)
    case "smallCaps": usesSmallCaps = wordOnOffValue(attributes)
    case "color": color = wordAttribute("val", in: attributes)
    case "highlight": highlight = wordAttribute("val", in: attributes)
    case "sz": fontSize = wordHalfPointLength(wordAttribute("val", in: attributes))
    case "szCs": complexScriptFontSize = wordHalfPointLength(wordAttribute("val", in: attributes))
    case "rFonts":
      asciiFont = wordAttribute("ascii", in: attributes)
      highAnsiFont = wordAttribute("hAnsi", in: attributes)
      eastAsianFont = wordAttribute("eastAsia", in: attributes)
      complexScriptFont = wordAttribute("cs", in: attributes)
      asciiThemeFont = wordAttribute("asciiTheme", in: attributes)
      highAnsiThemeFont = wordAttribute("hAnsiTheme", in: attributes)
      eastAsianThemeFont = wordAttribute("eastAsiaTheme", in: attributes)
      complexScriptThemeFont = wordAttribute("cstheme", in: attributes)
    case "lang":
      language = wordAttribute("val", in: attributes)
      eastAsianLanguage = wordAttribute("eastAsia", in: attributes)
      bidirectionalLanguage = wordAttribute("bidi", in: attributes)
    case "vertAlign": verticalAlignment = wordAttribute("val", in: attributes)
    default: break
    }
  }

  var value: OfficeWordRunProperties {
    OfficeWordRunProperties(
      styleIdentifier: styleIdentifier,
      isBold: isBold,
      isItalic: isItalic,
      underline: underline,
      isStruckThrough: isStruckThrough,
      isDoubleStruckThrough: isDoubleStruckThrough,
      usesAllCaps: usesAllCaps,
      usesSmallCaps: usesSmallCaps,
      color: color,
      highlight: highlight,
      fontSize: fontSize,
      complexScriptFontSize: complexScriptFontSize,
      asciiFont: asciiFont,
      highAnsiFont: highAnsiFont,
      eastAsianFont: eastAsianFont,
      complexScriptFont: complexScriptFont,
      asciiThemeFont: asciiThemeFont,
      highAnsiThemeFont: highAnsiThemeFont,
      eastAsianThemeFont: eastAsianThemeFont,
      complexScriptThemeFont: complexScriptThemeFont,
      language: language,
      eastAsianLanguage: eastAsianLanguage,
      bidirectionalLanguage: bidirectionalLanguage,
      verticalAlignment: verticalAlignment
    )
  }
}

package func wordHalfPointLength(_ text: String?) -> OfficeLength? {
  guard let text else { return nil }
  if let halfPoints = Double(text), halfPoints.isFinite {
    return OfficeLength(points: halfPoints / 2)
  }
  return wordUniversalMeasure(text)
}

package func wordUniversalMeasure(_ text: String) -> OfficeLength? {
  let units: [(suffix: String, emuPerUnit: Double)] = [
    ("mm", 36_000),
    ("cm", 360_000),
    ("in", 914_400),
    ("pt", 12_700),
    ("pc", 152_400),
    ("pi", 152_400),
  ]
  guard let unit = units.first(where: { text.hasSuffix($0.suffix) }) else { return nil }
  let number = text.dropLast(unit.suffix.count)
  guard !number.isEmpty, let value = Double(number), value.isFinite else { return nil }
  let emu = value * unit.emuPerUnit
  guard emu >= Double(Int64.min), emu <= Double(Int64.max) else { return nil }
  return OfficeLength(emu: Int64(emu.rounded()))
}
