/// The declaration at the beginning of an XML part.
public struct OfficeXMLDeclaration: Sendable, Hashable, Codable {
  /// The declared XML version, normally `1.0`.
  public let version: String

  /// The declared character encoding, or `nil` when XML encoding detection applies.
  public let encoding: String?

  /// The declared standalone value, or `nil` when it was omitted.
  public let isStandalone: Bool?

  /// Creates XML declaration metadata.
  public init(version: String, encoding: String? = nil, isStandalone: Bool? = nil) {
    self.version = version
    self.encoding = encoding
    self.isStandalone = isStandalone
  }
}

/// A location reported by the streaming XML parser.
public struct OfficeXMLLocation: Sendable, Hashable, Codable {
  /// The one-based source line, when available.
  public let line: Int

  /// The one-based source column, when available.
  public let column: Int

  /// Creates a source location.
  public init(line: Int, column: Int) {
    self.line = line
    self.column = column
  }
}

/// A namespace binding declared on an XML element.
public struct OfficeXMLNamespaceDeclaration: Sendable, Hashable, Codable {
  /// The bound prefix, or an empty string for the default namespace.
  public let prefix: String

  /// The namespace URI bound to the prefix.
  public let namespaceURI: String

  /// Creates a namespace declaration.
  public init(prefix: String, namespaceURI: String) {
    self.prefix = prefix
    self.namespaceURI = namespaceURI
  }
}

/// A namespace-aware XML name.
public struct OfficeXMLName: Sendable, Hashable, Codable, CustomStringConvertible {
  /// The namespace URI after known Strict OOXML spellings are canonicalized.
  public let namespaceURI: String?

  /// The local name without a namespace prefix.
  public let localName: String

  /// The source prefix, or `nil` when the name was unprefixed.
  public let prefix: String?

  /// Creates a namespace-aware name.
  public init(namespaceURI: String?, localName: String, prefix: String? = nil) {
    self.namespaceURI = namespaceURI.map(OfficeXMLNamespace.canonicalize)
    self.localName = localName
    self.prefix = prefix
  }

  /// The source-style qualified name.
  public var description: String {
    if let prefix, !prefix.isEmpty { return "\(prefix):\(localName)" }
    return localName
  }

  /// Compares expanded names; source prefixes do not affect XML name identity.
  public static func == (lhs: Self, rhs: Self) -> Bool {
    lhs.namespaceURI == rhs.namespaceURI && lhs.localName == rhs.localName
  }

  /// Hashes the expanded name independently of its source prefix.
  public func hash(into hasher: inout Hasher) {
    hasher.combine(namespaceURI)
    hasher.combine(localName)
  }
}

/// A namespace-aware attribute on an XML element.
public struct OfficeXMLAttribute: Sendable, Hashable, Codable {
  /// The attribute name.
  public let name: OfficeXMLName

  /// The attribute value after XML entity and character-reference decoding.
  public let value: String

  /// Creates an XML attribute.
  public init(name: OfficeXMLName, value: String) {
    self.name = name
    self.value = value
  }
}

/// One event emitted while reading an XML part.
public enum OfficeXMLEvent: Sendable, Hashable, Codable {
  /// The start of a document and its optional XML declaration.
  case startDocument(OfficeXMLDeclaration?)

  /// An opening element, its attributes, local namespace declarations, and source location.
  case startElement(
    name: OfficeXMLName,
    attributes: [OfficeXMLAttribute],
    namespaceDeclarations: [OfficeXMLNamespaceDeclaration],
    location: OfficeXMLLocation
  )

  /// Character data. Consecutive text may arrive in more than one event.
  case text(String, location: OfficeXMLLocation)

  /// A closing element and its source location.
  case endElement(name: OfficeXMLName, location: OfficeXMLLocation)

  /// The end of a successfully parsed document.
  case endDocument
}

/// Limits applied by the namespace-aware XML event reader.
public struct OfficeXMLParsingLimits: Sendable, Hashable, Codable {
  /// The maximum element nesting depth.
  public let maximumDepth: Int

  /// The maximum number of attributes accepted on one element.
  public let maximumAttributesPerElement: Int

  /// The maximum number of events emitted from one part.
  public let maximumEventCount: Int

  /// The maximum cumulative UTF-8 size of character data in one part.
  public let maximumTextSize: UInt64

  /// Creates XML parser limits.
  public init(
    maximumDepth: Int = 256,
    maximumAttributesPerElement: Int = 1_024,
    maximumEventCount: Int = 10_000_000,
    maximumTextSize: UInt64 = 256 * 1_024 * 1_024
  ) {
    self.maximumDepth = maximumDepth
    self.maximumAttributesPerElement = maximumAttributesPerElement
    self.maximumEventCount = maximumEventCount
    self.maximumTextSize = maximumTextSize
  }

  /// Limits suitable for normal Office XML parts.
  public static let standard = OfficeXMLParsingLimits()
}

/// Canonical namespace handling for Strict and Transitional Office Open XML.
public enum OfficeXMLNamespace {
  /// Canonicalizes a known Strict namespace URI to its Transitional semantic equivalent.
  ///
  /// Unknown and extension namespace URIs are preserved exactly.
  public static func canonicalize(_ namespaceURI: String) -> String {
    canonicalNamespaces[namespaceURI] ?? namespaceURI
  }

  private static let canonicalNamespaces = [
    "http://schemas.openxmlformats.org/wordprocessingml/2006/6/main":
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://schemas.openxmlformats.org/spreadsheetml/2006/7/main":
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://schemas.openxmlformats.org/presentationml/2006/3/main":
      "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/3/main":
      "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://schemas.microsoft.com/office/word/2010/11/wordml":
      "http://schemas.microsoft.com/office/word/2012/wordml",
    "http://purl.oclc.org/ooxml/descriptions/base":
      "http://descriptions.openxmlformats.org/description/base",
    "http://purl.oclc.org/ooxml/descriptions/full":
      "http://descriptions.openxmlformats.org/description/full",
    "http://purl.oclc.org/ooxml/wordprocessingml/main":
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/spreadsheetml/main":
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://purl.oclc.org/ooxml/presentationml/main":
      "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://purl.oclc.org/ooxml/drawingml/main":
      "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://purl.oclc.org/ooxml/drawingml/chart":
      "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://purl.oclc.org/ooxml/drawingml/chartDrawing":
      "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    "http://purl.oclc.org/ooxml/drawingml/diagram":
      "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "http://purl.oclc.org/ooxml/drawingml/picture":
      "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing":
      "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing":
      "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://purl.oclc.org/ooxml/drawingml/lockedCanvas":
      "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
    "http://purl.oclc.org/ooxml/drawingml/compatibility":
      "http://schemas.openxmlformats.org/drawingml/2006/compatibility",
    "http://purl.oclc.org/ooxml/officeDocument/bibliography":
      "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
    "http://purl.oclc.org/ooxml/officeDocument/customProperties":
      "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "http://purl.oclc.org/ooxml/officeDocument/customXml":
      "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
    "http://purl.oclc.org/ooxml/officeDocument/customXmlDataProps":
      "http://schemas.openxmlformats.org/officeDocument/2006/customXmlDataProps",
    "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes":
      "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    "http://purl.oclc.org/ooxml/officeDocument/extendedProperties":
      "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "http://purl.oclc.org/ooxml/officeDocument/math":
      "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://purl.oclc.org/ooxml/officeDocument/sharedTypes":
      "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
    "http://purl.oclc.org/ooxml/schemaLibrary/main":
      "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "http://purl.oclc.org/ooxml/officeDocument/relationships":
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  ]
}
