import Foundation

#if canImport(FoundationXML)
  import FoundationXML
#endif

/// A synchronous, namespace-aware event reader for one XML part.
///
/// The reader wraps Foundation's streaming SAX parser. It does not construct a DOM, resolve
/// external entities, or fetch network resources. A new parser is created for every `parse` call,
/// so the value can safely be reused across tasks when each task supplies its own event handler.
public struct OfficeXMLReader: Sendable {
  private enum Backing: Sendable {
    case data(Data)
    case file(URL)
  }

  private let backing: Backing
  private let source: String

  package var diagnosticSource: String { source }

  /// Resource limits enforced during parsing.
  public let limits: OfficeXMLParsingLimits

  /// XML declaration metadata found at the beginning of the part.
  public let declaration: OfficeXMLDeclaration?

  /// Creates a reader over XML bytes.
  ///
  /// - Parameters:
  ///   - data: The complete bytes for one XML part.
  ///   - source: A diagnostic name, normally an OPC part name.
  ///   - limits: Resource limits enforced while emitting events.
  public init(
    data: Data,
    source: String = "<memory>",
    limits: OfficeXMLParsingLimits = .standard
  ) {
    self.backing = .data(data)
    self.source = source
    self.limits = limits
    self.declaration = XMLDeclarationParser.parse(data)
  }

  /// Creates a reader that streams XML from a local file URL.
  ///
  /// The file is opened when `parse` begins and is not removed or modified by the reader.
  ///
  /// - Parameters:
  ///   - fileURL: A local file containing one XML document.
  ///   - source: A diagnostic name. The file path is used when omitted.
  ///   - limits: Resource limits enforced while emitting events.
  public init(
    contentsOf fileURL: URL,
    source: String? = nil,
    limits: OfficeXMLParsingLimits = .standard
  ) {
    self.backing = .file(fileURL)
    self.source = source ?? fileURL.path
    self.limits = limits
    self.declaration = try? XMLDeclarationParser.parse(contentsOf: fileURL)
  }

  /// Parses the XML and synchronously delivers events in document order.
  ///
  /// Errors thrown by `body` stop parsing immediately and are rethrown unchanged.
  public func parse(_ body: @escaping (OfficeXMLEvent) throws -> Void) throws {
    if try containsDocumentType() {
      throw OfficeKitError.invalidXML(
        part: source,
        message: "Document type declarations are not allowed."
      )
    }

    let delegate = OfficeXMLParserDelegate(
      source: source,
      declaration: declaration,
      limits: limits,
      body: body
    )
    let parser: XMLParser
    switch backing {
    case .data(let data):
      parser = XMLParser(data: data)
    case .file(let fileURL):
      guard let fileParser = XMLParser(contentsOf: fileURL) else {
        throw OfficeKitError.invalidXML(part: source, message: "The XML file could not be opened.")
      }
      parser = fileParser
    }
    parser.shouldProcessNamespaces = false
    parser.shouldReportNamespacePrefixes = false
    parser.shouldResolveExternalEntities = false
    parser.delegate = delegate
    let parsed = parser.parse()
    if let storedError = delegate.storedError { throw storedError }
    guard parsed else {
      throw OfficeKitError.invalidXML(
        part: source,
        message: parser.parserError?.localizedDescription ?? "Unknown XML parser error."
      )
    }
  }

  private func containsDocumentType() throws -> Bool {
    switch backing {
    case .data(let data):
      return XMLSecurityScanner.containsDocumentType(in: data)
    case .file(let fileURL):
      do {
        return try XMLSecurityScanner.containsDocumentType(inFileAt: fileURL)
      } catch {
        throw OfficeKitError.invalidXML(part: source, message: "The XML file could not be read.")
      }
    }
  }
}

private final class OfficeXMLParserDelegate: NSObject, XMLParserDelegate {
  let source: String
  let declaration: OfficeXMLDeclaration?
  let limits: OfficeXMLParsingLimits
  let body: (OfficeXMLEvent) throws -> Void

  var storedError: (any Error)?
  var namespaceScopes: [[String: String]] = [["xml": "http://www.w3.org/XML/1998/namespace"]]
  var depth = 0
  var eventCount = 0
  var textSize: UInt64 = 0

  init(
    source: String,
    declaration: OfficeXMLDeclaration?,
    limits: OfficeXMLParsingLimits,
    body: @escaping (OfficeXMLEvent) throws -> Void
  ) {
    self.source = source
    self.declaration = declaration
    self.limits = limits
    self.body = body
  }

  func parserDidStartDocument(_ parser: XMLParser) {
    emit(.startDocument(declaration), parser: parser)
  }

  func parser(
    _ parser: XMLParser,
    didStartElement elementName: String,
    namespaceURI _: String?,
    qualifiedName _: String?,
    attributes attributeDict: [String: String] = [:]
  ) {
    guard storedError == nil else { return }
    depth += 1
    guard depth <= limits.maximumDepth else {
      fail(
        .limitExceeded(
          limit: .xmlDepth,
          actual: UInt64(depth),
          maximum: UInt64(max(0, limits.maximumDepth))
        ),
        parser: parser
      )
      return
    }
    guard attributeDict.count <= limits.maximumAttributesPerElement else {
      fail(
        .limitExceeded(
          limit: .xmlAttributesPerElement,
          actual: UInt64(attributeDict.count),
          maximum: UInt64(max(0, limits.maximumAttributesPerElement))
        ),
        parser: parser
      )
      return
    }

    var namespaces = namespaceScopes.last ?? [:]
    let declarations = attributeDict.compactMap { qualifiedName, value in
      namespaceDeclaration(qualifiedName: qualifiedName, value: value)
    }.sorted { $0.prefix < $1.prefix }
    for declaration in declarations {
      if declaration.namespaceURI.isEmpty {
        namespaces.removeValue(forKey: declaration.prefix)
      } else {
        namespaces[declaration.prefix] = declaration.namespaceURI
      }
    }
    namespaceScopes.append(namespaces)

    var attributes: [OfficeXMLAttribute] = []
    var attributeNames: Set<OfficeXMLName> = []
    attributes.reserveCapacity(attributeDict.count - declarations.count)
    for (qualifiedName, value) in attributeDict {
      guard namespaceDeclaration(qualifiedName: qualifiedName, value: value) == nil else {
        continue
      }
      guard
        let attribute = makeAttribute(
          qualifiedName: qualifiedName,
          value: value,
          namespaces: namespaces
        ) else {
        fail(
          .invalidXML(
            part: source,
            message: "Attribute \(qualifiedName) uses an undeclared namespace prefix."
          ),
          parser: parser
        )
        return
      }
      guard attributeNames.insert(attribute.name).inserted else {
        fail(
          .invalidXML(
            part: source,
            message: "Duplicate expanded attribute \(attribute.name.localName)."
          ),
          parser: parser
        )
        return
      }
      attributes.append(attribute)
    }
    attributes.sort { lhs, rhs in
      if lhs.name.namespaceURI != rhs.name.namespaceURI {
        return (lhs.name.namespaceURI ?? "") < (rhs.name.namespaceURI ?? "")
      }
      return lhs.name.localName < rhs.name.localName
    }

    guard
      let name = makeName(
        qualifiedName: elementName,
        namespaces: namespaces,
        usesDefaultNamespace: true
      ) else {
      fail(
        .invalidXML(
          part: source,
          message: "Element \(elementName) uses an undeclared namespace prefix."
        ),
        parser: parser
      )
      return
    }
    emit(
      .startElement(
        name: name,
        attributes: attributes,
        namespaceDeclarations: declarations,
        location: location(of: parser)
      ),
      parser: parser
    )
  }

  func parser(_ parser: XMLParser, foundCharacters string: String) {
    guard storedError == nil, !string.isEmpty else { return }
    let size = UInt64(string.utf8.count)
    let (newSize, overflow) = textSize.addingReportingOverflow(size)
    guard !overflow, newSize <= limits.maximumTextSize else {
      fail(
        .limitExceeded(
          limit: .xmlTextSize,
          actual: overflow ? .max : newSize,
          maximum: limits.maximumTextSize
        ),
        parser: parser
      )
      return
    }
    textSize = newSize
    emit(.text(string, location: location(of: parser)), parser: parser)
  }

  func parser(
    _ parser: XMLParser,
    didEndElement elementName: String,
    namespaceURI _: String?,
    qualifiedName _: String?
  ) {
    guard storedError == nil else { return }
    guard
      let name = makeName(
        qualifiedName: elementName,
        namespaces: namespaceScopes.last ?? [:],
        usesDefaultNamespace: true
      ) else {
      fail(
        .invalidXML(
          part: source,
          message: "Element \(elementName) uses an undeclared namespace prefix."
        ),
        parser: parser
      )
      return
    }
    emit(
      .endElement(
        name: name,
        location: location(of: parser)
      ),
      parser: parser
    )
    if namespaceScopes.count > 1 { namespaceScopes.removeLast() }
    depth -= 1
  }

  func parserDidEndDocument(_ parser: XMLParser) {
    emit(.endDocument, parser: parser)
  }

  func parser(
    _ parser: XMLParser,
    foundInternalEntityDeclarationWithName name: String,
    value: String?
  ) {
    rejectEntityDeclaration(parser)
  }

  func parser(
    _ parser: XMLParser,
    foundExternalEntityDeclarationWithName name: String,
    publicID: String?,
    systemID: String?
  ) {
    rejectEntityDeclaration(parser)
  }

  private func rejectEntityDeclaration(_ parser: XMLParser) {
    fail(
      .invalidXML(part: source, message: "Entity declarations are not allowed."),
      parser: parser
    )
  }

  private func emit(_ event: OfficeXMLEvent, parser: XMLParser) {
    guard storedError == nil else { return }
    eventCount += 1
    if eventCount.isMultiple(of: 1_024), Task.isCancelled {
      storedError = CancellationError()
      parser.abortParsing()
      return
    }
    guard eventCount <= limits.maximumEventCount else {
      fail(
        .limitExceeded(
          limit: .xmlEventCount,
          actual: UInt64(eventCount),
          maximum: UInt64(max(0, limits.maximumEventCount))
        ),
        parser: parser
      )
      return
    }
    do {
      try body(event)
    } catch {
      storedError = error
      parser.abortParsing()
    }
  }

  private func fail(_ error: OfficeKitError, parser: XMLParser) {
    storedError = error
    parser.abortParsing()
  }

  private func location(of parser: XMLParser) -> OfficeXMLLocation {
    OfficeXMLLocation(line: parser.lineNumber, column: parser.columnNumber)
  }

  private func namespaceDeclaration(
    qualifiedName: String,
    value: String
  ) -> OfficeXMLNamespaceDeclaration? {
    if qualifiedName == "xmlns" {
      return OfficeXMLNamespaceDeclaration(prefix: "", namespaceURI: value)
    }
    let marker = "xmlns:"
    guard qualifiedName.hasPrefix(marker) else { return nil }
    return OfficeXMLNamespaceDeclaration(
      prefix: String(qualifiedName.dropFirst(marker.count)),
      namespaceURI: value
    )
  }

  private func makeName(
    qualifiedName: String,
    namespaces: [String: String],
    usesDefaultNamespace: Bool
  ) -> OfficeXMLName? {
    guard let colon = qualifiedName.firstIndex(of: ":") else {
      let namespaceURI = usesDefaultNamespace ? namespaces[""].flatMap(nonempty) : nil
      return OfficeXMLName(namespaceURI: namespaceURI, localName: qualifiedName)
    }
    let prefix = String(qualifiedName[..<colon])
    let localName = String(qualifiedName[qualifiedName.index(after: colon)...])
    guard let namespaceURI = namespaces[prefix] else { return nil }
    return OfficeXMLName(namespaceURI: namespaceURI, localName: localName, prefix: prefix)
  }

  private func makeAttribute(
    qualifiedName: String,
    value: String,
    namespaces: [String: String]
  ) -> OfficeXMLAttribute? {
    guard
      let name = makeName(
        qualifiedName: qualifiedName,
        namespaces: namespaces,
        usesDefaultNamespace: false
      ) else { return nil }
    return OfficeXMLAttribute(name: name, value: value)
  }

  private func nonempty(_ value: String) -> String? {
    value.isEmpty ? nil : value
  }
}

private enum XMLDeclarationParser {
  static func parse(_ data: Data) -> OfficeXMLDeclaration? {
    let prefix = data.prefix(1_024)
    guard var text = String(data: prefix, encoding: .utf8) else { return nil }
    if text.first == "\u{FEFF}" { text.removeFirst() }
    guard text.hasPrefix("<?xml"), let end = text.range(of: "?>") else { return nil }
    let declaration = String(text[text.index(text.startIndex, offsetBy: 5)..<end.lowerBound])
    guard let version = attribute(named: "version", in: declaration) else { return nil }
    let standalone: Bool?
    switch attribute(named: "standalone", in: declaration) {
    case "yes": standalone = true
    case "no": standalone = false
    default: standalone = nil
    }
    return OfficeXMLDeclaration(
      version: version,
      encoding: attribute(named: "encoding", in: declaration),
      isStandalone: standalone
    )
  }

  static func parse(contentsOf fileURL: URL) throws -> OfficeXMLDeclaration? {
    let handle = try FileHandle(forReadingFrom: fileURL)
    defer { try? handle.close() }
    return parse(try handle.read(upToCount: 1_024) ?? Data())
  }

  private static func attribute(named name: String, in declaration: String) -> String? {
    var searchStart = declaration.startIndex
    while let range = declaration.range(of: name, range: searchStart..<declaration.endIndex) {
      var index = range.upperBound
      skipWhitespace(in: declaration, index: &index)
      guard index < declaration.endIndex, declaration[index] == "=" else {
        searchStart = range.upperBound
        continue
      }
      index = declaration.index(after: index)
      skipWhitespace(in: declaration, index: &index)
      guard index < declaration.endIndex,
        declaration[index] == "\"" || declaration[index] == "'" else { return nil }
      let quote = declaration[index]
      let valueStart = declaration.index(after: index)
      guard let valueEnd = declaration[valueStart...].firstIndex(of: quote) else { return nil }
      return String(declaration[valueStart..<valueEnd])
    }
    return nil
  }

  private static func skipWhitespace(in text: String, index: inout String.Index) {
    while index < text.endIndex, text[index].isWhitespace {
      index = text.index(after: index)
    }
  }
}

package enum XMLSecurityScanner {
  private static let patterns = [
    encodedDocumentType(stride: 1, characterOffset: 0),
    encodedDocumentType(stride: 2, characterOffset: 0),
    encodedDocumentType(stride: 2, characterOffset: 1),
    encodedDocumentType(stride: 4, characterOffset: 0),
    encodedDocumentType(stride: 4, characterOffset: 3),
  ]

  static func containsDocumentType(in data: Data) -> Bool {
    patterns.contains { data.range(of: $0) != nil }
  }

  static func containsDocumentType(inFileAt fileURL: URL) throws -> Bool {
    let handle = try FileHandle(forReadingFrom: fileURL)
    defer { try? handle.close() }
    let overlapCount = (patterns.map(\.count).max() ?? 1) - 1
    var overlap = Data()
    while let chunk = try handle.read(upToCount: 64 * 1_024), !chunk.isEmpty {
      var window = overlap
      window.append(chunk)
      if containsDocumentType(in: window) { return true }
      overlap = Data(window.suffix(overlapCount))
    }
    return false
  }

  private static func encodedDocumentType(
    stride: Int,
    characterOffset: Int
  ) -> Data {
    var result = Data()
    for character in "<!DOCTYPE".utf8 {
      for offset in 0..<stride {
        result.append(offset == characterOffset ? character : 0)
      }
    }
    return result
  }
}
