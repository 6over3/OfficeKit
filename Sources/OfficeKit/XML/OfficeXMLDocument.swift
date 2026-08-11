import Foundation

/// One node in a structured Office XML document.
public indirect enum OfficeXMLNode: Sendable, Hashable, Codable {
  /// A namespace-aware XML element.
  case element(OfficeXMLElement)

  /// Decoded character data and its source location.
  case text(String, location: OfficeXMLLocation)
}

/// A namespace-aware XML element with its complete authored content.
public struct OfficeXMLElement: Sendable, Hashable, Codable {
  /// The expanded and source-qualified element name.
  public let name: OfficeXMLName

  /// Attributes in deterministic expanded-name order.
  public let attributes: [OfficeXMLAttribute]

  /// Namespace bindings declared directly on this element.
  public let namespaceDeclarations: [OfficeXMLNamespaceDeclaration]

  /// Child elements and character data in document order.
  public let children: [OfficeXMLNode]

  /// The source location of the opening tag.
  public let location: OfficeXMLLocation

  /// Returns the first attribute with the requested expanded name.
  public func attribute(named localName: String, namespaceURI: String? = nil) -> String? {
    let canonicalNamespace = namespaceURI.map(OfficeXMLNamespace.canonicalize)
    return attributes.first {
      $0.name.localName == localName && $0.name.namespaceURI == canonicalNamespace
    }?.value
  }

  /// Direct child elements in document order.
  public var childElements: [OfficeXMLElement] {
    children.compactMap { node in
      guard case .element(let element) = node else { return nil }
      return element
    }
  }

  /// Decoded text from this element and all descendants in document order.
  public var textContent: String {
    var text = ""
    appendText(to: &text)
    return text
  }
}

/// A complete, structured view of one XML document.
///
/// This model retains all schema and extension elements, attributes, namespace declarations, and
/// character data. Use ``OfficePackage/parseXML(in:maximumPartSize:limits:compatibility:_:)`` for
/// large parts that should remain streaming rather than materializing this tree.
public struct OfficeXMLDocument: Sendable, Hashable, Codable {
  /// The source XML declaration, when present.
  public let declaration: OfficeXMLDeclaration?

  /// The document element.
  public let root: OfficeXMLElement

  package init(declaration: OfficeXMLDeclaration?, root: OfficeXMLElement) {
    self.declaration = declaration
    self.root = root
  }

  /// Materializes all events emitted by `reader` as a structured document.
  public init(reading reader: OfficeXMLReader) throws {
    var builder = OfficeXMLDocumentBuilder(source: reader.diagnosticSource)
    try reader.parse { event in
      try builder.consume(event)
    }
    self = try builder.document()
  }

  /// Visits every element in document order without allocating a flattened element collection.
  ///
  /// The root path is empty. Each path component is an index into its parent's `children` array,
  /// so paths retain intervening text nodes exactly.
  public func traverseElements(
    _ body: (OfficeXMLElement, [Int]) throws -> Void
  ) rethrows {
    var path: [Int] = []
    try root.traverseElements(path: &path, body)
  }
}

/// One relationship-valued attribute found in a structured XML part.
public struct OfficeXMLRelationshipReference: Sendable {
  /// The path to the declaring element within the structured XML document.
  public let elementPath: [Int]

  /// The declaring element's expanded name.
  public let elementName: OfficeXMLName

  /// The relationship-valued attribute.
  public let attribute: OfficeXMLAttribute

  /// The referenced relationship identifier.
  public let relationshipID: OfficeRelationshipID

  /// The matching relationship, or `nil` for a dangling identifier.
  public let relationship: OfficeRelationship?

  /// A lazy URL-backed attachment for the relationship, when it exists.
  public let attachment: OfficeAttachment?
}

/// A structured XML part together with its complete relationship context.
public struct OfficeParsedXMLPart: Sendable {
  /// The source package part.
  public let part: OfficePart

  /// Every element, attribute, namespace declaration, and text node in the part.
  public let document: OfficeXMLDocument

  /// All relationships owned by the part.
  public let relationships: [OfficeRelationship]

  /// Lazy URL-backed attachments for all relationships owned by the part.
  public let attachments: [OfficeAttachment]

  /// Relationship-valued attributes found anywhere in the XML tree.
  public let relationshipReferences: [OfficeXMLRelationshipReference]
}

extension OfficePackage {
  /// Parses any XML package part without discarding unknown schema or extension content.
  ///
  /// Relationships are resolved after XML parsing, avoiding reentrant parser work. Binary
  /// resources remain lazy and are exposed through URL-backed attachments.
  public func parsedXMLPart(
    _ part: OfficePart,
    maximumPartSize: UInt64? = nil,
    limits: OfficeXMLParsingLimits? = nil,
    compatibility: OfficeXMLCompatibilityOptions? = nil
  ) throws -> OfficeParsedXMLPart {
    var builder = OfficeXMLDocumentBuilder(source: part.name.rawValue)
    try parseXML(
      in: part,
      maximumPartSize: maximumPartSize,
      limits: limits,
      compatibility: compatibility
    ) { event in
      try builder.consume(event)
    }
    let document = try builder.document()
    let relationships = try relationships(from: .part(part.name))
    var relationshipsByID: [OfficeRelationshipID: OfficeRelationship] = [:]
    for relationship in relationships where relationshipsByID[relationship.id] == nil {
      relationshipsByID[relationship.id] = relationship
    }
    var references: [OfficeXMLRelationshipReference] = []
    document.traverseElements { element, path in
      for attribute in element.attributes {
        guard Self.isRelationshipAttribute(attribute) else { continue }
        let id = OfficeRelationshipID(rawValue: attribute.value)
        let relationship = relationshipsByID[id]
        references.append(
          OfficeXMLRelationshipReference(
            elementPath: path,
            elementName: element.name,
            attribute: attribute,
            relationshipID: id,
            relationship: relationship,
            attachment: relationship.map(attachment(referencedBy:))
          )
        )
      }
    }
    return OfficeParsedXMLPart(
      part: part,
      document: document,
      relationships: relationships,
      attachments: relationships.map(attachment(referencedBy:)),
      relationshipReferences: references
    )
  }

  private static func isRelationshipAttribute(_ attribute: OfficeXMLAttribute) -> Bool {
    attribute.name.namespaceURI
      == "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  }
}

extension OfficeXMLElement {
  fileprivate func appendText(to result: inout String) {
    for child in children {
      switch child {
      case .element(let element): element.appendText(to: &result)
      case .text(let text, _): result.append(text)
      }
    }
  }

  fileprivate func traverseElements(
    path: inout [Int],
    _ body: (OfficeXMLElement, [Int]) throws -> Void
  ) rethrows {
    try body(self, path)
    for (index, child) in children.enumerated() {
      guard case .element(let element) = child else { continue }
      path.append(index)
      try element.traverseElements(path: &path, body)
      path.removeLast()
    }
  }
}

private struct OfficeXMLDocumentBuilder {
  private struct ElementBuilder {
    let name: OfficeXMLName
    let attributes: [OfficeXMLAttribute]
    let namespaceDeclarations: [OfficeXMLNamespaceDeclaration]
    let location: OfficeXMLLocation
    var children: [OfficeXMLNode] = []

    func element() -> OfficeXMLElement {
      OfficeXMLElement(
        name: name,
        attributes: attributes,
        namespaceDeclarations: namespaceDeclarations,
        children: children,
        location: location
      )
    }
  }

  let source: String
  var declaration: OfficeXMLDeclaration?
  private var stack: [ElementBuilder] = []
  var root: OfficeXMLElement?
  var reachedEnd = false

  fileprivate init(source: String) {
    self.source = source
  }

  mutating func consume(_ event: OfficeXMLEvent) throws {
    switch event {
    case .startDocument(let declaration):
      self.declaration = declaration
    case .startElement(let name, let attributes, let declarations, let location):
      stack.append(
        ElementBuilder(
          name: name,
          attributes: attributes,
          namespaceDeclarations: declarations,
          location: location
        )
      )
    case .text(let text, let location):
      guard !stack.isEmpty else { return }
      stack[stack.count - 1].children.append(.text(text, location: location))
    case .endElement(let name, _):
      guard let completed = stack.popLast(), completed.name == name else {
        throw OfficeKitError.invalidXML(part: source, message: "Mismatched element boundary.")
      }
      let element = completed.element()
      if stack.isEmpty {
        guard root == nil else {
          throw OfficeKitError.invalidXML(part: source, message: "Multiple document elements.")
        }
        root = element
      } else {
        stack[stack.count - 1].children.append(.element(element))
      }
    case .endDocument:
      reachedEnd = true
    }
  }

  func document() throws -> OfficeXMLDocument {
    guard reachedEnd, stack.isEmpty, let root else {
      throw OfficeKitError.invalidXML(part: source, message: "Incomplete XML document.")
    }
    return OfficeXMLDocument(declaration: declaration, root: root)
  }
}
