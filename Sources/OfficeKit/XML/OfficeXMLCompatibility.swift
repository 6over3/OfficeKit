/// Namespace capabilities used when processing OOXML markup-compatibility branches.
public struct OfficeXMLCompatibilityOptions: Sendable, Hashable, Codable {
  /// Namespace URIs understood by the consuming semantic parser.
  public let supportedNamespaces: Set<String>

  /// Creates a supported-namespace policy.
  ///
  /// Strict namespace spellings are canonicalized before storage and comparison.
  public init(supportedNamespaces: Set<String>) {
    self.supportedNamespaces = Set(supportedNamespaces.map(OfficeXMLNamespace.canonicalize))
  }

  /// Namespaces handled by OfficeKit's common Word, Excel, PowerPoint, and DrawingML readers.
  public static let commonOffice = OfficeXMLCompatibilityOptions(supportedNamespaces: [
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "http://schemas.microsoft.com/office/word/2008/6/28/wordprocessingCanvas",
    "http://schemas.microsoft.com/office/word/2008/6/28/wordprocessingGroup",
    "http://schemas.microsoft.com/office/word/2008/6/28/wordprocessingInk",
    "http://schemas.microsoft.com/office/word/2008/6/28/wordprocessingShape",
    "http://schemas.microsoft.com/office/word/2008/9/16/wordprocessingDrawing",
    "http://schemas.microsoft.com/office/word/2008/9/12/wordml",
    "http://schemas.microsoft.com/office/word/2010/wordml",
    "http://schemas.microsoft.com/office/drawing/2017/model3d",
  ])

  package func supports(_ namespaceURI: String) -> Bool {
    supportedNamespaces.contains(OfficeXMLNamespace.canonicalize(namespaceURI))
  }
}

extension OfficeXMLReader {
  /// Parses effective XML after applying supported OOXML markup-compatibility rules.
  ///
  /// `mc:AlternateContent`, `mc:Choice`, and `mc:Fallback` wrappers are omitted. For each block,
  /// the first choice whose `Requires` prefixes all resolve to supported namespaces is emitted;
  /// otherwise the fallback is emitted. Unsupported elements in an `mc:Ignorable` namespace are
  /// omitted, or unwrapped when selected by `mc:ProcessContent`. The transformation is streaming
  /// and does not build a DOM.
  public func parseCompatible(
    using options: OfficeXMLCompatibilityOptions = .commonOffice,
    _ body: @escaping (OfficeXMLEvent) throws -> Void
  ) throws {
    let filter = OfficeXMLCompatibilityFilter(
      source: diagnosticSource,
      options: options,
      body: body
    )
    try parse(filter.consume)
  }
}

private final class OfficeXMLCompatibilityFilter {
  private static let markupCompatibilityNamespace =
    "http://schemas.openxmlformats.org/markup-compatibility/2006"

  private struct AlternateState {
    let depth: Int
    var selectedBranch = false
    var branchDepth: Int?
    var emitsBranch = false
  }

  private struct QualifiedNamePattern {
    let namespaceURI: String?
    let localName: String?

    func matches(_ name: OfficeXMLName) -> Bool {
      (namespaceURI == nil || namespaceURI == name.namespaceURI)
        && (localName == nil || localName == name.localName)
    }
  }

  private struct CompatibilityScope {
    var ignorableNamespaces: Set<String> = []
    var processContent: [QualifiedNamePattern] = []
    var preserveElements: [QualifiedNamePattern] = []
    var preserveAttributes: [QualifiedNamePattern] = []
  }

  private struct IgnoredElementState {
    let depth: Int
    let processesContent: Bool
  }

  private let options: OfficeXMLCompatibilityOptions
  private let source: String
  private let body: (OfficeXMLEvent) throws -> Void
  private var namespaceScopes: [[String: String]] = [
    ["xml": "http://www.w3.org/XML/1998/namespace"]
  ]
  private var compatibilityScopes: [CompatibilityScope] = [CompatibilityScope()]
  private var alternates: [AlternateState] = []
  private var ignoredElements: [IgnoredElementState] = []
  private var depth = 0

  init(
    source: String,
    options: OfficeXMLCompatibilityOptions,
    body: @escaping (OfficeXMLEvent) throws -> Void
  ) {
    self.source = source
    self.options = options
    self.body = body
  }

  func consume(_ event: OfficeXMLEvent) throws {
    switch event {
    case .startDocument, .endDocument:
      try body(event)
    case .startElement(
      let name,
      let attributes,
      let namespaceDeclarations,
      let location
    ):
      depth += 1
      var namespaces = namespaceScopes.last ?? [:]
      for declaration in namespaceDeclarations {
        namespaces[declaration.prefix] = OfficeXMLNamespace.canonicalize(
          declaration.namespaceURI
        )
      }
      namespaceScopes.append(namespaces)
      var compatibility = compatibilityScopes.last ?? CompatibilityScope()
      try applyCompatibilityAttributes(
        attributes,
        namespaces: namespaces,
        location: location,
        to: &compatibility
      )
      compatibilityScopes.append(compatibility)

      if isCompatibilityElement(name, localName: "AlternateContent") {
        alternates.append(AlternateState(depth: depth))
        return
      }

      if let alternateIndex = directAlternateIndex(),
        isCompatibilityElement(name, localName: "Choice")
      {
        let emits =
          !alternates[alternateIndex].selectedBranch
          && requirementsAreSupported(attributes: attributes, namespaces: namespaces)
        alternates[alternateIndex].branchDepth = depth
        alternates[alternateIndex].emitsBranch = emits
        if emits { alternates[alternateIndex].selectedBranch = true }
        return
      }

      if let alternateIndex = directAlternateIndex(),
        isCompatibilityElement(name, localName: "Fallback")
      {
        let emits = !alternates[alternateIndex].selectedBranch
        alternates[alternateIndex].branchDepth = depth
        alternates[alternateIndex].emitsBranch = emits
        if emits { alternates[alternateIndex].selectedBranch = true }
        return
      }

      if isUnsupportedIgnorable(name, compatibility: compatibility) {
        ignoredElements.append(
          IgnoredElementState(
            depth: depth,
            processesContent: compatibility.processContent.contains { $0.matches(name) }
          )
        )
        return
      }

      if branchesEmit && ignoredContentEmits {
        let filteredAttributes = attributes.filter { attribute in
          guard attribute.name.namespaceURI != Self.markupCompatibilityNamespace else {
            return false
          }
          guard let namespaceURI = attribute.name.namespaceURI else { return true }
          return options.supports(namespaceURI)
            || !compatibility.ignorableNamespaces.contains(namespaceURI)
            || compatibility.preserveAttributes.contains { $0.matches(attribute.name) }
        }
        try body(
          .startElement(
            name: name,
            attributes: filteredAttributes,
            namespaceDeclarations: namespaceDeclarations,
            location: location
          )
        )
      }

    case .text:
      if branchesEmit && ignoredContentEmits { try body(event) }

    case .endElement(let name, _):
      if let alternateIndex = directAlternateIndex(),
        alternates[alternateIndex].branchDepth == depth,
        isCompatibilityElement(name, localName: "Choice")
          || isCompatibilityElement(name, localName: "Fallback")
      {
        alternates[alternateIndex].branchDepth = nil
        alternates[alternateIndex].emitsBranch = false
      } else if let alternate = alternates.last,
        alternate.depth == depth,
        isCompatibilityElement(name, localName: "AlternateContent")
      {
        alternates.removeLast()
      } else if ignoredElements.last?.depth == depth {
        ignoredElements.removeLast()
      } else if branchesEmit && ignoredContentEmits {
        try body(event)
      }
      if namespaceScopes.count > 1 { namespaceScopes.removeLast() }
      if compatibilityScopes.count > 1 { compatibilityScopes.removeLast() }
      depth -= 1
    }
  }

  private var branchesEmit: Bool {
    alternates.allSatisfy { alternate in
      guard depth > alternate.depth else { return true }
      return alternate.branchDepth != nil && alternate.emitsBranch
    }
  }

  private var ignoredContentEmits: Bool {
    ignoredElements.allSatisfy { ignored in
      depth <= ignored.depth || ignored.processesContent
    }
  }

  private func isUnsupportedIgnorable(
    _ name: OfficeXMLName,
    compatibility: CompatibilityScope
  ) -> Bool {
    guard let namespaceURI = name.namespaceURI else { return false }
    return compatibility.ignorableNamespaces.contains(namespaceURI)
      && !options.supports(namespaceURI)
      && !compatibility.preserveElements.contains { $0.matches(name) }
  }

  private func applyCompatibilityAttributes(
    _ attributes: [OfficeXMLAttribute],
    namespaces: [String: String],
    location: OfficeXMLLocation,
    to compatibility: inout CompatibilityScope
  ) throws {
    if let mustUnderstand = compatibilityAttribute("MustUnderstand", in: attributes) {
      for prefix in mustUnderstand.split(whereSeparator: \.isWhitespace) {
        guard let namespaceURI = namespaces[String(prefix)], options.supports(namespaceURI) else {
          throw OfficeKitError.invalidXML(
            part: source,
            message: "Unsupported mc:MustUnderstand namespace prefix \(prefix) at "
              + "line \(location.line), column \(location.column)."
          )
        }
      }
    }
    if let ignorable = compatibilityAttribute("Ignorable", in: attributes) {
      for prefix in ignorable.split(whereSeparator: \.isWhitespace) {
        if let namespaceURI = namespaces[String(prefix)] {
          compatibility.ignorableNamespaces.insert(namespaceURI)
        }
      }
    }
    if let processContent = compatibilityAttribute("ProcessContent", in: attributes) {
      for qualifiedName in processContent.split(whereSeparator: \.isWhitespace) {
        if let pattern = qualifiedNamePattern(String(qualifiedName), namespaces: namespaces) {
          compatibility.processContent.append(pattern)
        }
      }
    }
    appendPatterns(
      from: compatibilityAttribute("PreserveElements", in: attributes),
      namespaces: namespaces,
      to: &compatibility.preserveElements
    )
    appendPatterns(
      from: compatibilityAttribute("PreserveAttributes", in: attributes),
      namespaces: namespaces,
      to: &compatibility.preserveAttributes
    )
  }

  private func appendPatterns(
    from value: String?,
    namespaces: [String: String],
    to patterns: inout [QualifiedNamePattern]
  ) {
    guard let value else { return }
    for qualifiedName in value.split(whereSeparator: \.isWhitespace) {
      if let pattern = qualifiedNamePattern(String(qualifiedName), namespaces: namespaces) {
        patterns.append(pattern)
      }
    }
  }

  private func compatibilityAttribute(
    _ localName: String,
    in attributes: [OfficeXMLAttribute]
  ) -> String? {
    attributes.first {
      $0.name.namespaceURI == Self.markupCompatibilityNamespace
        && $0.name.localName == localName
    }?.value
  }

  private func qualifiedNamePattern(
    _ value: String,
    namespaces: [String: String]
  ) -> QualifiedNamePattern? {
    if value == "*" { return QualifiedNamePattern(namespaceURI: nil, localName: nil) }
    let components = value.split(separator: ":", maxSplits: 1, omittingEmptySubsequences: false)
    guard components.count == 2, let namespaceURI = namespaces[String(components[0])] else {
      return nil
    }
    let localName = components[1] == "*" ? nil : String(components[1])
    return QualifiedNamePattern(namespaceURI: namespaceURI, localName: localName)
  }

  private func directAlternateIndex() -> Int? {
    alternates.indices.last { alternates[$0].depth + 1 == depth }
  }

  private func requirementsAreSupported(
    attributes: [OfficeXMLAttribute],
    namespaces: [String: String]
  ) -> Bool {
    guard
      let requirements = attributes.first(where: {
        $0.name.namespaceURI == nil && $0.name.localName == "Requires"
      })?.value else { return false }
    let prefixes = requirements.split(whereSeparator: \.isWhitespace)
    guard !prefixes.isEmpty else { return false }
    return prefixes.allSatisfy { prefix in
      guard let namespaceURI = namespaces[String(prefix)] else { return false }
      return options.supports(namespaceURI)
    }
  }

  private func isCompatibilityElement(_ name: OfficeXMLName, localName: String) -> Bool {
    name.namespaceURI == Self.markupCompatibilityNamespace && name.localName == localName
  }
}
