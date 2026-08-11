import Foundation

package final class WordAuxiliaryContentParser {
  private static let wordNamespace =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  private static let mathNamespace =
    "http://schemas.openxmlformats.org/officeDocument/2006/math"
  private static let relationshipNamespace =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  private static let officeNamespace = "urn:schemas-microsoft-com:office:office"
  private static let vmlNamespace = "urn:schemas-microsoft-com:vml"

  private struct EquationBuilder {
    let startDepth: Int
    let sourceOrder: Int
    let inlinePosition: OfficeWordInlinePosition?
    let isDisplay: Bool
    var text = ""
  }

  private struct VMLShapeBuilder {
    let startDepth: Int
    let identifier: String?
    let typeIdentifier: String?
    let style: String?
    let alternativeText: String?
    let fillColor: String?
    let strokeColor: String?
    var relationshipIDs: [OfficeRelationshipID] = []
  }

  private let part: OfficePart
  private let package: OfficePackage
  private let relationshipsByID: [OfficeRelationshipID: OfficeRelationship]
  private var depth = 0
  private var mathParagraphDepth: Int?
  private var equation: EquationBuilder?
  private var mathTextDepth: Int?
  private var shapeStack: [VMLShapeBuilder] = []
  package private(set) var equations: [OfficeWordEquation] = []
  package private(set) var alternativeFormatImports: [OfficeWordAlternativeFormatImport] = []
  package private(set) var embeddedObjects: [OfficeWordEmbeddedObject] = []
  package private(set) var legacyShapes: [OfficeWordVMLShape] = []

  package init(
    part: OfficePart,
    package: OfficePackage,
    relationships: [OfficeRelationship]
  ) {
    self.part = part
    self.package = package
    self.relationshipsByID = Dictionary(
      relationships.map { ($0.id, $0) },
      uniquingKeysWith: { first, _ in first }
    )
  }

  package func consume(
    _ event: OfficeXMLEvent,
    sourceOrder: Int,
    inlinePosition: OfficeWordInlinePosition?
  ) throws {
    switch event {
    case .startElement(let name, let attributes, _, _):
      depth += 1
      try consumeStart(
        name,
        attributes: attributes,
        sourceOrder: sourceOrder,
        inlinePosition: inlinePosition
      )
    case .text(let text, _):
      if mathTextDepth != nil { equation?.text.append(text) }
    case .endElement(let name, _):
      consumeEnd(name)
      depth -= 1
    case .startDocument, .endDocument:
      break
    }
  }

  private func consumeStart(
    _ name: OfficeXMLName,
    attributes: [OfficeXMLAttribute],
    sourceOrder: Int,
    inlinePosition: OfficeWordInlinePosition?
  ) throws {
    switch (name.namespaceURI, name.localName) {
    case (Self.mathNamespace, "oMathPara"):
      mathParagraphDepth = depth
    case (Self.mathNamespace, "oMath"):
      if equation == nil {
        equation = EquationBuilder(
          startDepth: depth,
          sourceOrder: sourceOrder,
          inlinePosition: inlinePosition,
          isDisplay: mathParagraphDepth != nil
        )
      }
    case (Self.mathNamespace, "t"):
      if equation != nil { mathTextDepth = depth }
    case (Self.wordNamespace, "altChunk"):
      try appendAlternativeFormatImport(attributes)
    case (Self.officeNamespace, "OLEObject"):
      try appendEmbeddedObject(attributes)
    case (Self.vmlNamespace, "shape"):
      shapeStack.append(
        VMLShapeBuilder(
          startDepth: depth,
          identifier: attribute("id", in: attributes),
          typeIdentifier: attribute("type", in: attributes),
          style: attribute("style", in: attributes),
          alternativeText: attribute("alt", in: attributes),
          fillColor: attribute("fillcolor", in: attributes),
          strokeColor: attribute("strokecolor", in: attributes)
        )
      )
    default:
      collectVMLRelationshipIDs(attributes)
    }
  }

  private func consumeEnd(_ name: OfficeXMLName) {
    if name.namespaceURI == Self.mathNamespace, name.localName == "t",
      mathTextDepth == depth
    {
      mathTextDepth = nil
    }
    if name.namespaceURI == Self.mathNamespace, name.localName == "oMath",
      equation?.startDepth == depth, let completed = equation
    {
      equations.append(
        OfficeWordEquation(
          index: equations.count,
          sourceOrder: completed.sourceOrder,
          inlinePosition: completed.inlinePosition,
          isDisplay: completed.isDisplay,
          text: completed.text,
          sourcePart: part
        )
      )
      equation = nil
    }
    if name.namespaceURI == Self.mathNamespace, name.localName == "oMathPara",
      mathParagraphDepth == depth
    {
      mathParagraphDepth = nil
    }
    if name.namespaceURI == Self.vmlNamespace, name.localName == "shape",
      shapeStack.last?.startDepth == depth
    {
      finishVMLShape(shapeStack.removeLast())
    }
  }

  private func appendAlternativeFormatImport(
    _ attributes: [OfficeXMLAttribute]
  ) throws {
    guard let relationshipID = relationshipID(in: attributes),
      let relationship = relationshipsByID[relationshipID] else {
      throw OfficeKitError.invalidXML(
        part: part.name.rawValue,
        message: "w:altChunk requires a resolvable r:id relationship."
      )
    }
    alternativeFormatImports.append(
      OfficeWordAlternativeFormatImport(
        index: alternativeFormatImports.count,
        relationshipID: relationshipID,
        attachment: package.attachment(referencedBy: relationship),
        sourcePart: part
      )
    )
  }

  private func appendEmbeddedObject(
    _ attributes: [OfficeXMLAttribute]
  ) throws {
    guard let relationshipID = relationshipID(in: attributes),
      let relationship = relationshipsByID[relationshipID] else {
      throw OfficeKitError.invalidXML(
        part: part.name.rawValue,
        message: "o:OLEObject requires a resolvable r:id relationship."
      )
    }
    embeddedObjects.append(
      OfficeWordEmbeddedObject(
        index: embeddedObjects.count,
        relationshipID: relationshipID,
        type: attribute("Type", in: attributes),
        programIdentifier: attribute("ProgID", in: attributes),
        drawingAspect: attribute("DrawAspect", in: attributes),
        shapeIdentifier: attribute("ShapeID", in: attributes),
        objectIdentifier: attribute("ObjectID", in: attributes),
        attachment: package.attachment(referencedBy: relationship),
        sourcePart: part
      )
    )
  }

  private func collectVMLRelationshipIDs(_ attributes: [OfficeXMLAttribute]) {
    guard !shapeStack.isEmpty else { return }
    let index = shapeStack.count - 1
    for attribute in attributes where attribute.name.namespaceURI == Self.relationshipNamespace {
      let identifier = OfficeRelationshipID(rawValue: attribute.value)
      if relationshipsByID[identifier] != nil,
        !shapeStack[index].relationshipIDs.contains(identifier)
      {
        shapeStack[index].relationshipIDs.append(identifier)
      }
    }
  }

  private func finishVMLShape(_ shape: VMLShapeBuilder) {
    let attachments = shape.relationshipIDs.compactMap { identifier in
      relationshipsByID[identifier].map(package.attachment(referencedBy:))
    }
    let style = vmlStyle(shape.style)
    let hasRelativePercent = style.keys.contains { $0.hasSuffix("-percent") }
    let x = vmlLength(style["margin-left"])
    let y = vmlLength(style["margin-top"])
    let width = vmlLength(style["width"])
    let height = vmlLength(style["height"])
    let frame: OfficeRect?
    if !hasRelativePercent, let width, let height {
      frame = OfficeRect(x: x ?? 0, y: y ?? 0, width: width, height: height)
    } else {
      frame = nil
    }
    let coordinateSpace = vmlCoordinateSpace(style)
    let rotationDegrees = style["rotation"].flatMap(Double.init) ?? 0
    let flips = Set((style["flip"] ?? "").split(separator: " ").map(String.init))
    legacyShapes.append(
      OfficeWordVMLShape(
        index: legacyShapes.count,
        identifier: shape.identifier,
        typeIdentifier: shape.typeIdentifier,
        style: shape.style,
        alternativeText: shape.alternativeText,
        fillColor: shape.fillColor,
        strokeColor: shape.strokeColor,
        attachments: attachments,
        spatialInfo: OfficeSpatialInfo(
          coordinateSpace: coordinateSpace,
          geometrySourcePart: part.name,
          frame: frame,
          rotation: rotationDegrees * .pi / 180,
          isFlippedHorizontally: flips.contains("x"),
          isFlippedVertically: flips.contains("y"),
          zIndex: style["z-index"].flatMap(Int.init),
          resolution: frame != nil && x != nil && y != nil
            ? .exact
            : .unresolved(
              reason: "VML size is authored, but placement uses relative or unsupported values."
            )
        ),
        sourcePart: part
      )
    )
  }

  private func relationshipID(
    in attributes: [OfficeXMLAttribute]
  ) -> OfficeRelationshipID? {
    attributes.first {
      $0.name.namespaceURI == Self.relationshipNamespace && $0.name.localName == "id"
    }.map { OfficeRelationshipID(rawValue: $0.value) }
  }

  private func attribute(
    _ localName: String,
    in attributes: [OfficeXMLAttribute]
  ) -> String? {
    attributes.first {
      $0.name.localName == localName && $0.name.namespaceURI == nil
    }?.value
  }

  private func vmlStyle(_ rawValue: String?) -> [String: String] {
    guard let rawValue else { return [:] }
    return rawValue.split(separator: ";").reduce(into: [:]) { result, declaration in
      let components = declaration.split(separator: ":", maxSplits: 1)
      guard components.count == 2 else { return }
      result[String(components[0]).trimmingCharacters(in: .whitespaces)] =
        String(components[1]).trimmingCharacters(in: .whitespaces)
    }
  }

  private func vmlLength(_ rawValue: String?) -> Double? {
    guard let rawValue else { return nil }
    let units: [(suffix: String, pointsPerUnit: Double)] = [
      ("pt", 1),
      ("in", 72),
      ("cm", 72 / 2.54),
      ("mm", 72 / 25.4),
      ("px", 72 / 96),
    ]
    for unit in units where rawValue.hasSuffix(unit.suffix) {
      return Double(rawValue.dropLast(unit.suffix.count)).map { $0 * unit.pointsPerUnit }
    }
    return Double(rawValue)
  }

  private func vmlCoordinateSpace(_ style: [String: String]) -> OfficeCoordinateSpace {
    let relative =
      style["mso-position-horizontal-relative"]
      ?? style["mso-position-vertical-relative"]
    return switch relative {
    case "page": .page
    case "margin": .margin
    case "text": .paragraph
    case "char": .character
    default: .unknown
    }
  }
}
