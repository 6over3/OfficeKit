import Foundation

package final class WordDrawingParser {
  private static let wordProcessingDrawingNamespace =
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  private static let drawingNamespace =
    "http://schemas.openxmlformats.org/drawingml/2006/main"
  private static let pictureNamespace =
    "http://schemas.openxmlformats.org/drawingml/2006/picture"
  private static let chartNamespace =
    "http://schemas.openxmlformats.org/drawingml/2006/chart"
  private static let diagramNamespace =
    "http://schemas.openxmlformats.org/drawingml/2006/diagram"
  private static let relationshipNamespace =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  private static let wordShapeNamespace =
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  private static let wordGroupNamespace =
    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  private static let legacyWordShapeNamespace =
    "http://schemas.microsoft.com/office/word/2008/6/28/wordprocessingShape"
  private static let legacyWordGroupNamespace =
    "http://schemas.microsoft.com/office/word/2008/6/28/wordprocessingGroup"
  private static let chartGraphicURI =
    "http://schemas.openxmlformats.org/drawingml/2006/chart"
  private static let strictChartGraphicURI = "http://purl.oclc.org/ooxml/drawingml/chart"
  private static let diagramGraphicURI =
    "http://schemas.openxmlformats.org/drawingml/2006/diagram"
  private static let strictDiagramGraphicURI = "http://purl.oclc.org/ooxml/drawingml/diagram"

  private struct PositionBuilder {
    let relativeFrom: String
    var offset: OfficeLength?
    var alignment: String?

    var value: OfficeWordRelativePosition {
      OfficeWordRelativePosition(
        relativeFrom: relativeFrom,
        offset: offset,
        alignment: alignment
      )
    }
  }

  private struct Builder {
    let startDepth: Int
    let sourceOrder: Int
    let inlinePosition: OfficeWordInlinePosition?
    let placement: OfficeWordDrawingPlacement
    var width: OfficeLength?
    var height: OfficeLength?
    var horizontal: PositionBuilder?
    var vertical: PositionBuilder?
    var simplePosition: OfficePointEMU?
    var distances: OfficeWordTextDistances
    var wrap: String?
    var wrapText: String?
    var relativeHeight: UInt32?
    var isBehindDocument: Bool?
    var allowsOverlap: Bool?
    var laysOutInCell: Bool?
    var identifier: UInt32?
    var name: String?
    var alternativeText: String?
    var title: String?
    var graphicDataURI: String?
    var relationshipIDs: [OfficeRelationshipID] = []
    var isPicture = false
    var isShape = false
    var isGroup = false
  }

  private enum PositionAxis {
    case horizontal
    case vertical
  }

  private enum PositionValue {
    case offset
    case alignment
  }

  private struct TextCapture {
    let depth: Int
    let axis: PositionAxis
    let value: PositionValue
    var text = ""
  }

  private let part: OfficePart
  private let package: OfficePackage
  private let relationshipsByID: [OfficeRelationshipID: OfficeRelationship]
  private var depth = 0
  private var builder: Builder?
  private var currentPositionAxis: PositionAxis?
  private var textCapture: TextCapture?
  package private(set) var drawings: [OfficeWordDrawing] = []

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
      textCapture?.text.append(text)
    case .endElement(let name, _):
      try consumeEnd(name)
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
    if name.namespaceURI == Self.wordProcessingDrawingNamespace,
      name.localName == "inline" || name.localName == "anchor"
    {
      guard builder == nil else {
        throw OfficeKitError.invalidXML(
          part: part.name.rawValue,
          message: "A Word drawing anchor cannot contain another drawing anchor."
        )
      }
      beginAnchor(
        name.localName,
        attributes: attributes,
        sourceOrder: sourceOrder,
        inlinePosition: inlinePosition
      )
      return
    }
    guard builder != nil else { return }
    collectRelationshipIDs(attributes)
    switch (name.namespaceURI, name.localName) {
    case (Self.wordProcessingDrawingNamespace, "extent"):
      builder?.width = lengthAttribute("cx", in: attributes)
      builder?.height = lengthAttribute("cy", in: attributes)
    case (Self.wordProcessingDrawingNamespace, "docPr"):
      builder?.identifier = attribute("id", in: attributes).flatMap(UInt32.init)
      builder?.name = attribute("name", in: attributes)
      builder?.alternativeText = attribute("descr", in: attributes)
      builder?.title = attribute("title", in: attributes)
    case (Self.wordProcessingDrawingNamespace, "simplePos"):
      if let x = lengthAttribute("x", in: attributes),
        let y = lengthAttribute("y", in: attributes)
      {
        builder?.simplePosition = OfficePointEMU(x: x, y: y)
      }
    case (Self.wordProcessingDrawingNamespace, "positionH"):
      let relativeFrom = attribute("relativeFrom", in: attributes) ?? "unknown"
      builder?.horizontal = PositionBuilder(relativeFrom: relativeFrom)
      currentPositionAxis = .horizontal
    case (Self.wordProcessingDrawingNamespace, "positionV"):
      let relativeFrom = attribute("relativeFrom", in: attributes) ?? "unknown"
      builder?.vertical = PositionBuilder(relativeFrom: relativeFrom)
      currentPositionAxis = .vertical
    case (Self.wordProcessingDrawingNamespace, "posOffset"):
      beginTextCapture(.offset)
    case (Self.wordProcessingDrawingNamespace, "align"):
      beginTextCapture(.alignment)
    case (Self.wordProcessingDrawingNamespace, let localName) where localName.hasPrefix("wrap"):
      builder?.wrap = localName
      builder?.wrapText = attribute("wrapText", in: attributes)
    case (Self.drawingNamespace, "graphicData"):
      builder?.graphicDataURI = attribute("uri", in: attributes)
    case (Self.pictureNamespace, "pic"):
      builder?.isPicture = true
    case (Self.wordShapeNamespace, "wsp"), (Self.legacyWordShapeNamespace, "wsp"):
      builder?.isShape = true
    case (Self.wordGroupNamespace, "wgp"), (Self.legacyWordGroupNamespace, "wgp"):
      builder?.isGroup = true
    default:
      break
    }
  }

  private func consumeEnd(_ name: OfficeXMLName) throws {
    if textCapture?.depth == depth,
      name.namespaceURI == Self.wordProcessingDrawingNamespace,
      name.localName == "posOffset" || name.localName == "align"
    {
      finishTextCapture()
    }
    if name.namespaceURI == Self.wordProcessingDrawingNamespace,
      name.localName == "positionH" || name.localName == "positionV"
    {
      currentPositionAxis = nil
    }
    guard name.namespaceURI == Self.wordProcessingDrawingNamespace,
      name.localName == "inline" || name.localName == "anchor",
      builder?.startDepth == depth else { return }
    try finishAnchor()
  }

  private func beginAnchor(
    _ localName: String,
    attributes: [OfficeXMLAttribute],
    sourceOrder: Int,
    inlinePosition: OfficeWordInlinePosition?
  ) {
    let zero = OfficeLength(emu: 0)
    builder = Builder(
      startDepth: depth,
      sourceOrder: sourceOrder,
      inlinePosition: inlinePosition,
      placement: localName == "inline" ? .inline : .floating,
      distances: OfficeWordTextDistances(
        top: lengthAttribute("distT", in: attributes) ?? zero,
        bottom: lengthAttribute("distB", in: attributes) ?? zero,
        left: lengthAttribute("distL", in: attributes) ?? zero,
        right: lengthAttribute("distR", in: attributes) ?? zero
      ),
      relativeHeight: attribute("relativeHeight", in: attributes).flatMap(UInt32.init),
      isBehindDocument: attribute("behindDoc", in: attributes).flatMap(OfficeValueDecoder.boolean),
      allowsOverlap: attribute("allowOverlap", in: attributes).flatMap(OfficeValueDecoder.boolean),
      laysOutInCell: attribute("layoutInCell", in: attributes).flatMap(OfficeValueDecoder.boolean)
    )
  }

  private func beginTextCapture(_ value: PositionValue) {
    guard let currentPositionAxis else { return }
    textCapture = TextCapture(depth: depth, axis: currentPositionAxis, value: value)
  }

  private func finishTextCapture() {
    guard let capture = textCapture else { return }
    let text = capture.text.trimmingCharacters(in: .whitespacesAndNewlines)
    switch (capture.axis, capture.value) {
    case (.horizontal, .offset):
      builder?.horizontal?.offset = Int64(text).map(OfficeLength.init(emu:))
    case (.vertical, .offset): builder?.vertical?.offset = Int64(text).map(OfficeLength.init(emu:))
    case (.horizontal, .alignment): builder?.horizontal?.alignment = text
    case (.vertical, .alignment): builder?.vertical?.alignment = text
    }
    textCapture = nil
  }

  private func collectRelationshipIDs(_ attributes: [OfficeXMLAttribute]) {
    guard builder != nil else { return }
    for attribute in attributes where attribute.name.namespaceURI == Self.relationshipNamespace {
      let identifier = OfficeRelationshipID(rawValue: attribute.value)
      if relationshipsByID[identifier] != nil,
        builder?.relationshipIDs.contains(identifier) == false
      {
        builder?.relationshipIDs.append(identifier)
      }
    }
  }

  private func finishAnchor() throws {
    guard let completed = builder else { return }
    builder = nil
    guard let width = completed.width, let height = completed.height else {
      throw OfficeKitError.invalidXML(
        part: part.name.rawValue,
        message: "A Word drawing anchor requires wp:extent."
      )
    }
    let attachments = completed.relationshipIDs.compactMap { identifier in
      relationshipsByID[identifier].map(package.attachment(referencedBy:))
    }
    let imageAttachments = attachments.filter {
      $0.relationship.type.isEquivalent(to: .image)
        || $0.relationship.type.isEquivalent(to: .highDefinitionPhoto)
    }
    let picture =
      imageAttachments.isEmpty
      ? nil
      : OfficePicture(
        primaryImage: imageAttachments.first, images: imageAttachments, cropRectangle: nil)
    let graphicContent = makeGraphicContent(
      uri: completed.graphicDataURI,
      attachments: attachments
    )
    let kind: OfficeWordDrawingKind
    if completed.isPicture || picture != nil {
      kind = .picture
    } else if attachments.contains(where: { $0.relationship.type.isEquivalent(to: .chart) }) {
      kind = .chart
    } else if completed.graphicDataURI == Self.diagramGraphicURI
      || completed.graphicDataURI == Self.strictDiagramGraphicURI
    {
      kind = .diagram
    } else if completed.isGroup {
      kind = .group
    } else if completed.isShape {
      kind = .shape
    } else if completed.graphicDataURI != nil {
      kind = .graphic
    } else {
      kind = .unknown
    }
    let sourceTransform = OfficeDrawingTransform(
      x: completed.horizontal?.offset ?? OfficeLength(emu: 0),
      y: completed.vertical?.offset ?? OfficeLength(emu: 0),
      width: width,
      height: height
    )
    drawings.append(
      OfficeWordDrawing(
        index: drawings.count,
        sourceOrder: completed.sourceOrder,
        inlinePosition: completed.inlinePosition,
        kind: kind,
        identifier: completed.identifier,
        name: completed.name,
        alternativeText: completed.alternativeText,
        title: completed.title,
        anchor: OfficeWordDrawingAnchor(
          placement: completed.placement,
          width: width,
          height: height,
          horizontalPosition: completed.horizontal?.value,
          verticalPosition: completed.vertical?.value,
          simplePosition: completed.simplePosition,
          textDistances: completed.distances,
          wrap: completed.wrap,
          wrapText: completed.wrapText,
          relativeHeight: completed.relativeHeight,
          isBehindDocument: completed.isBehindDocument,
          allowsOverlap: completed.allowsOverlap,
          laysOutInCell: completed.laysOutInCell
        ),
        picture: picture,
        graphicContent: graphicContent,
        attachments: attachments,
        spatialInfo: OfficeSpatialInfo(
          coordinateSpace: coordinateSpace(for: completed),
          sourceTransform: sourceTransform,
          geometrySourcePart: part.name,
          frame: nil,
          rotation: 0,
          zIndex: drawings.count,
          resolution: .unresolved(
            reason: completed.placement == .inline
              ? "Inline drawing placement depends on Word text layout and pagination."
              : "Floating drawing offsets are relative to layout-dependent Word reference rectangles."
          )
        ),
        sourcePart: part
      )
    )
  }

  private func makeGraphicContent(
    uri: String?,
    attachments: [OfficeAttachment]
  ) -> OfficeGraphicContent? {
    if uri == Self.chartGraphicURI || uri == Self.strictChartGraphicURI,
      let attachment = attachments.first(where: {
        $0.relationship.type.isEquivalent(to: .chart)
      }),
      let chartPart = attachment.part
    {
      return .chart(
        OfficeChartReference(package: package, attachment: attachment, part: chartPart)
      )
    }
    if uri == Self.diagramGraphicURI || uri == Self.strictDiagramGraphicURI {
      return .diagram(OfficeDiagramReference(attachments: attachments))
    }
    return uri.map(OfficeGraphicContent.related(uri:))
  }

  private func coordinateSpace(for builder: Builder) -> OfficeCoordinateSpace {
    guard builder.placement == .floating else { return .character }
    let reference = builder.horizontal?.relativeFrom ?? builder.vertical?.relativeFrom
    return switch reference {
    case "page": .page
    case "margin", "leftMargin", "rightMargin", "insideMargin", "outsideMargin": .margin
    case "column": .column
    case "paragraph", "line": .paragraph
    case "character": .character
    default: .unknown
    }
  }

  private func lengthAttribute(
    _ localName: String,
    in attributes: [OfficeXMLAttribute]
  ) -> OfficeLength? {
    attribute(localName, in: attributes).flatMap(Int64.init).map(OfficeLength.init(emu:))
  }

  private func attribute(
    _ localName: String,
    in attributes: [OfficeXMLAttribute]
  ) -> String? {
    attributes.first {
      $0.name.localName == localName && $0.name.namespaceURI == nil
    }?.value
  }
}
