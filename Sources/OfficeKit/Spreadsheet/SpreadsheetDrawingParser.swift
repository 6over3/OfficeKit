import Foundation

package enum SpreadsheetDrawingParser {
  package static func parse(part: OfficePart, package: OfficePackage) throws
    -> OfficeWorksheetDrawing
  {
    let relationships = try package.relationships(from: .part(part.name))
    let relationshipsByID = Dictionary(
      relationships.map { ($0.id, $0) },
      uniquingKeysWith: { first, _ in first }
    )
    let session = SpreadsheetDrawingParsingSession(
      part: part,
      package: package,
      relationshipsByID: relationshipsByID
    )
    try package.parseXML(in: part, compatibility: .commonOffice, session.consume)
    return OfficeWorksheetDrawing(
      part: part,
      elements: try session.result(),
      attachments: relationships.map(package.attachment(referencedBy:))
    )
  }
}

private final class SpreadsheetDrawingParsingSession {
  private static let spreadsheetDrawingNamespace =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  private static let drawingNamespace =
    "http://schemas.openxmlformats.org/drawingml/2006/main"
  private static let relationshipNamespace =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  private static let diagramGraphicURI =
    "http://schemas.openxmlformats.org/drawingml/2006/diagram"
  private static let strictDiagramGraphicURI =
    "http://purl.oclc.org/ooxml/drawingml/diagram"
  private static let model3DGraphicURI =
    "http://schemas.microsoft.com/office/drawing/2017/model3d"

  private enum AnchorKind {
    case twoCell(editBehavior: String?)
    case oneCell
    case absolute
  }

  private enum MarkerTarget {
    case from
    case to
  }

  private struct RawMarker {
    var columnIndex: Int?
    var columnOffset: Int64?
    var rowIndex: Int?
    var rowOffset: Int64?

    var value: OfficeWorksheetCellMarker? {
      guard let columnIndex, let columnOffset, let rowIndex, let rowOffset else { return nil }
      return OfficeWorksheetCellMarker(
        columnIndex: columnIndex,
        columnOffset: OfficeLength(emu: columnOffset),
        rowIndex: rowIndex,
        rowOffset: OfficeLength(emu: rowOffset)
      )
    }
  }

  private struct RawTransform {
    var x: Int64?
    var y: Int64?
    var width: Int64?
    var height: Int64?
    var rotation: Int64 = 0
    var flipH = false
    var flipV = false

    var value: OfficeDrawingTransform? {
      guard let x, let y, let width, let height, width > 0, height > 0 else { return nil }
      return OfficeDrawingTransform(
        x: OfficeLength(emu: x),
        y: OfficeLength(emu: y),
        width: OfficeLength(emu: width),
        height: OfficeLength(emu: height),
        rotationUnits: rotation,
        isFlippedHorizontally: flipH,
        isFlippedVertically: flipV
      )
    }
  }

  private struct Builder {
    let kind: AnchorKind
    let startDepth: Int
    var elementKind: OfficeWorksheetDrawingElementKind?
    var identifier: UInt32?
    var name: String?
    var alternativeText: String?
    var text = ""
    var from = RawMarker()
    var to = RawMarker()
    var positionX: Int64?
    var positionY: Int64?
    var extentWidth: Int64?
    var extentHeight: Int64?
    var relationshipIDs: [OfficeRelationshipID] = []
    var graphicDataURI: String?
    var transform = RawTransform()
    var transformDepth: Int?
    var hasReadTransform = false
    var textDepth: Int?
  }

  private let part: OfficePart
  private let package: OfficePackage
  private let relationshipsByID: [OfficeRelationshipID: OfficeRelationship]
  private var depth = 0
  private var builders: [Builder] = []
  private var markerTarget: MarkerTarget?
  private var markerValueName: String?
  private var markerValueDepth: Int?
  private var markerText = ""

  init(
    part: OfficePart,
    package: OfficePackage,
    relationshipsByID: [OfficeRelationshipID: OfficeRelationship]
  ) {
    self.part = part
    self.package = package
    self.relationshipsByID = relationshipsByID
  }

  func consume(_ event: OfficeXMLEvent) throws {
    switch event {
    case .startElement(let name, let attributes, _, _):
      depth += 1
      if name.namespaceURI == Self.spreadsheetDrawingNamespace {
        switch name.localName {
        case "twoCellAnchor":
          builders.append(
            Builder(
              kind: .twoCell(editBehavior: drawingAttribute("editAs", in: attributes)),
              startDepth: depth
            )
          )
        case "oneCellAnchor":
          builders.append(Builder(kind: .oneCell, startDepth: depth))
        case "absoluteAnchor":
          builders.append(Builder(kind: .absolute, startDepth: depth))
        default:
          break
        }
      }
      guard !builders.isEmpty else { return }

      if name.namespaceURI == Self.spreadsheetDrawingNamespace {
        switch name.localName {
        case "from": markerTarget = .from
        case "to": markerTarget = .to
        case "col", "colOff", "row", "rowOff":
          if markerTarget != nil {
            markerValueName = name.localName
            markerValueDepth = depth
            markerText = ""
          }
        case "pos":
          builders[builders.count - 1].positionX = drawingAttribute("x", in: attributes)
            .flatMap(Int64.init)
          builders[builders.count - 1].positionY = drawingAttribute("y", in: attributes)
            .flatMap(Int64.init)
        case "ext":
          builders[builders.count - 1].extentWidth = drawingAttribute("cx", in: attributes)
            .flatMap(Int64.init)
          builders[builders.count - 1].extentHeight = drawingAttribute("cy", in: attributes)
            .flatMap(Int64.init)
        case "sp": builders[builders.count - 1].elementKind = .shape
        case "pic": builders[builders.count - 1].elementKind = .picture
        case "cxnSp": builders[builders.count - 1].elementKind = .connector
        case "graphicFrame": builders[builders.count - 1].elementKind = .graphicFrame
        case "grpSp": builders[builders.count - 1].elementKind = .group
        case "cNvPr" where builders[builders.count - 1].identifier == nil:
          builders[builders.count - 1].identifier = drawingAttribute("id", in: attributes)
            .flatMap(UInt32.init)
          builders[builders.count - 1].name = drawingAttribute("name", in: attributes)
          builders[builders.count - 1].alternativeText = drawingAttribute("descr", in: attributes)
        case "xfrm" where !builders[builders.count - 1].hasReadTransform:
          builders[builders.count - 1].hasReadTransform = true
          builders[builders.count - 1].transformDepth = depth
          builders[builders.count - 1].transform.rotation =
            drawingAttribute("rot", in: attributes)
            .flatMap(Int64.init) ?? 0
          builders[builders.count - 1].transform.flipH =
            drawingAttribute("flipH", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false
          builders[builders.count - 1].transform.flipV =
            drawingAttribute("flipV", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false
        default:
          break
        }
      }

      if name.namespaceURI == Self.drawingNamespace {
        if name.localName == "graphicData" {
          builders[builders.count - 1].graphicDataURI = drawingAttribute("uri", in: attributes)
        } else if name.localName == "xfrm", !builders[builders.count - 1].hasReadTransform {
          builders[builders.count - 1].hasReadTransform = true
          builders[builders.count - 1].transformDepth = depth
          builders[builders.count - 1].transform.rotation =
            drawingAttribute("rot", in: attributes)
            .flatMap(Int64.init) ?? 0
          builders[builders.count - 1].transform.flipH =
            drawingAttribute("flipH", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false
          builders[builders.count - 1].transform.flipV =
            drawingAttribute("flipV", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false
        } else if builders[builders.count - 1].transformDepth != nil, name.localName == "off" {
          builders[builders.count - 1].transform.x = drawingAttribute("x", in: attributes)
            .flatMap(Int64.init)
          builders[builders.count - 1].transform.y = drawingAttribute("y", in: attributes)
            .flatMap(Int64.init)
        } else if builders[builders.count - 1].transformDepth != nil, name.localName == "ext" {
          builders[builders.count - 1].transform.width = drawingAttribute("cx", in: attributes)
            .flatMap(Int64.init)
          builders[builders.count - 1].transform.height = drawingAttribute("cy", in: attributes)
            .flatMap(Int64.init)
        } else if name.localName == "t" {
          builders[builders.count - 1].textDepth = depth
        }
      }

      for attribute in attributes
      where attribute.name.namespaceURI == Self.relationshipNamespace && !attribute.value.isEmpty {
        let identifier = OfficeRelationshipID(rawValue: attribute.value)
        if !builders[builders.count - 1].relationshipIDs.contains(identifier) {
          builders[builders.count - 1].relationshipIDs.append(identifier)
        }
      }

    case .text(let text, _):
      if markerValueDepth != nil { markerText.append(text) }
      if builders.last?.textDepth != nil { builders[builders.count - 1].text.append(text) }

    case .endElement(let name, _):
      guard !builders.isEmpty else {
        depth -= 1
        return
      }
      if markerValueDepth == depth, let markerValueName {
        assignMarkerValue(name: markerValueName, text: markerText)
        self.markerValueName = nil
        markerValueDepth = nil
        markerText = ""
      }
      if name.namespaceURI == Self.spreadsheetDrawingNamespace,
        name.localName == "from" || name.localName == "to"
      {
        markerTarget = nil
      }
      if builders[builders.count - 1].textDepth == depth,
        name.namespaceURI == Self.drawingNamespace,
        name.localName == "t"
      {
        builders[builders.count - 1].textDepth = nil
      }
      if builders[builders.count - 1].transformDepth == depth, name.localName == "xfrm" {
        builders[builders.count - 1].transformDepth = nil
      }
      depth -= 1

    case .startDocument, .endDocument:
      break
    }
  }

  func result() throws -> [OfficeWorksheetDrawingElement] {
    try builders.enumerated().map { index, builder in
      let anchor = try makeAnchor(from: builder)
      let attachments = try builder.relationshipIDs.map { identifier in
        guard let relationship = relationshipsByID[identifier] else {
          throw OfficeKitError.invalidPackage(
            "Drawing object references missing relationship \(identifier.rawValue)."
          )
        }
        return package.attachment(referencedBy: relationship)
      }
      let graphicContent: OfficeGraphicContent?
      if let chartAttachment = attachments.first(where: {
        $0.relationship.type.isEquivalent(to: .chart)
      }), let chartPart = chartAttachment.part {
        graphicContent = .chart(
          OfficeChartReference(package: package, attachment: chartAttachment, part: chartPart)
        )
      } else if builder.graphicDataURI == Self.diagramGraphicURI
        || builder.graphicDataURI == Self.strictDiagramGraphicURI
      {
        graphicContent = .diagram(OfficeDiagramReference(attachments: attachments))
      } else if builder.graphicDataURI == Self.model3DGraphicURI {
        graphicContent = .model3D(OfficeModel3DReference(attachments: attachments))
      } else if let uri = builder.graphicDataURI {
        graphicContent = .related(uri: uri)
      } else {
        graphicContent = nil
      }
      let images = attachments.filter {
        $0.relationship.type.isEquivalent(to: .image)
          || $0.relationship.type.isEquivalent(to: .highDefinitionPhoto)
      }
      let picture =
        builder.elementKind == .picture
        ? OfficePicture(primaryImage: images.first, images: images, cropRectangle: nil)
        : nil
      return OfficeWorksheetDrawingElement(
        kind: builder.elementKind ?? .shape,
        identifier: builder.identifier,
        name: builder.name,
        alternativeText: builder.alternativeText,
        text: builder.text,
        anchor: anchor,
        graphicContent: graphicContent,
        picture: picture,
        spatialInfo: spatialInfo(for: builder, anchor: anchor, zIndex: index),
        attachments: attachments
      )
    }
  }

  private func assignMarkerValue(name: String, text: String) {
    guard let markerTarget else { return }
    switch markerTarget {
    case .from:
      assignMarkerValue(name: name, text: text, marker: &builders[builders.count - 1].from)
    case .to: assignMarkerValue(name: name, text: text, marker: &builders[builders.count - 1].to)
    }
  }

  private func assignMarkerValue(name: String, text: String, marker: inout RawMarker) {
    switch name {
    case "col": marker.columnIndex = Int(text)
    case "colOff": marker.columnOffset = Int64(text)
    case "row": marker.rowIndex = Int(text)
    case "rowOff": marker.rowOffset = Int64(text)
    default: break
    }
  }

  private func makeAnchor(from builder: Builder) throws -> OfficeWorksheetDrawingAnchor {
    switch builder.kind {
    case .twoCell(let editBehavior):
      guard let from = builder.from.value, let to = builder.to.value else {
        throw OfficeKitError.invalidPackage("Two-cell drawing anchor is incomplete.")
      }
      return .twoCell(from: from, to: to, editBehavior: editBehavior)
    case .oneCell:
      guard let from = builder.from.value,
        let width = builder.extentWidth,
        let height = builder.extentHeight else {
        throw OfficeKitError.invalidPackage("One-cell drawing anchor is incomplete.")
      }
      return .oneCell(
        from: from,
        width: OfficeLength(emu: width),
        height: OfficeLength(emu: height)
      )
    case .absolute:
      guard let x = builder.positionX,
        let y = builder.positionY,
        let width = builder.extentWidth,
        let height = builder.extentHeight else {
        throw OfficeKitError.invalidPackage("Absolute drawing anchor is incomplete.")
      }
      return .absolute(
        x: OfficeLength(emu: x),
        y: OfficeLength(emu: y),
        width: OfficeLength(emu: width),
        height: OfficeLength(emu: height)
      )
    }
  }

  private func spatialInfo(
    for builder: Builder,
    anchor: OfficeWorksheetDrawingAnchor,
    zIndex: Int
  ) -> OfficeSpatialInfo {
    let transform: OfficeDrawingTransform?
    if let value = builder.transform.value {
      transform = value
    } else if case .absolute(let x, let y, let width, let height) = anchor {
      transform = OfficeDrawingTransform(x: x, y: y, width: width, height: height)
    } else {
      transform = nil
    }
    guard let transform else {
      return OfficeSpatialInfo(
        coordinateSpace: .worksheet,
        geometrySourcePart: part.name,
        frame: nil,
        zIndex: zIndex,
        resolution: .unresolved(
          reason: "Cell-anchor bounds require worksheet row and column layout metrics."
        )
      )
    }
    return OfficeSpatialInfo(
      coordinateSpace: .worksheet,
      sourceTransform: transform,
      geometrySourcePart: part.name,
      frame: OfficeRect(
        x: transform.x.points,
        y: transform.y.points,
        width: transform.width.points,
        height: transform.height.points
      ),
      rotation: transform.rotationRadians,
      isFlippedHorizontally: transform.isFlippedHorizontally,
      isFlippedVertically: transform.isFlippedVertically,
      zIndex: zIndex,
      resolution: .exact
    )
  }
}

private func drawingAttribute(
  _ localName: String,
  in attributes: [OfficeXMLAttribute]
) -> String? {
  attributes.first { $0.name.localName == localName && $0.name.namespaceURI == nil }?.value
}
