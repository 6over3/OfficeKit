private struct PlaceholderGeometry: Sendable {
  let placeholder: OfficePlaceholder
  let transform: OfficeDrawingTransform
  let sourcePart: OfficePartName
}

package enum PlaceholderGeometryResolver {
  package struct ResolvedGeometry: Sendable {
    fileprivate let layout: [PlaceholderGeometry]
    fileprivate let master: [PlaceholderGeometry]
    package let layoutPart: OfficePart?
    package let masterPart: OfficePart?

    package static let empty = ResolvedGeometry(
      layout: [],
      master: [],
      layoutPart: nil,
      masterPart: nil
    )

    package var masterOnly: ResolvedGeometry {
      ResolvedGeometry(
        layout: [],
        master: master,
        layoutPart: nil,
        masterPart: masterPart
      )
    }

    package func geometry(matching placeholder: OfficePlaceholder) -> (
      transform: OfficeDrawingTransform,
      sourcePart: OfficePartName
    )? {
      guard
        let geometry = bestMatch(for: placeholder, in: layout)
          ?? bestMatch(for: placeholder, in: master) else { return nil }
      return (geometry.transform, geometry.sourcePart)
    }

    private func bestMatch(
      for placeholder: OfficePlaceholder,
      in geometries: [PlaceholderGeometry]
    ) -> PlaceholderGeometry? {
      if let exact = geometries.first(where: {
        $0.placeholder.index == placeholder.index && $0.placeholder.type == placeholder.type
      }) {
        return exact
      }
      if let type = placeholder.type,
        let typed = geometries.first(where: { $0.placeholder.type == type })
      {
        return typed
      }
      return geometries.first { $0.placeholder.index == placeholder.index }
    }
  }

  package static func resolve(
    for slidePart: OfficePart,
    package: OfficePackage
  ) throws -> ResolvedGeometry {
    guard
      let layoutPart = try relatedPart(
        ofType: .slideLayout,
        from: slidePart,
        package: package
      ) else {
      return .empty
    }

    let layout = try parsePlaceholders(in: layoutPart, package: package)
    let masterPart = try relatedPart(
      ofType: .slideMaster,
      from: layoutPart,
      package: package
    )
    let master = try masterPart.map { try parsePlaceholders(in: $0, package: package) } ?? []
    return ResolvedGeometry(
      layout: layout,
      master: master,
      layoutPart: layoutPart,
      masterPart: masterPart
    )
  }

  package static func resolveForNotes(
    notesPart: OfficePart,
    package: OfficePackage
  ) throws -> ResolvedGeometry {
    guard
      let masterPart = try relatedPart(
        ofType: .notesMaster,
        from: notesPart,
        package: package
      ) else { return .empty }
    return ResolvedGeometry(
      layout: [],
      master: try parsePlaceholders(in: masterPart, package: package),
      layoutPart: nil,
      masterPart: masterPart
    )
  }

  private static func relatedPart(
    ofType type: OfficeRelationshipType,
    from sourcePart: OfficePart,
    package: OfficePackage
  ) throws -> OfficePart? {
    let relationship = try package.relationships(from: .part(sourcePart.name))
      .first { $0.type.isEquivalent(to: type) }
    guard let relationship else { return nil }
    guard let part = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }
    return part
  }

  private static func parsePlaceholders(
    in part: OfficePart,
    package: OfficePackage
  ) throws -> [PlaceholderGeometry] {
    let session = PlaceholderParsingSession(sourcePart: part.name)
    try package.parseXML(in: part, compatibility: .commonOffice, session.consume)
    return session.geometries
  }
}

private final class PlaceholderParsingSession {
  private static let presentationNamespace =
    "http://schemas.openxmlformats.org/presentationml/2006/main"
  private static let drawingNamespace =
    "http://schemas.openxmlformats.org/drawingml/2006/main"

  struct RawTransform {
    var x: Int64?
    var y: Int64?
    var width: Int64?
    var height: Int64?
    var rotation: Int64 = 0
    var flipH = false
    var flipV = false

    var value: OfficeDrawingTransform? {
      guard let x, let y, let width, let height else { return nil }
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

  struct Builder {
    let startDepth: Int
    var placeholder: OfficePlaceholder?
    var transform = RawTransform()
    var transformDepth: Int?
  }

  let sourcePart: OfficePartName
  var depth = 0
  var builders: [Builder] = []
  var geometries: [PlaceholderGeometry] = []

  init(sourcePart: OfficePartName) {
    self.sourcePart = sourcePart
  }

  func consume(_ event: OfficeXMLEvent) {
    switch event {
    case .startElement(let name, let attributes, _, _):
      depth += 1
      if name.namespaceURI == Self.presentationNamespace, name.localName == "sp" {
        builders.append(Builder(startDepth: depth))
      }
      guard !builders.isEmpty else { return }

      updateCurrent { builder in
        if name.namespaceURI == Self.presentationNamespace, name.localName == "ph" {
          builder.placeholder = OfficePlaceholder(
            type: attribute("type", in: attributes),
            index: attribute("idx", in: attributes).flatMap(UInt32.init) ?? 0
          )
        }
        if name.namespaceURI == Self.drawingNamespace, name.localName == "xfrm",
          builder.transformDepth == nil
        {
          builder.transformDepth = depth
          builder.transform.rotation = attribute("rot", in: attributes).flatMap(Int64.init) ?? 0
          builder.transform.flipH =
            attribute("flipH", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false
          builder.transform.flipV =
            attribute("flipV", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false
        } else if builder.transformDepth != nil, name.namespaceURI == Self.drawingNamespace {
          switch name.localName {
          case "off":
            builder.transform.x = attribute("x", in: attributes).flatMap(Int64.init)
            builder.transform.y = attribute("y", in: attributes).flatMap(Int64.init)
          case "ext":
            builder.transform.width = attribute("cx", in: attributes).flatMap(Int64.init)
            builder.transform.height = attribute("cy", in: attributes).flatMap(Int64.init)
          default:
            break
          }
        }
      }

    case .endElement(let name, _):
      if !builders.isEmpty {
        updateCurrent { builder in
          if builder.transformDepth == depth, name.namespaceURI == Self.drawingNamespace,
            name.localName == "xfrm"
          {
            builder.transformDepth = nil
          }
        }
        if builders.last?.startDepth == depth, name.namespaceURI == Self.presentationNamespace,
          name.localName == "sp"
        {
          let builder = builders.removeLast()
          if let placeholder = builder.placeholder, let transform = builder.transform.value {
            geometries.append(
              PlaceholderGeometry(
                placeholder: placeholder,
                transform: transform,
                sourcePart: sourcePart
              )
            )
          }
        }
      }
      depth -= 1

    case .startDocument, .text, .endDocument:
      break
    }
  }

  private func updateCurrent(_ update: (inout Builder) -> Void) {
    guard !builders.isEmpty else { return }
    update(&builders[builders.count - 1])
  }

  private func attribute(
    _ localName: String,
    in attributes: [OfficeXMLAttribute]
  ) -> String? {
    attributes.first { $0.name.namespaceURI == nil && $0.name.localName == localName }?.value
  }
}
