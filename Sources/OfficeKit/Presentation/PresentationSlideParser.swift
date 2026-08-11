import Foundation

package enum PresentationSlideParser {
  static func parse(
    reference: OfficeSlideReference,
    package: OfficePackage,
    commentAuthors: [OfficePresentationCommentAuthor]
  ) throws -> OfficeSlide {
    let inheritedGeometry = try PlaceholderGeometryResolver.resolve(
      for: reference.part,
      package: package
    )
    let masterSession = try inheritedGeometry.masterPart.map { part in
      try parseSession(
        part: part,
        coordinateSpace: .slideMaster,
        inheritedGeometry: .empty,
        package: package
      )
    }
    let layoutSession = try inheritedGeometry.layoutPart.map { part in
      try parseSession(
        part: part,
        coordinateSpace: .slideLayout,
        inheritedGeometry: inheritedGeometry.masterOnly,
        package: package
      )
    }
    let slideSession = try parseSession(
      part: reference.part,
      coordinateSpace: .slide,
      inheritedGeometry: inheritedGeometry,
      package: package
    )
    let masterIsVisible =
      slideSession.showsMasterShapes
      && (layoutSession?.showsMasterShapes ?? true)
    let notes = try parseNotes(for: reference.part, package: package)
    let authorsByID = Dictionary(
      commentAuthors.map { ($0.identifier, $0) },
      uniquingKeysWith: { first, _ in first }
    )
    let comments = try PresentationCommentParser.comments(
      slidePart: reference.part,
      authorsByID: authorsByID,
      package: package
    )
    let theme = try inheritedGeometry.masterPart.flatMap {
      try PresentationThemeParser.parse(masterPart: $0, package: package)
    }
    return OfficeSlide(
      reference: reference,
      isHidden: slideSession.isHidden,
      transition: slideSession.transition,
      timing: slideSession.timing,
      background: slideSession.background,
      theme: theme,
      elements: try slideSession.result(),
      layoutLayer: try layoutSession.map { session in
        OfficeSlideLayer(
          part: session.sourcePart,
          isVisible: true,
          background: session.background,
          elements: try session.result(),
          relatedAttachments: try payloadAttachments(for: session.sourcePart, package: package)
        )
      },
      masterLayer: try masterSession.map { session in
        OfficeSlideLayer(
          part: session.sourcePart,
          isVisible: masterIsVisible,
          background: session.background,
          elements: try session.result(),
          relatedAttachments: try payloadAttachments(for: session.sourcePart, package: package)
        )
      },
      notes: notes,
      comments: comments,
      relatedAttachments: try payloadAttachments(for: reference.part, package: package)
    )
  }

  private static func parseNotes(
    for slidePart: OfficePart,
    package: OfficePackage
  ) throws -> OfficeSlideNotes? {
    guard
      let relationship = try package.relationships(
        from: .part(slidePart.name),
        ofType: .notesSlide
      ).first else { return nil }
    guard let notesPart = package.part(referencedBy: relationship) else {
      throw OfficeKitError.missingPart(relationship.rawTarget)
    }

    let inheritedGeometry = try PlaceholderGeometryResolver.resolveForNotes(
      notesPart: notesPart,
      package: package
    )
    let masterSession = try inheritedGeometry.masterPart.map { part in
      try parseSession(
        part: part,
        coordinateSpace: .notesMaster,
        inheritedGeometry: .empty,
        package: package
      )
    }
    let notesSession = try parseSession(
      part: notesPart,
      coordinateSpace: .notesPage,
      inheritedGeometry: inheritedGeometry.masterOnly,
      package: package
    )
    return OfficeSlideNotes(
      part: notesPart,
      elements: try notesSession.result(),
      masterLayer: try masterSession.map { session in
        OfficeSlideLayer(
          part: session.sourcePart,
          isVisible: notesSession.showsMasterShapes,
          background: session.background,
          elements: try session.result(),
          relatedAttachments: try payloadAttachments(for: session.sourcePart, package: package)
        )
      },
      relatedAttachments: try payloadAttachments(for: notesPart, package: package),
      showsMasterPlaceholderAnimations: notesSession.showsMasterPlaceholderAnimations
    )
  }

  private static func payloadAttachments(
    for part: OfficePart,
    package: OfficePackage
  ) throws -> [OfficeAttachment] {
    try package.relationships(from: .part(part.name))
      .filter { !isNavigationRelationship($0.type) }
      .map(package.attachment(referencedBy:))
  }

  private static func isNavigationRelationship(_ type: OfficeRelationshipType) -> Bool {
    [
      .slide,
      .slideLayout,
      .slideMaster,
      .notesSlide,
      .notesMaster,
      .comments,
    ].contains { type.isEquivalent(to: $0) }
  }

  private static func parseSession(
    part: OfficePart,
    coordinateSpace: OfficeCoordinateSpace,
    inheritedGeometry: PlaceholderGeometryResolver.ResolvedGeometry,
    package: OfficePackage
  ) throws -> SlideParsingSession {
    let session = SlideParsingSession(
      sourcePart: part,
      coordinateSpace: coordinateSpace,
      package: package,
      inheritedGeometry: inheritedGeometry
    )
    try package.parseXML(in: part, compatibility: .commonOffice, session.consume)
    return session
  }
}

private final class SlideParsingSession {
  private static let presentationNamespace =
    "http://schemas.openxmlformats.org/presentationml/2006/main"
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

  struct RawTransform {
    var x: Int64?
    var y: Int64?
    var width: Int64?
    var height: Int64?
    var childX: Int64?
    var childY: Int64?
    var childWidth: Int64?
    var childHeight: Int64?
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
        childX: childX.map(OfficeLength.init(emu:)),
        childY: childY.map(OfficeLength.init(emu:)),
        childWidth: childWidth.map(OfficeLength.init(emu:)),
        childHeight: childHeight.map(OfficeLength.init(emu:)),
        rotationUnits: rotation,
        isFlippedHorizontally: flipH,
        isFlippedVertically: flipV
      )
    }
  }

  struct RawTableCell {
    let columnIndex: Int
    let columnSpan: Int
    let rowSpan: Int
    let isHorizontallyMerged: Bool
    let isVerticallyMerged: Bool
    var text = ""
  }

  struct RawTableRow {
    let height: Int64
    var cells: [RawTableCell] = []
  }

  struct RawTable {
    var columnWidths: [Int64] = []
    var rows: [RawTableRow] = []
    var styleIdentifier = ""
    var rowDepth: Int?
    var cellDepth: Int?
    var styleTextDepth: Int?
    var activeRowIndex: Int?
    var activeCellIndex: Int?
  }

  struct RawCropRectangle {
    let left: Int64
    let top: Int64
    let right: Int64
    let bottom: Int64
  }

  struct RawTextRunProperties {
    var language: String?
    var alternativeLanguage: String?
    var fontSizeInPoints: Double?
    var isBold: Bool?
    var isItalic: Bool?
    var underline: String?
    var strike: String?
    var capitalization: String?
    var baseline: Double?
    var latinTypeface: String?
    var eastAsianTypeface: String?
    var complexScriptTypeface: String?
    var color: OfficeDrawingColor?

    var value: OfficeSlideTextRunProperties {
      OfficeSlideTextRunProperties(
        language: language,
        alternativeLanguage: alternativeLanguage,
        fontSizeInPoints: fontSizeInPoints,
        isBold: isBold,
        isItalic: isItalic,
        underline: underline,
        strike: strike,
        capitalization: capitalization,
        baseline: baseline,
        latinTypeface: latinTypeface,
        eastAsianTypeface: eastAsianTypeface,
        complexScriptTypeface: complexScriptTypeface,
        color: color
      )
    }
  }

  struct RawTextRun {
    let startDepth: Int
    let fieldIdentifier: String?
    let fieldType: String?
    var text = ""
    var properties = RawTextRunProperties()
    var propertiesDepth: Int?
  }

  struct RawTextParagraph {
    let startDepth: Int
    var level = 0
    var alignment: String?
    var leadingMargin: OfficeLength?
    var indent: OfficeLength?
    var isRightToLeft: Bool?
    var bullet: OfficeSlideBullet?
    var runs: [RawTextRun] = []
    var activeRun: RawTextRun?
  }

  struct RawLine {
    var width: OfficeLength?
    var fill: OfficeDrawingFill?
    var dash: String?
    var cap: String?
    var compound: String?
    var alignment: String?

    var value: OfficeDrawingLine {
      OfficeDrawingLine(
        width: width,
        fill: fill,
        dash: dash,
        cap: cap,
        compound: compound,
        alignment: alignment
      )
    }
  }

  struct Builder {
    let kind: OfficeSlideElementKind
    let startDepth: Int
    var identifier: UInt32?
    var name: String?
    var alternativeText: String?
    var title: String?
    var text = ""
    var relationshipIDs: [OfficeRelationshipID] = []
    var placeholder: OfficePlaceholder?
    var graphicDataURI: String?
    var table: RawTable?
    var cropRectangle: RawCropRectangle?
    var transform = RawTransform()
    var transformDepth: Int?
    var hasReadTransform = false
    var textDepth: Int?
    var textParagraphs: [RawTextParagraph] = []
    var activeParagraph: RawTextParagraph?
    var shapePropertiesDepth: Int?
    var geometry: OfficeDrawingGeometry?
    var customGeometryDepth: Int?
    var customPathCount = 0
    var customGuideCount = 0
    var fill: OfficeDrawingFill?
    var solidFillDepth: Int?
    var line: RawLine?
    var lineDepth: Int?
    var lineSolidFillDepth: Int?
    var children: [Builder] = []
  }

  let sourcePart: OfficePart
  let coordinateSpace: OfficeCoordinateSpace
  let package: OfficePackage
  let inheritedGeometry: PlaceholderGeometryResolver.ResolvedGeometry
  var depth = 0
  var shapeTreeDepth: Int?
  var builders: [Builder] = []
  var roots: [Builder] = []
  var isHidden = false
  var showsMasterShapes = true
  var showsMasterPlaceholderAnimations = true
  var transition: OfficeSlideTransition?
  var transitionDepth: Int?
  var timingDepth: Int?
  var timingNodeCount = 0
  var timingBehaviorCount = 0
  var background: OfficeSlideBackground?
  var backgroundDepth: Int?
  var backgroundSolidFillDepth: Int?
  var timing: OfficeSlideTiming? {
    guard timingNodeCount != 0 || timingBehaviorCount != 0 else { return nil }
    return OfficeSlideTiming(
      timeNodeCount: timingNodeCount,
      behaviorCount: timingBehaviorCount
    )
  }

  init(
    sourcePart: OfficePart,
    coordinateSpace: OfficeCoordinateSpace,
    package: OfficePackage,
    inheritedGeometry: PlaceholderGeometryResolver.ResolvedGeometry
  ) {
    self.sourcePart = sourcePart
    self.coordinateSpace = coordinateSpace
    self.package = package
    self.inheritedGeometry = inheritedGeometry
  }

  func consume(_ event: OfficeXMLEvent) throws {
    switch event {
    case .startElement(let name, let attributes, _, _):
      depth += 1
      if name.namespaceURI == Self.presentationNamespace, name.localName == "sld" {
        if let show = attribute("show", in: attributes),
          let isShown = OfficeValueDecoder.boolean(show)
        {
          isHidden = !isShown
        }
      }
      if name.namespaceURI == Self.presentationNamespace,
        name.localName == "sld" || name.localName == "sldLayout" || name.localName == "notes"
      {
        showsMasterShapes =
          attribute("showMasterSp", in: attributes)
          .flatMap(OfficeValueDecoder.boolean) ?? true
      }
      if name.namespaceURI == Self.presentationNamespace, name.localName == "notes" {
        showsMasterPlaceholderAnimations =
          attribute("showMasterPhAnim", in: attributes)
          .flatMap(OfficeValueDecoder.boolean) ?? true
      }
      if name.namespaceURI == Self.presentationNamespace, name.localName == "transition" {
        transitionDepth = depth
        transition = OfficeSlideTransition(
          kind: nil,
          speed: attribute("spd", in: attributes),
          durationMilliseconds: attribute("dur", in: attributes).flatMap(UInt32.init),
          advancesOnClick: attribute("advClick", in: attributes)
            .flatMap(OfficeValueDecoder.boolean),
          advanceAfterMilliseconds: attribute("advTm", in: attributes).flatMap(UInt32.init)
        )
      } else if transitionDepth != nil,
        name.namespaceURI == Self.presentationNamespace,
        name.localName != "sndAc" && name.localName != "stSnd" && name.localName != "endSnd"
      {
        transition = OfficeSlideTransition(
          kind: name.localName,
          speed: transition?.speed,
          durationMilliseconds: transition?.durationMilliseconds,
          advancesOnClick: transition?.advancesOnClick,
          advanceAfterMilliseconds: transition?.advanceAfterMilliseconds
        )
      }
      if name.namespaceURI == Self.presentationNamespace, name.localName == "timing" {
        timingDepth = depth
      } else if timingDepth != nil, name.namespaceURI == Self.presentationNamespace {
        if name.localName == "cTn" { timingNodeCount += 1 }
        if ["anim", "animClr", "animEffect", "animMotion", "animRot", "animScale", "cmd", "set"]
          .contains(name.localName)
        {
          timingBehaviorCount += 1
        }
      }
      if name.namespaceURI == Self.presentationNamespace, name.localName == "bg" {
        backgroundDepth = depth
        background = OfficeSlideBackground(fill: nil, styleIndex: nil)
      } else if backgroundDepth != nil,
        name.namespaceURI == Self.presentationNamespace,
        name.localName == "bgRef"
      {
        background = OfficeSlideBackground(
          fill: background?.fill,
          styleIndex: attribute("idx", in: attributes).flatMap(UInt32.init)
        )
      } else if backgroundDepth != nil, name.namespaceURI == Self.drawingNamespace {
        readBackgroundStart(name: name, attributes: attributes)
      }
      if name.namespaceURI == Self.presentationNamespace, name.localName == "spTree" {
        shapeTreeDepth = depth
        return
      }
      if shapeTreeDepth != nil, let kind = elementKind(for: name) {
        builders.append(Builder(kind: kind, startDepth: depth))
      }
      guard !builders.isEmpty else { return }

      updateCurrent { builder in
        if name.namespaceURI == Self.presentationNamespace, name.localName == "cNvPr",
          builder.identifier == nil
        {
          builder.identifier = attribute("id", in: attributes).flatMap(UInt32.init)
          builder.name = attribute("name", in: attributes)
          builder.alternativeText = attribute("descr", in: attributes)
          builder.title = attribute("title", in: attributes)
        }

        if name.namespaceURI == Self.presentationNamespace, name.localName == "ph" {
          builder.placeholder = OfficePlaceholder(
            type: attribute("type", in: attributes),
            index: attribute("idx", in: attributes).flatMap(UInt32.init) ?? 0
          )
        }

        if name.namespaceURI == Self.presentationNamespace,
          name.localName == "spPr" || name.localName == "grpSpPr"
        {
          builder.shapePropertiesDepth = depth
        }
        if builder.shapePropertiesDepth != nil, name.namespaceURI == Self.drawingNamespace {
          readShapeStyleStart(name: name, attributes: attributes, builder: &builder)
        }

        readGraphicContentStart(name: name, attributes: attributes, builder: &builder)

        if name.namespaceURI == Self.drawingNamespace {
          switch name.localName {
          case "p":
            builder.activeParagraph = RawTextParagraph(startDepth: depth)
          case "pPr":
            if builder.activeParagraph != nil {
              builder.activeParagraph?.level =
                attribute("lvl", in: attributes).flatMap(Int.init) ?? 0
              builder.activeParagraph?.alignment = attribute("algn", in: attributes)
              builder.activeParagraph?.leadingMargin = attribute("marL", in: attributes)
                .flatMap(Int64.init).map(OfficeLength.init(emu:))
              builder.activeParagraph?.indent = attribute("indent", in: attributes)
                .flatMap(Int64.init).map(OfficeLength.init(emu:))
              builder.activeParagraph?.isRightToLeft = attribute("rtl", in: attributes)
                .flatMap(OfficeValueDecoder.boolean)
            }
          case "buNone": builder.activeParagraph?.bullet = OfficeSlideBullet.none
          case "buChar":
            if let character = attribute("char", in: attributes) {
              builder.activeParagraph?.bullet = .character(character)
            }
          case "buAutoNum":
            if let numbering = attribute("type", in: attributes) {
              builder.activeParagraph?.bullet = .automatic(
                numbering: numbering,
                startAt: attribute("startAt", in: attributes).flatMap(Int.init)
              )
            }
          case "r", "fld":
            if builder.activeParagraph != nil {
              builder.activeParagraph?.activeRun = RawTextRun(
                startDepth: depth,
                fieldIdentifier: name.localName == "fld" ? attribute("id", in: attributes) : nil,
                fieldType: name.localName == "fld" ? attribute("type", in: attributes) : nil
              )
            }
          case "rPr":
            if var activeRun = builder.activeParagraph?.activeRun {
              activeRun.propertiesDepth = depth
              readTextRunProperties(attributes, into: &activeRun)
              builder.activeParagraph?.activeRun = activeRun
            }
          case "latin":
            if builder.activeParagraph?.activeRun?.propertiesDepth != nil {
              builder.activeParagraph?.activeRun?.properties.latinTypeface = attribute(
                "typeface", in: attributes)
            }
          case "ea":
            if builder.activeParagraph?.activeRun?.propertiesDepth != nil {
              builder.activeParagraph?.activeRun?.properties.eastAsianTypeface = attribute(
                "typeface", in: attributes)
            }
          case "cs":
            if builder.activeParagraph?.activeRun?.propertiesDepth != nil {
              builder.activeParagraph?.activeRun?.properties.complexScriptTypeface = attribute(
                "typeface", in: attributes)
            }
          case "srgbClr":
            if builder.activeParagraph?.activeRun?.propertiesDepth != nil,
              let value = attribute("val", in: attributes)
            {
              builder.activeParagraph?.activeRun?.properties.color = .sRGB(value)
            }
          case "schemeClr":
            if builder.activeParagraph?.activeRun?.propertiesDepth != nil,
              let value = attribute("val", in: attributes)
            {
              builder.activeParagraph?.activeRun?.properties.color = .scheme(value)
            }
          case "sysClr":
            if builder.activeParagraph?.activeRun?.propertiesDepth != nil,
              let value = attribute("val", in: attributes)
            {
              builder.activeParagraph?.activeRun?.properties.color = .system(
                name: value,
                lastColor: attribute("lastClr", in: attributes)
              )
            }
          case "prstClr":
            if builder.activeParagraph?.activeRun?.propertiesDepth != nil,
              let value = attribute("val", in: attributes)
            {
              builder.activeParagraph?.activeRun?.properties.color = .preset(value)
            }
          default: break
          }
        }

        if builder.kind == .picture, name.namespaceURI == Self.drawingNamespace,
          name.localName == "srcRect"
        {
          builder.cropRectangle = RawCropRectangle(
            left: attribute("l", in: attributes).flatMap(Int64.init) ?? 0,
            top: attribute("t", in: attributes).flatMap(Int64.init) ?? 0,
            right: attribute("r", in: attributes).flatMap(Int64.init) ?? 0,
            bottom: attribute("b", in: attributes).flatMap(Int64.init) ?? 0
          )
        }

        if name.localName == "xfrm",
          name.namespaceURI == Self.drawingNamespace
            || name.namespaceURI == Self.presentationNamespace,
          !builder.hasReadTransform
        {
          builder.hasReadTransform = true
          builder.transformDepth = depth
          builder.transform.rotation = attribute("rot", in: attributes).flatMap(Int64.init) ?? 0
          builder.transform.flipH =
            attribute("flipH", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false
          builder.transform.flipV =
            attribute("flipV", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false
        } else if builder.transformDepth != nil, name.namespaceURI == Self.drawingNamespace {
          readTransformChild(name: name, attributes: attributes, transform: &builder.transform)
        }

        if name.namespaceURI == Self.drawingNamespace, name.localName == "t" {
          builder.textDepth = depth
        } else if name.namespaceURI == Self.drawingNamespace, name.localName == "br" {
          builder.text.append("\n")
        }

        for attribute in attributes
        where attribute.name.namespaceURI == Self.relationshipNamespace
          && !attribute.value.isEmpty
        {
          let id = OfficeRelationshipID(rawValue: attribute.value)
          if !builder.relationshipIDs.contains(id) { builder.relationshipIDs.append(id) }
        }
      }

    case .text(let text, _):
      guard builders.last?.textDepth != nil || builders.last?.table?.styleTextDepth != nil else {
        return
      }
      updateCurrent { builder in
        if builder.textDepth != nil {
          builder.text.append(text)
          builder.activeParagraph?.activeRun?.text.append(text)
          if let rowIndex = builder.table?.activeRowIndex,
            let cellIndex = builder.table?.activeCellIndex
          {
            builder.table?.rows[rowIndex].cells[cellIndex].text.append(text)
          }
        }
        if builder.table?.styleTextDepth != nil {
          builder.table?.styleIdentifier.append(text)
        }
      }

    case .endElement(let name, _):
      if !builders.isEmpty {
        updateCurrent { builder in
          if builder.textDepth == depth, name.namespaceURI == Self.drawingNamespace,
            name.localName == "t"
          {
            builder.textDepth = nil
          }
          if name.namespaceURI == Self.drawingNamespace, name.localName == "rPr",
            builder.activeParagraph?.activeRun?.propertiesDepth == depth
          {
            builder.activeParagraph?.activeRun?.propertiesDepth = nil
          }
          if name.namespaceURI == Self.drawingNamespace,
            name.localName == "r" || name.localName == "fld",
            builder.activeParagraph?.activeRun?.startDepth == depth,
            let run = builder.activeParagraph?.activeRun
          {
            builder.activeParagraph?.runs.append(run)
            builder.activeParagraph?.activeRun = nil
          }
          if builder.transformDepth == depth, name.localName == "xfrm" {
            builder.transformDepth = nil
          }
          if builder.solidFillDepth == depth, name.namespaceURI == Self.drawingNamespace,
            name.localName == "solidFill"
          {
            builder.solidFillDepth = nil
          }
          if builder.lineSolidFillDepth == depth, name.namespaceURI == Self.drawingNamespace,
            name.localName == "solidFill"
          {
            builder.lineSolidFillDepth = nil
          }
          if builder.lineDepth == depth, name.namespaceURI == Self.drawingNamespace,
            name.localName == "ln"
          {
            builder.lineDepth = nil
          }
          if builder.customGeometryDepth == depth, name.namespaceURI == Self.drawingNamespace,
            name.localName == "custGeom"
          {
            builder.geometry = .custom(
              pathCount: builder.customPathCount,
              guideCount: builder.customGuideCount
            )
            builder.customGeometryDepth = nil
          }
          if builder.shapePropertiesDepth == depth,
            name.namespaceURI == Self.presentationNamespace,
            name.localName == "spPr" || name.localName == "grpSpPr"
          {
            builder.shapePropertiesDepth = nil
          }
          if name.namespaceURI == Self.drawingNamespace, name.localName == "p",
            !builder.text.isEmpty, !builder.text.hasSuffix("\n")
          {
            builder.text.append("\n")
            if let rowIndex = builder.table?.activeRowIndex,
              let cellIndex = builder.table?.activeCellIndex,
              builder.table?.rows[rowIndex].cells[cellIndex].text.isEmpty == false,
              builder.table?.rows[rowIndex].cells[cellIndex].text.hasSuffix("\n") == false
            {
              builder.table?.rows[rowIndex].cells[cellIndex].text.append("\n")
            }
          }
          if name.namespaceURI == Self.drawingNamespace, name.localName == "p",
            builder.activeParagraph?.startDepth == depth,
            let paragraph = builder.activeParagraph
          {
            builder.textParagraphs.append(paragraph)
            builder.activeParagraph = nil
          }
          if builder.table?.cellDepth == depth, name.namespaceURI == Self.drawingNamespace,
            name.localName == "tc"
          {
            builder.table?.cellDepth = nil
            builder.table?.activeCellIndex = nil
          }
          if builder.table?.rowDepth == depth, name.namespaceURI == Self.drawingNamespace,
            name.localName == "tr"
          {
            builder.table?.rowDepth = nil
            builder.table?.activeRowIndex = nil
          }
          if builder.table?.styleTextDepth == depth,
            name.namespaceURI == Self.drawingNamespace,
            name.localName == "tableStyleId"
          {
            builder.table?.styleTextDepth = nil
          }
        }

        if builders.last?.startDepth == depth {
          let completed = builders.removeLast()
          if builders.isEmpty {
            roots.append(completed)
          } else {
            updateCurrent { $0.children.append(completed) }
          }
        }
      }
      if shapeTreeDepth == depth, name.namespaceURI == Self.presentationNamespace,
        name.localName == "spTree"
      {
        shapeTreeDepth = nil
      }
      if transitionDepth == depth, name.namespaceURI == Self.presentationNamespace,
        name.localName == "transition"
      {
        transitionDepth = nil
      }
      if timingDepth == depth, name.namespaceURI == Self.presentationNamespace,
        name.localName == "timing"
      {
        timingDepth = nil
      }
      if backgroundSolidFillDepth == depth, name.namespaceURI == Self.drawingNamespace,
        name.localName == "solidFill"
      {
        backgroundSolidFillDepth = nil
      }
      if backgroundDepth == depth, name.namespaceURI == Self.presentationNamespace,
        name.localName == "bg"
      {
        backgroundDepth = nil
      }
      depth -= 1

    case .startDocument, .endDocument:
      break
    }
  }

  func result() throws -> [OfficeSlideElement] {
    try roots.enumerated().map { index, builder in
      try makeElement(
        from: builder,
        zIndex: index,
        coordinateSpace: coordinateSpace,
        parentTransform: .identity
      )
    }
  }

  private func readTextRunProperties(
    _ attributes: [OfficeXMLAttribute],
    into run: inout RawTextRun
  ) {
    run.properties.language = attribute("lang", in: attributes)
    run.properties.alternativeLanguage = attribute("altLang", in: attributes)
    run.properties.fontSizeInPoints = attribute("sz", in: attributes)
      .flatMap(Double.init).map { $0 / 100 }
    run.properties.isBold = attribute("b", in: attributes)
      .flatMap(OfficeValueDecoder.boolean)
    run.properties.isItalic = attribute("i", in: attributes)
      .flatMap(OfficeValueDecoder.boolean)
    run.properties.underline = attribute("u", in: attributes)
    run.properties.strike = attribute("strike", in: attributes)
    run.properties.capitalization = attribute("cap", in: attributes)
    run.properties.baseline = attribute("baseline", in: attributes)
      .flatMap(Double.init).map { $0 / 1000 }
  }

  private func readBackgroundStart(
    name: OfficeXMLName,
    attributes: [OfficeXMLAttribute]
  ) {
    switch name.localName {
    case "noFill": setBackgroundFill(OfficeDrawingFill.none)
    case "solidFill":
      backgroundSolidFillDepth = depth
      setBackgroundFill(.solid(nil))
    case "gradFill": setBackgroundFill(.gradient)
    case "pattFill": setBackgroundFill(.pattern(attribute("prst", in: attributes)))
    case "blipFill": setBackgroundFill(.picture)
    case "grpFill": setBackgroundFill(.group)
    case "srgbClr" where backgroundSolidFillDepth != nil:
      if let value = attribute("val", in: attributes) {
        setBackgroundFill(.solid(.sRGB(value)))
      }
    case "schemeClr" where backgroundSolidFillDepth != nil:
      if let value = attribute("val", in: attributes) {
        setBackgroundFill(.solid(.scheme(value)))
      }
    case "sysClr" where backgroundSolidFillDepth != nil:
      if let value = attribute("val", in: attributes) {
        setBackgroundFill(
          .solid(
            .system(
              name: value,
              lastColor: attribute("lastClr", in: attributes)
            )))
      }
    case "prstClr" where backgroundSolidFillDepth != nil:
      if let value = attribute("val", in: attributes) {
        setBackgroundFill(.solid(.preset(value)))
      }
    default: break
    }
  }

  private func setBackgroundFill(_ fill: OfficeDrawingFill?) {
    background = OfficeSlideBackground(fill: fill, styleIndex: background?.styleIndex)
  }

  private func readShapeStyleStart(
    name: OfficeXMLName,
    attributes: [OfficeXMLAttribute],
    builder: inout Builder
  ) {
    switch name.localName {
    case "prstGeom":
      if let preset = attribute("prst", in: attributes) { builder.geometry = .preset(preset) }
    case "custGeom":
      builder.customGeometryDepth = depth
      builder.customPathCount = 0
      builder.customGuideCount = 0
    case "path" where builder.customGeometryDepth != nil:
      builder.customPathCount += 1
    case "gd" where builder.customGeometryDepth != nil:
      builder.customGuideCount += 1
    case "ln":
      builder.lineDepth = depth
      builder.line = RawLine(
        width: attribute("w", in: attributes).flatMap(Int64.init).map(OfficeLength.init(emu:)),
        cap: attribute("cap", in: attributes),
        compound: attribute("cmpd", in: attributes),
        alignment: attribute("algn", in: attributes)
      )
    case "prstDash" where builder.lineDepth != nil:
      builder.line?.dash = attribute("val", in: attributes)
    case "noFill":
      if builder.lineDepth != nil {
        builder.line?.fill = OfficeDrawingFill.none
      } else {
        builder.fill = OfficeDrawingFill.none
      }
    case "solidFill":
      if builder.lineDepth != nil {
        builder.line?.fill = .solid(nil)
        builder.lineSolidFillDepth = depth
      } else {
        builder.fill = .solid(nil)
        builder.solidFillDepth = depth
      }
    case "gradFill":
      if builder.lineDepth != nil {
        builder.line?.fill = .gradient
      } else {
        builder.fill = .gradient
      }
    case "pattFill":
      let fill = OfficeDrawingFill.pattern(attribute("prst", in: attributes))
      if builder.lineDepth != nil { builder.line?.fill = fill } else { builder.fill = fill }
    case "blipFill":
      if builder.lineDepth != nil { builder.line?.fill = .picture } else { builder.fill = .picture }
    case "grpFill":
      if builder.lineDepth != nil { builder.line?.fill = .group } else { builder.fill = .group }
    case "srgbClr":
      if let value = attribute("val", in: attributes) {
        updateSolidColor(.sRGB(value), builder: &builder)
      }
    case "schemeClr":
      if let value = attribute("val", in: attributes) {
        updateSolidColor(.scheme(value), builder: &builder)
      }
    case "sysClr":
      if let value = attribute("val", in: attributes) {
        updateSolidColor(
          .system(name: value, lastColor: attribute("lastClr", in: attributes)),
          builder: &builder
        )
      }
    case "prstClr":
      if let value = attribute("val", in: attributes) {
        updateSolidColor(.preset(value), builder: &builder)
      }
    default: break
    }
  }

  private func updateSolidColor(_ color: OfficeDrawingColor, builder: inout Builder) {
    if builder.lineSolidFillDepth != nil {
      builder.line?.fill = .solid(color)
    } else if builder.solidFillDepth != nil {
      builder.fill = .solid(color)
    }
  }

  private func makeElement(
    from builder: Builder,
    zIndex: Int,
    coordinateSpace: OfficeCoordinateSpace,
    parentTransform: OfficeAffineTransform?
  ) throws -> OfficeSlideElement {
    let localSourceTransform = builder.transform.value
    let inherited = builder.placeholder.flatMap(inheritedGeometry.geometry(matching:))
    let sourceTransform = localSourceTransform ?? inherited?.transform
    let geometrySourcePart =
      localSourceTransform == nil
      ? inherited?.sourcePart
      : sourcePart.name
    let localTransform = sourceTransform.map(transform(from:))
    let documentTransform: OfficeAffineTransform?
    if let localTransform, let parentTransform {
      documentTransform = localTransform.followed(by: parentTransform)
    } else {
      documentTransform = nil
    }

    let spatialInfo: OfficeSpatialInfo
    if let sourceTransform, let localTransform, let documentTransform {
      let localRect: OfficeRect
      if builder.kind == .group,
        let childX = sourceTransform.childX,
        let childY = sourceTransform.childY,
        let childWidth = sourceTransform.childWidth,
        let childHeight = sourceTransform.childHeight
      {
        localRect = OfficeRect(
          x: childX.points,
          y: childY.points,
          width: childWidth.points,
          height: childHeight.points
        )
      } else {
        localRect = OfficeRect(
          x: 0,
          y: 0,
          width: sourceTransform.width.points,
          height: sourceTransform.height.points
        )
      }
      spatialInfo = OfficeSpatialInfo(
        coordinateSpace: coordinateSpace,
        sourceTransform: sourceTransform,
        geometrySourcePart: geometrySourcePart,
        frame: localRect.applying(documentTransform),
        transformToParent: localTransform,
        transformToDocument: documentTransform,
        rotation: sourceTransform.rotationRadians,
        isFlippedHorizontally: sourceTransform.isFlippedHorizontally,
        isFlippedVertically: sourceTransform.isFlippedVertically,
        zIndex: zIndex,
        resolution: coordinateSpace != .group && localSourceTransform != nil ? .exact : .derived
      )
    } else {
      spatialInfo = OfficeSpatialInfo(
        coordinateSpace: coordinateSpace,
        sourceTransform: sourceTransform,
        geometrySourcePart: geometrySourcePart,
        frame: nil,
        transformToParent: localTransform,
        transformToDocument: nil,
        rotation: sourceTransform?.rotationRadians ?? 0,
        isFlippedHorizontally: sourceTransform?.isFlippedHorizontally ?? false,
        isFlippedVertically: sourceTransform?.isFlippedVertically ?? false,
        zIndex: zIndex,
        resolution: .unresolved(
          reason: "No local transform is declared; layout inheritance is required.")
      )
    }

    let attachments = try builder.relationshipIDs.map { id in
      guard
        let relationship = try package.relationship(
          identifiedBy: id,
          from: .part(sourcePart.name)
        ) else {
        throw OfficeKitError.invalidPackage(
          "Drawing element in \(sourcePart.name.rawValue) references missing relationship "
            + "\(id.rawValue)."
        )
      }
      return package.attachment(referencedBy: relationship)
    }
    let childParentTransform = builder.kind == .group ? documentTransform : parentTransform
    let children = try builder.children.enumerated().map { index, child in
      try makeElement(
        from: child,
        zIndex: index,
        coordinateSpace: .group,
        parentTransform: childParentTransform
      )
    }
    let graphicContent: OfficeGraphicContent?
    if let table = builder.table {
      graphicContent = .table(makeTable(from: table, documentTransform: documentTransform))
    } else if let chartAttachment = attachments.first(where: {
      $0.relationship.type.isEquivalent(to: .chart)
    }), let chartPart = chartAttachment.part {
      graphicContent = .chart(
        OfficeChartReference(
          package: package,
          attachment: chartAttachment,
          part: chartPart
        )
      )
    } else if builder.graphicDataURI == Self.diagramGraphicURI
      || builder.graphicDataURI == Self.strictDiagramGraphicURI
    {
      graphicContent = .diagram(OfficeDiagramReference(attachments: attachments))
    } else if builder.graphicDataURI == Self.model3DGraphicURI {
      graphicContent = .model3D(OfficeModel3DReference(attachments: attachments))
    } else if let graphicDataURI = builder.graphicDataURI {
      graphicContent = .related(uri: graphicDataURI)
    } else {
      graphicContent = nil
    }
    let picture: OfficePicture?
    if builder.kind == .picture {
      let images = attachments.filter { $0.relationship.type.isEquivalent(to: .image) }
      picture = OfficePicture(
        primaryImage: images.first,
        images: images,
        cropRectangle: builder.cropRectangle.map {
          OfficeCropRectangle(
            left: OfficePercentage(rawValue: $0.left),
            top: OfficePercentage(rawValue: $0.top),
            right: OfficePercentage(rawValue: $0.right),
            bottom: OfficePercentage(rawValue: $0.bottom)
          )
        }
      )
    } else {
      picture = nil
    }
    return OfficeSlideElement(
      kind: builder.kind,
      identifier: builder.identifier,
      name: builder.name,
      alternativeText: builder.alternativeText,
      title: builder.title,
      text: builder.text.trimmingCharacters(in: .newlines),
      textBody: builder.textParagraphs.isEmpty
        ? nil
        : OfficeSlideTextBody(
          paragraphs: builder.textParagraphs.map { paragraph in
            OfficeSlideTextParagraph(
              level: paragraph.level,
              alignment: paragraph.alignment,
              leadingMargin: paragraph.leadingMargin,
              indent: paragraph.indent,
              isRightToLeft: paragraph.isRightToLeft,
              bullet: paragraph.bullet,
              runs: paragraph.runs.map { run in
                OfficeSlideTextRun(
                  text: run.text,
                  properties: run.properties.value,
                  fieldIdentifier: run.fieldIdentifier,
                  fieldType: run.fieldType
                )
              }
            )
          }
        ),
      placeholder: builder.placeholder,
      graphicContent: graphicContent,
      picture: picture,
      geometry: builder.geometry,
      fill: builder.fill,
      line: builder.line?.value,
      spatialInfo: spatialInfo,
      attachments: attachments,
      children: children
    )
  }

  private func readGraphicContentStart(
    name: OfficeXMLName,
    attributes: [OfficeXMLAttribute],
    builder: inout Builder
  ) {
    guard builder.kind == .graphicFrame, name.namespaceURI == Self.drawingNamespace else {
      return
    }
    switch name.localName {
    case "graphicData":
      builder.graphicDataURI = attribute("uri", in: attributes)
    case "tbl":
      builder.table = RawTable()
    case "gridCol":
      if let width = attribute("w", in: attributes).flatMap(Int64.init) {
        builder.table?.columnWidths.append(width)
      }
    case "tr":
      let rowIndex = builder.table?.rows.count ?? 0
      builder.table?.rows.append(
        RawTableRow(height: attribute("h", in: attributes).flatMap(Int64.init) ?? 0)
      )
      builder.table?.rowDepth = depth
      builder.table?.activeRowIndex = rowIndex
    case "tc":
      guard let rowIndex = builder.table?.activeRowIndex else { return }
      let columnIndex =
        builder.table?.rows[rowIndex].cells.reduce(into: 0) {
          $0 += max($1.columnSpan, 1)
        } ?? 0
      let cellIndex = builder.table?.rows[rowIndex].cells.count ?? 0
      builder.table?.rows[rowIndex].cells.append(
        RawTableCell(
          columnIndex: columnIndex,
          columnSpan: max(attribute("gridSpan", in: attributes).flatMap(Int.init) ?? 1, 1),
          rowSpan: max(attribute("rowSpan", in: attributes).flatMap(Int.init) ?? 1, 1),
          isHorizontallyMerged: attribute("hMerge", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false,
          isVerticallyMerged: attribute("vMerge", in: attributes)
            .flatMap(OfficeValueDecoder.boolean) ?? false
        )
      )
      builder.table?.cellDepth = depth
      builder.table?.activeCellIndex = cellIndex
    case "tableStyleId":
      builder.table?.styleTextDepth = depth
    default:
      break
    }
  }

  private func makeTable(
    from source: RawTable,
    documentTransform: OfficeAffineTransform?
  ) -> OfficeTable {
    let columnWidths = source.columnWidths.map(OfficeLength.init(emu:))
    let rowHeights = source.rows.map { OfficeLength(emu: $0.height) }
    let dimensionsAreResolvable =
      !columnWidths.isEmpty
      && columnWidths.allSatisfy { $0.emu > 0 }
      && !rowHeights.isEmpty
      && rowHeights.allSatisfy { $0.emu > 0 }

    let rows = source.rows.enumerated().map { rowIndex, row in
      OfficeTableRow(
        index: rowIndex,
        height: OfficeLength(emu: row.height),
        cells: row.cells.map { cell in
          let spatialInfo: OfficeSpatialInfo
          if dimensionsAreResolvable, let documentTransform,
            cell.columnIndex < columnWidths.count,
            cell.columnIndex + cell.columnSpan <= columnWidths.count,
            rowIndex < rowHeights.count,
            rowIndex + cell.rowSpan <= rowHeights.count
          {
            let x = columnWidths.prefix(cell.columnIndex).reduce(0) { $0 + $1.points }
            let y = rowHeights.prefix(rowIndex).reduce(0) { $0 + $1.points }
            let width =
              columnWidths
              .dropFirst(cell.columnIndex)
              .prefix(cell.columnSpan)
              .reduce(0) { $0 + $1.points }
            let height =
              rowHeights
              .dropFirst(rowIndex)
              .prefix(cell.rowSpan)
              .reduce(0) { $0 + $1.points }
            let transformToParent = OfficeAffineTransform.translation(x: x, y: y)
            let transformToDocument = transformToParent.followed(by: documentTransform)
            spatialInfo = OfficeSpatialInfo(
              coordinateSpace: .tableCell,
              frame: OfficeRect(x: 0, y: 0, width: width, height: height)
                .applying(transformToDocument),
              transformToParent: transformToParent,
              transformToDocument: transformToDocument,
              resolution: .derived
            )
          } else {
            spatialInfo = OfficeSpatialInfo(
              coordinateSpace: .tableCell,
              frame: nil,
              resolution: .unresolved(
                reason: "Table cell geometry requires positive authored row and column dimensions."
              )
            )
          }
          return OfficeTableCell(
            rowIndex: rowIndex,
            columnIndex: cell.columnIndex,
            columnSpan: cell.columnSpan,
            rowSpan: cell.rowSpan,
            isHorizontallyMerged: cell.isHorizontallyMerged,
            isVerticallyMerged: cell.isVerticallyMerged,
            text: cell.text.trimmingCharacters(in: .newlines),
            spatialInfo: spatialInfo
          )
        }
      )
    }
    let trimmedStyleIdentifier = source.styleIdentifier.trimmingCharacters(
      in: .whitespacesAndNewlines)
    return OfficeTable(
      columnWidths: columnWidths,
      rows: rows,
      styleIdentifier: trimmedStyleIdentifier.isEmpty ? nil : trimmedStyleIdentifier
    )
  }

  private func transform(from source: OfficeDrawingTransform) -> OfficeAffineTransform {
    let width = source.width.points
    let height = source.height.points
    let centerX = source.x.points + width / 2
    let centerY = source.y.points + height / 2
    let flipAndRotation = OfficeAffineTransform.translation(x: -centerX, y: -centerY)
      .followed(
        by: .scale(
          x: source.isFlippedHorizontally ? -1 : 1,
          y: source.isFlippedVertically ? -1 : 1
        )
      )
      .followed(by: .rotation(radians: source.rotationRadians))
      .followed(by: .translation(x: centerX, y: centerY))

    if let childX = source.childX?.points,
      let childY = source.childY?.points,
      let childWidth = source.childWidth?.points,
      let childHeight = source.childHeight?.points,
      childWidth != 0, childHeight != 0
    {
      return OfficeAffineTransform.translation(x: -childX, y: -childY)
        .followed(by: .scale(x: width / childWidth, y: height / childHeight))
        .followed(by: .translation(x: source.x.points, y: source.y.points))
        .followed(by: flipAndRotation)
    }
    return OfficeAffineTransform.translation(x: source.x.points, y: source.y.points)
      .followed(by: flipAndRotation)
  }

  private func updateCurrent(_ update: (inout Builder) -> Void) {
    guard !builders.isEmpty else { return }
    update(&builders[builders.count - 1])
  }

  private func elementKind(for name: OfficeXMLName) -> OfficeSlideElementKind? {
    guard name.namespaceURI == Self.presentationNamespace else { return nil }
    return switch name.localName {
    case "sp": .shape
    case "pic": .picture
    case "cxnSp": .connector
    case "graphicFrame": .graphicFrame
    case "grpSp": .group
    default: nil
    }
  }

  private func readTransformChild(
    name: OfficeXMLName,
    attributes: [OfficeXMLAttribute],
    transform: inout RawTransform
  ) {
    switch name.localName {
    case "off":
      transform.x = attribute("x", in: attributes).flatMap(Int64.init)
      transform.y = attribute("y", in: attributes).flatMap(Int64.init)
    case "ext":
      transform.width = attribute("cx", in: attributes).flatMap(Int64.init)
      transform.height = attribute("cy", in: attributes).flatMap(Int64.init)
    case "chOff":
      transform.childX = attribute("x", in: attributes).flatMap(Int64.init)
      transform.childY = attribute("y", in: attributes).flatMap(Int64.init)
    case "chExt":
      transform.childWidth = attribute("cx", in: attributes).flatMap(Int64.init)
      transform.childHeight = attribute("cy", in: attributes).flatMap(Int64.init)
    default:
      break
    }
  }

  private func attribute(
    _ localName: String,
    in attributes: [OfficeXMLAttribute]
  ) -> String? {
    attributes.first { $0.name.namespaceURI == nil && $0.name.localName == localName }?.value
  }
}
