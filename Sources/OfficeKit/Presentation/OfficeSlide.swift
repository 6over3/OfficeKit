/// The semantic kind of an element in a PowerPoint slide shape tree.
public enum OfficeSlideElementKind: String, Sendable, Hashable, Codable {
  /// A text box or geometric shape.
  case shape
  /// A picture or media poster frame.
  case picture
  /// A connector between shapes.
  case connector
  /// A table, chart, diagram, or other graphic frame.
  case graphicFrame
  /// A group containing other drawing elements.
  case group
}

/// A PresentationML placeholder identity used for layout and master inheritance.
public struct OfficePlaceholder: Sendable, Hashable, Codable {
  /// The placeholder role exactly as declared, such as `title`, `body`, or `subTitle`.
  public let type: String?

  /// The placeholder index. PresentationML defines an omitted index as zero.
  public let index: UInt32

  /// Creates a placeholder identity from its source role and index.
  public init(type: String?, index: UInt32 = 0) {
    self.type = type
    self.index = index
  }
}

/// A DrawingML color token retained without resolving it through a theme.
public enum OfficeDrawingColor: Sendable, Hashable, Codable {
  /// An eight-digit or six-digit sRGB hexadecimal value.
  case sRGB(String)
  /// A theme color role such as `accent1` or `tx1`.
  case scheme(String)
  /// A system color name and optional last-known sRGB fallback.
  case system(name: String, lastColor: String?)
  /// A DrawingML preset color name.
  case preset(String)
}

/// Shape geometry retained from DrawingML.
public enum OfficeDrawingGeometry: Sendable, Hashable, Codable {
  /// A named preset such as `rect`, `ellipse`, or `roundRect`.
  case preset(String)
  /// A custom path geometry; path and guide counts summarize its authored complexity.
  case custom(pathCount: Int, guideCount: Int)
}

/// The direct fill authored on a DrawingML object.
public enum OfficeDrawingFill: Sendable, Hashable, Codable {
  /// No fill is painted.
  case none
  /// A single color, or an unresolved color token.
  case solid(OfficeDrawingColor?)
  /// A gradient fill whose detailed stops remain in package XML.
  case gradient
  /// A named pattern fill.
  case pattern(String?)
  /// A picture or texture fill.
  case picture
  /// A fill inherited from the containing group.
  case group
}

/// The direct outline authored on a DrawingML object.
public struct OfficeDrawingLine: Sendable, Hashable, Codable {
  /// Authored line width in exact EMUs.
  public let width: OfficeLength?
  /// Authored line fill.
  public let fill: OfficeDrawingFill?
  /// Dash preset token.
  public let dash: String?
  /// Line-cap token.
  public let cap: String?
  /// Compound-line token.
  public let compound: String?
  /// Stroke-alignment token.
  public let alignment: String?
}

/// A slide background authored directly or through a background style reference.
public struct OfficeSlideBackground: Sendable, Hashable, Codable {
  /// Directly authored background fill.
  public let fill: OfficeDrawingFill?
  /// One-based background style reference into the theme format scheme.
  public let styleIndex: UInt32?
}

/// One named color in a DrawingML theme color scheme.
public struct OfficeThemeColor: Sendable, Hashable, Codable {
  /// Theme role name such as `accent1`.
  public let name: String
  /// Color token assigned to the role.
  public let value: OfficeDrawingColor
}

/// One supplemental theme typeface selected for a writing system.
public struct OfficeThemeSupplementalFont: Sendable, Hashable, Codable {
  /// ISO 15924 script code.
  public let script: String
  /// Typeface selected for the script.
  public let typeface: String
}

/// Major or minor fonts from a DrawingML theme.
public struct OfficeThemeFontSet: Sendable, Hashable, Codable {
  /// Latin typeface.
  public let latinTypeface: String?
  /// East Asian typeface.
  public let eastAsianTypeface: String?
  /// Complex-script typeface.
  public let complexScriptTypeface: String?
  /// Script-specific supplemental typefaces.
  public let supplementalFonts: [OfficeThemeSupplementalFont]
}

/// The DrawingML theme related from a slide master.
public struct OfficeTheme: Sendable {
  /// Producer-authored theme name.
  public let name: String?
  /// Color-scheme name.
  public let colorSchemeName: String?
  /// Named theme colors in source order.
  public let colors: [OfficeThemeColor]
  /// Font-scheme name.
  public let fontSchemeName: String?
  /// Major-font set, commonly used for headings.
  public let majorFonts: OfficeThemeFontSet
  /// Minor-font set, commonly used for body text.
  public let minorFonts: OfficeThemeFontSet
  /// Theme XML part.
  public let sourcePart: OfficePart
  /// Lazy URL-backed theme resource.
  public let attachment: OfficeAttachment

  /// Finds a theme color by role, such as `accent1` or `dk1`.
  public func color(named name: String) -> OfficeDrawingColor? {
    colors.first { $0.name == name }?.value
  }
}

/// Direct character formatting for one PresentationML text run.
public struct OfficeSlideTextRunProperties: Sendable, Hashable, Codable {
  /// Primary BCP 47 language tag.
  public let language: String?
  /// Alternative BCP 47 language tag.
  public let alternativeLanguage: String?
  /// Font size measured in points.
  public let fontSizeInPoints: Double?
  /// Direct bold state.
  public let isBold: Bool?
  /// Direct italic state.
  public let isItalic: Bool?
  /// Underline style token.
  public let underline: String?
  /// Strike-through style token.
  public let strike: String?
  /// Capitalization style token.
  public let capitalization: String?
  /// Baseline adjustment expressed as an OOXML percentage.
  public let baseline: Double?
  /// Direct Latin typeface.
  public let latinTypeface: String?
  /// Direct East Asian typeface.
  public let eastAsianTypeface: String?
  /// Direct complex-script typeface.
  public let complexScriptTypeface: String?
  /// Direct text color.
  public let color: OfficeDrawingColor?
}

/// One authored run or field in PresentationML text.
public struct OfficeSlideTextRun: Sendable, Hashable, Codable {
  /// Visible run or field text.
  public let text: String
  /// Direct character formatting.
  public let properties: OfficeSlideTextRunProperties
  /// Field identifier, when the run represents a field.
  public let fieldIdentifier: String?
  /// Field type token, when the run represents a field.
  public let fieldType: String?
}

/// Bullet metadata for a PresentationML paragraph.
public enum OfficeSlideBullet: Sendable, Hashable, Codable {
  /// An explicitly unbulleted paragraph.
  case none
  /// A paragraph introduced by the supplied bullet character.
  case character(String)
  /// A paragraph using the authored automatic numbering scheme and optional starting value.
  case automatic(numbering: String, startAt: Int?)
}

/// One authored PresentationML text paragraph.
public struct OfficeSlideTextParagraph: Sendable, Hashable, Codable {
  /// Zero-based paragraph outline level.
  public let level: Int
  /// Paragraph alignment token.
  public let alignment: String?
  /// Leading text margin in exact EMUs.
  public let leadingMargin: OfficeLength?
  /// First-line indentation in exact EMUs.
  public let indent: OfficeLength?
  /// Direct right-to-left state.
  public let isRightToLeft: Bool?
  /// Bullet configuration.
  public let bullet: OfficeSlideBullet?
  /// Runs and fields in source order.
  public let runs: [OfficeSlideTextRun]

  /// Visible paragraph text assembled from its runs.
  public var text: String { runs.map(\.text).joined() }
}

/// Structured DrawingML text attached to a slide element.
public struct OfficeSlideTextBody: Sendable, Hashable, Codable {
  /// Paragraphs in source order.
  public let paragraphs: [OfficeSlideTextParagraph]

  /// Visible text assembled using paragraph separators.
  public var text: String { paragraphs.map(\.text).joined(separator: "\n") }
}

/// The semantic role of a slide media relationship.
public enum OfficeSlideMediaKind: String, Sendable, Hashable, Codable {
  /// Audio content.
  case audio
  /// Video content.
  case video
  /// Generic media content whose audiovisual subtype is not declared.
  case media
}

/// A typed audio/video relationship with lazy URL access through its attachment.
public struct OfficeSlideMedia: Sendable {
  /// Semantic media role.
  public let kind: OfficeSlideMediaKind
  /// Relationship-backed media with lazy URL access.
  public let attachment: OfficeAttachment
}

/// Authored slide-transition metadata without playback behavior.
public struct OfficeSlideTransition: Sendable, Hashable, Codable {
  /// Transition element name, such as `wipe` or `fade`.
  public let kind: String?
  /// Authored speed token.
  public let speed: String?
  /// Explicit transition duration in milliseconds.
  public let durationMilliseconds: UInt32?
  /// Whether a click advances the slide.
  public let advancesOnClick: Bool?
  /// Automatic advance delay in milliseconds.
  public let advanceAfterMilliseconds: UInt32?
}

/// Inspectable summary of a slide timing tree.
public struct OfficeSlideTiming: Sendable, Hashable, Codable {
  /// Number of timing nodes in the authored animation tree.
  public let timeNodeCount: Int
  /// Number of animation behavior elements.
  public let behaviorCount: Int
}

/// One shape-tree element from a PowerPoint slide.
public struct OfficeSlideElement: Sendable {
  /// The element kind.
  public let kind: OfficeSlideElementKind

  /// The non-visual DrawingML identifier, when declared.
  public let identifier: UInt32?

  /// The producer-assigned non-visual name.
  public let name: String?

  /// Accessibility-oriented alternative text from `p:cNvPr/@descr`.
  public let alternativeText: String?

  /// The producer-assigned non-visual title from `p:cNvPr/@title`.
  public let title: String?

  /// Plain text found in DrawingML text runs, in source order.
  public let text: String

  /// Structured paragraphs and runs, including direct language and font metadata.
  public let textBody: OfficeSlideTextBody?

  /// The placeholder identity, when this element participates in layout inheritance.
  public let placeholder: OfficePlaceholder?

  /// The typed payload of a graphic frame, when recognized.
  public let graphicContent: OfficeGraphicContent?

  /// Picture-specific image and crop information.
  public let picture: OfficePicture?

  /// Direct shape geometry, when authored on this element.
  public let geometry: OfficeDrawingGeometry?

  /// Direct shape fill, before theme color resolution.
  public let fill: OfficeDrawingFill?

  /// Direct shape outline, before theme color resolution.
  public let line: OfficeDrawingLine?

  /// Authored and resolved spatial information.
  public let spatialInfo: OfficeSpatialInfo

  /// Relationship-backed resources referenced by this element.
  public let attachments: [OfficeAttachment]

  /// Nested elements for a group.
  public let children: [OfficeSlideElement]

  /// Audio and video payloads referenced by this element.
  public var media: [OfficeSlideMedia] {
    attachments.compactMap { attachment in
      let kind: OfficeSlideMediaKind
      if attachment.relationship.type.isEquivalent(to: .audio) {
        kind = .audio
      } else if attachment.relationship.type.isEquivalent(to: .video) {
        kind = .video
      } else if attachment.relationship.type.isEquivalent(to: .media) {
        kind = .media
      } else {
        return nil
      }
      return OfficeSlideMedia(kind: kind, attachment: attachment)
    }
  }
}

/// A related layout or master shape tree that contributes content to a slide.
public struct OfficeSlideLayer: Sendable {
  /// The XML part that owns the layer.
  public let part: OfficePart

  /// Whether PresentationML requests that this layer be shown for the slide.
  public let isVisible: Bool

  /// The layer's directly authored background.
  public let background: OfficeSlideBackground?

  /// Elements in the layer's own back-to-front shape-tree order.
  public let elements: [OfficeSlideElement]

  /// Payload relationships owned by this layer's part, including resources not directly cited
  /// by an element.
  public let relatedAttachments: [OfficeAttachment]

  /// Every relationship-backed resource referenced by this layer.
  public var attachments: [OfficeAttachment] {
    uniqueOfficeAttachments(relatedAttachments + elements.flatMap(attachments(in:)))
  }

  private func attachments(in element: OfficeSlideElement) -> [OfficeAttachment] {
    element.attachments + element.children.flatMap(attachments(in:))
  }
}

/// A slide's speaker notes and their inherited notes-master layer.
public struct OfficeSlideNotes: Sendable {
  /// The notes XML part related from the slide.
  public let part: OfficePart

  /// Notes-page elements in back-to-front shape-tree order.
  public let elements: [OfficeSlideElement]

  /// The related notes-master layer, when present.
  public let masterLayer: OfficeSlideLayer?

  /// Payload relationships owned directly by the notes page.
  public let relatedAttachments: [OfficeAttachment]

  /// Whether placeholder animation inherited from the master is enabled.
  public let showsMasterPlaceholderAnimations: Bool

  /// Plain speaker-note text from body placeholders, in shape-tree order.
  public var text: String {
    elements
      .filter { $0.placeholder?.type == "body" }
      .map(\.text)
      .filter { !$0.isEmpty }
      .joined(separator: "\n")
  }

  /// Every attachment referenced by the notes page and notes master.
  public var attachments: [OfficeAttachment] {
    uniqueOfficeAttachments(
      (masterLayer?.attachments ?? [])
        + relatedAttachments
        + elements.flatMap(attachments(in:))
    )
  }

  private func attachments(in element: OfficeSlideElement) -> [OfficeAttachment] {
    element.attachments + element.children.flatMap(attachments(in:))
  }
}

/// The parsed semantic content of one PowerPoint slide.
public struct OfficeSlide: Sendable {
  /// The ordered presentation reference that led to this slide.
  public let reference: OfficeSlideReference

  /// Whether the slide is hidden from the normal slide show.
  public let isHidden: Bool

  /// Authored transition metadata, when present.
  public let transition: OfficeSlideTransition?

  /// Summary of the authored animation timing tree, when present.
  public let timing: OfficeSlideTiming?

  /// The slide's directly authored background.
  public let background: OfficeSlideBackground?

  /// The theme related from the slide's master, when present.
  public let theme: OfficeTheme?

  /// Top-level shape-tree elements in back-to-front z-order.
  public let elements: [OfficeSlideElement]

  /// The slide-layout shape tree, when related by the package graph.
  public let layoutLayer: OfficeSlideLayer?

  /// The slide-master shape tree, when related through the slide layout.
  public let masterLayer: OfficeSlideLayer?

  /// Speaker notes related from this slide, when present.
  public let notes: OfficeSlideNotes?

  /// Legacy comments related from this slide, in comment-list order.
  public let comments: [OfficePresentationComment]

  /// Payload relationships owned directly by the slide part.
  ///
  /// This includes resources such as rendered SmartArt drawing parts that are related from the
  /// slide but are not explicitly cited by a shape attribute.
  public let relatedAttachments: [OfficeAttachment]

  /// Every payload attachment owned by or referenced directly from the slide shape tree.
  public var localAttachments: [OfficeAttachment] {
    uniqueOfficeAttachments(relatedAttachments + elements.flatMap(attachments(in:)))
  }

  /// Every attachment referenced by the slide, layout, and master shape trees.
  public var attachments: [OfficeAttachment] {
    uniqueOfficeAttachments(
      (masterLayer?.attachments ?? [])
        + (layoutLayer?.attachments ?? [])
        + (notes?.attachments ?? [])
        + localAttachments
    )
  }

  private func attachments(in element: OfficeSlideElement) -> [OfficeAttachment] {
    element.attachments + element.children.flatMap(attachments(in:))
  }
}

private func uniqueOfficeAttachments(_ attachments: [OfficeAttachment]) -> [OfficeAttachment] {
  var relationships = Set<OfficeRelationship>()
  return attachments.filter { relationships.insert($0.relationship).inserted }
}
