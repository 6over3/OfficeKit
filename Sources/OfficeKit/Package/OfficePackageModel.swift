import UniformTypeIdentifiers

/// A MIME content type declared by an OPC package.
public struct OfficeContentType: RawRepresentable, Sendable, Hashable, Codable,
  CustomStringConvertible
{
  /// The content type exactly as declared by the package.
  public let rawValue: String

  /// Creates a content type from its MIME spelling.
  public init(rawValue: String) {
    self.rawValue = rawValue
  }

  /// A canonical spelling used when matching known equivalent real-world MIME types.
  ///
  /// The original spelling remains available through `rawValue`.
  public var canonicalValue: String {
    switch rawValue.lowercased() {
    case "model/gltf.binary": "model/gltf-binary"
    default: rawValue.lowercased()
    }
  }

  /// The system uniform type corresponding to this MIME type, when one can be formed.
  ///
  /// Use `rawValue` when exact producer spelling or an unknown extension MIME type matters.
  public var uniformType: UTType? {
    UTType(mimeType: canonicalValue)
  }

  /// Whether this content type identifies XML handled by OfficeKit's event or tree readers.
  public var isXML: Bool {
    canonicalValue.hasSuffix("+xml") || canonicalValue.hasSuffix("/xml")
      || canonicalValue == "application/vnd.openxmlformats-officedocument.vmldrawing"
  }

  /// The content type exactly as declared by the package.
  public var description: String { rawValue }
}

/// Metadata for one file-like part in an OPC package.
public struct OfficePart: Sendable, Hashable, Codable {
  /// The package-absolute part name.
  public let name: OfficePartName

  /// The MIME content type assigned by `[Content_Types].xml`.
  public let contentType: OfficeContentType

  /// The system uniform type corresponding to `contentType`, when available.
  public var uniformType: UTType? { contentType.uniformType }

  /// Whether this part contains XML that can be consumed by OfficeKit's event or tree readers.
  public var isXML: Bool { contentType.isXML }

  /// The declared uncompressed byte size in the ZIP central directory.
  public let uncompressedSize: UInt64

  package init(name: OfficePartName, contentType: OfficeContentType, uncompressedSize: UInt64) {
    self.name = name
    self.contentType = contentType
    self.uncompressedSize = uncompressedSize
  }
}

/// The source that owns an OPC relationship.
public enum OfficeRelationshipSource: Sendable, Hashable, Codable, CustomStringConvertible {
  /// A relationship owned by the package itself.
  case package

  /// A relationship owned by a part.
  case part(OfficePartName)

  /// A stable description of the relationship source.
  public var description: String {
    switch self {
    case .package: "/"
    case .part(let name): name.rawValue
    }
  }

  package var relationshipArchivePath: String {
    switch self {
    case .package:
      return "_rels/.rels"
    case .part(let name):
      let path = name.archivePath
      let finalSlash = path.lastIndex(of: "/")
      if let finalSlash {
        let directory = path[..<finalSlash]
        let filename = path[path.index(after: finalSlash)...]
        return "\(directory)/_rels/\(filename).rels"
      }
      return "_rels/\(path).rels"
    }
  }
}

/// The identifier of a relationship within its source relationship part.
public struct OfficeRelationshipID: RawRepresentable, Sendable, Hashable, Codable,
  CustomStringConvertible
{
  /// The identifier exactly as declared by the relationship XML.
  public let rawValue: String

  /// Creates a relationship identifier.
  public init(rawValue: String) { self.rawValue = rawValue }

  /// The identifier exactly as declared by the relationship XML.
  public var description: String { rawValue }
}

/// The semantic URI that describes an OPC relationship.
public struct OfficeRelationshipType: RawRepresentable, Sendable, Hashable, Codable,
  CustomStringConvertible
{
  /// The relationship type URI.
  public let rawValue: String

  /// Creates a relationship type from its URI.
  public init(rawValue: String) { self.rawValue = rawValue }

  /// The relationship type URI.
  public var description: String { rawValue }

  /// A canonical URI used to compare equivalent Strict and Transitional relationship types.
  ///
  /// The source spelling remains available through `rawValue`.
  public var canonicalValue: String {
    let strictPrefix = "http://purl.oclc.org/ooxml/officeDocument/relationships/"
    guard rawValue.hasPrefix(strictPrefix) else { return rawValue }
    if rawValue == strictPrefix + "customXml" {
      return "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
    }
    return rawValue.replacingOccurrences(
      of: strictPrefix,
      with: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    )
  }

  /// Reports whether two relationship types have the same Strict/Transitional semantic role.
  public func isEquivalent(to other: Self) -> Bool {
    canonicalValue == other.canonicalValue
  }

  /// The Transitional relationship from a package to its main Office document part.
  public static let officeDocument = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
  )

  /// The Strict relationship from a package to its main Office document part.
  public static let strictOfficeDocument = Self(
    rawValue: "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
  )

  /// The relationship from a workbook to its calculation chain.
  public static let calculationChain = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
  )

  /// A relationship from a workbook to a worksheet.
  public static let worksheet = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  )

  /// A relationship from a workbook to a chart sheet.
  public static let chartSheet = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
  )

  /// A relationship from a workbook to a legacy dialog sheet.
  public static let dialogSheet = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet"
  )

  /// A relationship from a workbook to a legacy macro sheet.
  public static let macroSheet = Self(
    rawValue: "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet"
  )

  /// A relationship from a workbook to an international legacy macro sheet.
  public static let internationalMacroSheet = Self(
    rawValue: "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet"
  )

  /// A relationship from a worksheet to its DrawingML drawing part.
  public static let drawing = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
  )

  /// A relationship from a workbook to its shared-string table.
  public static let sharedStrings = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  )

  /// A relationship from a workbook to its cell-style table.
  public static let styles = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  )

  /// The Microsoft relationship from a drawing to a 3D model.
  public static let model3D = Self(
    rawValue: "http://schemas.microsoft.com/office/2017/06/relationships/model3d"
  )

  /// The package relationship to Dublin Core package properties.
  public static let coreProperties = Self(
    rawValue:
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
  )

  /// The package relationship to extended application properties.
  public static let extendedProperties = Self(
    rawValue:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
  )

  /// The package relationship to custom document properties.
  public static let customProperties = Self(
    rawValue:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
  )

  /// The package relationship to a thumbnail image.
  public static let thumbnail = Self(
    rawValue: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
  )

  /// A relationship from a presentation to an ordered slide.
  public static let slide = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
  )

  /// A relationship from a slide to its layout.
  public static let slideLayout = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
  )

  /// A relationship from a slide layout to its master.
  public static let slideMaster = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
  )

  /// A relationship from presentation content to its theme definition.
  public static let theme = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
  )

  /// A relationship from a slide to its speaker-notes page.
  public static let notesSlide = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
  )

  /// A relationship from a notes page to its notes master.
  public static let notesMaster = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
  )

  /// A relationship from an Office document part to its legacy comment list.
  public static let comments = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
  )

  /// A relationship from a Word document to a header story.
  public static let header = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
  )

  /// A relationship from a Word document to a footer story.
  public static let footer = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
  )

  /// A relationship from a Word document to its footnote collection.
  public static let footnotes = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
  )

  /// A relationship from a Word document to its endnote collection.
  public static let endnotes = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
  )

  /// A relationship from a Word document to its numbering definitions.
  public static let numbering = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
  )

  /// A relationship from a Word document to its document settings.
  public static let settings = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
  )

  /// A relationship from a worksheet to a structured table definition.
  public static let table = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
  )

  /// A relationship from a worksheet to a legacy VML drawing.
  public static let vmlDrawing = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
  )

  /// A relationship from a worksheet to an inert OLE payload.
  public static let oleObject = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
  )

  /// A relationship from a worksheet to an inert ActiveX control descriptor.
  public static let control = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control"
  )

  /// A relationship from a presentation to its comment-author list.
  public static let commentAuthors = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
  )

  /// A relationship from a drawing frame to a chart part.
  public static let chart = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
  )

  /// A relationship from a SmartArt frame to its diagram data.
  public static let diagramData = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
  )

  /// A relationship from a SmartArt frame to its diagram layout definition.
  public static let diagramLayout = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"
  )

  /// A relationship from a SmartArt frame to its diagram quick style.
  public static let diagramQuickStyle = Self(
    rawValue:
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle"
  )

  /// A relationship from a SmartArt frame to its diagram color definition.
  public static let diagramColors = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"
  )

  /// A Microsoft relationship from a slide to rendered SmartArt drawing data.
  public static let diagramDrawing = Self(
    rawValue: "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"
  )

  /// A relationship to an embedded package, commonly a chart's source workbook.
  public static let embeddedPackage = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
  )

  /// A relationship to an image part.
  public static let image = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
  )

  /// A Microsoft relationship to a JPEG XR / HD Photo image part.
  public static let highDefinitionPhoto = Self(
    rawValue: "http://schemas.microsoft.com/office/2007/relationships/hdphoto"
  )

  /// A relationship to an audio part.
  public static let audio = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
  )

  /// A relationship to a video part.
  public static let video = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
  )

  /// A Microsoft relationship to a media payload used by modern PowerPoint.
  public static let media = Self(
    rawValue: "http://schemas.microsoft.com/office/2007/relationships/media"
  )

  /// A relationship to an external hyperlink.
  public static let hyperlink = Self(
    rawValue: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  )
}

/// The resolved target of an OPC relationship.
public enum OfficeRelationshipTarget: Sendable, Hashable, Codable {
  /// A part in the same package, with an optional URI fragment.
  case internalPart(OfficePartName, fragment: String?)

  /// A target outside the package, preserved without network access or normalization.
  case external(String)
}

/// A resolved relationship owned by the package or one of its parts.
public struct OfficeRelationship: Sendable, Hashable, Codable {
  /// The package or part that owns this relationship.
  public let source: OfficeRelationshipSource

  /// The identifier used by XML in the source part.
  public let id: OfficeRelationshipID

  /// The URI describing the relationship's semantic role.
  public let type: OfficeRelationshipType

  /// The resolved internal part or preserved external URI.
  public let target: OfficeRelationshipTarget

  /// The target string exactly as it appeared in the relationship XML.
  public let rawTarget: String
}
