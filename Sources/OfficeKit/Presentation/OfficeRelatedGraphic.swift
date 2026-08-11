/// The related parts that define a SmartArt diagram.
public struct OfficeDiagramReference: Sendable {
  /// The diagram data model.
  public let data: OfficeAttachment?

  /// The diagram layout definition.
  public let layout: OfficeAttachment?

  /// The diagram quick-style definition.
  public let quickStyle: OfficeAttachment?

  /// The diagram color definition.
  public let colors: OfficeAttachment?

  /// Every relationship explicitly referenced by the diagram frame.
  public let attachments: [OfficeAttachment]

  package init(attachments: [OfficeAttachment]) {
    self.data = attachments.first {
      $0.relationship.type.isEquivalent(to: .diagramData)
    }
    self.layout = attachments.first {
      $0.relationship.type.isEquivalent(to: .diagramLayout)
    }
    self.quickStyle = attachments.first {
      $0.relationship.type.isEquivalent(to: .diagramQuickStyle)
    }
    self.colors = attachments.first {
      $0.relationship.type.isEquivalent(to: .diagramColors)
    }
    self.attachments = attachments
  }
}

/// The packaged 3D model and its raster fallback image.
public struct OfficeModel3DReference: Sendable {
  /// The packaged 3D model, commonly a binary glTF (`.glb`) part.
  public let model: OfficeAttachment?

  /// The raster fallback or preview image displayed by clients without 3D support.
  public let posterImage: OfficeAttachment?

  /// Every relationship explicitly referenced by the 3D graphic frame.
  public let attachments: [OfficeAttachment]

  package init(attachments: [OfficeAttachment]) {
    self.model = attachments.first {
      $0.relationship.type.isEquivalent(to: .model3D)
    }
    self.posterImage = attachments.first {
      $0.relationship.type.isEquivalent(to: .image)
    }
    self.attachments = attachments
  }
}
