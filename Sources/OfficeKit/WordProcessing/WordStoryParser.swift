package enum WordStoryParser {
  package static func parse(
    kind: OfficeWordStoryKind,
    package: OfficePackage,
    relationships: [OfficeRelationship]
  ) throws -> [OfficeWordStory] {
    let relationshipType: OfficeRelationshipType = kind == .header ? .header : .footer
    var seenParts: Set<OfficePartName> = []
    var stories: [OfficeWordStory] = []
    for relationship in relationships where relationship.type.isEquivalent(to: relationshipType) {
      guard let part = package.part(referencedBy: relationship) else {
        throw OfficeKitError.missingPart(relationship.rawTarget)
      }
      guard seenParts.insert(part.name).inserted else { continue }
      let storyRelationships = try package.relationships(from: .part(part.name))
      stories.append(
        OfficeWordStory(
          kind: kind,
          part: part,
          content: try WordDocumentParser.parse(
            part: part,
            package: package,
            relationships: storyRelationships
          ),
          attachments: storyRelationships.map(package.attachment(referencedBy:))
        )
      )
    }
    return stories
  }
}
