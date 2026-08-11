import Testing

@testable import OfficeKit

@Test func presentationIndexPreservesSlideOrderAndExactCanvasDimensions() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Presentation.pptx")
  )

  #expect(presentation.slideSize == OfficeSize(width: 960, height: 540))
  #expect(presentation.notesSize == OfficeSize(width: 540, height: 720))
  #expect(presentation.slides.map(\.identifier) == [256, 257])
  #expect(presentation.slides.map(\.relationshipID.rawValue) == ["rId2", "rId3"])
  #expect(
    presentation.slides.map(\.part.name.rawValue) == [
      "/ppt/slides/slide1.xml",
      "/ppt/slides/slide2.xml",
    ])
}

@Test func presentationIndexDoesNotInflateSlideOrAttachmentParts() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/mediareference.pptx")
  )
  #expect(presentation.slides.count == 2)
  let firstSlideRelationships = try presentation.document.package.relationships(
    from: .part(presentation.slides[0].part.name)
  )
  #expect(firstSlideRelationships.contains { $0.type.isEquivalent(to: .audio) })
  #expect(firstSlideRelationships.contains { $0.type.isEquivalent(to: .image) })
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/OpenXmlDomTest/DocumentTraverseTest.cs,
// TraversePPTDocument.
@Test(arguments: [
  "Text_withExtrusion_200chars+Animation (Fly In, all at once).pptx",
  "Text_withExtrusion_200chars+Animation (Fly In, by letter).pptx",
  "Text_withExtrusion_200chars.pptx",
])
func traversePPTDocument(filename: String) throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/smallset/" + filename
    )
  )
  #expect(!presentation.slides.isEmpty)

  for index in presentation.slides.indices {
    let slide = try presentation.slide(at: index)
    #expect(!slide.elements.isEmpty)
    for element in slide.elements {
      traverse(element)
    }
  }
}

private func traverse(_ element: OfficeSlideElement) {
  for child in element.children {
    traverse(child)
  }
}
