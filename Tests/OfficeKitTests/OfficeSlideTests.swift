import Foundation
import Testing
import UniformTypeIdentifiers

@testable import OfficeKit

@Test func strictSlideAndMasterRetainAuthoredBackgroundDeclarations() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/O14ISOStrict/PowerPoint/4.6.10 bg.pptx"
    )
  )
  let slide = try presentation.slide(at: 0)

  #expect(presentation.document.conformance == .strict)
  #expect(slide.background != nil)
  #expect(slide.masterLayer?.background != nil)
}

@Test func slideTextRunsTransitionsAndTimingRemainStructuredAndInspectable() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Presentation.pptx")
  )
  let firstSlide = try presentation.slide(at: 0)
  let title = try #require(firstSlide.elements.first { $0.text.contains("Title") })
  let titleBody = try #require(title.textBody)
  let titleRun = try #require(titleBody.paragraphs.first?.runs.first)
  let theme = try #require(firstSlide.theme)
  let secondSlide = try presentation.slide(at: 1)

  #expect(titleBody.text == title.text)
  #expect(titleRun.text == "Title")
  #expect(titleRun.properties.language == "en-US")
  #expect(theme.name == "Organic")
  #expect(theme.colorSchemeName == "Organic")
  #expect(theme.color(named: "accent1") == .sRGB("83992A"))
  #expect(theme.majorFonts.latinTypeface == "Garamond")
  #expect(
    theme.majorFonts.supplementalFonts.first { $0.script == "Jpan" }?.typeface
      == "ＭＳ Ｐゴシック")
  #expect(try theme.attachment.url().isFileURL)
  #expect(secondSlide.transition?.kind == "wipe")
  #expect(secondSlide.transition?.speed == "slow")
  #expect(secondSlide.timing?.timeNodeCount ?? 0 > 0)

  let animatedPresentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/smallset/"
        + "Text_withExtrusion_200chars+Animation (Fly In, all at once).pptx"
    )
  )
  let animatedSlide = try animatedPresentation.slide(at: 0)
  #expect(animatedSlide.timing?.timeNodeCount == 24)
  #expect(animatedSlide.timing?.behaviorCount == 12)
}

@Test func slideShapeTreeRetainsZOrderExactTransformsAndAttachmentURLs() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/mediareference.pptx")
  )
  let slide = try presentation.slide(at: 0)

  #expect(slide.elements.map(\.kind) == [.shape, .picture])
  #expect(slide.elements.map(\.spatialInfo.zIndex) == [0, 1])

  let title = slide.elements[0]
  #expect(title.identifier == 2)
  #expect(title.name == "Title 1")
  #expect(title.spatialInfo.sourceTransform?.x.emu == 611_560)
  #expect(title.spatialInfo.sourceTransform?.y.emu == 620_688)
  #expect(title.spatialInfo.sourceTransform?.width.emu == 7_772_400)
  #expect(title.spatialInfo.sourceTransform?.height.emu == 1_470_025)

  let media = slide.elements[1]
  #expect(media.identifier == 4)
  #expect(media.name == "Windows Navigation Start.wav")
  #expect(media.geometry == .preset("rect"))
  #expect(media.attachments.count == 3)
  #expect(media.media.map(\.kind) == [.audio, .media])
  #expect(try media.media.allSatisfy { try $0.attachment.url().isFileURL })
  #expect(Set(media.attachments.compactMap(\.contentType)) == Set([.wav, .png]))
  #expect(
    media.attachments.contains { attachment in
      attachment.part?.name.rawValue == "/ppt/media/image1.png"
        && (try? attachment.url().isFileURL) == true
    })
  #expect(slide.localAttachments.count == 3)
  #expect(slide.attachments.count == 4)
  #expect(
    slide.attachments.contains {
      $0.relationship.type.isEquivalent(to: .theme)
    })
}

@Test func placeholderTransformIsInheritedFromSlideLayout() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Presentation.pptx")
  )
  let slide = try presentation.slide(at: 0)
  let title = try #require(slide.elements.first)
  #expect(title.text == "Title")
  #expect(title.placeholder == OfficePlaceholder(type: "ctrTitle", index: 0))
  #expect(title.spatialInfo.sourceTransform?.x.emu == 2_692_398)
  #expect(title.spatialInfo.sourceTransform?.y.emu == 1_871_131)
  #expect(title.spatialInfo.sourceTransform?.width.emu == 6_815_669)
  #expect(title.spatialInfo.sourceTransform?.height.emu == 1_515_533)
  #expect(title.spatialInfo.geometrySourcePart?.rawValue == "/ppt/slideLayouts/slideLayout1.xml")
  #expect(title.spatialInfo.frame != nil)
  #expect(title.spatialInfo.resolution == .derived)

  let subtitle = try #require(slide.elements.dropFirst().first)
  #expect(subtitle.placeholder == OfficePlaceholder(type: "subTitle", index: 1))
  #expect(subtitle.spatialInfo.sourceTransform?.x.emu == 2_692_398)
  #expect(subtitle.spatialInfo.sourceTransform?.y.emu == 3_657_597)
}

@Test func slideTraversalIncludesRelationshipBackedLayoutAndMasterLayers() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Presentation.pptx")
  )
  let slide = try presentation.slide(at: 0)
  let layout = try #require(slide.layoutLayer)
  let master = try #require(slide.masterLayer)

  #expect(layout.part.name.rawValue == "/ppt/slideLayouts/slideLayout1.xml")
  #expect(master.part.name.rawValue == "/ppt/slideMasters/slideMaster1.xml")
  #expect(layout.isVisible)
  #expect(!master.isVisible)
  #expect(!layout.elements.isEmpty)
  #expect(!master.elements.isEmpty)
  #expect(!layout.attachments.isEmpty)
  #expect(!master.attachments.isEmpty)
  #expect(slide.localAttachments.isEmpty)
  #expect(slide.attachments.count == layout.attachments.count + master.attachments.count)
  #expect(try slide.attachments.allSatisfy { try $0.url().isFileURL })
}

@Test func speakerNotesRetainTextMasterGeometryAndPartIdentity() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/smallset/"
        + "Text_withExtrusion_200chars.pptx"
    )
  )
  let slide = try presentation.slide(at: 0)
  let notes = try #require(slide.notes)
  let master = try #require(notes.masterLayer)
  let body = try #require(notes.elements.first { $0.placeholder?.type == "body" })

  #expect(notes.part.name.rawValue == "/ppt/notesSlides/notesSlide1.xml")
  #expect(master.part.name.rawValue == "/ppt/notesMasters/notesMaster1.xml")
  #expect(master.isVisible)
  #expect(notes.showsMasterPlaceholderAnimations)
  #expect(notes.text.contains("Created in"))
  #expect(notes.text.contains("3-D Format"))
  #expect(body.spatialInfo.sourceTransform?.x.emu == 685_800)
  #expect(body.spatialInfo.sourceTransform?.y.emu == 4_343_400)
  #expect(body.spatialInfo.sourceTransform?.width.emu == 5_486_400)
  #expect(body.spatialInfo.sourceTransform?.height.emu == 4_114_800)
  #expect(body.spatialInfo.geometrySourcePart?.rawValue == "/ppt/notesMasters/notesMaster1.xml")
  #expect(body.spatialInfo.resolution == .derived)
}

@Test func strictPresentationCommentsResolveAuthorsTextAndExactPositions() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/O14ISOStrict/PowerPoint/"
        + "[HC]viewPr-PresentationViewProperties-showComments-1.pptx"
    )
  )
  let author = try #require(presentation.commentAuthors.first)
  let slide = try presentation.slide(at: 0)

  #expect(presentation.commentAuthors.count == 1)
  #expect(author.identifier == 0)
  #expect(author.name == "Harshal Doshi")
  #expect(author.initials == "HD")
  #expect(author.colorIndex == 0)
  #expect(slide.comments.count == 2)

  let first = slide.comments[0]
  #expect(first.index == 1)
  #expect(first.authorID == 0)
  #expect(first.author == author)
  #expect(first.dateTime == "2008-02-01T13:48:50.465")
  #expect(first.text == "PowerPoint Rocks!!!")
  #expect(first.position.x.emu == 10)
  #expect(first.position.y.emu == 10)
  #expect(first.sourcePart.rawValue == "/ppt/comments/comment1.xml")

  let second = slide.comments[1]
  #expect(second.index == 2)
  #expect(second.text == "PowerPoint Logo!")
  #expect(second.position.x.emu == 5_244)
  #expect(second.position.y.emu == 384)
}

// Read-only regression extracted from dotnet/Open-XML-SDK @
// cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/OpenXmlDomTest/CodeGenSanityTest.cs,
// Bug225919_MitigateNamespaceIssue with Block_hyperlink_crash.
@Test func hyperlinkRegressionResolvesExternalURLWithoutFetching() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/block_hyperlink_crash.pptx"
    )
  )
  let slide = try presentation.slide(at: 0)
  let linkedShape = try #require(slide.elements.first { $0.text.contains("www.foo.com") })
  let hyperlink = try #require(
    linkedShape.attachments.first {
      $0.relationship.type.isEquivalent(to: .hyperlink)
    })

  #expect(hyperlink.part == nil)
  #expect(hyperlink.relationship.target == .external("http://www.foo.com/"))
  #expect(try hyperlink.url() == URL(string: "http://www.foo.com/"))
}

@Test func drawingTableRetainsGridTextStyleAndHonestCellGeometry() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/Table_Small.pptx"
    )
  )
  let slide = try presentation.slide(at: 0)
  let frame = try #require(
    slide.elements.first { element in
      guard case .table = element.graphicContent else { return false }
      return true
    })
  guard case .table(let table) = frame.graphicContent else {
    Issue.record("Expected an inline DrawingML table.")
    return
  }

  #expect(frame.identifier == 6)
  #expect(frame.name == "Table 5")
  #expect(frame.text.contains("Major"))
  #expect(!frame.text.contains("5C22544A"))
  #expect(frame.spatialInfo.sourceTransform?.x.emu == 457_200)
  #expect(frame.spatialInfo.sourceTransform?.y.emu == 1_600_200)
  #expect(table.columnWidths.map(\.emu) == [2_743_200, 2_743_200, 2_743_200])
  #expect(table.rows.count == 6)
  #expect(table.rows.allSatisfy { $0.cells.count == 3 })
  #expect(table.styleIdentifier == "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
  #expect(table.rows[0].cells.map(\.text) == ["", "Major", "Minor"])
  #expect(table.rows[1].cells.map(\.text) == ["Office Theme", "Cambria", "Calibri"])
  #expect(table.rows.allSatisfy { $0.height.emu == 0 })
  #expect(table.rows.allSatisfy { row in row.cells.allSatisfy { $0.spatialInfo.frame == nil } })
  guard case .unresolved = table.rows[0].cells[0].spatialInfo.resolution else {
    Issue.record("Zero authored row heights must not produce invented cell frames.")
    return
  }
}

@Test func chartFrameResolvesLazyChartPartAndExactFrame() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/v2FxTestFiles/presentation/test.pptx"
    )
  )
  let slide = try presentation.slide(at: 0)
  let frame = try #require(
    slide.elements.first { element in
      guard case .chart = element.graphicContent else { return false }
      return true
    })
  guard case .chart(let chartReference) = frame.graphicContent else {
    Issue.record("Expected a lazy chart reference.")
    return
  }
  let chart = try #require(
    frame.attachments.first {
      $0.relationship.type.isEquivalent(to: .chart)
    })

  #expect(frame.identifier == 2_052)
  #expect(frame.name == "Object 4")
  #expect(frame.spatialInfo.sourceTransform?.x.emu == 1_971_675)
  #expect(frame.spatialInfo.sourceTransform?.y.emu == 1_681_163)
  #expect(frame.spatialInfo.sourceTransform?.width.emu == 5_200_650)
  #expect(frame.spatialInfo.sourceTransform?.height.emu == 3_495_675)
  #expect(chart.part?.name.rawValue == "/ppt/charts/chart1.xml")
  #expect(try chart.url().isFileURL)

  let parsed = try chartReference.chart()
  #expect(parsed.kind == .bar)
  #expect(parsed.sourceKind == "bar3DChart")
  #expect(parsed.styleIdentifier == 13)
  #expect(parsed.grouping == "clustered")
  #expect(parsed.barDirection == "col")
  #expect(parsed.plotAreaLayout?.target == "inner")
  #expect(parsed.plotAreaLayout?.horizontalMode == "edge")
  #expect(parsed.plotAreaLayout?.x == 0.113_805_970_149_253_75)
  #expect(parsed.axes.map(\.kind) == [.category, .value])
  #expect(parsed.axes.map(\.identifier) == [39_394_304, 49_064_192])
  #expect(parsed.axes.map(\.position) == [.bottom, .left])
  #expect(parsed.axes[0].crossingAxisIdentifier == 49_064_192)
  #expect(parsed.axes[1].crossingAxisIdentifier == 39_394_304)
  #expect(parsed.axes[1].crossingValue == 1)
  #expect(parsed.legend?.position == "r")
  #expect(parsed.legend?.layout?.horizontalMode == "edge")
  #expect(parsed.plotsVisibleCellsOnly == true)
  #expect(parsed.displayBlanksAs == "gap")
  #expect(parsed.series.count == 3)
  #expect(parsed.series.map(\.name) == ["East", "West", "North"])
  #expect(parsed.series[0].nameFormula == "Sheet1!$A$2")
  #expect(parsed.series[0].categoryFormula == "Sheet1!$B$1:$E$1")
  #expect(parsed.series[0].categories == ["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"])
  #expect(parsed.series[0].valueFormula == "Sheet1!$B$2:$E$2")
  #expect(parsed.series[0].values == [20.4, 27.4, 90, 20.4])
  let workbook = try #require(
    parsed.attachments.first {
      $0.relationship.type.isEquivalent(to: .embeddedPackage)
    })
  #expect(
    workbook.part?.name.rawValue
      == "/ppt/embeddings/Microsoft_Office_Excel_Worksheet1.xlsx"
  )
  #expect(try workbook.url().isFileURL)
}

@Test func strictPictureRetainsAltTextExactCropAndLazyImageURL() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(
      for: "Open-XML-SDK/TestDataStorage/O14ISOStrict/PowerPoint/"
        + "webPr-WebProperties-ImageTargetResolution-imgsz-1024x786.pptx"
    )
  )
  let slide = try presentation.slide(at: 0)
  let element = try #require(slide.elements.first { $0.kind == .picture })
  let picture = try #require(element.picture)
  let image = try #require(picture.primaryImage)
  let crop = try #require(picture.cropRectangle)

  #expect(element.identifier == 1_026)
  #expect(element.name == "Picture 2")
  #expect(element.alternativeText == "F:\\Users\\hdoshi\\Pictures\\powerpoint-logo.jpg")
  #expect(element.spatialInfo.sourceTransform?.x.emu == 7_689_850)
  #expect(element.spatialInfo.sourceTransform?.y.emu == 381_000)
  #expect(element.spatialInfo.sourceTransform?.width.emu == 1_073_150)
  #expect(element.spatialInfo.sourceTransform?.height.emu == 990_600)
  #expect(crop.left.rawValue == 0)
  #expect(crop.top.rawValue == 0)
  #expect(crop.right.rawValue == 0)
  #expect(crop.bottom.rawValue == 7_692)
  #expect(abs(crop.bottom.fraction - 0.076_92) < 0.000_000_1)
  #expect(picture.images.count == 1)
  #expect(image.contentType == .jpeg)
  #expect(image.part?.name.rawValue == "/ppt/media/image1.jpeg")
  #expect(try image.url().isFileURL)
}

// Upstream fixture: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/Presentation.pptx.
@Test func smartArtFrameResolvesAllDefiningPartsAndUnreferencedDrawingAttachment() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/Presentation.pptx")
  )
  let slide = try presentation.slide(at: 1)
  let frame = try #require(
    slide.elements.first { element in
      guard case .diagram = element.graphicContent else { return false }
      return true
    })
  guard case .diagram(let diagram) = frame.graphicContent else {
    Issue.record("Expected a typed SmartArt diagram reference.")
    return
  }

  #expect(frame.identifier == 6)
  #expect(frame.name == "Diagram 5")
  #expect(frame.spatialInfo.sourceTransform?.x.emu == 6_163_294)
  #expect(frame.spatialInfo.sourceTransform?.y.emu == 4_773_881)
  #expect(frame.spatialInfo.sourceTransform?.width.emu == 3_996_706)
  #expect(frame.spatialInfo.sourceTransform?.height.emu == 1_364_452)
  #expect(diagram.data?.part?.name.rawValue == "/ppt/diagrams/data1.xml")
  #expect(diagram.layout?.part?.name.rawValue == "/ppt/diagrams/layout1.xml")
  #expect(diagram.quickStyle?.part?.name.rawValue == "/ppt/diagrams/quickStyle1.xml")
  #expect(diagram.colors?.part?.name.rawValue == "/ppt/diagrams/colors1.xml")
  #expect(diagram.attachments.count == 4)
  #expect(try diagram.attachments.allSatisfy { try $0.url().isFileURL })

  let drawing = try #require(
    slide.localAttachments.first {
      $0.relationship.type.isEquivalent(to: .diagramDrawing)
    })
  #expect(drawing.part?.name.rawValue == "/ppt/diagrams/drawing1.xml")
  #expect(try drawing.url().isFileURL)
  #expect(slide.localAttachments.count == 6)
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Packaging.Tests/OpenXmlPackageTests.cs,
// TestOpenModel3DWrittenByPowerPoint_DotMime.
@Test func model3DFrameResolvesLazyModelAndPosterURLs() throws {
  let presentation = try OfficePresentation(
    contentsOf: FixtureCatalog.url(for: "Open-XML-SDK/TestFiles/3dtestdot.pptx")
  )
  let slide = try presentation.slide(at: 0)
  let frame = try #require(
    slide.elements.first { element in
      guard case .model3D = element.graphicContent else { return false }
      return true
    })
  guard case .model3D(let model3D) = frame.graphicContent else {
    Issue.record("Expected a typed 3D model reference.")
    return
  }
  let model = try #require(model3D.model)
  let poster = try #require(model3D.posterImage)

  #expect(frame.identifier == 4)
  #expect(frame.name == "3D Model 3")
  #expect(frame.alternativeText == "Smiling Face")
  #expect(frame.spatialInfo.sourceTransform?.x.emu == 4_606_986)
  #expect(frame.spatialInfo.sourceTransform?.y.emu == 1_920_956)
  #expect(frame.spatialInfo.sourceTransform?.width.emu == 2_978_028)
  #expect(frame.spatialInfo.sourceTransform?.height.emu == 3_016_087)
  #expect(model.part?.name.rawValue == "/ppt/media/model3d1.glb")
  #expect(model.contentType?.identifier == "org.khronos.glb")
  #expect(poster.part?.name.rawValue == "/ppt/media/image1.png")
  #expect(poster.contentType == .png)
  #expect(try model.url().isFileURL)
  #expect(try poster.url().isFileURL)
  #expect(slide.localAttachments.count == 2)
}
