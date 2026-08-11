import Foundation
import Testing

@testable import OfficeKit

@Test func emuConversionsRetainExactSourceUnits() {
  #expect(OfficeLength(emu: 914_400).inches == 1)
  #expect(OfficeLength(emu: 12_700).points == 1)
  #expect(OfficeLength(points: 72).emu == 914_400)
  #expect(OfficeLength(emu: 914_400).pixels(atDPI: 96) == 96)
}

@Test func affineTransformsComposeInApplicationOrder() {
  let scale = OfficeAffineTransform.scale(x: 2, y: 3)
  let translation = OfficeAffineTransform.translation(x: 10, y: 20)
  let composed = scale.followed(by: translation)
  #expect(composed.applying(to: OfficePoint(x: 4, y: 5)) == OfficePoint(x: 18, y: 35))
}

@Test func rotatedRectangleProducesAxisAlignedDocumentBounds() {
  let rectangle = OfficeRect(x: 0, y: 0, width: 20, height: 10)
  let bounds = rectangle.applying(.rotation(radians: .pi / 2))
  #expect(abs(bounds.minX + 10) < 0.000_001)
  #expect(abs(bounds.minY) < 0.000_001)
  #expect(abs(bounds.size.width - 10) < 0.000_001)
  #expect(abs(bounds.size.height - 20) < 0.000_001)
}
