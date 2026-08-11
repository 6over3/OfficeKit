import Foundation
import Testing

@testable import OfficeKit

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/Spreadsheet/CellValueTests.cs,
// CellDoubleTest and CellDoubleTestExponential.
@Test(arguments: [
  ("-1.5", -1.5),
  ("-1.0", -1.0),
  ("0.0", 0.0),
  ("1.0", 1.0),
  ("1.5", 1.5),
  ("987.6E+30", 9.876E+32),
  ("-12.34E-20", -1.234E-19),
])
func cellDoubleTest(text: String, expected: Double) {
  #expect(OfficeValueDecoder.double(text) == expected)
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/Spreadsheet/CellValueTests.cs,
// CellDoubleTestFalse.
@Test(arguments: ["", "other", "1,5", " 1", "1 "])
func cellDoubleTestFalse(text: String) {
  #expect(OfficeValueDecoder.double(text) == nil)
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/Spreadsheet/CellValueTests.cs,
// CellIntTest and CellIntTestExponential.
@Test(arguments: [Int.min, Int.min + 1, -1, 0, 1, Int.max - 1, Int.max])
func cellIntTest(value: Int) {
  #expect(OfficeValueDecoder.integer(String(value)) == value)
}

@Test func cellIntTestExponential() {
  #expect(OfficeValueDecoder.integer("987E+5") == 98_700_000)
  #expect(OfficeValueDecoder.integer("1.5E+1") == 15)
  #expect(OfficeValueDecoder.integer("1.5") == nil)
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/Spreadsheet/CellValueTests.cs,
// CellDecimalTestExponential and CellDecimalTestNegative.
@Test func cellDecimalTestExponential() {
  #expect(OfficeValueDecoder.decimal("987.6E+8") == Decimal(string: "98760000000"))
  #expect(OfficeValueDecoder.decimal("-12.34E-7") == Decimal(string: "-0.000001234"))
  #expect(OfficeValueDecoder.decimal("other") == nil)
  #expect(OfficeValueDecoder.decimal("") == nil)
  #expect(OfficeValueDecoder.decimal("1,2") == nil)
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/Spreadsheet/CellValueTests.cs,
// CellBooleanTest and CellBooleanTestNegative.
@Test(arguments: [("0", false), ("false", false), ("1", true), ("true", true)])
func cellBooleanTest(text: String, expected: Bool) {
  #expect(OfficeValueDecoder.boolean(text) == expected)
}

@Test(arguments: ["", "other", "False", "True"])
func cellBooleanTestNegative(text: String) {
  #expect(OfficeValueDecoder.boolean(text) == nil)
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/SimpleTypes/HexBinaryValueTests.cs,
// ValidateValue and GetBytes.
@Test func hexBinaryValueGetBytes() {
  #expect(OfficeValueDecoder.hexadecimalData("") == Data())
  #expect(OfficeValueDecoder.hexadecimalData("00") == Data([0]))
  #expect(OfficeValueDecoder.hexadecimalData("01") == Data([1]))
  #expect(OfficeValueDecoder.hexadecimalData("FF") == Data([0xFF]))
  #expect(OfficeValueDecoder.hexadecimalData("FF01") == Data([0xFF, 0x01]))
  #expect(OfficeValueDecoder.hexadecimalData("FFF") == nil)
  #expect(OfficeValueDecoder.hexadecimalData("zz") == nil)
}

@Test func xmlSchemaDoubleSpecialValues() {
  #expect(OfficeValueDecoder.double("INF") == .infinity)
  #expect(OfficeValueDecoder.double("-INF") == -.infinity)
  #expect(OfficeValueDecoder.double("NaN")?.isNaN == true)
}
