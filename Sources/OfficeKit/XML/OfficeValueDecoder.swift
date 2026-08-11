import Foundation

/// Locale-independent decoders for lexical values used by Office Open XML.
public enum OfficeValueDecoder {
  /// Decodes the XML Schema boolean spellings `true`, `false`, `1`, and `0`.
  public static func boolean(_ text: String) -> Bool? {
    if let value = Bool(text) { return value }
    if text == "1" { return true }
    if text == "0" { return false }
    return nil
  }

  /// Decodes a finite or XML Schema special double without consulting the current locale.
  public static func double(_ text: String) -> Double? {
    Double(text)
  }

  /// Decodes a base-ten integer, including integral exponent notation used by spreadsheet cells.
  public static func integer(_ text: String) -> Int? {
    if let integer = Int(text) { return integer }
    guard text.contains("e") || text.contains("E"),
      let value = Double(text),
      value.isFinite,
      value.rounded(.towardZero) == value,
      value >= Double(Int.min),
      value < Double(Int.max) else { return nil }
    let integer = Int(value)
    guard let decimal = Decimal(string: text, locale: Locale(identifier: "en_US_POSIX")) else {
      return nil
    }
    guard Decimal(integer) == decimal else { return nil }
    return integer
  }

  /// Decodes a base-ten decimal, including exponent notation, using POSIX punctuation.
  public static func decimal(_ text: String) -> Decimal? {
    guard let lexicalValue = Double(text), lexicalValue.isFinite else { return nil }
    return Decimal(string: text, locale: Locale(identifier: "en_US_POSIX"))
  }

  /// Decodes an even-length hexadecimal byte string.
  public static func hexadecimalData(_ text: String) -> Data? {
    guard text.utf8.count.isMultiple(of: 2) else { return nil }
    var data = Data()
    data.reserveCapacity(text.utf8.count / 2)
    var highNibble: UInt8?
    for byte in text.utf8 {
      guard let nibble = hexadecimalNibble(byte) else { return nil }
      if let high = highNibble {
        data.append(high << 4 | nibble)
        highNibble = nil
      } else {
        highNibble = nibble
      }
    }
    return data
  }

  /// Decodes a base-64 byte string without accepting unknown characters.
  public static func base64Data(_ text: String) -> Data? {
    Data(base64Encoded: text)
  }

  private static func hexadecimalNibble(_ byte: UInt8) -> UInt8? {
    switch byte {
    case 0x30...0x39: byte - 0x30
    case 0x41...0x46: byte - 0x41 + 10
    case 0x61...0x66: byte - 0x61 + 10
    default: nil
    }
  }

}
