import Foundation
import Testing

@testable import OfficeKit

private let compatibilityNamespace =
  "http://schemas.openxmlformats.org/markup-compatibility/2006"

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/OpenXmlDomTest/MarkupCompatibilityTest.cs,
// NoChoice_NoFallback_FullMode.
@Test func noChoiceNoFallbackFullMode() throws {
  let xml = """
    <root xmlns:mc="\(compatibilityNamespace)">
      <mc:AlternateContent/>
      <after/>
    </root>
    """
  #expect(try effectiveElementNames(xml) == ["root", "after"])
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/OpenXmlDomTest/MarkupCompatibilityTest.cs,
// OneChoice_NoFallback_FullMode.
@Test func oneChoiceNoFallbackFullMode() throws {
  let xml = """
    <root xmlns:mc="\(compatibilityNamespace)" xmlns:s="urn:supported">
      <mc:AlternateContent>
        <mc:Choice Requires="s"><s:selected/></mc:Choice>
      </mc:AlternateContent>
    </root>
    """
  #expect(
    try effectiveElementNames(xml, supportedNamespaces: ["urn:supported"])
      == ["root", "selected"]
  )
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/OpenXmlDomTest/MarkupCompatibilityTest.cs,
// MultipleChoice_NoMatches_OneFallback_FullMode.
@Test func multipleChoiceNoMatchesOneFallbackFullMode() throws {
  let xml = """
    <root xmlns:mc="\(compatibilityNamespace)" xmlns:u1="urn:unknown:1" xmlns:u2="urn:unknown:2">
      <mc:AlternateContent>
        <mc:Choice Requires="u1"><u1:first/></mc:Choice>
        <mc:Choice Requires="u2"><u2:second/></mc:Choice>
        <mc:Fallback><fallback/></mc:Fallback>
      </mc:AlternateContent>
    </root>
    """
  #expect(try effectiveElementNames(xml) == ["root", "fallback"])
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/OpenXmlDomTest/MarkupCompatibilityTest.cs,
// MultipleChoice_OneFallback_FullMode.
@Test func multipleChoiceOneFallbackFullMode() throws {
  let xml = """
    <root xmlns:mc="\(compatibilityNamespace)" xmlns:s="urn:supported" xmlns:u="urn:unknown">
      <mc:AlternateContent>
        <mc:Choice Requires="u"><u:unavailable/></mc:Choice>
        <mc:Choice Requires="s"><s:firstSupported/></mc:Choice>
        <mc:Choice Requires="s"><s:laterSupported/></mc:Choice>
        <mc:Fallback><fallback/></mc:Fallback>
      </mc:AlternateContent>
    </root>
    """
  #expect(
    try effectiveElementNames(xml, supportedNamespaces: ["urn:supported"])
      == ["root", "firstSupported"]
  )
}

@Test func nestedAlternateContentPreservesOnlyEffectiveBranches() throws {
  let xml = """
    <root xmlns:mc="\(compatibilityNamespace)" xmlns:s="urn:supported" xmlns:u="urn:unknown">
      <mc:AlternateContent>
        <mc:Choice Requires="s">
          <outer>
            <mc:AlternateContent>
              <mc:Choice Requires="u"><wrong/></mc:Choice>
              <mc:Fallback><nestedFallback/></mc:Fallback>
            </mc:AlternateContent>
          </outer>
        </mc:Choice>
        <mc:Fallback><outerFallback/></mc:Fallback>
      </mc:AlternateContent>
    </root>
    """
  #expect(
    try effectiveElementNames(xml, supportedNamespaces: ["urn:supported"])
      == ["root", "outer", "nestedFallback"]
  )
}

@Test func rawReaderStillExposesCompatibilityWrappers() throws {
  let xml = """
    <root xmlns:mc="\(compatibilityNamespace)">
      <mc:AlternateContent><mc:Fallback><fallback/></mc:Fallback></mc:AlternateContent>
    </root>
    """
  var names: [String] = []
  try OfficeXMLReader(data: Data(xml.utf8)).parse { event in
    if case .startElement(let name, _, _, _) = event { names.append(name.localName) }
  }
  #expect(names == ["root", "AlternateContent", "Fallback", "fallback"])
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/OpenXmlDomTest/MarkupCompatibilityTest.cs,
// Ignored_UnknownElement_FullMode.
@Test func ignoredUnknownElementFullMode() throws {
  let xml = """
    <root xmlns:mc="\(compatibilityNamespace)" xmlns:u="urn:unknown" mc:Ignorable="u">
      <before/>
      <u:discard><childThatMustAlsoBeDiscarded/></u:discard>
      <after/>
    </root>
    """
  #expect(try effectiveElementNames(xml) == ["root", "before", "after"])
}

// Upstream: dotnet/Open-XML-SDK @ cd2b359ef824737edb93f1c6157c19551aae1e52,
// test/DocumentFormat.OpenXml.Tests/OpenXmlDomTest/MarkupCompatibilityTest.cs,
// ProcessContent_Ignored_UnknownElement_FullMode.
@Test func processContentIgnoredUnknownElementFullMode() throws {
  let xml = """
    <root xmlns:mc="\(compatibilityNamespace)" xmlns:u="urn:unknown"
      mc:Ignorable="u" mc:ProcessContent="u:wrapper">
      <before/>
      <u:wrapper><preservedChild/></u:wrapper>
      <u:discard><discardedChild/></u:discard>
      <after/>
    </root>
    """
  #expect(
    try effectiveElementNames(xml) == ["root", "before", "preservedChild", "after"]
  )
}

@Test func supportedIgnorableNamespaceRemainsAvailableToSemanticReaders() throws {
  let xml = """
    <root xmlns:mc="\(compatibilityNamespace)" xmlns:s="urn:supported" mc:Ignorable="s">
      <s:extension><child/></s:extension>
    </root>
    """
  #expect(
    try effectiveElementNames(xml, supportedNamespaces: ["urn:supported"])
      == ["root", "extension", "child"]
  )
}

@Test func mustUnderstandRejectsUnsupportedNamespaceAndAcceptsSupportedNamespace() throws {
  let xml = """
    <root xmlns:mc="\(compatibilityNamespace)" xmlns:x="urn:required"
      mc:MustUnderstand="x"><x:value/></root>
    """
  #expect {
    try effectiveElementNames(xml)
  } throws: { error in
    guard case .invalidXML(let part, let message) = error as? OfficeKitError else {
      return false
    }
    return part == "<memory>" && message.contains("MustUnderstand")
  }
  #expect(
    try effectiveElementNames(xml, supportedNamespaces: ["urn:required"])
      == ["root", "value"]
  )
}

@Test func preserveElementsAndAttributesKeepSelectedIgnorableMarkup() throws {
  let xml = """
    <root xmlns:mc="\(compatibilityNamespace)" xmlns:u="urn:unknown"
      mc:Ignorable="u" mc:PreserveElements="u:keep" mc:PreserveAttributes="u:flag">
      <u:keep><child/></u:keep>
      <u:discard><lost/></u:discard>
      <normal u:flag="yes" u:drop="no"/>
    </root>
    """
  var names: [String] = []
  var normalAttributes: [OfficeXMLAttribute] = []
  try OfficeXMLReader(data: Data(xml.utf8)).parseCompatible(
    using: OfficeXMLCompatibilityOptions(supportedNamespaces: [])
  ) { event in
    guard case .startElement(let name, let attributes, _, _) = event else { return }
    names.append(name.localName)
    if name.localName == "normal" { normalAttributes = attributes }
  }

  #expect(names == ["root", "keep", "child", "normal"])
  #expect(normalAttributes.map(\.name.localName) == ["flag"])
  #expect(normalAttributes.first?.value == "yes")
}

private func effectiveElementNames(
  _ xml: String,
  supportedNamespaces: Set<String> = []
) throws -> [String] {
  var names: [String] = []
  let options = OfficeXMLCompatibilityOptions(supportedNamespaces: supportedNamespaces)
  try OfficeXMLReader(data: Data(xml.utf8)).parseCompatible(using: options) { event in
    if case .startElement(let name, _, _, _) = event { names.append(name.localName) }
  }
  return names
}
