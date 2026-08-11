import Foundation

package enum PresentationChartParser {
  package static func parse(part: OfficePart, package: OfficePackage) throws -> OfficeChart {
    let session = ChartParsingSession()
    try package.parseXML(in: part, compatibility: .commonOffice, session.consume)
    let attachments = try package.relationships(from: .part(part.name)).map {
      package.attachment(referencedBy: $0)
    }
    return OfficeChart(
      sourcePart: part,
      kind: session.kind,
      sourceKind: session.sourceKind,
      title: session.title.isEmpty ? nil : session.title,
      styleIdentifier: session.styleIdentifier,
      grouping: session.grouping,
      barDirection: session.barDirection,
      plotAreaLayout: session.plotAreaLayout,
      axes: session.axes,
      legend: session.legend,
      plotsVisibleCellsOnly: session.plotsVisibleCellsOnly,
      displayBlanksAs: session.displayBlanksAs,
      series: session.series,
      attachments: attachments
    )
  }
}

private final class ChartParsingSession {
  private static let chartNamespace =
    "http://schemas.openxmlformats.org/drawingml/2006/chart"
  private static let drawingNamespace =
    "http://schemas.openxmlformats.org/drawingml/2006/main"

  enum SeriesContext {
    case name
    case categories
    case values
  }

  enum TextTarget {
    case formula
    case value
    case title
    case axisTitle
  }

  enum LayoutTarget {
    case plotArea
    case legend
  }

  struct LayoutBuilder {
    let depth: Int
    let target: LayoutTarget
    var layoutTarget: String?
    var horizontalMode: String?
    var verticalMode: String?
    var x: Double?
    var y: Double?
    var width: Double?
    var height: Double?

    var value: OfficeChartLayout {
      OfficeChartLayout(
        target: layoutTarget,
        horizontalMode: horizontalMode,
        verticalMode: verticalMode,
        x: x,
        y: y,
        width: width,
        height: height
      )
    }
  }

  struct AxisBuilder {
    let depth: Int
    let kind: OfficeChartAxisKind
    var identifier: UInt32?
    var position: OfficeChartAxisPosition?
    var crossingAxisIdentifier: UInt32?
    var title = ""
    var numberFormatCode: String?
    var numberFormatIsSourceLinked: Bool?
    var minimum: Double?
    var maximum: Double?
    var logarithmBase: Double?
    var isReversed = false
    var majorUnit: Double?
    var minorUnit: Double?
    var crossing: String?
    var crossingValue: Double?

    var value: OfficeChartAxis {
      OfficeChartAxis(
        kind: kind,
        identifier: identifier,
        position: position,
        crossingAxisIdentifier: crossingAxisIdentifier,
        title: title.isEmpty ? nil : title,
        numberFormatCode: numberFormatCode,
        numberFormatIsSourceLinked: numberFormatIsSourceLinked,
        minimum: minimum,
        maximum: maximum,
        logarithmBase: logarithmBase,
        isReversed: isReversed,
        majorUnit: majorUnit,
        minorUnit: minorUnit,
        crossing: crossing,
        crossingValue: crossingValue
      )
    }
  }

  struct SeriesBuilder {
    let depth: Int
    var index: UInt32?
    var order: UInt32?
    var name: String?
    var nameFormula: String?
    var categories: [String] = []
    var categoryFormula: String?
    var values: [Double?] = []
    var valueFormula: String?
  }

  var depth = 0
  var plotAreaDepth: Int?
  var chartDepth: Int?
  var titleDepth: Int?
  var axisTitleDepth: Int?
  var title = ""
  var styleIdentifier: UInt32?
  var kind: OfficeChartKind = .unknown
  var sourceKind: String?
  var chartKindDepth: Int?
  var grouping: String?
  var barDirection: String?
  var plotAreaLayout: OfficeChartLayout?
  var axes: [OfficeChartAxis] = []
  var axisBuilder: AxisBuilder?
  var legendDepth: Int?
  var legendPosition: String?
  var legendOverlaysPlotArea: Bool?
  var legendLayout: OfficeChartLayout?
  var legend: OfficeChartLegend?
  var layoutBuilder: LayoutBuilder?
  var plotsVisibleCellsOnly: Bool?
  var displayBlanksAs: String?
  var series: [OfficeChartSeries] = []
  var seriesBuilder: SeriesBuilder?
  var context: SeriesContext?
  var contextDepth: Int?
  var textTarget: TextTarget?
  var textDepth: Int?
  var textBuffer = ""

  func consume(_ event: OfficeXMLEvent) {
    switch event {
    case .startElement(let name, let attributes, _, _):
      depth += 1
      if name.namespaceURI == Self.chartNamespace, name.localName == "chart" {
        chartDepth = depth
      } else if name.namespaceURI == Self.chartNamespace, name.localName == "style",
        chartDepth == nil
      {
        styleIdentifier = attribute("val", in: attributes).flatMap(UInt32.init)
      } else if name.namespaceURI == Self.chartNamespace, name.localName == "plotArea" {
        plotAreaDepth = depth
      } else if plotAreaDepth != nil, name.namespaceURI == Self.chartNamespace,
        let chartKind = chartKind(for: name.localName)
      {
        kind = chartKind
        sourceKind = name.localName
        chartKindDepth = depth
      }

      if name.namespaceURI == Self.chartNamespace, name.localName == "title" {
        if axisBuilder == nil {
          titleDepth = depth
        } else {
          axisTitleDepth = depth
        }
      }
      if axisTitleDepth != nil, name.namespaceURI == Self.drawingNamespace,
        name.localName == "t"
      {
        startText(.axisTitle)
      } else if titleDepth != nil, name.namespaceURI == Self.drawingNamespace,
        name.localName == "t"
      {
        startText(.title)
      }

      if plotAreaDepth != nil, name.namespaceURI == Self.chartNamespace,
        let axisKind = axisKind(for: name.localName)
      {
        axisBuilder = AxisBuilder(depth: depth, kind: axisKind)
      }
      consumeAxisElement(name: name, attributes: attributes)

      if name.namespaceURI == Self.chartNamespace, name.localName == "legend" {
        legendDepth = depth
      } else if legendDepth != nil, name.namespaceURI == Self.chartNamespace,
        name.localName == "legendPos"
      {
        legendPosition = attribute("val", in: attributes)
      } else if legendDepth != nil, name.namespaceURI == Self.chartNamespace,
        name.localName == "overlay"
      {
        legendOverlaysPlotArea = attribute("val", in: attributes)
          .flatMap(OfficeValueDecoder.boolean)
      }

      if name.namespaceURI == Self.chartNamespace, name.localName == "manualLayout" {
        if legendDepth != nil {
          layoutBuilder = LayoutBuilder(depth: depth, target: .legend)
        } else if plotAreaDepth != nil {
          layoutBuilder = LayoutBuilder(depth: depth, target: .plotArea)
        }
      }
      consumeLayoutElement(name: name, attributes: attributes)

      if chartKindDepth != nil, name.namespaceURI == Self.chartNamespace,
        name.localName == "grouping"
      {
        grouping = attribute("val", in: attributes)
      } else if chartKindDepth != nil, name.namespaceURI == Self.chartNamespace,
        name.localName == "barDir"
      {
        barDirection = attribute("val", in: attributes)
      } else if chartDepth != nil, name.namespaceURI == Self.chartNamespace,
        name.localName == "plotVisOnly"
      {
        plotsVisibleCellsOnly = attribute("val", in: attributes)
          .flatMap(OfficeValueDecoder.boolean)
      } else if chartDepth != nil, name.namespaceURI == Self.chartNamespace,
        name.localName == "dispBlanksAs"
      {
        displayBlanksAs = attribute("val", in: attributes)
      }

      if name.namespaceURI == Self.chartNamespace, name.localName == "ser" {
        seriesBuilder = SeriesBuilder(depth: depth)
      }
      guard seriesBuilder != nil else { return }
      if name.namespaceURI == Self.chartNamespace, name.localName == "idx" {
        seriesBuilder?.index = attribute("val", in: attributes).flatMap(UInt32.init)
      } else if name.namespaceURI == Self.chartNamespace, name.localName == "order" {
        seriesBuilder?.order = attribute("val", in: attributes).flatMap(UInt32.init)
      } else if name.namespaceURI == Self.chartNamespace, name.localName == "tx" {
        context = .name
        contextDepth = depth
      } else if name.namespaceURI == Self.chartNamespace, name.localName == "cat" {
        context = .categories
        contextDepth = depth
      } else if name.namespaceURI == Self.chartNamespace, name.localName == "val" {
        context = .values
        contextDepth = depth
      } else if name.namespaceURI == Self.chartNamespace, name.localName == "f",
        context != nil
      {
        startText(.formula)
      } else if name.namespaceURI == Self.chartNamespace, name.localName == "v",
        context != nil
      {
        startText(.value)
      }

    case .text(let text, _):
      guard textTarget != nil else { return }
      textBuffer.append(text)

    case .endElement(let name, _):
      if textDepth == depth, let textTarget {
        finishText(textTarget)
      }
      if axisTitleDepth == depth, name.namespaceURI == Self.chartNamespace,
        name.localName == "title"
      {
        axisTitleDepth = nil
      }
      if contextDepth == depth, name.namespaceURI == Self.chartNamespace {
        context = nil
        contextDepth = nil
      }
      if let completed = seriesBuilder, completed.depth == depth,
        name.namespaceURI == Self.chartNamespace, name.localName == "ser"
      {
        series.append(
          OfficeChartSeries(
            index: completed.index,
            order: completed.order,
            name: completed.name,
            nameFormula: completed.nameFormula,
            categories: completed.categories,
            categoryFormula: completed.categoryFormula,
            values: completed.values,
            valueFormula: completed.valueFormula
          )
        )
        seriesBuilder = nil
      }
      if let completed = layoutBuilder, completed.depth == depth,
        name.namespaceURI == Self.chartNamespace, name.localName == "manualLayout"
      {
        switch completed.target {
        case .plotArea: plotAreaLayout = completed.value
        case .legend: legendLayout = completed.value
        }
        layoutBuilder = nil
      }
      if let completed = axisBuilder, completed.depth == depth,
        name.namespaceURI == Self.chartNamespace,
        axisKind(for: name.localName) != nil
      {
        axes.append(completed.value)
        axisBuilder = nil
      }
      if legendDepth == depth, name.namespaceURI == Self.chartNamespace,
        name.localName == "legend"
      {
        legend = OfficeChartLegend(
          position: legendPosition,
          overlaysPlotArea: legendOverlaysPlotArea,
          layout: legendLayout
        )
        legendDepth = nil
      }
      if titleDepth == depth, name.namespaceURI == Self.chartNamespace,
        name.localName == "title"
      {
        titleDepth = nil
      }
      if plotAreaDepth == depth, name.namespaceURI == Self.chartNamespace,
        name.localName == "plotArea"
      {
        plotAreaDepth = nil
      }
      if chartKindDepth == depth, name.namespaceURI == Self.chartNamespace {
        chartKindDepth = nil
      }
      if chartDepth == depth, name.namespaceURI == Self.chartNamespace,
        name.localName == "chart"
      {
        chartDepth = nil
      }
      depth -= 1

    case .startDocument, .endDocument:
      break
    }
  }

  private func startText(_ target: TextTarget) {
    textTarget = target
    textDepth = depth
    textBuffer = ""
  }

  private func finishText(_ target: TextTarget) {
    let value = textBuffer.trimmingCharacters(in: .whitespacesAndNewlines)
    switch target {
    case .title:
      title.append(value)
    case .axisTitle:
      axisBuilder?.title.append(value)
    case .formula:
      switch context {
      case .name: seriesBuilder?.nameFormula = value
      case .categories: seriesBuilder?.categoryFormula = value
      case .values: seriesBuilder?.valueFormula = value
      case nil: break
      }
    case .value:
      switch context {
      case .name: seriesBuilder?.name = value
      case .categories: seriesBuilder?.categories.append(value)
      case .values: seriesBuilder?.values.append(Double(value))
      case nil: break
      }
    }
    textTarget = nil
    textDepth = nil
    textBuffer = ""
  }

  private func chartKind(for localName: String) -> OfficeChartKind? {
    switch localName {
    case "areaChart", "area3DChart": .area
    case "barChart", "bar3DChart": .bar
    case "bubbleChart": .bubble
    case "doughnutChart": .doughnut
    case "lineChart", "line3DChart": .line
    case "pieChart", "pie3DChart", "ofPieChart": .pie
    case "radarChart": .radar
    case "scatterChart": .scatter
    case "stockChart": .stock
    case "surfaceChart", "surface3DChart": .surface
    default: nil
    }
  }

  private func axisKind(for localName: String) -> OfficeChartAxisKind? {
    switch localName {
    case "catAx": .category
    case "dateAx": .date
    case "serAx": .series
    case "valAx": .value
    default: nil
    }
  }

  private func consumeAxisElement(
    name: OfficeXMLName,
    attributes: [OfficeXMLAttribute]
  ) {
    guard axisBuilder != nil, name.namespaceURI == Self.chartNamespace else { return }
    let value = attribute("val", in: attributes)
    switch name.localName {
    case "axId": axisBuilder?.identifier = value.flatMap(UInt32.init)
    case "axPos": axisBuilder?.position = value.flatMap(OfficeChartAxisPosition.init(rawValue:))
    case "crossAx": axisBuilder?.crossingAxisIdentifier = value.flatMap(UInt32.init)
    case "numFmt":
      axisBuilder?.numberFormatCode = attribute("formatCode", in: attributes)
      axisBuilder?.numberFormatIsSourceLinked = attribute("sourceLinked", in: attributes)
        .flatMap(OfficeValueDecoder.boolean)
    case "min": axisBuilder?.minimum = value.flatMap(Double.init)
    case "max": axisBuilder?.maximum = value.flatMap(Double.init)
    case "logBase": axisBuilder?.logarithmBase = value.flatMap(Double.init)
    case "orientation": axisBuilder?.isReversed = value == "maxMin"
    case "majorUnit": axisBuilder?.majorUnit = value.flatMap(Double.init)
    case "minorUnit": axisBuilder?.minorUnit = value.flatMap(Double.init)
    case "crosses": axisBuilder?.crossing = value
    case "crossesAt": axisBuilder?.crossingValue = value.flatMap(Double.init)
    default: break
    }
  }

  private func consumeLayoutElement(
    name: OfficeXMLName,
    attributes: [OfficeXMLAttribute]
  ) {
    guard layoutBuilder != nil, name.namespaceURI == Self.chartNamespace else { return }
    let value = attribute("val", in: attributes)
    switch name.localName {
    case "layoutTarget": layoutBuilder?.layoutTarget = value
    case "xMode": layoutBuilder?.horizontalMode = value
    case "yMode": layoutBuilder?.verticalMode = value
    case "x": layoutBuilder?.x = value.flatMap(Double.init)
    case "y": layoutBuilder?.y = value.flatMap(Double.init)
    case "w": layoutBuilder?.width = value.flatMap(Double.init)
    case "h": layoutBuilder?.height = value.flatMap(Double.init)
    default: break
    }
  }

  private func attribute(
    _ localName: String,
    in attributes: [OfficeXMLAttribute]
  ) -> String? {
    attributes.first { $0.name.namespaceURI == nil && $0.name.localName == localName }?.value
  }
}
