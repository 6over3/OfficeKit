/// A relationship-backed chart that is parsed only when requested.
public struct OfficeChartReference: Sendable {
  private let package: OfficePackage

  /// The relationship-backed chart attachment.
  public let attachment: OfficeAttachment

  /// The chart XML part.
  public let part: OfficePart

  package init(package: OfficePackage, attachment: OfficeAttachment, part: OfficePart) {
    self.package = package
    self.attachment = attachment
    self.part = part
  }

  /// Parses chart series and cached values without loading related workbooks into memory.
  public func chart() throws -> OfficeChart {
    try PresentationChartParser.parse(part: part, package: package)
  }
}

/// A broad semantic chart family.
public enum OfficeChartKind: String, Sendable, Hashable, Codable {
  /// An area chart.
  case area
  /// A bar or column chart.
  case bar
  /// A bubble chart.
  case bubble
  /// A doughnut chart.
  case doughnut
  /// A line chart.
  case line
  /// A pie chart.
  case pie
  /// A radar chart.
  case radar
  /// An XY scatter chart.
  case scatter
  /// A stock chart.
  case stock
  /// A surface chart.
  case surface
  /// A chart family that OfficeKit does not currently classify.
  case unknown
}

/// The semantic role of an axis in a chart plot area.
public enum OfficeChartAxisKind: String, Sendable, Hashable, Codable {
  /// A categorical axis.
  case category
  /// A date-scaled axis.
  case date
  /// A series axis used by three-dimensional charts.
  case series
  /// A numeric value axis.
  case value
}

/// The authored edge on which a chart axis is positioned.
public enum OfficeChartAxisPosition: String, Sendable, Hashable, Codable {
  /// The bottom edge of the plot area.
  case bottom = "b"
  /// The left edge of the plot area.
  case left = "l"
  /// The right edge of the plot area.
  case right = "r"
  /// The top edge of the plot area.
  case top = "t"
}

/// A chart axis with its authored scale and crossing information.
public struct OfficeChartAxis: Sendable, Hashable, Codable {
  /// The axis role.
  public let kind: OfficeChartAxisKind

  /// The identifier used to connect axes within the plot area.
  public let identifier: UInt32?

  /// The edge on which the axis is positioned.
  public let position: OfficeChartAxisPosition?

  /// The identifier of the axis crossed by this axis.
  public let crossingAxisIdentifier: UInt32?

  /// The axis title as plain text, when present.
  public let title: String?

  /// The authored number-format code, when present.
  public let numberFormatCode: String?

  /// Whether the number format follows linked source data.
  public let numberFormatIsSourceLinked: Bool?

  /// An explicitly authored lower scale bound.
  public let minimum: Double?

  /// An explicitly authored upper scale bound.
  public let maximum: Double?

  /// An explicitly authored logarithm base.
  public let logarithmBase: Double?

  /// Whether the axis runs from maximum to minimum.
  public let isReversed: Bool

  /// The authored major unit.
  public let majorUnit: Double?

  /// The authored minor unit.
  public let minorUnit: Double?

  /// The producer spelling for automatic crossing, such as `autoZero`, when present.
  public let crossing: String?

  /// An explicit numeric crossing value.
  public let crossingValue: Double?
}

/// A normalized manual chart layout expressed as fractions of its containing chart.
public struct OfficeChartLayout: Sendable, Hashable, Codable {
  /// The producer spelling for the layout target, such as `inner` or `outer`.
  public let target: String?

  /// The producer spelling for horizontal positioning, such as `edge` or `factor`.
  public let horizontalMode: String?

  /// The producer spelling for vertical positioning, such as `edge` or `factor`.
  public let verticalMode: String?

  /// The authored horizontal coordinate.
  public let x: Double?

  /// The authored vertical coordinate.
  public let y: Double?

  /// The authored width.
  public let width: Double?

  /// The authored height.
  public let height: Double?
}

/// An authored chart legend.
public struct OfficeChartLegend: Sendable, Hashable, Codable {
  /// The producer spelling for its position, such as `r`, `l`, `t`, or `b`.
  public let position: String?

  /// Whether the legend overlays the plot area.
  public let overlaysPlotArea: Bool?

  /// The legend's manual fractional layout, when authored.
  public let layout: OfficeChartLayout?
}

/// A parsed DrawingML chart with cached series data and lazy related resources.
public struct OfficeChart: Sendable {
  /// The chart XML part.
  public let sourcePart: OfficePart

  /// The broad chart family.
  public let kind: OfficeChartKind

  /// The exact plot element name, such as `bar3DChart`.
  public let sourceKind: String?

  /// Plain chart-title text, when present.
  public let title: String?

  /// The numeric chart style identifier authored by the producer.
  public let styleIdentifier: UInt32?

  /// The producer spelling for series grouping, such as `clustered` or `stacked`.
  public let grouping: String?

  /// The producer spelling for bar direction, such as `bar` or `col`.
  public let barDirection: String?

  /// The plot area's manual fractional layout, when authored.
  public let plotAreaLayout: OfficeChartLayout?

  /// Axes in authored order.
  public let axes: [OfficeChartAxis]

  /// The chart legend, when present.
  public let legend: OfficeChartLegend?

  /// Whether only visible source cells contribute to the plot.
  public let plotsVisibleCellsOnly: Bool?

  /// The producer spelling for blank-cell display, such as `gap`, `zero`, or `span`.
  public let displayBlanksAs: String?

  /// Series in authored order.
  public let series: [OfficeChartSeries]

  /// Relationship-backed resources such as an embedded source workbook.
  public let attachments: [OfficeAttachment]
}

/// One chart series with source formulas and cached display values.
public struct OfficeChartSeries: Sendable {
  /// The producer's series index, when declared.
  public let index: UInt32?

  /// The producer's plot order, when declared.
  public let order: UInt32?

  /// The cached series name.
  public let name: String?

  /// The source formula for the series name.
  public let nameFormula: String?

  /// Cached category labels in source order.
  public let categories: [String]

  /// The source formula for category labels.
  public let categoryFormula: String?

  /// Cached numeric values in source order. Invalid lexical values remain `nil`.
  public let values: [Double?]

  /// The source formula for numeric values.
  public let valueFormula: String?
}
