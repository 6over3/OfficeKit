/// Severity of a recoverable Office package or semantic parsing defect.
public enum OfficeDiagnosticSeverity: String, Sendable, Hashable, Codable {
  /// Parsing continued without inventing or misassociating content.
  case warning
  /// The defect prevented reliable interpretation of affected content.
  case error
}

/// Stable identifiers for recoverable conditions reported by OfficeKit.
public enum OfficeDiagnosticCode: String, Sendable, Hashable, Codable {
  /// An internal OPC relationship points to a part that is not present.
  case danglingRelationship = "opc.danglingRelationship"
}

/// Source context for a package or semantic diagnostic.
public struct OfficeSourceReference: Sendable, Hashable, Codable {
  /// Package root (`/`) or the source part name.
  public let part: String

  /// Relationship identifier involved in the condition, when applicable.
  public let relationshipID: OfficeRelationshipID?

  /// XML qualified name involved in the condition, when available.
  public let qualifiedName: String?
}

/// A recoverable defect with stable machine-readable identity and source context.
public struct OfficeDiagnostic: Sendable, Hashable, Codable {
  /// Severity assigned by the validation policy.
  public let severity: OfficeDiagnosticSeverity
  /// Stable machine-readable diagnostic identifier.
  public let code: OfficeDiagnosticCode
  /// Caller-facing explanation of the defect.
  public let message: String
  /// Package or XML source that produced the diagnostic.
  public let source: OfficeSourceReference
  /// Raw relationship or semantic target involved, when applicable.
  public let target: String?
}

/// Whether validation throws immediately or reports conditions that are safe to recover from.
public enum OfficeValidationPolicy: String, Sendable, Hashable, Codable {
  /// Throw immediately on the first invalid condition.
  case strict
  /// Return diagnostics for defects that are safe to isolate.
  case recovering
}
