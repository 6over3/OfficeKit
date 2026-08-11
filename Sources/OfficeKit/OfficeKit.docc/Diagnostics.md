# Diagnostics

Use ``OfficePackage/validateRelationshipTargets(policy:)`` with
``OfficeValidationPolicy/recovering`` to report missing relationship targets. Each
``OfficeDiagnostic`` has a code, severity, source part, optional relationship ID, and target.

Strict validation throws the first ``OfficeKitError``. Recovering validation reports all missing
targets but does not guess replacements. Invalid XML, unsafe paths, resource-limit failures, and
ambiguous document structures always throw errors.
