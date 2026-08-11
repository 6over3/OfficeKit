# ``OfficeKit``

Read Microsoft Office Open XML documents.

## Overview

Open a file with ``OfficeDocument``. Use the format-specific types for parsed content or
``OfficePackage`` for parts, relationships, and raw XML. ``OfficeAttachment/url()`` provides access
to related files without loading them into memory.

## Topics

### Package access

- ``OfficeDocument``
- ``OfficePackage``
- ``OfficePart``
- ``OfficeRelationship``
- ``OfficeAttachment``
- ``OfficeXMLDocument``
- ``OfficeParsedXMLPart``
- ``OfficeDiagnostic``

### Semantic formats

- ``OfficePresentation``
- ``OfficeSlide``
- ``OfficeWorkbook``
- ``OfficeWorksheet``
- ``OfficeWordDocument``
- ``OfficeDocumentVisitor``

### Geometry and limits

- ``OfficeSpatialInfo``
- ``OfficeLength``
- ``OfficeParsingLimits``
- ``OfficeXMLParsingLimits``

### Guides

- <doc:SupportedFeatures>
- <doc:Diagnostics>
- <doc:CorpusValidation>
