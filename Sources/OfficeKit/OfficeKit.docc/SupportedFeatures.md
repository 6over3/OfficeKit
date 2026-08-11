# Supported Features

OfficeKit reads ZIP and Flat OPC Word, Excel, and PowerPoint files. The typed APIs expose document
content, relationships, attachments, and authored geometry.

All XML parts are also available as events or complete trees. Unknown elements and attributes stay
available through these APIs. Attachments use local file URLs and are not loaded into `Data` first.

OfficeKit does not edit, render, evaluate formulas, run active content, decrypt files, or download
external resources.
