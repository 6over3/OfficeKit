#!/bin/sh
set -eu

output=${1:-.build/million-rows.xlsx}
row_count=${2:-1000000}

case "$row_count" in
  ''|*[!0-9]*) echo "ROW_COUNT must be a positive integer." >&2; exit 2 ;;
esac
if [ "$row_count" -lt 1 ]; then
  echo "ROW_COUNT must be a positive integer." >&2
  exit 2
fi

work_directory=$(mktemp -d "${TMPDIR:-/tmp}/officekit-large-sheet.XXXXXX")
trap 'rm -rf "$work_directory"' EXIT HUP INT TERM
mkdir -p "$work_directory/_rels" "$work_directory/xl/_rels" \
  "$work_directory/xl/worksheets"

printf '%s\n' '<?xml version="1.0" encoding="UTF-8"?>' \
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' \
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' \
  '<Default Extension="xml" ContentType="application/xml"/>' \
  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' \
  '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' \
  '</Types>' > "$work_directory/[Content_Types].xml"

printf '%s\n' '<?xml version="1.0" encoding="UTF-8"?>' \
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' \
  '</Relationships>' > "$work_directory/_rels/.rels"

printf '%s\n' '<?xml version="1.0" encoding="UTF-8"?>' \
  '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' \
  '<sheets><sheet name="Million Rows" sheetId="1" r:id="rId1"/></sheets>' \
  '</workbook>' > "$work_directory/xl/workbook.xml"

printf '%s\n' '<?xml version="1.0" encoding="UTF-8"?>' \
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' \
  '</Relationships>' > "$work_directory/xl/_rels/workbook.xml.rels"

awk -v rows="$row_count" 'BEGIN {
  print "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
  print "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
  for (row = 1; row <= rows; row++) {
    printf "<row r=\"%d\"><c r=\"A%d\"><v>%d</v></c></row>\n", row, row, row
  }
  print "</sheetData></worksheet>"
}' > "$work_directory/xl/worksheets/sheet1.xml"

output_directory=$(dirname "$output")
mkdir -p "$output_directory"
output=$(cd "$output_directory" && pwd)/$(basename "$output")
(cd "$work_directory" && zip -q -r "$output" .)
printf '%s\n' "$output"
