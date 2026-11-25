"use strict";

const spreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const documentRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
const xlsxMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function escapeExcelText(text) {
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function excelTripDate(dateTime) {
  const [date, time] = dateTime.split(" ");
  const days = (Date.parse(date) - Date.UTC(1899, 11, 31)) / 86_400_000;
  // Excel's 1900 calendar includes a fictitious February 29.
  const serial = days + (date >= "1900-03-01" ? 1 : 0);
  return time ? serial + minutesOfDay(time) / 1440 : serial;
}

function tripWorksheet({ description, items, notes, sum, departure, returnDate }) {
  const rows = [];
  const merged = [];

  function addRow(cells, { height = 20, merge = false } = {}) {
    const rowNumber = rows.length + 1;
    const rowCells = cells.map(([value, style = 0], index) => {
      const reference = `${index === 0 ? "A" : "B"}${rowNumber}`;
      if (typeof value === "number") return `<c r="${reference}" s="${style}"><v>${value}</v></c>`;
      return `<c r="${reference}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${escapeExcelText(value)}</t></is></c>`;
    });
    rows.push(`<row r="${rowNumber}" ht="${height}" customHeight="1">${rowCells.join("")}</row>`);
    if (merge) merged.push(`<mergeCell ref="A${rowNumber}:B${rowNumber}"/>`);
  }

  addRow([["예상 출장여비", 1]], { height: 32, merge: true });
  addRow([[description, 8]], { height: 30, merge: true });
  addRow([["출장일"], [excelTripDate(departure), departure.includes(" ") ? 7 : 6]]);
  addRow([["복귀일"], [excelTripDate(returnDate), returnDate.includes(" ") ? 7 : 6]]);
  addRow([["예상 합계 (원)", 4], [sum, 5]], { height: 25 });
  addRow([], { height: 10 });
  addRow([["항목", 2], ["금액 (원)", 2]]);
  for (const [label, amount] of items) addRow([[label], [amount, 3]]);
  addRow([], { height: 10 });
  addRow([["적용 근거", 2]], { merge: true });
  addRow([["공무원 여비 규정 (2026. 10. 2. 시행)", 8]], { merge: true });
  for (const note of notes) {
    addRow([[note, 8]], { height: note.length > 52 ? 45 : 28, merge: true });
  }

  return `<worksheet xmlns="${spreadsheetNamespace}">
    <cols><col min="1" max="1" width="44" customWidth="1"/><col min="2" max="2" width="28" customWidth="1"/></cols>
    <sheetData>${rows.join("")}</sheetData>
    <mergeCells count="${merged.length}">${merged.join("")}</mergeCells>
  </worksheet>`;
}

const tripWorkbookStyles = `<styleSheet xmlns="${spreadsheetNamespace}">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
    <numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm"/>
  </numFmts>
  <fonts count="3">
    <font><sz val="11"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="17"/><color rgb="FF132F29"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="11"/><color rgb="FF132F29"/><name val="맑은 고딕"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE9F4EE"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD3EDE1"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="9">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="3" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

async function downloadTripWorkbook(result) {
  const zip = new JSZip();
  const contentTypes = "application/vnd.openxmlformats-officedocument.spreadsheetml";
  zip.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="${contentTypes}.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="${contentTypes}.worksheet+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="${contentTypes}.styles+xml"/>
  </Types>`);
  zip.file("_rels/.rels", `<Relationships xmlns="${packageRelationships}">
    <Relationship Id="rId1" Type="${documentRelationships}/officeDocument" Target="xl/workbook.xml"/>
  </Relationships>`);
  zip.file("xl/workbook.xml", `<workbook xmlns="${spreadsheetNamespace}" xmlns:r="${documentRelationships}">
    <sheets><sheet name="출장여비" sheetId="1" r:id="rId1"/></sheets>
  </workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<Relationships xmlns="${packageRelationships}">
    <Relationship Id="rId1" Type="${documentRelationships}/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="${documentRelationships}/styles" Target="styles.xml"/>
  </Relationships>`);
  zip.file("xl/worksheets/sheet1.xml", tripWorksheet(result));
  zip.file("xl/styles.xml", tripWorkbookStyles);

  const file = await zip.generateAsync({ type: "blob", mimeType: xlsxMime, compression: "DEFLATE" });
  const url = URL.createObjectURL(file);
  const link = document.createElement("a");
  link.href = url;
  link.download = `출장여비_${result.departure.slice(0, 10)}_${result.returnDate.slice(0, 10)}.xlsx`;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
