import ExcelJS from "exceljs";

export type CustomerListResponse = {
  fileName: string;
  formattedFileName: string;
  formattedBase64: string;
  summary: {
    sourceSheet: string;
    totalRows: number;
    possibleClosedRows: number;
    columns: number;
  };
  emailSubject: string;
  emailBody: string;
};

const CUSTOMER_NUMBER_HEADER = "Customer number";
const NOTES_HEADER = "Notes";
const CLOSED_COLUMN_HINT = "last day of svc";

function normalizeHeader(value: string) {
  return value.trim().replace(/\s+/g, " ").toLowerCase();
}

function cellToText(cell: ExcelJS.Cell) {
  if (cell.type === ExcelJS.ValueType.Date && cell.value instanceof Date) {
    return `${cell.value.getMonth() + 1}/${cell.value.getDate()}/${cell.value.getFullYear()}`;
  }

  if (typeof cell.value === "object" && cell.value && "text" in cell.value) {
    return String(cell.value.text ?? "");
  }

  return cell.text?.trim() ?? "";
}

function findWorksheet(workbook: ExcelJS.Workbook) {
  return workbook.getWorksheet("Unit List") ?? workbook.worksheets[0];
}

function findHeaderRow(worksheet: ExcelJS.Worksheet) {
  for (let rowNumber = 1; rowNumber <= Math.min(20, worksheet.rowCount); rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const values: string[] = [];
    row.eachCell({ includeEmpty: false }, (cell) => values.push(cellToText(cell)));
    const normalized = values.map(normalizeHeader);
    if (normalized.some((value) => value.includes("cost ctr")) || normalized.some((value) => value.includes("customer"))) {
      return rowNumber;
    }
  }

  return 1;
}

function buildColumns(headers: string[]) {
  const filtered = headers.filter((header) => {
    const normalized = normalizeHeader(header);
    return normalized !== normalizeHeader(CUSTOMER_NUMBER_HEADER) && normalized !== normalizeHeader(NOTES_HEADER);
  });

  return [filtered[0] || "Cost Ctr Nbr", CUSTOMER_NUMBER_HEADER, ...filtered.slice(1), NOTES_HEADER];
}

function styleHeader(row: ExcelJS.Row) {
  row.font = { bold: true, color: { argb: "FFFFFFFF" } };
  row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F2937" } };
  row.alignment = { vertical: "middle", wrapText: true };
}

function styleInputCell(cell: ExcelJS.Cell) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF7CC" } };
  cell.border = {
    top: { style: "thin", color: { argb: "FFEAB308" } },
    left: { style: "thin", color: { argb: "FFEAB308" } },
    bottom: { style: "thin", color: { argb: "FFEAB308" } },
    right: { style: "thin", color: { argb: "FFEAB308" } }
  };
}

function buildSupplierEmail(input: {
  fileName: string;
  totalRows: number;
  possibleClosedRows: number;
  formattedFileName: string;
}) {
  const subject = "Customer list review requested";
  const body = [
    "Hello,",
    "",
    `Please review the attached customer list prepared from ${input.fileName}.`,
    "",
    "We added two supplier review columns:",
    "- Customer number: please enter your customer/account number for each active location.",
    "- Notes: please mark any closed/inactive locations and include any comments needed for setup or cleanup.",
    "",
    "Please also add any missing or new customers/locations that should be included for ordering.",
    "",
    "Review summary:",
    `- Customer rows included: ${input.totalRows}`,
    `- Rows with a last day of service date: ${input.possibleClosedRows}`,
    `- File to return: ${input.formattedFileName}`,
    "",
    "Please return the completed workbook when finished. If a row is no longer active, leave the row in the file and note that it is closed in the Notes column.",
    "",
    "Thank you."
  ].join("\n");

  return { subject, body };
}

export async function formatCustomerListWorkbook(input: { fileName: string; buffer: ArrayBuffer }): Promise<CustomerListResponse> {
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.load(input.buffer);

  const sourceWorksheet = findWorksheet(sourceWorkbook);
  if (!sourceWorksheet) {
    throw new Error("The workbook does not contain any worksheets.");
  }

  const headerRowNumber = findHeaderRow(sourceWorksheet);
  const sourceHeaderRow = sourceWorksheet.getRow(headerRowNumber);
  const sourceHeaders: string[] = [];
  for (let colNumber = 1; colNumber <= sourceWorksheet.columnCount; colNumber += 1) {
    const header = cellToText(sourceHeaderRow.getCell(colNumber));
    if (header) sourceHeaders.push(header);
  }

  const outputHeaders = buildColumns(sourceHeaders);
  const sourceColumnByHeader = new Map<string, number>();
  sourceHeaders.forEach((header, index) => {
    sourceColumnByHeader.set(normalizeHeader(header), index + 1);
  });

  const outputWorkbook = new ExcelJS.Workbook();
  outputWorkbook.creator = "Catalog Validator";

  const listSheet = outputWorkbook.addWorksheet("Customer List");
  listSheet.addRow(outputHeaders);
  styleHeader(listSheet.getRow(1));
  listSheet.views = [{ state: "frozen", ySplit: 1, xSplit: 2 }];
  listSheet.autoFilter = {
    from: "A1",
    to: `${listSheet.getColumn(outputHeaders.length).letter}1`
  };

  const closedColumnIndex = outputHeaders.findIndex((header) => normalizeHeader(header).includes(CLOSED_COLUMN_HINT)) + 1;
  let totalRows = 0;
  let possibleClosedRows = 0;

  for (let rowNumber = headerRowNumber + 1; rowNumber <= sourceWorksheet.rowCount; rowNumber += 1) {
    const sourceRow = sourceWorksheet.getRow(rowNumber);
    const hasAnyValue = sourceHeaders.some((header) => {
      const sourceCol = sourceColumnByHeader.get(normalizeHeader(header));
      return sourceCol ? cellToText(sourceRow.getCell(sourceCol)) !== "" : false;
    });
    if (!hasAnyValue) continue;

    const outputValues = outputHeaders.map((header) => {
      const normalized = normalizeHeader(header);
      if (normalized === normalizeHeader(CUSTOMER_NUMBER_HEADER) || normalized === normalizeHeader(NOTES_HEADER)) {
        const sourceCol = sourceColumnByHeader.get(normalized);
        return sourceCol ? sourceRow.getCell(sourceCol).value ?? "" : "";
      }
      const sourceCol = sourceColumnByHeader.get(normalized);
      return sourceCol ? sourceRow.getCell(sourceCol).value ?? "" : "";
    });

    const outputRow = listSheet.addRow(outputValues);
    totalRows += 1;

    styleInputCell(outputRow.getCell(2));
    styleInputCell(outputRow.getCell(outputHeaders.length));
    outputRow.getCell(2).note = "Supplier: enter your customer/account number for this location.";
    outputRow.getCell(outputHeaders.length).note = "Supplier: mark closed/inactive locations here and add any missing setup notes.";

    if (closedColumnIndex > 0 && cellToText(outputRow.getCell(closedColumnIndex))) {
      possibleClosedRows += 1;
      outputRow.getCell(closedColumnIndex).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
      outputRow.getCell(outputHeaders.length).note = "This row has a last day of service date. Please confirm if closed/inactive in Notes.";
    }
  }

  listSheet.columns.forEach((column) => {
    const header = String(column.header ?? "");
    column.width = Math.max(14, Math.min(34, header.length + 6));
  });
  listSheet.getColumn(2).width = 22;
  listSheet.getColumn(outputHeaders.length).width = 36;

  const instructionsSheet = outputWorkbook.addWorksheet("Instructions");
  instructionsSheet.addRows([
    ["Customer List Review"],
    ["1", "Fill in Customer number for every active customer/location."],
    ["2", "Use Notes to identify closed or inactive locations."],
    ["3", "Add any missing or new customers/locations to the bottom of the Customer List tab."],
    ["4", "Return the completed workbook to Sodexo."]
  ]);
  instructionsSheet.getCell("A1").font = { bold: true, size: 16 };
  instructionsSheet.getColumn(1).width = 8;
  instructionsSheet.getColumn(2).width = 96;

  const safeName = input.fileName.replace(/\.xlsx$/i, "");
  const formattedFileName = `${safeName}-customer-review.xlsx`;
  const email = buildSupplierEmail({
    fileName: input.fileName,
    totalRows,
    possibleClosedRows,
    formattedFileName
  });

  return {
    fileName: input.fileName,
    formattedFileName,
    formattedBase64: Buffer.from(await outputWorkbook.xlsx.writeBuffer()).toString("base64"),
    summary: {
      sourceSheet: sourceWorksheet.name,
      totalRows,
      possibleClosedRows,
      columns: outputHeaders.length
    },
    emailSubject: email.subject,
    emailBody: email.body
  };
}
