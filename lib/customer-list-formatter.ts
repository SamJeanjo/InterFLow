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
const CATALOG_HEADER_HINTS = ["item sku", "item description", "unit price", "currency"];
const CUSTOMER_LIST_HEADER_HINTS = ["cost ctr nbr", "cost ctr nm"];

type SourceColumn = {
  header: string;
  sourceCol?: number;
  kind?: "customerNumber" | "notes";
};

function normalizeHeader(value: string) {
  return value.trim().replace(/\s+/g, " ").toLowerCase();
}

function cloneStyle<T>(value: T | undefined): T | undefined {
  if (value === undefined) return undefined;
  return JSON.parse(JSON.stringify(value)) as T;
}

function cloneCellValue(cell: ExcelJS.Cell): ExcelJS.CellValue {
  if (cell.value instanceof Date) {
    return new Date(cell.value.getTime());
  }

  if (typeof cell.value === "object" && cell.value !== null) {
    return JSON.parse(JSON.stringify(cell.value)) as ExcelJS.CellValue;
  }

  return cell.value;
}

function copyCellFormat(sourceCell: ExcelJS.Cell, targetCell: ExcelJS.Cell) {
  targetCell.style = cloneStyle(sourceCell.style) ?? {};
  if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt;
  const alignment = cloneStyle(sourceCell.alignment);
  const border = cloneStyle(sourceCell.border);
  const fill = cloneStyle(sourceCell.fill);
  const font = cloneStyle(sourceCell.font);
  const protection = cloneStyle(sourceCell.protection);
  if (alignment) targetCell.alignment = alignment;
  if (border) targetCell.border = border;
  if (fill) targetCell.fill = fill;
  if (font) targetCell.font = font;
  if (protection) targetCell.protection = protection;
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
  const unitList = workbook.getWorksheet("Unit List");
  if (unitList) return unitList;

  return workbook.worksheets.find((worksheet) => {
    const headerRowNumber = findHeaderRow(worksheet);
    const row = worksheet.getRow(headerRowNumber);
    const headers: string[] = [];
    row.eachCell({ includeEmpty: false }, (cell) => headers.push(normalizeHeader(cellToText(cell))));
    return CUSTOMER_LIST_HEADER_HINTS.some((hint) => headers.includes(hint));
  });
}

function workbookLooksLikeCatalog(workbook: ExcelJS.Workbook) {
  return workbook.worksheets.some((worksheet) => {
    for (let rowNumber = 1; rowNumber <= Math.min(20, worksheet.rowCount); rowNumber += 1) {
      const row = worksheet.getRow(rowNumber);
      const headers: string[] = [];
      row.eachCell({ includeEmpty: false }, (cell) => headers.push(normalizeHeader(cellToText(cell))));
      if (CATALOG_HEADER_HINTS.filter((hint) => headers.includes(hint)).length >= 3) {
        return true;
      }
    }

    return false;
  });
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

function buildColumns(headers: string[], sourceColumnByHeader: Map<string, number>): SourceColumn[] {
  const filtered = headers
    .map((header) => ({ header, sourceCol: sourceColumnByHeader.get(normalizeHeader(header)) }))
    .filter((column) => {
      const normalized = normalizeHeader(column.header);
      return normalized !== normalizeHeader(CUSTOMER_NUMBER_HEADER) && normalized !== normalizeHeader(NOTES_HEADER);
    });

  return [
    filtered[0] ?? { header: "Cost Ctr Nbr" },
    { header: CUSTOMER_NUMBER_HEADER, sourceCol: sourceColumnByHeader.get(normalizeHeader(CUSTOMER_NUMBER_HEADER)), kind: "customerNumber" },
    ...filtered.slice(1),
    { header: NOTES_HEADER, sourceCol: sourceColumnByHeader.get(normalizeHeader(NOTES_HEADER)), kind: "notes" }
  ];
}

function buildSupplierEmail() {
  const subject = "Customer list review requested";
  const body = [
    "Hi,",
    "",
    "Please review the attached customer list.",
    "",
    "I added two supplier review columns:",
    "",
    "Customer number: Please enter your customer/account number for each active location.",
    "Notes: Please mark any closed or inactive locations and include any comments needed for setup or cleanup.",
    "",
    "Please also add any missing or new customers/locations that should be included for ordering.",
    "",
    "Once completed, please return the workbook. If a row is no longer active, please leave it in the file and note that it is closed in the Notes column.",
    "",
    "Thank you."
  ].join("\n");

  return { subject, body };
}

export async function formatCustomerListWorkbook(input: { fileName: string; buffer: ArrayBuffer }): Promise<CustomerListResponse> {
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.load(input.buffer);

  if (workbookLooksLikeCatalog(sourceWorkbook)) {
    throw new Error("This looks like a catalog file. Please use the Catalog Validator section for catalog review files.");
  }

  const sourceWorksheet = findWorksheet(sourceWorkbook);
  if (!sourceWorksheet) {
    throw new Error("This does not look like an AP customer list. Upload a workbook with a Unit List sheet or Cost Ctr columns.");
  }

  const headerRowNumber = findHeaderRow(sourceWorksheet);
  const sourceHeaderRow = sourceWorksheet.getRow(headerRowNumber);
  const sourceHeaders: string[] = [];
  for (let colNumber = 1; colNumber <= sourceWorksheet.columnCount; colNumber += 1) {
    const header = cellToText(sourceHeaderRow.getCell(colNumber));
    if (header) sourceHeaders.push(header);
  }

  const sourceColumnByHeader = new Map<string, number>();
  sourceHeaders.forEach((header, index) => {
    sourceColumnByHeader.set(normalizeHeader(header), index + 1);
  });

  const normalizedHeaders = sourceHeaders.map(normalizeHeader);
  const looksLikeCustomerList = CUSTOMER_LIST_HEADER_HINTS.some((hint) => normalizedHeaders.includes(hint));

  if (!looksLikeCustomerList) {
    throw new Error("This does not look like an AP customer list. Expected Cost Ctr columns on the Unit List sheet.");
  }

  const outputColumns = buildColumns(sourceHeaders, sourceColumnByHeader);
  const outputHeaders = outputColumns.map((column) => column.header);

  const outputWorkbook = new ExcelJS.Workbook();
  outputWorkbook.creator = "Catalog Validator";

  const listSheet = outputWorkbook.addWorksheet("Customer List");
  const headerRow = listSheet.addRow(outputHeaders);
  headerRow.height = sourceHeaderRow.height;
  outputColumns.forEach((column, index) => {
    const targetCell = headerRow.getCell(index + 1);
    const templateSourceCol = column.sourceCol ?? (column.kind === "customerNumber" ? 1 : sourceWorksheet.columnCount);
    copyCellFormat(sourceHeaderRow.getCell(templateSourceCol), targetCell);
  });
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

    const outputRow = listSheet.addRow([]);
    outputRow.height = sourceRow.height;
    totalRows += 1;

    outputColumns.forEach((column, index) => {
      const targetCell = outputRow.getCell(index + 1);
      const templateSourceCol = column.sourceCol ?? (column.kind === "customerNumber" ? 1 : sourceWorksheet.columnCount);
      const templateCell = sourceRow.getCell(templateSourceCol);
      copyCellFormat(templateCell, targetCell);

      if (column.sourceCol) {
        targetCell.value = cloneCellValue(sourceRow.getCell(column.sourceCol));
      } else {
        targetCell.value = "";
      }
    });

    if (closedColumnIndex > 0 && cellToText(outputRow.getCell(closedColumnIndex))) {
      possibleClosedRows += 1;
    }
  }

  outputColumns.forEach((column, index) => {
    const targetColumn = listSheet.getColumn(index + 1);
    if (column.sourceCol) {
      targetColumn.width = sourceWorksheet.getColumn(column.sourceCol).width;
    } else if (column.kind === "customerNumber") {
      targetColumn.width = Math.max(18, sourceWorksheet.getColumn(1).width ?? 18);
    } else {
      targetColumn.width = Math.max(28, sourceWorksheet.getColumn(sourceWorksheet.columnCount).width ?? 28);
    }
  });

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
  const formattedFileName = `${safeName}-customer-list-review.xlsx`;
  const email = buildSupplierEmail();

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
