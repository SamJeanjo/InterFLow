import ExcelJS from "exceljs";
import {
  approvedUoms,
  CatalogField,
  CatalogIssue,
  cleanRestrictedText,
  cleanText,
  expectedColumns,
  hasForbiddenDescriptionChars,
  isBlank,
  isDigits,
  isNumeric,
  isValidDateValue,
  makeIssue,
  normalizeHeader,
  SupplierCurrency,
  ValidationResponse,
  ValidationSummary
} from "@/lib/catalog-rules";
import { allowedLeadTimes } from "@/lib/catalog-rules";

type HeaderMatch = {
  rowNumber: number;
  mapping: Map<CatalogField, number>;
};

type RowContext = {
  rowNumber: number;
  values: Record<CatalogField, string>;
  cleaned: Record<CatalogField, string>;
  autoFix: boolean;
  supplierCurrency: SupplierCurrency;
};

type CatalogReviewRow = {
  rowNumber: number;
  values: Record<CatalogField, string>;
};

function cellToText(cell: ExcelJS.Cell) {
  if (cell.type === ExcelJS.ValueType.Date && cell.value instanceof Date) {
    return `${cell.value.getMonth() + 1}/${cell.value.getDate()}/${cell.value.getFullYear()}`;
  }

  if (typeof cell.value === "object" && cell.value && "text" in cell.value) {
    return String(cell.value.text ?? "");
  }

  return cell.text?.trim() ?? "";
}

function findHeaderRow(worksheet: ExcelJS.Worksheet): HeaderMatch | null {
  const expected = new Map(expectedColumns.map((column) => [normalizeHeader(column), column]));

  for (let rowNumber = 1; rowNumber <= Math.min(30, worksheet.rowCount); rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const mapping = new Map<CatalogField, number>();

    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const matched = expected.get(normalizeHeader(cellToText(cell)));
      if (matched) mapping.set(matched, colNumber);
    });

    if (mapping.size >= Math.ceil(expectedColumns.length * 0.75)) {
      return { rowNumber, mapping };
    }
  }

  return null;
}

function addIssue(
  issues: CatalogIssue[],
  ctx: RowContext,
  field: CatalogField,
  severity: "error" | "warning",
  issue: string,
  suggestedFix: string
) {
  issues.push(makeIssue(ctx.rowNumber, field, ctx.values[field], severity, issue, suggestedFix));
}

function validateRow(ctx: RowContext) {
  const issues: CatalogIssue[] = [];
  const value = (field: CatalogField) => ctx.values[field].trim();

  if (value("Action") && value("Action") !== "D") {
    addIssue(issues, ctx, "Action", "error", "Action can only be blank or D.", "Clear the value or use D only for deletions.");
  }

  if (isBlank(value("Item SKU"))) {
    addIssue(issues, ctx, "Item SKU", "error", "Item SKU is required.", "Enter a SKU up to 15 characters.");
  } else if (value("Item SKU").length > 15 || !/^[A-Za-z0-9-]+$/.test(value("Item SKU"))) {
    addIssue(issues, ctx, "Item SKU", "error", "Item SKU must be 15 characters or fewer and use only letters, numbers, and dashes.", "Remove spaces and punctuation except dashes.");
  }

  if (isBlank(value("Item Description"))) {
    addIssue(issues, ctx, "Item Description", "error", "Item Description is required.", "Enter a description up to 500 characters.");
  } else {
    if (value("Item Description").length > 500) {
      addIssue(issues, ctx, "Item Description", "error", "Item Description exceeds 500 characters.", "Shorten the description to 500 characters or fewer.");
    }
    if (hasForbiddenDescriptionChars(ctx.values["Item Description"])) {
      addIssue(issues, ctx, "Item Description", "error", "Item Description contains forbidden characters or double spaces.", "Remove commas, semicolons, quotes, stars, tildes, greater-than signs, pipes, and double spaces.");
    }
  }

  if (isBlank(value("UOM"))) {
    addIssue(issues, ctx, "UOM", "error", "UOM is required.", "Enter a valid unit of measure.");
  } else if (!approvedUoms.includes(value("UOM").toUpperCase())) {
    addIssue(issues, ctx, "UOM", "warning", "UOM is not in the approved MVP list.", `Use one of: ${approvedUoms.join(", ")}.`);
  }

  if (!isBlank(value("Items Per Case")) && !isNumeric(value("Items Per Case"))) {
    addIssue(issues, ctx, "Items Per Case", "error", "Items Per Case must be numeric when provided.", "Enter a number such as 1, 6, or 12.");
  }

  if (isBlank(value("Pack Size"))) {
    addIssue(issues, ctx, "Pack Size", "warning", "Pack Size is blank.", "Enter item size or weight when available.");
  }

  if (!isBlank(value("Catchweight")) && value("Catchweight") !== "1") {
    addIssue(issues, ctx, "Catchweight", "error", "Catchweight can only be blank or 1.", "Clear the value or enter 1 for catch weight items.");
  }

  if (value("Catchweight") === "1" && isBlank(value("Avg Case Weight"))) {
    addIssue(issues, ctx, "Avg Case Weight", "error", "Avg Case Weight is required when Catchweight is 1.", "Enter the average case weight.");
  } else if (!isBlank(value("Avg Case Weight")) && !isNumeric(value("Avg Case Weight"))) {
    addIssue(issues, ctx, "Avg Case Weight", "error", "Avg Case Weight must be numeric.", "Enter a numeric weight.");
  }

  if (isBlank(value("Unit Price"))) {
    addIssue(issues, ctx, "Unit Price", "error", "Unit Price is required.", "Enter a numeric price without a currency sign.");
  } else if (!isNumeric(value("Unit Price")) || /[$€£]/.test(value("Unit Price"))) {
    addIssue(issues, ctx, "Unit Price", "error", "Unit Price must be numeric and cannot include currency signs.", "Remove currency symbols and enter numbers only.");
  }

  if (isBlank(value("Currency"))) {
    const severity = ctx.autoFix ? "warning" : "error";
    addIssue(issues, ctx, "Currency", severity, "Currency is required.", `Enter ${ctx.supplierCurrency}.`);
    if (ctx.autoFix) ctx.cleaned.Currency = ctx.supplierCurrency;
  } else if (value("Currency") !== ctx.supplierCurrency) {
    const isLowercaseExpected = value("Currency").toLowerCase() === ctx.supplierCurrency.toLowerCase();
    const severity = isLowercaseExpected ? "warning" : "error";
    addIssue(issues, ctx, "Currency", severity, `Currency must be ${ctx.supplierCurrency}.`, `Use ${ctx.supplierCurrency}.`);
    if (isLowercaseExpected) ctx.cleaned.Currency = ctx.supplierCurrency;
  }

  if (!isBlank(value("Product Expiration Date")) && !isValidDateValue(value("Product Expiration Date"))) {
    addIssue(issues, ctx, "Product Expiration Date", "error", "Product Expiration Date must be a valid date.", "Use MM/DD/YYYY.");
  }

  if (isBlank(value("Price Effective Date"))) {
    addIssue(issues, ctx, "Price Effective Date", "error", "Price Effective Date is required.", "Use MM/DD/YYYY.");
  } else if (!isValidDateValue(value("Price Effective Date"))) {
    addIssue(issues, ctx, "Price Effective Date", "error", "Price Effective Date must be a valid date.", "Use MM/DD/YYYY.");
  }

  if (!isBlank(value("Price Expiration Date")) && !isValidDateValue(value("Price Expiration Date"))) {
    addIssue(issues, ctx, "Price Expiration Date", "error", "Price Expiration Date must be a valid date.", "Use MM/DD/YYYY.");
  }

  if (!isBlank(value("Long Description"))) {
    if (value("Long Description").length > 4000) {
      addIssue(issues, ctx, "Long Description", "error", "Long Description exceeds 4000 characters.", "Shorten the long description.");
    }
    if (hasForbiddenDescriptionChars(ctx.values["Long Description"])) {
      addIssue(issues, ctx, "Long Description", "error", "Long Description contains forbidden characters or double spaces.", "Remove forbidden characters and double spaces.");
    }
  }

  if (isBlank(value("Lead Time"))) {
    addIssue(issues, ctx, "Lead Time", "warning", "Lead Time is blank.", "Enter 0, 1, 2, 3, or 5 when available.");
  } else if (!isNumeric(value("Lead Time")) || !allowedLeadTimes.includes(value("Lead Time"))) {
    addIssue(issues, ctx, "Lead Time", "error", "Lead Time must be one of the allowed values.", "Use 0, 1, 2, 3, or 5.");
  }

  if (!isBlank(value("UNSPSC")) && !isDigits(value("UNSPSC"))) {
    addIssue(issues, ctx, "UNSPSC", "error", "UNSPSC must be numeric when provided.", "Enter numbers only.");
  }

  if (!isBlank(value("Classification Code"))) {
    addIssue(issues, ctx, "Classification Code", "error", "Classification Code must be blank.", "Clear this field.");
  }

  for (const field of ["UPC", "GTIN"] as const) {
    if (!isBlank(value(field)) && !isDigits(value(field))) {
      addIssue(issues, ctx, field, "error", `${field} must contain numbers only.`, "Remove commas, spaces, and non-numeric characters.");
    }
  }

  if (value("Product Origin").includes(",")) {
    addIssue(issues, ctx, "Product Origin", "error", "Product Origin cannot contain commas.", "Remove the comma.");
  }

  if (isBlank(value("Break Case"))) {
    addIssue(issues, ctx, "Break Case", "error", "Break Case is required.", "Use 0 for no case break or 1 for case break.");
    if (ctx.autoFix) ctx.cleaned["Break Case"] = "0";
  } else if (!["0", "1"].includes(value("Break Case"))) {
    addIssue(issues, ctx, "Break Case", "error", "Break Case must be 0 or 1.", "Use 0 or 1.");
  }

  if (isBlank(value("Minimum Order Qty"))) {
    addIssue(issues, ctx, "Minimum Order Qty", "error", "Minimum Order Qty is required.", "Use 1 when there is no minimum.");
    if (ctx.autoFix) ctx.cleaned["Minimum Order Qty"] = "1";
  } else if (!isNumeric(value("Minimum Order Qty")) || Number(value("Minimum Order Qty")) === 0) {
    addIssue(issues, ctx, "Minimum Order Qty", "error", "Minimum Order Qty must be numeric and cannot be 0.", "Use 1 when there is no minimum.");
    if (ctx.autoFix && value("Minimum Order Qty") === "0") ctx.cleaned["Minimum Order Qty"] = "1";
  }

  if (!isBlank(value("Image"))) {
    addIssue(issues, ctx, "Image", "warning", "Image should be blank.", "Clear this field.");
  }

  if (!isBlank(value("Item Status"))) {
    addIssue(issues, ctx, "Item Status", "warning", "Item Status should be blank.", "Clear this field.");
  }

  return issues;
}

function summarize(totalRows: number, issues: CatalogIssue[]): ValidationSummary {
  const errors = new Set<number>();
  const warnings = new Set<number>();

  for (const issue of issues) {
    if (issue.severity === "error") errors.add(issue.rowNumber);
    if (issue.severity === "warning") warnings.add(issue.rowNumber);
  }

  return {
    totalRows,
    passedRows: Math.max(0, totalRows - new Set([...errors, ...warnings]).size),
    rowsWithWarnings: warnings.size,
    rowsWithErrors: errors.size,
    totalIssues: issues.length
  };
}

function styleHeaderRow(row: ExcelJS.Row) {
  row.font = { bold: true, color: { argb: "FFFFFFFF" } };
  row.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F2937" }
  };
  row.alignment = { vertical: "middle", wrapText: true };
}

function applyIssueStyle(cell: ExcelJS.Cell, issues: CatalogIssue[]) {
  const hasError = issues.some((issue) => issue.severity === "error");
  const fillColor = hasError ? "FFFEE2E2" : "FFFEF3C7";
  const borderColor = hasError ? "FFDC2626" : "FFD97706";

  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: fillColor }
  };
  cell.border = {
    top: { style: "thin", color: { argb: borderColor } },
    left: { style: "thin", color: { argb: borderColor } },
    bottom: { style: "thin", color: { argb: borderColor } },
    right: { style: "thin", color: { argb: borderColor } }
  };
  cell.note = issues
    .map((issue) => `${issue.severity.toUpperCase()}: ${issue.issue}\nSuggested fix: ${issue.suggestedFix}`)
    .join("\n\n");
}

async function buildReportWorkbook(fileName: string, summary: ValidationSummary, issues: CatalogIssue[], catalogRows: CatalogReviewRow[]) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Catalog Validator";

  const summarySheet = workbook.addWorksheet("Summary");
  summarySheet.addRows([
    ["Catalog Validator Report"],
    ["Source File", fileName],
    ["Total Rows", summary.totalRows],
    ["Passed Rows", summary.passedRows],
    ["Rows With Warnings", summary.rowsWithWarnings],
    ["Rows With Errors", summary.rowsWithErrors],
    ["Total Issues", summary.totalIssues]
  ]);
  summarySheet.getColumn(1).width = 24;
  summarySheet.getColumn(2).width = 36;
  summarySheet.getCell("A1").font = { bold: true, size: 16 };
  summarySheet.addRow([]);
  summarySheet.addRow(["Report Tabs"]);
  summarySheet.addRow(["Issues", "Complete issue list by row, field, current value, severity, issue, and suggested fix."]);
  summarySheet.addRow(["Catalog Review", "Catalog rows with issue cells highlighted red for errors and amber for warnings."]);
  summarySheet.getCell("A9").font = { bold: true };

  const issueSheet = workbook.addWorksheet("Issues");
  issueSheet.columns = [
    { header: "Row Number", key: "rowNumber", width: 12 },
    { header: "Field", key: "field", width: 24 },
    { header: "Current Value", key: "currentValue", width: 30 },
    { header: "Severity", key: "severity", width: 12 },
    { header: "Issue", key: "issue", width: 46 },
    { header: "Suggested Fix", key: "suggestedFix", width: 46 }
  ];
  issueSheet.addRows(issues);
  styleHeaderRow(issueSheet.getRow(1));
  issueSheet.views = [{ state: "frozen", ySplit: 1 }];
  issueSheet.autoFilter = {
    from: "A1",
    to: "F1"
  };
  issueSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const severity = String(row.getCell("severity").value ?? "").toLowerCase();
    if (severity === "error") {
      row.getCell("severity").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
      row.getCell("severity").font = { bold: true, color: { argb: "FF991B1B" } };
    }
    if (severity === "warning") {
      row.getCell("severity").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3C7" } };
      row.getCell("severity").font = { bold: true, color: { argb: "FF92400E" } };
    }
  });

  const reviewSheet = workbook.addWorksheet("Catalog Review");
  reviewSheet.columns = [
    { header: "Excel Row", key: "rowNumber", width: 12 },
    ...expectedColumns.map((field) => ({ header: field, key: field, width: Math.max(16, Math.min(34, field.length + 4)) }))
  ];
  styleHeaderRow(reviewSheet.getRow(1));
  reviewSheet.views = [{ state: "frozen", ySplit: 1, xSplit: 1 }];
  reviewSheet.autoFilter = {
    from: "A1",
    to: `${reviewSheet.getColumn(expectedColumns.length + 1).letter}1`
  };

  const issuesByCell = new Map<string, CatalogIssue[]>();
  for (const issue of issues) {
    if (issue.field === "Workbook") continue;
    const key = `${issue.rowNumber}::${issue.field}`;
    const existing = issuesByCell.get(key) ?? [];
    existing.push(issue);
    issuesByCell.set(key, existing);
  }

  for (const catalogRow of catalogRows) {
    const row = reviewSheet.addRow({
      rowNumber: catalogRow.rowNumber,
      ...catalogRow.values
    });
    row.getCell(1).font = { bold: true };

    expectedColumns.forEach((field, index) => {
      const cellIssues = issuesByCell.get(`${catalogRow.rowNumber}::${field}`);
      if (!cellIssues?.length) return;
      applyIssueStyle(row.getCell(index + 2), cellIssues);
    });
  }

  if (catalogRows.length === 0) {
    reviewSheet.addRow(["No catalog data rows were detected."]);
  }

  return Buffer.from(await workbook.xlsx.writeBuffer()).toString("base64");
}

async function workbookToBase64(workbook: ExcelJS.Workbook) {
  return Buffer.from(await workbook.xlsx.writeBuffer()).toString("base64");
}

export async function validateCatalogWorkbook(input: {
  fileName: string;
  buffer: ArrayBuffer;
  autoFix: boolean;
  supplierCurrency: SupplierCurrency;
}): Promise<ValidationResponse> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(input.buffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("The workbook does not contain any worksheets.");
  }

  const header = findHeaderRow(worksheet);
  if (!header) {
    const reportIssue = makeIssue(0, "Workbook", "", "error", "Could not detect a valid Collaboration Portal header row.", "Use the expected Collaboration Portal catalog template.");
    const summary = summarize(0, [reportIssue]);
    const reportBase64 = await buildReportWorkbook(input.fileName, summary, [reportIssue], []);
    return {
      fileName: input.fileName,
      summary,
      issues: [reportIssue],
      fields: [...expectedColumns],
      reportFileName: "catalog-validation-report.xlsx",
      reportBase64,
      cleanedFileName: "cleaned-catalog.xlsx",
      cleanedBase64: await workbookToBase64(workbook)
    };
  }

  const missingColumns = expectedColumns.filter((column) => !header.mapping.has(column));
  const allIssues: CatalogIssue[] = missingColumns.map((column) =>
    makeIssue(header.rowNumber, "Workbook", "", "error", `Missing expected column: ${column}.`, "Add the missing column to the header row.")
  );

  let totalRows = 0;
  const catalogRows: CatalogReviewRow[] = [];
  for (let rowNumber = header.rowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const hasAnyValue = expectedColumns.some((field) => {
      const columnNumber = header.mapping.get(field);
      return columnNumber ? cellToText(row.getCell(columnNumber)).trim() !== "" : false;
    });

    if (!hasAnyValue) continue;
    totalRows += 1;

    const values = {} as Record<CatalogField, string>;
    const cleaned = {} as Record<CatalogField, string>;

    for (const field of expectedColumns) {
      const columnNumber = header.mapping.get(field);
      const raw = columnNumber ? cellToText(row.getCell(columnNumber)) : "";
      values[field] = raw;
      cleaned[field] = field === "Item Description" || field === "Long Description" ? cleanRestrictedText(raw) : cleanText(raw);
    }
    catalogRows.push({ rowNumber, values: { ...values } });

    const ctx: RowContext = {
      rowNumber,
      values,
      cleaned,
      autoFix: input.autoFix,
      supplierCurrency: input.supplierCurrency
    };
    allIssues.push(...validateRow(ctx));

    for (const field of expectedColumns) {
      const columnNumber = header.mapping.get(field);
      if (!columnNumber) continue;
      row.getCell(columnNumber).value = cleaned[field] === "" ? null : cleaned[field];
    }
    row.commit();
  }

  const summary = summarize(totalRows, allIssues);
  const reportBase64 = await buildReportWorkbook(input.fileName, summary, allIssues, catalogRows);
  const cleanedBase64 = await workbookToBase64(workbook);
  const safeName = input.fileName.replace(/\.xlsx$/i, "");

  return {
    fileName: input.fileName,
    summary,
    issues: allIssues,
    fields: [...expectedColumns],
    reportFileName: `${safeName}-validation-report.xlsx`,
    reportBase64,
    cleanedFileName: `${safeName}-cleaned.xlsx`,
    cleanedBase64
  };
}
