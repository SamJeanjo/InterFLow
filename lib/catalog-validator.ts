import ExcelJS from "exceljs";
import {
  approvedUomSet,
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
  Severity,
  SupplierCurrency,
  ValidationResponse,
  ValidationSummary
} from "@/lib/catalog-rules";
import { allowedLeadTimes } from "@/lib/catalog-rules";

const CATCH_WEIGHT_REVIEW_GROUP = "Catch Weight / UOM Consistency Review";
const WEIGHT_BASED_UOMS = new Set(["LBR", "LB", "KG", "G"]);

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
  cells: Record<
    CatalogField,
    {
      value: ExcelJS.CellValue;
      numFmt?: string;
    }
  >;
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

function cloneCellValue(cell: ExcelJS.Cell): ExcelJS.CellValue {
  if (cell.value instanceof Date) {
    return new Date(cell.value.getTime());
  }

  if (typeof cell.value === "object" && cell.value !== null) {
    return JSON.parse(JSON.stringify(cell.value)) as ExcelJS.CellValue;
  }

  return cell.value;
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
  severity: Severity,
  issue: string,
  suggestedFix: string,
  group?: string
) {
  issues.push(makeIssue(ctx.rowNumber, field, ctx.values[field], severity, issue, suggestedFix, group));
}

function isCatchWeightTrue(value: string) {
  return ["1", "TRUE", "YES", "Y"].includes(value.trim().toUpperCase());
}

function isCatchWeightBlankOrFalse(value: string) {
  return isBlank(value) || ["0", "FALSE", "NO", "N"].includes(value.trim().toUpperCase());
}

function hasEachLanguage(value: string) {
  return /\b(EACH|EA|PC|PIECE)\b/i.test(value);
}

function hasCasePackLanguage(value: string) {
  return /\b(CASE|CS|BOX|BAG|TUB|PAIL|LOAF|FILLET|WHOLE|RANDOM|AVG|APPROX)\b/i.test(value);
}

function appearsWeightPriced(value: string) {
  return /(?:\/|\bPER\s+)\s*(LBR|LB|LBS|POUND|POUNDS|KG|KGS|KILOGRAM|KILOGRAMS)\b/i.test(value) ||
    /\b(LBR|LB|LBS|POUND|POUNDS|KG|KGS|KILOGRAM|KILOGRAMS)\b/i.test(value);
}

function hasFixedWeightPattern(value: string) {
  return /\b\d+(?:\.\d+)?\s*(?:X|x|×|\/)\s*\d+(?:\.\d+)?\s*(?:OZ|OUNCE|OUNCES|LB|LBR|LBS|KG|KGS|G|GRAM|GRAMS)\b/i.test(value) ||
    /\b\d+(?:\.\d+)?\s*(?:OZ|OUNCE|OUNCES|LB|LBR|LBS|KG|KGS|G|GRAM|GRAMS)\s*(?:CASE|CS|BOX|BAG)\b/i.test(value);
}

function applyCatchWeightUomConsistencyReview(issues: CatalogIssue[], ctx: RowContext) {
  const value = (field: CatalogField) => ctx.values[field].trim();
  const uom = value("UOM").toUpperCase();
  const description = value("Item Description");
  const packSize = value("Pack Size");
  const catchweight = value("Catchweight");
  const combinedProductText = `${description} ${packSize}`;
  const isWeightBasedUom = WEIGHT_BASED_UOMS.has(uom);
  const catchWeightTrue = isCatchWeightTrue(catchweight);
  const catchWeightBlankOrFalse = isCatchWeightBlankOrFalse(catchweight);
  const fixedWeightPack = hasFixedWeightPattern(combinedProductText);

  if ((uom === "LBR" || uom === "LB") && hasEachLanguage(description) && catchWeightBlankOrFalse) {
    addIssue(
      issues,
      ctx,
      "UOM",
      "warning",
      "Possible UOM conflict. Description indicates the item may be sold by each, but UOM is set to pound. Please confirm whether this item is sold by pound or by each.",
      "If sold by pound: keep UOM as LBR and update description if needed. If sold by each but priced by pound: review Catch Weight and Average Case Weight requirements.",
      CATCH_WEIGHT_REVIEW_GROUP
    );
  }

  if (
    hasCasePackLanguage(description) &&
    !isWeightBasedUom &&
    !fixedWeightPack &&
    !isBlank(value("Unit Price")) &&
    appearsWeightPriced(combinedProductText) &&
    catchWeightBlankOrFalse
  ) {
    addIssue(
      issues,
      ctx,
      "Catchweight",
      "warning",
      "Possible Catch Weight item. This item appears to be sold as a case/pack but may be priced by weight. If the case weight can vary, set Catch Weight to True and provide Average Case Weight.",
      "Confirm the selling logic. If the case weight can vary, set Catch Weight to 1 and provide Average Case Weight.",
      CATCH_WEIGHT_REVIEW_GROUP
    );
  }

  if (catchWeightTrue && isBlank(value("Avg Case Weight"))) {
    addIssue(
      issues,
      ctx,
      "Avg Case Weight",
      "error",
      "Catch Weight is True, but Average Case Weight is missing. Please provide the average case weight so the system can calculate the sellable unit price.",
      "Enter the average case weight.",
      CATCH_WEIGHT_REVIEW_GROUP
    );
  }

  if (catchWeightTrue && isWeightBasedUom) {
    addIssue(
      issues,
      ctx,
      "Catchweight",
      "warning",
      "Catch Weight may not be required because the UOM is already weight-based. If the item is sold directly by pound/kg, set Catch Weight to False. If the item is sold as a case but priced by pound/kg, confirm the correct UOM and Average Case Weight.",
      "If sold directly by pound/kg, clear Catchweight. If sold as a case but priced by pound/kg, confirm the correct UOM and Average Case Weight.",
      CATCH_WEIGHT_REVIEW_GROUP
    );
  }

  if (catchWeightTrue && hasFixedWeightPattern(packSize)) {
    addIssue(
      issues,
      ctx,
      "Catchweight",
      "warning",
      "Catch Weight is marked True, but the pack size appears fixed/premeasured. Please confirm whether the case weight actually varies.",
      "Confirm whether the case weight varies. If fixed/premeasured, clear Catchweight.",
      CATCH_WEIGHT_REVIEW_GROUP
    );
  }
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
  } else if (!approvedUomSet.has(value("UOM").toUpperCase())) {
    addIssue(
      issues,
      ctx,
      "UOM",
      "warning",
      "UOM is not in the approved UNUOM code list.",
      "Use a valid UNUOM Revision 9 unit of measure code."
    );
  }

  if (!isBlank(value("Items Per Case")) && !isNumeric(value("Items Per Case"))) {
    addIssue(issues, ctx, "Items Per Case", "error", "Items Per Case must be numeric when provided.", "Enter a number such as 1, 6, or 12.");
  }

  if (isBlank(value("Pack Size"))) {
    addIssue(issues, ctx, "Pack Size", "warning", "Pack Size is blank.", "Enter item size or weight when available.");
  }

  if (!isCatchWeightBlankOrFalse(value("Catchweight")) && !isCatchWeightTrue(value("Catchweight"))) {
    addIssue(
      issues,
      ctx,
      "Catchweight",
      "error",
      "Catchweight can only be blank/False or 1/True.",
      "Clear the value for non-catch-weight items or enter 1 for catch weight items."
    );
  }

  if (!isBlank(value("Avg Case Weight")) && !isNumeric(value("Avg Case Weight"))) {
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

  applyCatchWeightUomConsistencyReview(issues, ctx);

  return issues;
}

function summarize(totalRows: number, issues: CatalogIssue[]): ValidationSummary {
  const errors = new Set<number>();
  const warnings = new Set<number>();
  const suggestions = new Set<number>();

  for (const issue of issues) {
    if (issue.severity === "error") errors.add(issue.rowNumber);
    if (issue.severity === "warning") warnings.add(issue.rowNumber);
    if (issue.severity === "suggestion") suggestions.add(issue.rowNumber);
  }

  return {
    totalRows,
    passedRows: Math.max(0, totalRows - new Set([...errors, ...warnings, ...suggestions]).size),
    rowsWithSuggestions: suggestions.size,
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
  const hasWarning = issues.some((issue) => issue.severity === "warning");
  const fillColor = hasError ? "FFFEE2E2" : hasWarning ? "FFFEF3C7" : "FFEFF6FF";
  const borderColor = hasError ? "FFDC2626" : hasWarning ? "FFD97706" : "FF2563EB";

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
    .map((issue) => `${issue.severity.toUpperCase()}${issue.group ? ` - ${issue.group}` : ""}: ${issue.issue}\nSuggested fix: ${issue.suggestedFix}`)
    .join("\n\n");
}

function styleIssueSeverityCell(row: ExcelJS.Row) {
  const severity = String(row.getCell("severity").value ?? "").toLowerCase();
  if (severity === "error") {
    row.getCell("severity").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
    row.getCell("severity").font = { bold: true, color: { argb: "FF991B1B" } };
  }
  if (severity === "warning") {
    row.getCell("severity").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3C7" } };
    row.getCell("severity").font = { bold: true, color: { argb: "FF92400E" } };
  }
  if (severity === "suggestion") {
    row.getCell("severity").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF6FF" } };
    row.getCell("severity").font = { bold: true, color: { argb: "FF1D4ED8" } };
  }
}

function addIssueWorksheet(workbook: ExcelJS.Workbook, name: string, issues: CatalogIssue[]) {
  const sheet = workbook.addWorksheet(name);
  sheet.columns = [
    { header: "Row Number", key: "rowNumber", width: 12 },
    { header: "Field", key: "field", width: 24 },
    { header: "Current Value", key: "currentValue", width: 30 },
    { header: "Severity", key: "severity", width: 12 },
    { header: "Rule Group", key: "group", width: 34 },
    { header: "Issue", key: "issue", width: 46 },
    { header: "Suggested Fix", key: "suggestedFix", width: 46 }
  ];
  sheet.addRows(
    issues.length
      ? issues
      : [
          {
            rowNumber: "",
            field: "",
            currentValue: "",
            severity: "",
            group: "",
            issue: `No ${name.toLowerCase()} found.`,
            suggestedFix: ""
          }
        ]
  );
  styleHeaderRow(sheet.getRow(1));
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = {
    from: "A1",
    to: "G1"
  };
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    styleIssueSeverityCell(row);
  });

  return sheet;
}

async function buildReportWorkbook(fileName: string, summary: ValidationSummary, issues: CatalogIssue[], catalogRows: CatalogReviewRow[]) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Catalog Validator";
  workbook.views = [{ x: 0, y: 0, width: 12000, height: 20000, firstSheet: 0, activeTab: 1, visibility: "visible" }];

  const summarySheet = workbook.addWorksheet("Summary");
  summarySheet.addRows([
    ["Catalog Validator Report"],
    ["Source File", fileName],
    ["Total Rows", summary.totalRows],
    ["Passed Rows", summary.passedRows],
    ["Rows With Suggestions", summary.rowsWithSuggestions],
    ["Rows With Warnings", summary.rowsWithWarnings],
    ["Rows With Errors", summary.rowsWithErrors],
    ["Total Issues", summary.totalIssues]
  ]);
  summarySheet.getColumn(1).width = 24;
  summarySheet.getColumn(2).width = 36;
  summarySheet.getCell("A1").font = { bold: true, size: 16 };
  summarySheet.addRow([]);
  summarySheet.addRow(["Report Tabs"]);
  summarySheet.addRow(["Errors", "Blocking issues that must be corrected before submission. This tab opens first by default."]);
  summarySheet.addRow(["Warnings", "Likely supplier-confirmation items and important review items."]);
  summarySheet.addRow(["Suggestions", "Cleanup or improvement items that are not blockers."]);
  summarySheet.addRow(["All Issues", "Complete issue list by row, field, current value, severity, rule group, issue, and suggested fix."]);
  summarySheet.addRow(["Catalog Review", "Catalog rows with issue cells highlighted red for errors, amber for warnings, and blue for suggestions."]);
  summarySheet.getCell("A10").font = { bold: true };

  addIssueWorksheet(workbook, "Errors", issues.filter((issue) => issue.severity === "error"));
  addIssueWorksheet(workbook, "Warnings", issues.filter((issue) => issue.severity === "warning"));
  addIssueWorksheet(workbook, "Suggestions", issues.filter((issue) => issue.severity === "suggestion"));
  addIssueWorksheet(workbook, "All Issues", issues);

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
    const row = reviewSheet.addRow([catalogRow.rowNumber]);
    row.getCell(1).font = { bold: true };

    expectedColumns.forEach((field, index) => {
      const cell = row.getCell(index + 2);
      const sourceCell = catalogRow.cells[field];
      cell.value = sourceCell.value;
      if (sourceCell.numFmt) {
        cell.numFmt = sourceCell.numFmt;
      } else if (field === "Unit Price" && typeof sourceCell.value === "number") {
        cell.numFmt = "0.00";
      }

      const cellIssues = issuesByCell.get(`${catalogRow.rowNumber}::${field}`);
      if (!cellIssues?.length) return;
      applyIssueStyle(cell, cellIssues);
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
    const reviewCells = {} as CatalogReviewRow["cells"];

    for (const field of expectedColumns) {
      const columnNumber = header.mapping.get(field);
      const sourceCell = columnNumber ? row.getCell(columnNumber) : undefined;
      const raw = sourceCell ? cellToText(sourceCell) : "";
      values[field] = raw;
      cleaned[field] = field === "Item Description" || field === "Long Description" ? cleanRestrictedText(raw) : cleanText(raw);
      reviewCells[field] = {
        value: sourceCell ? cloneCellValue(sourceCell) : null,
        numFmt: sourceCell?.numFmt
      };
    }
    catalogRows.push({ rowNumber, cells: reviewCells });

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
      if (cleaned[field] !== values[field]) {
        row.getCell(columnNumber).value = cleaned[field] === "" ? null : cleaned[field];
      }
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
