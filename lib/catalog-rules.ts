import { z } from "zod";
import { approvedUnuomCodes, approvedUnuomCodeSet } from "@/lib/unuom-codes";

export const expectedColumns = [
  "Action",
  "Item SKU",
  "Item Description",
  "UOM",
  "Items Per Case",
  "Pack Size",
  "Catchweight",
  "Avg Case Weight",
  "Unit Price",
  "Currency",
  "Product Expiration Date",
  "Price Effective Date",
  "Price Expiration Date",
  "Brand Name",
  "Mfg Name",
  "Mfg Part Number",
  "Long Description",
  "Lead Time",
  "UNSPSC",
  "Classification Code",
  "Auxiliary Part Id",
  "UPC",
  "GTIN",
  "Product Origin",
  "Break Case",
  "Minimum Order Qty",
  "Dimensions",
  "Image",
  "Item Status"
] as const;

export type CatalogField = (typeof expectedColumns)[number];
export type Severity = "error" | "warning" | "suggestion";
export type SupplierCurrency = "USD" | "CAD";

export type CatalogIssue = {
  rowNumber: number;
  field: CatalogField | "Workbook";
  currentValue: string;
  severity: Severity;
  group?: string;
  issue: string;
  suggestedFix: string;
};

export type ValidationSummary = {
  totalRows: number;
  passedRows: number;
  rowsWithSuggestions: number;
  rowsWithWarnings: number;
  rowsWithErrors: number;
  totalIssues: number;
};

export type ValidationResponse = {
  fileName: string;
  summary: ValidationSummary;
  issues: CatalogIssue[];
  fields: string[];
  reportFileName: string;
  reportBase64: string;
  cleanedFileName: string;
  cleanedBase64: string;
};

export const uploadOptionsSchema = z.object({
  autoFix: z.boolean().default(false),
  supplierCurrency: z.enum(["USD", "CAD"]).default("USD")
});

export const approvedUoms = approvedUnuomCodes;
export const approvedUomSet = approvedUnuomCodeSet;
export const allowedLeadTimes = ["0", "1", "2", "3", "5"];
export const forbiddenTextPattern = /[;,"*~>|]/;

export function normalizeHeader(value: string) {
  return value.trim().replace(/\s+/g, " ").toLowerCase();
}

export function cleanText(value: string) {
  return value.trim().replace(/\s{2,}/g, " ");
}

export function cleanRestrictedText(value: string) {
  return cleanText(value).replace(/[;,"*~>|]/g, "");
}

export function hasForbiddenDescriptionChars(value: string) {
  return forbiddenTextPattern.test(value) || /\s{2,}/.test(value);
}

export function isBlank(value: string) {
  return value.trim() === "";
}

export function isNumeric(value: string) {
  if (isBlank(value)) return false;
  return /^-?\d+(\.\d+)?$/.test(value.trim());
}

export function isDigits(value: string) {
  return /^\d+$/.test(value.trim());
}

export function isValidDateValue(value: string) {
  const trimmed = value.trim();
  if (!trimmed) return false;
  const mmddyyyy = /^(0?[1-9]|1[0-2])\/(0?[1-9]|[12]\d|3[01])\/\d{4}$/;
  if (mmddyyyy.test(trimmed)) {
    const [month, day, year] = trimmed.split("/").map(Number);
    const date = new Date(year, month - 1, day);
    return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day;
  }

  const parsed = new Date(trimmed);
  return !Number.isNaN(parsed.getTime());
}

export function makeIssue(
  rowNumber: number,
  field: CatalogIssue["field"],
  currentValue: string,
  severity: Severity,
  issue: string,
  suggestedFix: string,
  group?: string
): CatalogIssue {
  return {
    rowNumber,
    field,
    currentValue,
    severity,
    group,
    issue,
    suggestedFix
  };
}
