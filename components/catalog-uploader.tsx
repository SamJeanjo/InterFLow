"use client";

import { ChangeEvent, DragEvent, useMemo, useRef, useState } from "react";
import {
  AlertCircle,
  CheckCircle2,
  Clipboard,
  Download,
  FileSpreadsheet,
  Mail,
  Search,
  UploadCloud,
  Wand2
} from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import type { CatalogIssue, ValidationResponse } from "@/lib/catalog-rules";
import { cn } from "@/lib/utils";

type Status = "idle" | "validating" | "done" | "error";
type SupplierCurrency = "USD" | "CAD";
const INLINE_EMAIL_ISSUE_LIMIT = 20;

function downloadBase64Xlsx(base64: string, fileName: string) {
  const bytes = Uint8Array.from(atob(base64), (char) => char.charCodeAt(0));
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function buildSupplierEmail(result: ValidationResponse) {
  const errorCount = result.issues.filter((issue) => issue.severity === "error").length;
  const warningCount = result.issues.filter((issue) => issue.severity === "warning").length;
  const hasAttachment = result.issues.length > INLINE_EMAIL_ISSUE_LIMIT;
  const visibleIssues = hasAttachment ? [] : result.issues;

  const issueLines = visibleIssues.map((issue, index) => {
    const row = issue.rowNumber || "Workbook";
    return [
      `${index + 1}. Row: ${row}`,
      `   Column / Field: ${issue.field}`,
      `   Severity: ${issue.severity.toUpperCase()}`,
      `   Current value: ${issue.currentValue || "(blank)"}`,
      `   Issue: ${issue.issue}`,
      `   Recommended fix: ${issue.suggestedFix}`
    ].join("\n");
  });

  return [
    "Subject: Supplier catalog corrections required",
    "",
    "Hello,",
    "",
    `Thank you for submitting your catalog file: ${result.fileName}.`,
    "",
    "We completed a validation review and found items that need to be corrected before the catalog can be submitted.",
    "",
    "Validation summary:",
    `- Total rows reviewed: ${result.summary.totalRows}`,
    `- Rows with errors: ${result.summary.rowsWithErrors}`,
    `- Rows with warnings: ${result.summary.rowsWithWarnings}`,
    `- Total issues found: ${result.summary.totalIssues}`,
    `- Error count: ${errorCount}`,
    `- Warning count: ${warningCount}`,
    "",
    hasAttachment
      ? [
          `Because there are more than ${INLINE_EMAIL_ISSUE_LIMIT} issues, please review the attached Excel validation report.`,
          "The report lists each item by row number, column / field, current value, severity, issue, and recommended fix."
        ].join("\n")
      : [
          "Please correct the issues listed below:",
          "",
          issueLines.join("\n\n") || "No row-level issues were found."
        ].join("\n"),
    "",
    "Please update the catalog and resend the corrected file. Focus first on rows marked ERROR, because those must be corrected before submission.",
    "",
    "Thank you."
  ].join("\n");
}

export function CatalogUploader() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [autoFix, setAutoFix] = useState(true);
  const [supplierCurrency, setSupplierCurrency] = useState<SupplierCurrency>("USD");
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [result, setResult] = useState<ValidationResponse | null>(null);
  const [severity, setSeverity] = useState<"all" | "error" | "warning">("all");
  const [field, setField] = useState("all");
  const [query, setQuery] = useState("");
  const [emailCopied, setEmailCopied] = useState(false);

  async function validateFile(file: File) {
    setError("");
    setResult(null);

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setStatus("error");
      setError("Only .xlsx files are supported for the MVP.");
      return;
    }

    setStatus("validating");
    setFileName(file.name);
    setEmailCopied(false);

    const body = new FormData();
    body.append("file", file);
    body.append("autoFix", String(autoFix));
    body.append("supplierCurrency", supplierCurrency);

    const response = await fetch("/api/validate", {
      method: "POST",
      body
    });

    const payload = await response.json();
    if (!response.ok) {
      setStatus("error");
      setError(payload.error ?? "Unable to validate this workbook.");
      return;
    }

    setResult(payload);
    setStatus("done");
  }

  function handleInput(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) void validateFile(file);
    event.target.value = "";
  }

  function handleDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    const file = event.dataTransfer.files?.[0];
    if (file) void validateFile(file);
  }

  const filteredIssues = useMemo(() => {
    const lowerQuery = query.trim().toLowerCase();
    return (
      result?.issues.filter((issue) => {
        const matchesSeverity = severity === "all" || issue.severity === severity;
        const matchesField = field === "all" || issue.field === field;
        const searchable = `${issue.rowNumber} ${issue.field} ${issue.currentValue} ${issue.issue} ${issue.suggestedFix}`.toLowerCase();
        return matchesSeverity && matchesField && (!lowerQuery || searchable.includes(lowerQuery));
      }) ?? []
    );
  }, [field, query, result?.issues, severity]);

  const issueFields = useMemo(() => {
    const fields = new Set(result?.issues.map((issue) => issue.field) ?? []);
    return [...fields].sort();
  }, [result?.issues]);

  const supplierEmail = useMemo(() => (result ? buildSupplierEmail(result) : ""), [result]);

  async function copySupplierEmail() {
    if (!supplierEmail) return;
    await navigator.clipboard.writeText(supplierEmail);
    setEmailCopied(true);
    window.setTimeout(() => setEmailCopied(false), 2200);
  }

  return (
    <div className="space-y-6">
      <section className="grid gap-5 lg:grid-cols-[minmax(0,1fr)_320px]">
        <Card>
          <CardHeader>
            <CardTitle>Upload Catalog</CardTitle>
            <CardDescription>Drop a Collaboration Portal supplier catalog workbook and validate the first worksheet.</CardDescription>
          </CardHeader>
          <CardContent>
            <div
              className={cn(
                "flex min-h-[260px] flex-col items-center justify-center rounded-lg border border-dashed border-slate-300 bg-white px-5 py-8 text-center transition-colors",
                status === "validating" && "border-primary bg-emerald-50/60"
              )}
              onDragOver={(event) => event.preventDefault()}
              onDrop={handleDrop}
            >
              <div className="mb-4 flex h-14 w-14 items-center justify-center rounded-md bg-primary/10 text-primary">
                <UploadCloud className="h-7 w-7" aria-hidden="true" />
              </div>
              <h2 className="text-xl font-semibold">Drop your .xlsx file here</h2>
              <p className="mt-2 max-w-md text-sm text-muted-foreground">
                Files are processed through a server-side route handler and are not stored permanently.
              </p>
              <div className="mt-5 flex flex-wrap items-center justify-center gap-3">
                <Button type="button" onClick={() => inputRef.current?.click()} disabled={status === "validating"}>
                  <FileSpreadsheet className="h-4 w-4" aria-hidden="true" />
                  Choose file
                </Button>
                <label className="flex items-center gap-2 rounded-md border bg-background px-3 py-2 text-sm">
                  <input
                    type="checkbox"
                    checked={autoFix}
                    onChange={(event) => setAutoFix(event.target.checked)}
                    className="h-4 w-4 accent-emerald-700"
                  />
                  Auto-fix safe fields
                </label>
              </div>
              <div className="mt-4 w-full max-w-sm text-left">
                <label className="text-sm font-semibold text-slate-700" htmlFor="supplier-currency">
                  Supplier currency
                </label>
                <select
                  id="supplier-currency"
                  className="focus-ring mt-2 h-10 w-full rounded-md border border-input bg-background px-3 text-sm"
                  value={supplierCurrency}
                  onChange={(event) => setSupplierCurrency(event.target.value as SupplierCurrency)}
                  disabled={status === "validating"}
                >
                  <option value="USD">USD - United States supplier</option>
                  <option value="CAD">CAD - Canadian supplier</option>
                </select>
                <p className="mt-2 text-xs leading-5 text-muted-foreground">
                  Currency validation will use this selection for every row in the uploaded catalog.
                </p>
              </div>
              <input ref={inputRef} className="hidden" type="file" accept=".xlsx" onChange={handleInput} />
              {fileName ? <p className="mt-4 text-sm font-medium text-slate-700">{fileName}</p> : null}
              {status === "validating" ? <p className="mt-3 text-sm text-primary">Validating workbook...</p> : null}
              {status === "error" ? (
                <p className="mt-3 flex items-center gap-2 text-sm font-medium text-red-700">
                  <AlertCircle className="h-4 w-4" aria-hidden="true" />
                  {error}
                </p>
              ) : null}
            </div>
          </CardContent>
        </Card>

        <aside className="space-y-5">
          <Card>
            <CardHeader>
              <CardTitle>Catalog Type</CardTitle>
              <CardDescription>Collaboration Portal validation is active for this MVP.</CardDescription>
            </CardHeader>
            <CardContent>
              <select className="focus-ring h-10 w-full rounded-md border border-input bg-background px-3 text-sm" defaultValue="collaboration">
                <option value="collaboration">Collaboration Portal</option>
                <option value="edi" disabled>
                  EDI Enablement - coming soon
                </option>
              </select>
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle>Cleaning</CardTitle>
              <CardDescription>Cleaned downloads trim text, collapse double spaces, and remove forbidden description characters.</CardDescription>
            </CardHeader>
            <CardContent>
              <div className="flex items-start gap-3 rounded-md bg-accent p-3 text-sm text-accent-foreground">
                <Wand2 className="mt-0.5 h-4 w-4 flex-none" aria-hidden="true" />
                <span>
                  Optional auto-fix fills blank Currency with the selected supplier currency, normalizes lowercase currency, fills blank Break Case with 0,
                  and sets blank or zero Minimum Order Qty to 1.
                </span>
              </div>
            </CardContent>
          </Card>
        </aside>
      </section>

      {result ? (
        <section className="space-y-5">
          <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-5">
            {[
              ["Total rows", result.summary.totalRows],
              ["Passed rows", result.summary.passedRows],
              ["Warnings", result.summary.rowsWithWarnings],
              ["Errors", result.summary.rowsWithErrors],
              ["Total issues", result.summary.totalIssues]
            ].map(([label, value]) => (
              <Card key={label}>
                <CardContent className="p-4">
                  <p className="text-sm text-muted-foreground">{label}</p>
                  <p className="mt-2 text-3xl font-semibold">{value}</p>
                </CardContent>
              </Card>
            ))}
          </div>

          <Card>
            <CardHeader className="gap-4 lg:flex lg:flex-row lg:items-start lg:justify-between">
              <div>
                <CardTitle>Validation Results</CardTitle>
                <CardDescription>Every issue by row, field, current value, severity, and recommended fix.</CardDescription>
              </div>
              <div className="flex flex-wrap gap-2">
                <Button variant="outline" onClick={() => downloadBase64Xlsx(result.reportBase64, result.reportFileName)}>
                  <Download className="h-4 w-4" aria-hidden="true" />
                  Validation Report
                </Button>
                <Button onClick={() => downloadBase64Xlsx(result.cleanedBase64, result.cleanedFileName)}>
                  <Download className="h-4 w-4" aria-hidden="true" />
                  Cleaned Catalog
                </Button>
              </div>
            </CardHeader>
            <CardContent>
              <div className="mb-4 grid gap-3 lg:grid-cols-[1fr_160px_220px]">
                <div className="relative">
                  <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" aria-hidden="true" />
                  <Input className="pl-9" placeholder="Search issues, fields, or values" value={query} onChange={(event) => setQuery(event.target.value)} />
                </div>
                <select className="focus-ring h-10 rounded-md border border-input bg-background px-3 text-sm" value={severity} onChange={(event) => setSeverity(event.target.value as "all" | "error" | "warning")}>
                  <option value="all">All</option>
                  <option value="error">Errors</option>
                  <option value="warning">Warnings</option>
                </select>
                <select className="focus-ring h-10 rounded-md border border-input bg-background px-3 text-sm" value={field} onChange={(event) => setField(event.target.value)}>
                  <option value="all">All fields</option>
                  {issueFields.map((issueField) => (
                    <option key={issueField} value={issueField}>
                      {issueField}
                    </option>
                  ))}
                </select>
              </div>

              {filteredIssues.length ? (
                <div className="overflow-hidden rounded-lg border">
                  <div className="max-h-[560px] overflow-auto">
                    <table className="w-full min-w-[980px] border-collapse text-left text-sm">
                      <thead className="sticky top-0 bg-slate-100 text-xs uppercase tracking-wide text-slate-600">
                        <tr>
                          <th className="px-3 py-3">Row</th>
                          <th className="px-3 py-3">Field</th>
                          <th className="px-3 py-3">Current value</th>
                          <th className="px-3 py-3">Severity</th>
                          <th className="px-3 py-3">Issue</th>
                          <th className="px-3 py-3">Suggested fix</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y bg-white">
                        {filteredIssues.map((issue: CatalogIssue, index) => (
                          <tr key={`${issue.rowNumber}-${issue.field}-${index}`} className="align-top">
                            <td className="px-3 py-3 font-medium">{issue.rowNumber || "-"}</td>
                            <td className="px-3 py-3">{issue.field}</td>
                            <td className="max-w-[220px] break-words px-3 py-3 text-muted-foreground">{issue.currentValue || "-"}</td>
                            <td className="px-3 py-3">
                              <Badge tone={issue.severity === "error" ? "error" : "warning"}>{issue.severity}</Badge>
                            </td>
                            <td className="max-w-[320px] px-3 py-3">{issue.issue}</td>
                            <td className="max-w-[320px] px-3 py-3 text-muted-foreground">{issue.suggestedFix}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ) : (
                <div className="flex min-h-[180px] items-center justify-center rounded-lg border bg-emerald-50 text-center text-emerald-800">
                  <div>
                    <CheckCircle2 className="mx-auto mb-2 h-8 w-8" aria-hidden="true" />
                    <p className="font-semibold">No issues match the current filters.</p>
                  </div>
                </div>
              )}
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="gap-4 lg:flex lg:flex-row lg:items-start lg:justify-between">
              <div>
                <CardTitle>Supplier Email</CardTitle>
                <CardDescription>
                  Premium supplier-ready message with clear row and column guidance. If there are more than {INLINE_EMAIL_ISSUE_LIMIT} issues, attach the
                  Excel validation report.
                </CardDescription>
              </div>
              <div className="flex flex-wrap gap-2">
                <Button variant="outline" onClick={copySupplierEmail}>
                  {emailCopied ? <CheckCircle2 className="h-4 w-4" aria-hidden="true" /> : <Clipboard className="h-4 w-4" aria-hidden="true" />}
                  {emailCopied ? "Copied" : "Copy Email"}
                </Button>
                {result.summary.totalIssues > INLINE_EMAIL_ISSUE_LIMIT ? (
                  <Button onClick={() => downloadBase64Xlsx(result.reportBase64, result.reportFileName)}>
                    <Download className="h-4 w-4" aria-hidden="true" />
                    Attachment
                  </Button>
                ) : null}
              </div>
            </CardHeader>
            <CardContent>
              <div className="mb-4 grid gap-3 sm:grid-cols-3">
                <div className="rounded-md border bg-slate-50 p-3">
                  <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Email mode</p>
                  <p className="mt-1 text-sm font-semibold">
                    {result.summary.totalIssues > INLINE_EMAIL_ISSUE_LIMIT ? "Attach Excel report" : "Inline issue list"}
                  </p>
                </div>
                <div className="rounded-md border bg-slate-50 p-3">
                  <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Inline threshold</p>
                  <p className="mt-1 text-sm font-semibold">{INLINE_EMAIL_ISSUE_LIMIT} issues or fewer</p>
                </div>
                <div className="rounded-md border bg-slate-50 p-3">
                  <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Attachment file</p>
                  <p className="mt-1 break-words text-sm font-semibold">{result.reportFileName}</p>
                </div>
              </div>
              <div className="rounded-lg border bg-white">
                <div className="flex items-center gap-2 border-b bg-slate-50 px-4 py-3 text-sm font-semibold">
                  <Mail className="h-4 w-4 text-primary" aria-hidden="true" />
                  Supplier correction request
                </div>
                <pre className="max-h-[460px] overflow-auto whitespace-pre-wrap p-4 text-sm leading-6 text-slate-800">{supplierEmail}</pre>
              </div>
            </CardContent>
          </Card>
        </section>
      ) : null}
    </div>
  );
}
