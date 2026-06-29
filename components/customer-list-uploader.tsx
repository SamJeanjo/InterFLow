"use client";

import { ChangeEvent, DragEvent, useRef, useState } from "react";
import { AlertCircle, CheckCircle2, Clipboard, Download, FileSpreadsheet, Mail, UploadCloud } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { cn } from "@/lib/utils";

type Status = "idle" | "formatting" | "done" | "error";

type CustomerListResponse = {
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

export function CustomerListUploader() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [result, setResult] = useState<CustomerListResponse | null>(null);
  const [emailCopied, setEmailCopied] = useState(false);

  async function formatFile(file: File) {
    setError("");
    setResult(null);
    setEmailCopied(false);

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setStatus("error");
      setError("Only .xlsx files are supported.");
      return;
    }

    setStatus("formatting");
    setFileName(file.name);

    const body = new FormData();
    body.append("file", file);

    const response = await fetch("/api/customers", {
      method: "POST",
      body
    });
    const payload = await response.json();

    if (!response.ok) {
      setStatus("error");
      setError(payload.error ?? "Unable to format this customer list.");
      return;
    }

    setResult(payload);
    setStatus("done");
  }

  function handleInput(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) void formatFile(file);
    event.target.value = "";
  }

  function handleDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    const file = event.dataTransfer.files?.[0];
    if (file) void formatFile(file);
  }

  async function copyEmail() {
    if (!result) return;
    await navigator.clipboard.writeText(`Subject: ${result.emailSubject}\n\n${result.emailBody}`);
    setEmailCopied(true);
    window.setTimeout(() => setEmailCopied(false), 2200);
  }

  return (
    <div className="space-y-6">
      <Card>
        <CardHeader>
          <CardTitle>Upload AP Customer List</CardTitle>
          <CardDescription>
            Upload the workbook from AP. The app will create a supplier review workbook with Customer number in column B and Notes at the end.
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div
            className={cn(
              "flex min-h-[260px] flex-col items-center justify-center rounded-lg border border-dashed border-slate-300 bg-white px-5 py-8 text-center transition-colors",
              status === "formatting" && "border-primary bg-emerald-50/60"
            )}
            onDragOver={(event) => event.preventDefault()}
            onDrop={handleDrop}
          >
            <div className="mb-4 flex h-14 w-14 items-center justify-center rounded-md bg-primary/10 text-primary">
              <UploadCloud className="h-7 w-7" aria-hidden="true" />
            </div>
            <h2 className="text-xl font-semibold">Drop the AP customer list here</h2>
            <p className="mt-2 max-w-lg text-sm text-muted-foreground">
              The output workbook is formatted for supplier review and includes an Instructions tab.
            </p>
            <Button className="mt-5" type="button" onClick={() => inputRef.current?.click()} disabled={status === "formatting"}>
              <FileSpreadsheet className="h-4 w-4" aria-hidden="true" />
              Choose file
            </Button>
            <input ref={inputRef} className="hidden" type="file" accept=".xlsx" onChange={handleInput} />
            {fileName ? <p className="mt-4 text-sm font-medium text-slate-700">{fileName}</p> : null}
            {status === "formatting" ? <p className="mt-3 text-sm text-primary">Formatting customer list...</p> : null}
            {status === "error" ? (
              <p className="mt-3 flex items-center gap-2 text-sm font-medium text-red-700">
                <AlertCircle className="h-4 w-4" aria-hidden="true" />
                {error}
              </p>
            ) : null}
          </div>
        </CardContent>
      </Card>

      {result ? (
        <section className="space-y-5">
          <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
            {[
              ["Source sheet", result.summary.sourceSheet],
              ["Customer rows", result.summary.totalRows],
              ["Possible closed", result.summary.possibleClosedRows],
              ["Columns", result.summary.columns]
            ].map(([label, value]) => (
              <Card key={label}>
                <CardContent className="p-4">
                  <p className="text-sm text-muted-foreground">{label}</p>
                  <p className="mt-2 break-words text-2xl font-semibold">{value}</p>
                </CardContent>
              </Card>
            ))}
          </div>

          <Card>
            <CardHeader className="gap-4 lg:flex lg:flex-row lg:items-start lg:justify-between">
              <div>
                <CardTitle>Supplier Package</CardTitle>
                <CardDescription>Download the formatted workbook and copy the email text for the supplier contact.</CardDescription>
              </div>
              <div className="flex flex-wrap gap-2">
                <Button variant="outline" onClick={copyEmail}>
                  {emailCopied ? <CheckCircle2 className="h-4 w-4" aria-hidden="true" /> : <Clipboard className="h-4 w-4" aria-hidden="true" />}
                  {emailCopied ? "Copied" : "Copy Email"}
                </Button>
                <Button onClick={() => downloadBase64Xlsx(result.formattedBase64, result.formattedFileName)}>
                  <Download className="h-4 w-4" aria-hidden="true" />
                  Customer Workbook
                </Button>
              </div>
            </CardHeader>
            <CardContent>
              <div className="rounded-lg border bg-white">
                <div className="flex items-center gap-2 border-b bg-slate-50 px-4 py-3 text-sm font-semibold">
                  <Mail className="h-4 w-4 text-primary" aria-hidden="true" />
                  {result.emailSubject}
                </div>
                <pre className="max-h-[440px] overflow-auto whitespace-pre-wrap p-4 text-sm leading-6 text-slate-800">{result.emailBody}</pre>
              </div>
            </CardContent>
          </Card>
        </section>
      ) : null}
    </div>
  );
}
