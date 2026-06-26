import { NextResponse } from "next/server";
import { uploadOptionsSchema } from "@/lib/catalog-rules";
import { validateCatalogWorkbook } from "@/lib/catalog-validator";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");

    if (!(file instanceof File)) {
      return NextResponse.json({ error: "Upload a .xlsx catalog file." }, { status: 400 });
    }

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      return NextResponse.json({ error: "Only .xlsx files are supported for the MVP." }, { status: 400 });
    }

    const options = uploadOptionsSchema.parse({
      autoFix: formData.get("autoFix") === "true",
      supplierCurrency: formData.get("supplierCurrency") || "USD"
    });

    const result = await validateCatalogWorkbook({
      fileName: file.name,
      buffer: await file.arrayBuffer(),
      autoFix: options.autoFix,
      supplierCurrency: options.supplierCurrency
    });

    return NextResponse.json(result);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unable to validate this workbook.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
