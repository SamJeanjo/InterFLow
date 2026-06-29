import { NextResponse } from "next/server";
import { formatCustomerListWorkbook } from "@/lib/customer-list-formatter";

export const runtime = "nodejs";

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");

    if (!(file instanceof File)) {
      return NextResponse.json({ error: "Upload a .xlsx customer list file." }, { status: 400 });
    }

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      return NextResponse.json({ error: "Only .xlsx files are supported." }, { status: 400 });
    }

    const result = await formatCustomerListWorkbook({
      fileName: file.name,
      buffer: await file.arrayBuffer()
    });

    return NextResponse.json(result);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unable to format this customer list.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
