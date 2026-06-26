import Link from "next/link";
import { ArrowRight, CheckCircle2, FileCheck2, ShieldCheck } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";

export default function HomePage() {
  return (
    <main className="min-h-screen">
      <section className="border-b bg-white">
        <div className="mx-auto flex max-w-7xl flex-col gap-8 px-5 py-10 sm:px-8 lg:flex-row lg:items-center lg:justify-between lg:py-14">
          <div className="max-w-3xl">
            <div className="mb-4 inline-flex items-center gap-2 rounded-md bg-emerald-50 px-3 py-1 text-sm font-semibold text-emerald-800 ring-1 ring-emerald-200">
              <ShieldCheck className="h-4 w-4" aria-hidden="true" />
              Collaboration Portal MVP
            </div>
            <h1 className="text-4xl font-semibold tracking-normal text-slate-950 sm:text-5xl">Catalog Validator</h1>
            <p className="mt-4 max-w-2xl text-lg leading-8 text-slate-600">
              Upload your supplier catalog and validate it before submission. The app flags every row-level issue and creates downloadable report and cleaned catalog workbooks.
            </p>
          </div>
          <Button asChild size="lg">
            <Link href="/validate">
              Upload Catalog
              <ArrowRight className="h-4 w-4" aria-hidden="true" />
            </Link>
          </Button>
        </div>
      </section>

      <section className="mx-auto max-w-7xl px-5 py-8 sm:px-8">
        <div className="mb-5 grid gap-4 lg:grid-cols-[minmax(0,1fr)_300px]">
          <Card>
            <CardContent className="grid gap-4 p-5 sm:grid-cols-3">
              {[
                ["Upload", "Accept .xlsx supplier catalog workbooks."],
                ["Validate", "Detect headers and apply Collaboration Portal rules."],
                ["Download", "Export issue reports and cleaned catalog files."]
              ].map(([title, description]) => (
                <div key={title} className="rounded-md border bg-slate-50 p-4">
                  <FileCheck2 className="mb-3 h-5 w-5 text-primary" aria-hidden="true" />
                  <h2 className="font-semibold">{title}</h2>
                  <p className="mt-2 text-sm leading-6 text-muted-foreground">{description}</p>
                </div>
              ))}
            </CardContent>
          </Card>

          <Card>
            <CardContent className="p-5">
              <label className="text-sm font-semibold text-slate-700" htmlFor="catalog-type">
                Catalog type
              </label>
              <select id="catalog-type" className="focus-ring mt-2 h-10 w-full rounded-md border border-input bg-background px-3 text-sm" defaultValue="collaboration">
                <option value="collaboration">Collaboration Portal</option>
                <option value="edi" disabled>
                  EDI Enablement - coming soon
                </option>
              </select>
              <div className="mt-4 flex items-start gap-2 text-sm text-emerald-800">
                <CheckCircle2 className="mt-0.5 h-4 w-4 flex-none" aria-hidden="true" />
                <span>No database is required for the MVP and uploads are processed without permanent storage.</span>
              </div>
            </CardContent>
          </Card>
        </div>
      </section>
    </main>
  );
}
