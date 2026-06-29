import Link from "next/link";
import { ArrowLeft, UsersRound } from "lucide-react";
import { CustomerListUploader } from "@/components/customer-list-uploader";
import { Button } from "@/components/ui/button";

export default function CustomersPage() {
  return (
    <main className="min-h-screen">
      <header className="border-b bg-white">
        <div className="mx-auto flex max-w-7xl flex-col gap-4 px-5 py-6 sm:px-8 lg:flex-row lg:items-center lg:justify-between">
          <div>
            <Button asChild variant="ghost" className="-ml-3 mb-2">
              <Link href="/">
                <ArrowLeft className="h-4 w-4" aria-hidden="true" />
                Dashboard
              </Link>
            </Button>
            <div className="flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-md bg-primary text-primary-foreground">
                <UsersRound className="h-5 w-5" aria-hidden="true" />
              </div>
              <div>
                <h1 className="text-2xl font-semibold">Customer List</h1>
                <p className="text-sm text-muted-foreground">AP workbook formatting and supplier review email</p>
              </div>
            </div>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-7xl px-5 py-6 sm:px-8">
        <CustomerListUploader />
      </div>
    </main>
  );
}
