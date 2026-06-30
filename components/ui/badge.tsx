import * as React from "react";
import { cn } from "@/lib/utils";

export function Badge({
  className,
  tone = "neutral",
  ...props
}: React.HTMLAttributes<HTMLSpanElement> & { tone?: "neutral" | "error" | "warning" | "suggestion" | "success" }) {
  const tones = {
    neutral: "bg-muted text-muted-foreground",
    error: "bg-red-50 text-red-700 ring-red-200",
    warning: "bg-amber-50 text-amber-800 ring-amber-200",
    suggestion: "bg-blue-50 text-blue-700 ring-blue-200",
    success: "bg-emerald-50 text-emerald-800 ring-emerald-200"
  };

  return (
    <span
      className={cn(
        "inline-flex items-center rounded-sm px-2 py-1 text-xs font-semibold ring-1 ring-inset",
        tones[tone],
        className
      )}
      {...props}
    />
  );
}
