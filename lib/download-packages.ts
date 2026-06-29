import JSZip from "jszip";

type PackageFile =
  | {
      kind: "base64";
      path: string;
      content: string;
    }
  | {
      kind: "text";
      path: string;
      content: string;
    };

export function base64ToBytes(base64: string) {
  return Uint8Array.from(atob(base64), (char) => char.charCodeAt(0));
}

export function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

export function downloadBase64Xlsx(base64: string, fileName: string) {
  downloadBlob(
    new Blob([base64ToBytes(base64)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }),
    fileName
  );
}

export function sanitizeFolderName(value: string) {
  return (
    value
      .trim()
      .replace(/\.[A-Za-z0-9]+$/, "")
      .replace(/^V?\d{5,}\s+/i, "")
      .replace(/\b\d{1,2}[-_/]\d{1,2}[-_/]\d{2,4}\b/g, "")
      .replace(/\s+/g, " ")
      .replace(/[<>:"/\\|?*]+/g, "-")
      .replace(/\.+$/g, "")
      .trim() || "Supplier"
  );
}

export function slugFileName(value: string) {
  return sanitizeFolderName(value).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "supplier";
}

export async function downloadSupplierZip(input: {
  supplierName: string;
  packageLabel: "catalog" | "customer-list";
  files: PackageFile[];
}) {
  const supplierFolder = sanitizeFolderName(input.supplierName);
  const zip = new JSZip();

  for (const file of input.files) {
    const fullPath = `${supplierFolder}/${file.path}`;
    if (file.kind === "base64") {
      zip.file(fullPath, base64ToBytes(file.content));
    } else {
      zip.file(fullPath, file.content);
    }
  }

  const blob = await zip.generateAsync({ type: "blob" });
  downloadBlob(blob, `${slugFileName(supplierFolder)}-${input.packageLabel}-package.zip`);
}
