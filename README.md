# Catalog Validator

Premium web-based supplier catalog validation app for Collaboration Portal catalog files.

## Features

- Upload `.xlsx` supplier catalog workbooks
- Detect the Collaboration Portal header row
- Validate row-level catalog rules
- Filter issues by severity, field, and search text
- Download a validation report workbook
- Download a cleaned catalog workbook
- Select expected supplier currency: USD or CAD

## Tech Stack

- Next.js App Router
- TypeScript
- Tailwind CSS
- shadcn-style UI primitives
- ExcelJS
- Zod

## Local Development

```bash
pnpm install
pnpm dev
```

Open `http://127.0.0.1:3000`.
