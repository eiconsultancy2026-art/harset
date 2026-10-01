# Harset Projects

Monorepo containing multiple independent applications.

## Vercel deployment

Create a separate Vercel project for each application and set its **Root Directory** to the corresponding folder under `projects/`.

Examples:
- `projects/bank-statement-to-excel`
- `projects/bank-to-tally`
- `projects/tally-bank-import`
- `projects/temple-accounting`
- `projects/email-automator`
- `projects/form-10bd`
- `projects/tally-schedule-3`

Do not deploy the repository root as a single Vercel application because it contains multiple independent applications.
