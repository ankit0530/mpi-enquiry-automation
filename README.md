# AI Enquiry Automation - Moti Packaging Industries (prototype)

Turns customer enquiries (from a Google Form or pasted rows) into organised sales leads:
cleans and validates each enquiry, classifies it, saves it in a structured sheet, writes a
customer reply draft, routes uncertain or unusual cases to a human, and shows a dashboard.

**Rules make every decision. AI (Google Gemini) only writes draft messages.**
If the AI fails or its draft fails a safety check, a fixed template is used.

## Files

| File | What it is |
|---|---|
| `Code.gs` | The complete Google Apps Script (rules engine, drafts, dashboard, form intake, auto-run) |
| `README.md` | This file |

The Google Sheet (view-only link in the submission document) has these tabs:
`Raw_Enquiries`, `Config`, `Structured_Leads`, `Drafts`, `Dashboard`, `Error_Log`.

## How information moves

Google Form or pasted row -> `Raw_Enquiries` -> rules engine (limits from `Config`)
-> `Structured_Leads` -> draft generator -> `Drafts` (human approval) -> `Dashboard`.
Failures are written to `Error_Log`.

## Setup (to run your own copy)

1. Create a Google Sheet with the tabs listed above (exact names).
2. Open Extensions > Apps Script, delete the default code, paste `Code.gs`, and save.
3. Project Settings > Script Properties > add `GEMINI_API_KEY` with your own key
   (from Google AI Studio). The key is never stored in the code or the sheet.
4. Run `setupConfig` once to create the `Config` tab (it resets all Config values).
5. Reload the sheet. A menu called **MPI Automation** appears.
6. Add enquiries to `Raw_Enquiries` (columns: enquiry_id, customer name, company, mobile, email,
   application, product category, can type, diameter, height, quantity, printing, delivery location,
   required date). Leave `enquiry_id` empty and it is assigned automatically.
7. Menu > **Run pipeline now**. Optional: **Turn ON auto-run (every 5 min)**.
8. Optional Google Form intake: menu > **Create enquiry form (one time)**, then **Turn ON form intake**.

## Changing business rules

Edit the `Config` tab (diameter and height range, MOQ, high-value quantity, minimum lead days,
review keywords) and run the pipeline again. No code change is needed.

## Status rules (priority order)

1. Possible duplicate enquiry - same mobile or email as an earlier enquiry
2. Not currently serviceable - not a composite can, size outside range, or quantity below MOQ
3. More information required - a mandatory field is missing or unclear
4. Ready for quotation

## Notes

- The limits in `Config` are assumptions. Replace them with the company's real manufacturing limits.
- Nothing is sent to customers automatically. Drafts need human approval.
- Do not commit API keys or real customer data to this repository.
