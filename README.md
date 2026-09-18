# PMS Rate Sheet Cleaner

A Google Apps Script toolkit that turns messy PMS export data into clean, upload-ready spreadsheets — adding a custom menu to Google Sheets with five tools for fixing merged cells, filling gaps, and validating data before it gets pushed anywhere.

Built around IQware's PMS export/import feature: IQware lets you export current rate/inventory data and import a prepared sheet back in bulk, but the raw export comes with merged header cells and gaps that need cleaning first, and the import format expects one row per room-type-per-date rather than the export's condensed layout. This toolkit automates that reshaping — export, clean, restructure, review, import — instead of manually retyping rates for every date.

## The five tools

Adds a **"Room Pricing Tools"** menu to any Google Sheet it's installed on:

1. **Unmerge and Fill** — breaks apart merged "unit type" header cells and fills the value down into every row it used to span, in one batched pass.
2. **Remove Blank Rows** — deletes rows where a chosen column is empty, without disturbing the header row.
3. **Add Missing Dates (Fill 0s)** — for each room type, fills in any missing date in the range with a zero-value row, so every room type has a complete, gapless daily series.
4. **Repeat Unit Types Across Dates** — builds a template from the unique room types present, cross-checks that ID and name columns agree everywhere (catching data entry mismatches before anything is written), then generates one row per room-type-per-date across a chosen range. Every written row is read back and compared against what was sent before the tool reports success.
5. **Check Unit Type / Date Integrity** — a full pre-upload audit: flags ID/name mismatches, blank fields, duplicate dates, and any room type whose date coverage doesn't match the rest. Writes a dedicated "Integrity Check" report sheet listing everything found.

## Why the verify-after-write pattern

Tool 4 doesn't just write rows and assume they landed — it re-reads every row it just wrote and diffs it against what was sent before declaring success. A silent write failure or partial update in a file headed for a PMS import is far worse than a script that's slow; this catches it before the file ever leaves the spreadsheet.

## Setup

This is a Google Apps Script, not a standalone program — it runs inside a Google Sheet, not on your computer.

1. Open the Google Sheet you want to clean
2. **Extensions → Apps Script**
3. Delete any starter code in the editor, paste in the contents of `RoomPricingTools.gs`
4. Save (Ctrl/Cmd + S), then close the Apps Script tab and reload the Sheet
5. A new **"Room Pricing Tools"** menu appears in the Sheet's menu bar

## Usage

Run the tools in order for a typical cleanup pass: **1 → 2 → 3 → 4 → 5**. Each tool prompts for the column letters or date range it needs via a dialog box — no code editing required per run.

## Tech stack

Google Apps Script (JavaScript running on Google's Sheets API)
