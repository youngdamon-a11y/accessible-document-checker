# Accessible Document Checker

A desktop app for Windows and Mac that checks Word (`.docx`), PowerPoint (`.pptx`) and PDF documents against **WCAG 2.1 level AA** (plus key **PDF/UA** rules for PDFs), fixes what can be fixed, and re-checks the result.

- **Dashboard**: add documents with a button or drag and drop. Each issue shows what is wrong, where it is (paragraph, slide or page, with the text highlighted) and the WCAG success criterion it relates to.
- **Fix** issues one at a time, or **Fix all automatic** in one go. **Undo** is available.
- **Claude suggestions** (optional): alt text, link text and slide titles can be suggested by Claude using your own API key. Nothing is applied until you approve it.
- **Save accessible copy**: fixes are written to a new file named `… (accessible).docx` (or `.pptx` / `.pdf`). **The original is never changed.** The saved copy is read back from disk and checked again automatically.
- **Human checks**: items software cannot judge (reading order, colour-only meaning, alt text quality, etc.) are listed so a person can confirm them.
- **Export report** as PDF, HTML or JSON, including fixes applied, anything remaining and the human-check status.
- Works **offline**. Only Claude suggestions use the internet.

## Install

Installers are built automatically by GitHub Actions for every push to `main` and every pull request:

1. Open the **Actions** tab, choose the latest **Test and build installers** run.
2. Download **installers-windows-latest** (`.exe` installer or `.zip`) or **installers-macos-latest** (`.dmg` or `.zip`).

To make a release with the installers attached, push a tag such as `v0.1.0`; a draft release is created.

The builds are **not code-signed** yet, so:

- **Windows**: SmartScreen may say “Windows protected your PC”. Choose *More info* > *Run anyway*.
- **Mac**: right-click the app and choose *Open* the first time (or allow it in *System Settings > Privacy & Security*).

Code signing needs an Apple Developer ID and a Windows signing certificate; it can be added to the workflow later.

## What is checked and fixed

| | Automatic fixes | Needs your text (Claude can suggest) | Flagged for a person |
|---|---|---|---|
| **Word** | Title, language, heading levels skipped, empty headings, table header rows, colour contrast, typed bullet/number lists, repeated blank lines | Image/chart alt text (or mark decorative), link text, bold text that should be a heading | Merged/nested tables, floating objects and text boxes, documents with no headings |
| **PowerPoint** | Title, language, table header rows, colour contrast (explicit colours), typed lists, title read first | Missing slide titles, duplicate slide titles, picture alt text, link text | Merged table cells, media captions, reading order of other objects |
| **PDF** | Title + show title in window, language, tagged flag, heading levels, table header cells, link descriptions, tab order, bookmarks from headings | Figure alt text (Claude sees the page) | Untagged PDFs, scanned pages (need OCR), fonts not embedded |

Every document also gets the human checks: reading order, alt text quality, use of colour, sensory instructions, images of text and link context.

### Honest limits

- **No tool can make a document fully WCAG compliant automatically.** “Passes the automated checks” means every testable rule passes; the document is only compliant once the human checks are done too.
- **Untagged PDFs are not auto-tagged.** Automatic tagging is usually wrong (reading order, tables, lists) and would hide problems. Fix the source Word or PowerPoint file with this app, then export to PDF with “Document structure tags for accessibility” turned on. For formal PDF/UA sign-off, also validate with PAC or veraPDF.
- PowerPoint text that inherits its colour from the template is not contrast-checked; it is covered by the human checks.
- Old `.doc` / `.ppt` files need saving as `.docx` / `.pptx` first.

## Claude API key

1. Create a key at <https://console.anthropic.com/>.
2. In the app open **Settings**, paste the key, select **Save key**, then **Test connection**.

The key is encrypted with the operating system keychain (Electron `safeStorage`) and stored only on your computer. When you ask for a suggestion, the image (or PDF page) and nearby text are sent to Anthropic. The default model is Claude Opus 5.5; Sonnet 5.5 and Haiku 5.5 can be chosen in Settings for lower cost.

## Accessibility of the app itself

The interface is built to WCAG 2.1 AA: full keyboard use, skip link, landmarks and headings, visible focus, labelled controls, status messages announced to screen readers, colour contrast of at least 4.5:1, support for zoom, dark mode, Windows high contrast and reduced motion. The end-to-end test runs axe-core against every screen.

## Development

Requires Node.js 22.

```bash
npm install
npm start            # run the app
npm test             # unit tests for the checking and fixing engine
npm run fixtures     # regenerate sample documents (needs Python with python-docx, python-pptx, reportlab, pillow)
node test/e2e.js     # drive the real app with Playwright, save copies, export a report and run axe-core
npm run dist:win     # build installers locally (dist:mac must run on a Mac)
```

### Project layout

```
src/core/      checking and fixing engine (plain Node, no Electron)
  docx.js        Word rules and fixes
  pptx.js        PowerPoint rules and fixes
  pdf.js         PDF rules and fixes (pdf-lib to edit, pdf.js to read text)
  contrast.js    WCAG contrast maths
  issues.js      issue model, WCAG criteria, human checks
  report.js      exportable report
  engine.js      format detection, audit, fix, fix all
src/main/      Electron main process (files, dialogs, settings, Claude)
src/renderer/  the dashboard UI (HTML, CSS, JS; no framework)
test/          unit tests, end-to-end test, sample inaccessible documents
```

## Branding

The colours are inspired by the British Council palette (deep indigo with cyan and pink accents). No British Council logos or trademarks are included; add official assets only with permission.
