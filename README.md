# Accessible Document Checker

A desktop app for Windows and Mac that checks Word (`.docx`), PowerPoint (`.pptx`) and PDF documents against **WCAG 2.1 level AA** (plus key **PDF/UA** rules for PDFs), fixes what can be fixed, and re-checks the result.

- **Dashboard**: add documents with a button or drag and drop. Each issue shows what is wrong, where it is (paragraph, slide or page, with the text highlighted) and the WCAG success criterion it relates to.
- **Fix** issues one at a time, **Fix all automatic** in one go, or **Fix all, including suggested text**, which also fills in alt text, link text and slide titles from suggestions and lists every one for you to review. **Undo** is available.
- **Suggestions without AI**: alt text for charts is written from the chart's own data, alt text for pictures is taken from a nearby caption (“Figure 1: …”), link text from the web address, and slide titles from the slide's first line.
- **AI suggestions** (optional) for pictures with no caption. Choose in Settings: a **local model on your computer** (Ollama, free and fully offline), **Claude** with an API key, or **OpenAI** with an API key.
- **Save accessible copy**: fixes are written to a new file named `… (accessible).docx` (or `.pptx` / `.pdf`). **The original is never changed.** The saved copy is read back from disk and checked again automatically.
- **Human checks**: items software cannot judge (reading order, colour-only meaning, alt text quality, etc.) are listed so a person can confirm them.
- **Export report** as PDF, HTML or JSON, including fixes applied, anything remaining and the human-check status.
- Works **offline**, including picture descriptions when you use the local model. Only the Claude and OpenAI options use the internet.

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
| **Word** | Title, language, heading levels skipped, empty headings, table header rows, colour contrast, typed bullet/number lists, repeated blank lines | Image alt text (from caption or AI), chart alt text (from chart data), link text (from address), bold text that should be a heading | Merged/nested tables, floating objects and text boxes, documents with no headings |
| **PowerPoint** | Title, language, table header rows, colour contrast (explicit colours), typed lists, title read first | Missing slide titles (from first line), duplicate slide titles, picture alt text (caption or AI), chart alt text (chart data), link text | Merged table cells, media captions, reading order of other objects |
| **PDF** | Title + show title in window, language, tagged flag, heading levels, table header cells, link descriptions, tab order, bookmarks from headings | Figure alt text (from caption, or the AI sees the page) | Untagged PDFs, scanned pages (need OCR), fonts not embedded |

Every document also gets the human checks: reading order, alt text quality, use of colour, sensory instructions, images of text and link context.

### Honest limits

- **No tool can make a document fully WCAG compliant automatically.** “Passes the automated checks” means every testable rule passes; the document is only compliant once the human checks are done too.
- **Untagged PDFs are not auto-tagged.** Automatic tagging is usually wrong (reading order, tables, lists) and would hide problems. Fix the source Word or PowerPoint file with this app, then export to PDF with “Document structure tags for accessibility” turned on. For formal PDF/UA sign-off, also validate with PAC or veraPDF.
- PowerPoint text that inherits its colour from the template is not contrast-checked; it is covered by the human checks.
- Old `.doc` / `.ppt` files need saving as `.docx` / `.pptx` first.

## Suggestions: offline, local model, Claude or OpenAI

Open **Settings** and choose who writes picture descriptions:

| Option | Internet | Cost | Set up |
|---|---|---|---|
| **A model on this computer (Ollama)** | No. Nothing leaves your computer | Free | Install [Ollama](https://ollama.com), then run `ollama pull gemma3` (about 3 GB). Other vision models work too, such as `qwen2.5vl` or `llama3.2-vision`. Select **Test connection**. |
| **Claude, with an API key** | Yes | Billed to your Anthropic account | Create a key at <https://console.anthropic.com/>, paste it, **Save key**, **Test connection**. Default model Claude Opus 5.5; Sonnet 5.5 and Haiku 5.5 cost less. |
| **OpenAI, with an API key** | Yes | Billed to your OpenAI account | Create a key at <https://platform.openai.com/>, paste it, **Save key**, **Test connection**. Any OpenAI model that accepts images (default `gpt-5`). |
| **No AI** | No | Free | Nothing. Offline suggestions still work; pictures with no caption need you to type the alt text. |

API keys are encrypted with the operating system keychain (Electron `safeStorage`) and stored only on your computer. With Claude or OpenAI, the picture (or the PDF page) and nearby text are sent to that company when a suggestion is made. With the local model, nothing is sent anywhere.

A local model is usually less accurate than Claude, so review its suggestions carefully.

### Why not “sign in with Claude” or “sign in with ChatGPT”?

- **Claude Pro / Max**: Anthropic's terms do not allow third-party apps to use a Claude subscription (consumer sign-in). Apps must use an API key. See Anthropic's [legal and compliance notes](https://code.claude.com/docs/en/legal-and-compliance) and [coverage of the policy](https://www.theregister.com/2026/02/20/anthropic_clarifies_ban_third_party_claude_access/).
- **ChatGPT Plus / Pro**: OpenAI's “Sign in with ChatGPT” with plan usage launched in September 2026 for a limited set of partners; other apps must apply through OpenAI's interest form ([OpenAI help article](https://help.openai.com/en/articles/20001542-using-your-chatgpt-plan-in-other-apps-and-sites)). If OpenAI approves this app, it can be added as a fourth option.

### What “Fix all, including suggested text” does

1. Applies every automatic fix.
2. Fills in text that can be worked out without AI: captions, chart data, link addresses, first line of a slide, “(continued)” for duplicate slide titles.
3. If an AI option is chosen, asks it for alt text for pictures that have no caption.
4. Lists every filled-in item under **Suggested text to review**, and in the report, so a person checks them.

Bold text that only *looks* like a heading is not changed automatically, because it may not be a heading.

## What can never be fixed automatically

Some WCAG 2.1 AA requirements need human judgement or a different tool, so no app (with or without AI) can fix every possible issue on its own:

- Whether alt text, link text and headings are **accurate and meaningful** (1.1.1, 2.4.4, 2.4.6). Suggestions help, but a person must confirm.
- **Reading order** of floating objects and complex layouts (1.3.2).
- Meaning conveyed by **colour, shape or position alone** (1.4.1, 1.3.3).
- **Captions and transcripts** for video and audio (1.2.x).
- **Tagging an untagged PDF** and **OCR for scanned PDFs**: automatic results are often wrong. Fix the source file and export again, or use a specialist tool.
- Contrast of PowerPoint text that inherits template colours, and text inside images.

These are listed as human checks in every report.

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
  describe.js    offline suggestions: chart descriptions, captions
  engine.js      format detection, audit, fix, fix all
src/main/      Electron main process (files, dialogs, settings, AI providers, PDF page rendering)
src/renderer/  the dashboard UI (HTML, CSS, JS; no framework)
test/          unit tests, end-to-end test, sample inaccessible documents
```

## Branding

The colours are inspired by the British Council palette (deep indigo with cyan and pink accents). No British Council logos or trademarks are included; add official assets only with permission.
