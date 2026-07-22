# Publish Pro

Version 0.9.0 Beta

Publish Pro is a local-first professional PDF publishing workspace. It supports PDF, DOCX, and PPTX import, project files, page assembly, annotations, comments, bookmarks, publishing marks, and PDF export.

## Run locally

```bash
npm install
npm run dev
```

## Desktop development

Publish Pro includes a first-phase Tauri desktop shell while preserving the browser workflow.

```bash
npm run desktop:dev
npm run desktop:build
```

See [desktop setup notes](docs/desktop.md) for prerequisites, packaging notes, app-data storage, and security permissions.

## Build

```bash
npm run build
```

The app is built with React, Vite, PDF.js, pdf-lib, and Tauri for desktop packaging.
