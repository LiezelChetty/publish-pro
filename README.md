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

## Windows beta build

Build Windows beta artifacts on a Windows machine or GitHub Actions Windows runner:

```bash
npm ci
npm run desktop:build:windows
```

Expected Windows outputs:

- `src-tauri/target/x86_64-pc-windows-msvc/release/publish-pro.exe`
- `src-tauri/target/x86_64-pc-windows-msvc/release/bundle/msi/*.msi`
- `src-tauri/target/x86_64-pc-windows-msvc/release/bundle/nsis/*-setup.exe`

Windows prerequisites include Node.js, npm, Rust, Visual Studio Build Tools with the Desktop development with C++ workload, WiX Toolset v3 for MSI packaging, NSIS for setup `.exe` packaging, and Microsoft Edge WebView2 Runtime for running the installed app. The Tauri installer configuration uses WebView2's downloaded bootstrapper when the runtime is missing.

## Build

```bash
npm run build
```

The app is built with React, Vite, PDF.js, pdf-lib, and Tauri for desktop packaging.
