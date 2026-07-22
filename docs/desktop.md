# Publish Pro Desktop

Publish Pro uses Tauri as a thin desktop shell around the existing React/Vite application. There is one frontend codebase; browser development remains available.

## Prerequisites

- Node.js and npm
- Rust toolchain with Cargo
- Tauri platform prerequisites for the target OS
- Windows packaging requires a Windows environment for final installer validation

Tauri v2 requires Rust 1.77.2 or newer.

## Commands

```sh
npm install
npm run dev
npm run build
npm run desktop:dev
npm run desktop:build
npm run desktop:build:debug
```

`npm run dev` and `npm run build` remain the browser/Vite workflow.

`npm run desktop:dev` starts the Tauri desktop shell and points it at the Vite dev server.

`npm run desktop:build` builds the frontend and then produces Tauri bundles in `src-tauri/target/release/bundle`.

## Application Identity

- Product: Publish Pro
- Publisher: Designovation
- Identifier: `com.designovation.publishpro`
- Version: inherited from the app version, currently `0.9.0 Beta`

## Runtime Architecture

The frontend calls `src/runtime/index.ts`, which selects:

- Browser runtime: keeps upload/download, IndexedDB autosave, and localStorage recents.
- Desktop runtime: uses Tauri native dialogs, scoped filesystem access, app-data autosave, app-data recents, native window title, and open/reveal helpers.

Tauri-specific checks are kept inside `src/runtime`.

## File Handling

Desktop native dialogs support:

- Open `.pproj`
- Open/import `.pdf`
- Import `.docx`
- Import `.pptx`
- Save `.pproj`
- Publish `.pdf`

Browser mode keeps the existing file input and download behavior.

Portable `.pproj` files still embed source data and remain portable. Source paths are optional desktop metadata only.

## File Associations

The Tauri bundle associates `.pproj` with Publish Pro.

PDF, DOCX, and PPTX are supported launch/import formats, but Publish Pro does not claim ownership of those extensions in this phase.

## Autosave And Recents

Browser mode:

- Autosave uses IndexedDB.
- Recents use localStorage.

Desktop mode:

- Autosave is stored in the OS app-data directory under `autosaves/latest.pproj`.
- Autosave metadata is stored in `autosaves/index.json`.
- Recent projects are stored in `recent-projects.json`.

Legacy browser autosaves are not deleted automatically.

## Security Permissions

Configured Tauri permissions are intentionally limited to:

- Core window/menu functionality
- Native dialogs
- Opener plugin for exported file open/reveal actions
- Filesystem plugin scoped to app-data and common user document locations

Publish Pro does not upload documents, signatures, comments, or project content to remote services. No telemetry or updater endpoint is configured.

## Native Menu

The native menu emits action IDs to the React app. React remains responsible for project state, document logic, undo/redo, and workspace switching.

Menus included:

- File
- Edit
- View
- Help

## Window Title And Close Guard

The desktop window title reflects the current project and unsaved state.

Close/exit events are intercepted and routed to an in-app unsaved-changes dialog with Save, Discard, and Cancel.

## Packaging Notes

Configured bundle targets:

- Windows NSIS
- Windows MSI
- macOS app
- macOS DMG

Final Windows installer validation must be performed on Windows.

macOS signing/notarization is not configured in this phase.

## Known Limitations

- No production updater is configured yet.
- No code signing is configured.
- Native recent file reopening depends on stored desktop paths and browser security still applies in browser mode.
- Full Windows installer validation requires a Windows build host.
- The desktop close guard is implemented for Tauri close/menu events; browser tab close remains browser-controlled.
