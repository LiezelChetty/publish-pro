import { Download, FilePlus2, FileText, FolderOpen, Keyboard, Monitor, Moon, MoreHorizontal, Printer, Redo2, Save, Sun, Undo2, X } from "lucide-react";

export type ThemeMode = "light" | "dark" | "system";

type AppChromeProps = {
  logoSrc: string;
  documentName: string;
  saveStatus: string;
  isBusy: boolean;
  hasUnsavedChanges: boolean;
  themeMode: ThemeMode;
  resolvedTheme: "light" | "dark";
  onThemeChange: (mode: ThemeMode) => void;
  onNewProject: () => void;
  onOpenProject: () => void;
  onOpen: () => void;
  onImportDocx: () => void;
  onImportPptx: () => void;
  canSaveProject: boolean;
  canUndo: boolean;
  canRedo: boolean;
  onSaveProject: () => void;
  onSaveProjectAs: () => void;
  onCloseProject: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onShowShortcuts: () => void;
  onExport: () => void;
};

const themeOptions: Array<{ mode: ThemeMode; label: string; icon: typeof Sun }> = [
  { mode: "light", label: "Light", icon: Sun },
  { mode: "dark", label: "Dark", icon: Moon },
  { mode: "system", label: "System", icon: Monitor },
];

export function AppChrome({
  logoSrc,
  documentName,
  saveStatus,
  isBusy,
  hasUnsavedChanges,
  themeMode,
  resolvedTheme,
  onThemeChange,
  onNewProject,
  onOpenProject,
  onOpen,
  onImportDocx,
  onImportPptx,
  canSaveProject,
  canUndo,
  canRedo,
  onSaveProject,
  onSaveProjectAs,
  onCloseProject,
  onUndo,
  onRedo,
  onShowShortcuts,
  onExport,
}: AppChromeProps) {
  return (
    <header className="app-chrome">
      <div className="app-identity">
        <img src={logoSrc} alt="" aria-hidden="true" />
        <span className="product-name">Publish Pro</span>
      </div>
      <div className="document-title" title={documentName}>
        <span>{documentName}</span>
        {hasUnsavedChanges ? <span className="unsaved-indicator" aria-label="Unsaved changes">*</span> : null}
        <small>{saveStatus}</small>
      </div>
      <div className="chrome-actions">
        <button className="chrome-button icon-only" onClick={onNewProject} disabled={isBusy} title="New project" aria-label="New project">
          <FilePlus2 size={16} />
        </button>
        <button className="chrome-button" onClick={onOpen} disabled={isBusy} title="Open PDF" aria-label="Open PDF">
          <FolderOpen size={16} />
          <span>Open</span>
        </button>
        <button className="chrome-button icon-only" onClick={onSaveProject} disabled={isBusy || !canSaveProject} title="Save project" aria-label="Save project">
          <Save size={16} />
        </button>
        <button className="chrome-button icon-only" onClick={onUndo} disabled={!canUndo} title="Undo" aria-label="Undo">
          <Undo2 size={16} />
        </button>
        <button className="chrome-button icon-only" onClick={onRedo} disabled={!canRedo} title="Redo" aria-label="Redo">
          <Redo2 size={16} />
        </button>
        <button className="chrome-button primary" onClick={onExport} disabled={isBusy} title="Export edited PDF" aria-label="Export edited PDF">
          <Download size={16} />
          <span>Publish</span>
        </button>
        <details className="chrome-menu">
          <summary className="chrome-button icon-only" title="More options" aria-label="More options">
            <MoreHorizontal size={16} />
          </summary>
          <div className="chrome-menu-panel">
            <button onClick={onNewProject} disabled={isBusy}><FilePlus2 size={15} />New Project</button>
            <button onClick={onOpenProject} disabled={isBusy}><FolderOpen size={15} />Open Project</button>
            <button onClick={onOpen} disabled={isBusy}><FolderOpen size={15} />Open PDF</button>
            <button onClick={onImportDocx} disabled={isBusy}><FileText size={15} />Import Word document</button>
            <button onClick={onImportPptx} disabled={isBusy}><FileText size={15} />Import PowerPoint presentation</button>
            <button onClick={onSaveProject} disabled={isBusy || !canSaveProject}><Save size={15} />Save Project</button>
            <button onClick={onSaveProjectAs} disabled={isBusy || !canSaveProject}><Save size={15} />Save Project As</button>
            <button onClick={onCloseProject} disabled={isBusy || !canSaveProject}><X size={15} />Close Project</button>
            <button onClick={onShowShortcuts}><Keyboard size={15} />Keyboard Shortcuts</button>
            <button disabled><Printer size={15} />Print</button>
            <div className="appearance-menu" role="group" aria-label="Appearance">
              <span>Appearance</span>
              {themeOptions.map((option) => {
                const Icon = option.icon;
                return (
                  <button
                    className={themeMode === option.mode ? "active" : ""}
                    key={option.mode}
                    onClick={() => onThemeChange(option.mode)}
                    aria-pressed={themeMode === option.mode}
                    title={`${option.label} appearance`}
                  >
                    <Icon size={15} />
                    {option.label}
                    {themeMode === option.mode ? <small>{option.mode === "system" ? resolvedTheme : "active"}</small> : null}
                  </button>
                );
              })}
            </div>
          </div>
        </details>
      </div>
    </header>
  );
}
