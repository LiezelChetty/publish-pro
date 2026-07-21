import { Download, FolderOpen, Monitor, Moon, MoreHorizontal, Printer, Save, Sun } from "lucide-react";

export type ThemeMode = "light" | "dark" | "system";

type AppChromeProps = {
  logoSrc: string;
  documentName: string;
  isBusy: boolean;
  hasUnsavedChanges: boolean;
  themeMode: ThemeMode;
  resolvedTheme: "light" | "dark";
  onThemeChange: (mode: ThemeMode) => void;
  onOpen: () => void;
  onExport: () => void;
};

const themeOptions: Array<{ mode: ThemeMode; label: string; icon: typeof Sun }> = [
  { mode: "light", label: "Light", icon: Sun },
  { mode: "dark", label: "Dark", icon: Moon },
  { mode: "system", label: "System", icon: Monitor },
];

export function AppChrome({ logoSrc, documentName, isBusy, hasUnsavedChanges, themeMode, resolvedTheme, onThemeChange, onOpen, onExport }: AppChromeProps) {
  return (
    <header className="app-chrome">
      <div className="app-identity">
        <img src={logoSrc} alt="" aria-hidden="true" />
        <span className="product-name">Publish Pro</span>
      </div>
      <div className="document-title" title={documentName}>
        <span>{documentName}</span>
        {hasUnsavedChanges ? <span className="unsaved-indicator" aria-label="Unsaved changes">*</span> : null}
      </div>
      <div className="chrome-actions">
        <button className="chrome-button" onClick={onOpen} disabled={isBusy} title="Open PDF" aria-label="Open PDF">
          <FolderOpen size={16} />
          <span>Open</span>
        </button>
        <button className="chrome-button icon-only" disabled title="Save project is not available yet" aria-label="Save project unavailable">
          <Save size={16} />
        </button>
        <button className="chrome-button primary" onClick={onExport} disabled={isBusy} title="Export edited PDF" aria-label="Export edited PDF">
          <Download size={16} />
          <span>Export</span>
        </button>
        <details className="chrome-menu">
          <summary className="chrome-button icon-only" title="More options" aria-label="More options">
            <MoreHorizontal size={16} />
          </summary>
          <div className="chrome-menu-panel">
            <button onClick={onOpen} disabled={isBusy}><FolderOpen size={15} />Open PDF</button>
            <button disabled><Save size={15} />Save project</button>
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
