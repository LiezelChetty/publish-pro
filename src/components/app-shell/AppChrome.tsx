import { Download, FolderOpen, MoreHorizontal, Printer, Save } from "lucide-react";

type AppChromeProps = {
  logoSrc: string;
  documentName: string;
  isBusy: boolean;
  hasUnsavedChanges: boolean;
  onOpen: () => void;
  onExport: () => void;
};

export function AppChrome({ logoSrc, documentName, isBusy, hasUnsavedChanges, onOpen, onExport }: AppChromeProps) {
  return (
    <>
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
          <button className="chrome-button" disabled title="Save project is not available yet" aria-label="Save project unavailable">
            <Save size={16} />
            <span>Save</span>
          </button>
          <button className="chrome-button primary" onClick={onExport} disabled={isBusy} title="Export edited PDF" aria-label="Export edited PDF">
            <Download size={16} />
            <span>Export</span>
          </button>
          <button className="chrome-button" disabled title="Print is not available yet" aria-label="Print unavailable">
            <Printer size={16} />
          </button>
          <button className="chrome-button icon-only" disabled title="More options" aria-label="More options unavailable">
            <MoreHorizontal size={16} />
          </button>
        </div>
      </header>
      <div className="document-tabs" role="tablist" aria-label="Open documents">
        <button className="document-tab active" role="tab" aria-selected="true" title={documentName}>
          <span>{documentName}</span>
          {hasUnsavedChanges ? <span aria-hidden="true">*</span> : null}
          <span className="tab-close" aria-hidden="true">x</span>
        </button>
      </div>
    </>
  );
}
