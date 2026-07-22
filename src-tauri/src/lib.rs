use tauri::{
    menu::{Menu, MenuBuilder, MenuItemBuilder, SubmenuBuilder},
    Emitter, Manager, WindowEvent,
};

const MENU_EVENT: &str = "desktop://menu-action";
const LAUNCH_EVENT: &str = "desktop://launch-files";

pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![force_close])
        .setup(|app| {
            let menu = build_menu(app)?;
            app.set_menu(menu)?;

            if let Some(window) = app.get_webview_window("main") {
                let launch_files = collect_launch_files();
                if !launch_files.is_empty() {
                    let window_for_launch = window.clone();
                    tauri::async_runtime::spawn(async move {
                        let _ = window_for_launch.emit(LAUNCH_EVENT, launch_files);
                    });
                }
            }

            Ok(())
        })
        .on_menu_event(|app, event| {
            if let Some(window) = app.get_webview_window("main") {
                let _ = window.emit(MENU_EVENT, event.id().as_ref());
            }
        })
        .on_window_event(|window, event| {
            if matches!(event, WindowEvent::CloseRequested { api, .. }) {
                api.prevent_close();
                let _ = window.emit(MENU_EVENT, "window-close-requested");
            }
        })
        .run(tauri::generate_context!())
        .expect("error while running Publish Pro");
}

#[tauri::command]
fn force_close(app: tauri::AppHandle) {
    app.exit(0);
}

fn build_menu<R: tauri::Runtime>(app: &tauri::AppHandle<R>) -> tauri::Result<Menu<R>> {
    let file = SubmenuBuilder::new(app, "File")
        .item(&menu_item(app, "new-project", "New Project")?)
        .item(&menu_item(app, "open-project", "Open Project")?)
        .separator()
        .item(&menu_item(app, "save-project", "Save")?)
        .item(&menu_item(app, "save-project-as", "Save As")?)
        .separator()
        .item(&menu_item(app, "open-pdf", "Open PDF")?)
        .item(&menu_item(app, "import-pdf", "Import PDF")?)
        .item(&menu_item(app, "import-docx", "Import Word Document")?)
        .item(&menu_item(app, "import-pptx", "Import PowerPoint Presentation")?)
        .separator()
        .item(&menu_item(app, "publish-pdf", "Publish PDF")?)
        .item(&menu_item(app, "close-project", "Close Project")?)
        .separator()
        .item(&menu_item(app, "exit", "Exit")?)
        .build()?;

    let edit = SubmenuBuilder::new(app, "Edit")
        .item(&menu_item(app, "undo", "Undo")?)
        .item(&menu_item(app, "redo", "Redo")?)
        .separator()
        .item(&menu_item(app, "copy", "Copy")?)
        .item(&menu_item(app, "paste", "Paste")?)
        .item(&menu_item(app, "delete", "Delete")?)
        .item(&menu_item(app, "select-all-pages", "Select All Pages")?)
        .build()?;

    let view = SubmenuBuilder::new(app, "View")
        .item(&menu_item(app, "workspace-import", "Import Workspace")?)
        .item(&menu_item(app, "workspace-assemble", "Assemble Workspace")?)
        .item(&menu_item(app, "workspace-review", "Review Workspace")?)
        .item(&menu_item(app, "workspace-publish", "Publish Workspace")?)
        .separator()
        .item(&menu_item(app, "zoom-in", "Zoom In")?)
        .item(&menu_item(app, "zoom-out", "Zoom Out")?)
        .item(&menu_item(app, "fit-page", "Fit Page")?)
        .item(&menu_item(app, "fit-width", "Fit Width")?)
        .separator()
        .item(&menu_item(app, "toggle-left-panel", "Toggle Left Panel")?)
        .item(&menu_item(app, "toggle-inspector", "Toggle Inspector")?)
        .build()?;

    let help = SubmenuBuilder::new(app, "Help")
        .item(&menu_item(app, "shortcuts", "Keyboard Shortcuts")?)
        .item(&menu_item(app, "about", "About Publish Pro")?)
        .build()?;

    MenuBuilder::new(app)
        .items(&[&file, &edit, &view, &help])
        .build()
}

fn menu_item<R: tauri::Runtime>(app: &tauri::AppHandle<R>, id: &str, text: &str) -> tauri::Result<tauri::menu::MenuItem<R>> {
    MenuItemBuilder::with_id(id, text).build(app)
}

fn collect_launch_files() -> Vec<String> {
    std::env::args()
        .skip(1)
        .filter(|arg| {
            let lower = arg.to_lowercase();
            lower.ends_with(".pproj") || lower.ends_with(".pdf") || lower.ends_with(".docx") || lower.ends_with(".pptx")
        })
        .collect()
}
