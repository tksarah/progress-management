mod excel_store;
mod models;
mod settings;

use crate::excel_store::{build_summary, create_item, delete_item, ensure_workbook, read_items, update_item};
use crate::models::{AppSettings, DashboardResponse, ProgressItem, ProgressPayload, StartupState};
use crate::settings::{load_settings, normalize_settings, save_settings, suggested_excel_path};
use std::fs;
use std::path::PathBuf;

fn configured_excel_path(settings: &AppSettings) -> Option<PathBuf> {
    let trimmed = settings.excel_file_path.trim();

    if trimmed.is_empty() {
        None
    } else {
        Some(PathBuf::from(trimmed))
    }
}

fn require_configured_excel_path(settings: &AppSettings) -> Result<PathBuf, String> {
    configured_excel_path(settings)
        .ok_or_else(|| "Excel ファイルが未設定です。起動ウィザードまたは設定から選択してください。".to_string())
}

fn current_excel_path(settings: &AppSettings) -> Result<PathBuf, String> {
    let path = require_configured_excel_path(settings)?;

    if !path.exists() {
        return Err("指定された Excel ファイルが見つかりません。起動ウィザードまたは設定から選択してください。".to_string());
    }

    ensure_workbook(&path, settings)?;
    Ok(path)
}

#[tauri::command]
fn get_startup_state() -> Result<StartupState, String> {
    let settings = load_settings()?;
    let configured_path = configured_excel_path(&settings);
    let configured_excel_exists = configured_path.as_ref().is_some_and(|path| path.exists());
    let suggested_new_excel_path = suggested_excel_path()?.to_string_lossy().to_string();

    Ok(StartupState {
        settings,
        has_configured_excel: configured_path.is_some(),
        configured_excel_exists,
        needs_onboarding: !configured_excel_exists,
        suggested_new_excel_path,
    })
}

#[tauri::command]
fn get_dashboard(query: Option<String>, status: Option<String>) -> Result<DashboardResponse, String> {
    let settings = load_settings()?;
    let path = current_excel_path(&settings)?;
    let query = query.unwrap_or_default().to_lowercase();
    let status = status.unwrap_or_default();

    let items: Vec<ProgressItem> = read_items(&path)?
        .into_iter()
        .filter(|item| {
            let matches_query = if query.is_empty() {
                true
            } else {
                format!(
                    "{} {} {} {} {} {}",
                    item.kpi_number,
                    item.assignee,
                    item.customer,
                    item.content,
                    item.next_action,
                    item.report_memo
                )
                .to_lowercase()
                .contains(&query)
            };
            let matches_status = if status.is_empty() {
                true
            } else {
                item.status == status
            };
            matches_query && matches_status
        })
        .collect();

    let summary = build_summary(&items);

    Ok(DashboardResponse {
        total: items.len(),
        items,
        summary,
        excel_file_path: settings.excel_file_path.clone(),
        settings,
    })
}

#[tauri::command]
fn set_excel_file_path(excel_file_path: String) -> Result<AppSettings, String> {
    let trimmed = excel_file_path.trim();

    if trimmed.is_empty() {
        return Err("excelFilePath は必須です。".to_string());
    }

    let mut settings = load_settings()?;
    let path = PathBuf::from(trimmed);
    settings.excel_file_path = path.to_string_lossy().to_string();
    let settings = normalize_settings(settings)?;
    ensure_workbook(&path, &settings)?;
    save_settings(&settings)?;
    Ok(settings)
}

#[tauri::command]
fn get_app_settings() -> Result<AppSettings, String> {
    load_settings()
}

#[tauri::command]
fn update_app_settings(settings: AppSettings) -> Result<AppSettings, String> {
    let settings = normalize_settings(settings)?;

    if let Some(path) = configured_excel_path(&settings) {
        ensure_workbook(&path, &settings)?;
    }

    save_settings(&settings)?;
    Ok(settings)
}

#[tauri::command]
fn create_progress(payload: ProgressPayload) -> Result<ProgressItem, String> {
    let settings = load_settings()?;
    let path = current_excel_path(&settings)?;
    create_item(&path, &settings, payload)
}

#[tauri::command]
fn update_progress(id: String, payload: ProgressPayload) -> Result<ProgressItem, String> {
    let settings = load_settings()?;
    let path = current_excel_path(&settings)?;
    update_item(&path, &settings, &id, payload)
}

#[tauri::command]
fn delete_progress(id: String) -> Result<(), String> {
    let settings = load_settings()?;
    let path = current_excel_path(&settings)?;
    delete_item(&path, &settings, &id)
}

#[tauri::command]
fn export_current_excel(export_file_path: String) -> Result<String, String> {
    let trimmed = export_file_path.trim();

    if trimmed.is_empty() {
        return Err("exportFilePath は必須です。".to_string());
    }

    let settings = load_settings()?;
    let source_path = current_excel_path(&settings)?;
    let export_path = PathBuf::from(trimmed);

    if let Some(parent) = export_path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }

    fs::copy(&source_path, &export_path).map_err(|error| error.to_string())?;
    Ok(export_path.to_string_lossy().to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            get_startup_state,
            get_dashboard,
            get_app_settings,
            update_app_settings,
            set_excel_file_path,
            create_progress,
            update_progress,
            delete_progress,
            export_current_excel
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
