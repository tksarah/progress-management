use crate::models::AppSettings;
use std::collections::BTreeSet;
use std::fs;
use std::path::PathBuf;

const DEFAULT_VISIBLE_COLUMNS: [&str; 6] = ["status", "kpiNumber", "category", "assignee", "updatedAt", "content"];

fn normalize_list(values: Vec<String>) -> Vec<String> {
    let mut seen = BTreeSet::new();

    values
        .into_iter()
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty())
        .filter(|value| seen.insert(value.clone()))
        .collect()
}

fn normalize_visible_columns(values: Vec<String>) -> Vec<String> {
    let allowed = BTreeSet::from([
        "status".to_string(),
        "kpiNumber".to_string(),
        "category".to_string(),
        "assignee".to_string(),
        "updatedAt".to_string(),
        "customer".to_string(),
        "rank".to_string(),
        "dealSize".to_string(),
        "content".to_string(),
        "nextAction".to_string(),
        "reportMemo".to_string(),
    ]);
    let mut seen = BTreeSet::new();
    let normalized: Vec<String> = values
        .into_iter()
        .map(|value| value.trim().to_string())
        .filter(|value| allowed.contains(value))
        .filter(|value| seen.insert(value.clone()))
        .collect();

    if normalized.is_empty() {
        return DEFAULT_VISIBLE_COLUMNS.iter().map(|value| value.to_string()).collect();
    }

    normalized
}

fn config_root() -> Result<PathBuf, String> {
    if let Some(path) = dirs::config_dir() {
        return Ok(path.join("ProgressTrackerPoc"));
    }

    std::env::current_dir()
        .map(|path| path.join("ProgressTrackerPoc"))
        .map_err(|error| error.to_string())
}

fn default_excel_path() -> Result<PathBuf, String> {
    if let Some(path) = dirs::document_dir() {
        return Ok(path.join("ProgressTrackerPoc").join("progress.xlsx"));
    }

    std::env::current_dir()
        .map(|path| path.join("progress.xlsx"))
        .map_err(|error| error.to_string())
}

fn settings_path() -> Result<PathBuf, String> {
    Ok(config_root()?.join("settings.json"))
}

pub fn normalize_settings(mut settings: AppSettings) -> Result<AppSettings, String> {
    settings.excel_file_path = if settings.excel_file_path.trim().is_empty() {
        default_excel_path()?.to_string_lossy().to_string()
    } else {
        settings.excel_file_path.trim().to_string()
    };
    settings.category_options = normalize_list(settings.category_options);
    settings.assignee_options = normalize_list(settings.assignee_options);
    settings.status_options = normalize_list(settings.status_options);
    settings.rank_options = normalize_list(settings.rank_options);
    settings.visible_columns = normalize_visible_columns(settings.visible_columns);

    Ok(settings)
}

pub fn load_settings() -> Result<AppSettings, String> {
    let path = settings_path()?;

    if path.exists() {
        let content = fs::read_to_string(&path).map_err(|error| error.to_string())?;
        let parsed: AppSettings = serde_json::from_str(&content).map_err(|error| error.to_string())?;
        let settings = normalize_settings(parsed.clone())?;

        if settings != parsed {
            save_settings(&settings)?;
        }

        return Ok(settings);
    }

    let settings = normalize_settings(AppSettings::default())?;
    save_settings(&settings)?;
    Ok(settings)
}

pub fn save_settings(settings: &AppSettings) -> Result<(), String> {
    let path = settings_path()?;

    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }

    let normalized = normalize_settings(settings.clone())?;
    let content = serde_json::to_string_pretty(&normalized).map_err(|error| error.to_string())?;
    fs::write(path, content).map_err(|error| error.to_string())
}
