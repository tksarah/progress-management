use crate::excel_store::{load_embedded_settings, save_embedded_settings};
use crate::models::AppSettings;
use serde::Serialize;
use std::collections::BTreeSet;
use std::fs;
use std::path::PathBuf;

const DEFAULT_VISIBLE_COLUMNS: [&str; 6] = ["status", "kpiNumber", "category", "assignee", "updatedAt", "content"];
const LEGACY_DEFAULT_LEAD_SOURCE_OPTIONS: [&str; 6] = [
    "TDW",
    "主催・共催イベント",
    "オフラインイベント",
    "アウトバウンド",
    "社内",
    "個別ネットワーキング",
];

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct StoredJsonSettings {
    excel_file_path: String,
}

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
        "title".to_string(),
        "status".to_string(),
        "kpiNumber".to_string(),
        "category".to_string(),
        "assignee".to_string(),
        "updatedAt".to_string(),
        "customer".to_string(),
        "rank".to_string(),
        "dealSize".to_string(),
        "leadSource".to_string(),
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

pub fn suggested_excel_path() -> Result<PathBuf, String> {
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

fn save_json_settings(settings: &AppSettings) -> Result<(), String> {
    let path = settings_path()?;

    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }

    let content = serde_json::to_string_pretty(&StoredJsonSettings {
        excel_file_path: settings.excel_file_path.clone(),
    })
    .map_err(|error| error.to_string())?;

    fs::write(path, content).map_err(|error| error.to_string())
}

pub fn normalize_settings(mut settings: AppSettings) -> Result<AppSettings, String> {
    settings.excel_file_path = settings.excel_file_path.trim().to_string();
    settings.category_options = normalize_list(settings.category_options);
    settings.assignee_options = normalize_list(settings.assignee_options);
    settings.status_options = normalize_list(settings.status_options);
    settings.rank_options = normalize_list(settings.rank_options);
    settings.lead_source_options = normalize_list(settings.lead_source_options);
    let legacy_default_lead_sources = LEGACY_DEFAULT_LEAD_SOURCE_OPTIONS
        .iter()
        .map(|value| value.to_string())
        .collect::<Vec<String>>();
    if settings.lead_source_options == legacy_default_lead_sources {
        settings.lead_source_options.push("ウェビナー".to_string());
    }
    // Migrate legacy single 'X' rank to new options if the user hasn't already customized
    let current_ranks = settings.rank_options.clone();
    let has_new_ranks = current_ranks.iter().any(|v| v == "X1" || v == "X2" || v == "1");
    if current_ranks.iter().any(|v| v == "X") && !has_new_ranks {
        let mut migrated = Vec::new();
        for v in current_ranks.into_iter() {
            if v == "X" {
                migrated.push("X1".to_string());
                migrated.push("X2".to_string());
                migrated.push("1".to_string());
            } else {
                migrated.push(v);
            }
        }
        settings.rank_options = normalize_list(migrated);
    }
    settings.visible_columns = normalize_visible_columns(settings.visible_columns);

    Ok(settings)
}

pub fn load_settings() -> Result<AppSettings, String> {
    let path = settings_path()?;

    let settings = if path.exists() {
        let content = fs::read_to_string(&path).map_err(|error| error.to_string())?;
        let parsed: AppSettings = serde_json::from_str(&content).map_err(|error| error.to_string())?;

        normalize_settings(parsed)?
    } else {
        normalize_settings(AppSettings::default())?
    };

    let resolved_settings = if settings.excel_file_path.trim().is_empty() {
        settings
    } else {
        let excel_path = PathBuf::from(&settings.excel_file_path);

        if !excel_path.exists() {
            settings
        } else if let Some(mut embedded_settings) = load_embedded_settings(&excel_path)? {
            embedded_settings.excel_file_path = settings.excel_file_path.clone();
            normalize_settings(embedded_settings)?
        } else {
            save_embedded_settings(&excel_path, &settings)?;
            settings
        }
    };

    save_json_settings(&resolved_settings)?;
    Ok(resolved_settings)
}

pub fn save_settings(settings: &AppSettings) -> Result<(), String> {
    let normalized = normalize_settings(settings.clone())?;

    if !normalized.excel_file_path.trim().is_empty() {
        let excel_path = PathBuf::from(&normalized.excel_file_path);

        if excel_path.exists() {
            save_embedded_settings(&excel_path, &normalized)?;
        }
    }

    save_json_settings(&normalized)
}
