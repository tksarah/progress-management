use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;

fn default_category_options() -> Vec<String> {
    vec!["営業".to_string(), "マーケティング".to_string()]
}

fn default_assignee_options() -> Vec<String> {
    vec!["山﨑".to_string(), "倉持".to_string(), "西田".to_string()]
}

fn default_status_options() -> Vec<String> {
    vec![
        "進捗中".to_string(),
        "計画中".to_string(),
        "クローズ".to_string(),
        "保留".to_string(),
    ]
}

fn default_rank_options() -> Vec<String> {
    vec![
        "A".to_string(),
        "B".to_string(),
        "C".to_string(),
        "D".to_string(),
        "X1".to_string(),
        "X2".to_string(),
        "1".to_string(),
    ]
}

fn default_visible_columns() -> Vec<String> {
    vec![
        "status".to_string(),
        "kpiNumber".to_string(),
        "category".to_string(),
        "assignee".to_string(),
        "updatedAt".to_string(),
        "content".to_string(),
    ]
}

fn default_lead_source_options() -> Vec<String> {
    vec![
        "TDW".to_string(),
        "主催・共催イベント".to_string(),
        "オフラインイベント".to_string(),
        "アウトバウンド".to_string(),
        "社内".to_string(),
        "個別ネットワーキング".to_string(),
        "ウェビナー".to_string(),
    ]
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProgressItem {
    pub id: String,
    #[serde(default)]
    pub title: String,
    pub kpi_number: String,
    pub category: String,
    pub assignee: String,
    pub created_at: String,
    pub updated_at: String,
    pub status: String,
    pub rank: String,
    pub deal_size: String,
    #[serde(default)]
    pub lead_source: String,
    pub external_stakeholders: String,
    pub internal_departments: String,
    pub customer: String,
    pub content: String,
    pub next_action: String,
    pub report_memo: String,
    pub updated_by: String,
    pub version: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ProgressPayload {
    pub id: Option<String>,
    #[serde(default)]
    pub title: String,
    pub kpi_number: String,
    pub category: String,
    pub assignee: String,
    pub created_at: Option<String>,
    pub updated_at: Option<String>,
    pub status: String,
    pub rank: String,
    pub deal_size: String,
    #[serde(default)]
    pub lead_source: String,
    pub external_stakeholders: String,
    pub internal_departments: String,
    pub customer: String,
    pub content: String,
    pub next_action: String,
    pub report_memo: String,
    pub updated_by: String,
    pub version: Option<u32>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Summary {
    pub total: usize,
    pub by_status: BTreeMap<String, usize>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct DashboardResponse {
    pub items: Vec<ProgressItem>,
    pub total: usize,
    pub summary: Summary,
    pub excel_file_path: String,
    pub settings: AppSettings,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StartupState {
    pub settings: AppSettings,
    pub has_configured_excel: bool,
    pub configured_excel_exists: bool,
    pub needs_onboarding: bool,
    pub suggested_new_excel_path: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase", default)]
pub struct AppSettings {
    pub excel_file_path: String,
    #[serde(default = "default_category_options")]
    pub category_options: Vec<String>,
    #[serde(default = "default_assignee_options")]
    pub assignee_options: Vec<String>,
    #[serde(default = "default_status_options")]
    pub status_options: Vec<String>,
    #[serde(default = "default_rank_options")]
    pub rank_options: Vec<String>,
    #[serde(default = "default_visible_columns")]
    pub visible_columns: Vec<String>,
    #[serde(default = "default_lead_source_options")]
    pub lead_source_options: Vec<String>,
}

impl Default for AppSettings {
    fn default() -> Self {
        Self {
            excel_file_path: String::new(),
            category_options: default_category_options(),
            assignee_options: default_assignee_options(),
            status_options: default_status_options(),
            rank_options: default_rank_options(),
            visible_columns: default_visible_columns(),
            lead_source_options: default_lead_source_options(),
        }
    }
}
