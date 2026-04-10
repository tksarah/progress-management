use crate::models::{AppSettings, ProgressItem, ProgressPayload, Summary};
use calamine::{open_workbook_auto, Reader};
use chrono::Utc;
use rust_xlsxwriter::{Color, Format, Workbook};
use std::collections::BTreeMap;
use std::collections::HashMap;
use std::fs;
use std::path::Path;
use uuid::Uuid;

const SHEET_NAME: &str = "Progress";
const HEADERS: [&str; 17] = [
    "RowID",
    "KPI番号",
    "カテゴリー",
    "担当者名",
    "登録日",
    "更新日",
    "ステータス",
    "ランク",
    "ディールサイズ",
    "社外関係者",
    "社内関連部署",
    "顧客名",
    "内容",
    "NextAction",
    "報告メモ",
    "更新者",
    "Version",
];

const LEGACY_HEADERS: [&str; 13] = [
    "RowID",
    "KPI番号",
    "担当者名",
    "登録日",
    "更新日",
    "ステータス",
    "社外関係者",
    "社内関連部署",
    "顧客名",
    "内容",
    "NextAction",
    "更新者",
    "Version",
];

fn parse_deal_size_units(value: &str) -> Option<u64> {
    let trimmed = value.trim();

    if trimmed.is_empty() {
        return None;
    }

    let normalized = trimmed.replace(',', "").replace(' ', "");

    if normalized.chars().all(|character| character.is_ascii_digit()) {
        return normalized.parse::<u64>().ok();
    }

    if let Some(units) = normalized.strip_suffix("万円").or_else(|| normalized.strip_suffix('万')) {
        return units.parse::<u64>().ok();
    }

    if let Some(yen) = normalized.strip_suffix('円') {
        let parsed = yen.parse::<u64>().ok()?;

        if parsed % 10_000 == 0 {
            return Some(parsed / 10_000);
        }
    }

    None
}

fn normalize_deal_size(value: &str) -> Result<String, String> {
    let trimmed = value.trim();

    if trimmed.is_empty() {
        return Ok(String::new());
    }

    parse_deal_size_units(trimmed)
        .map(|units| units.to_string())
        .ok_or_else(|| "ディールサイズ は1万円単位の整数で入力してください。".to_string())
}

fn normalize_sales_fields(category: &str, rank: String, deal_size: String) -> Result<(String, String), String> {
    if category == "営業" {
        Ok((rank, normalize_deal_size(&deal_size)?))
    } else {
        Ok((String::new(), String::new()))
    }
}

fn normalize_content(content: &str) -> String {
    let trimmed = content.trim();

    if trimmed.is_empty() {
        "空です".to_string()
    } else {
        trimmed.to_string()
    }
}

fn write_workbook(path: &Path, items: &[ProgressItem]) -> Result<(), String> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let header_format = Format::new()
        .set_bold()
        .set_background_color(Color::RGB(0xDCEAF7));

    worksheet.set_name(SHEET_NAME).map_err(|error| error.to_string())?;
    worksheet
        .set_freeze_panes(1, 0)
        .map_err(|error| error.to_string())?;

    for (column, header) in HEADERS.iter().enumerate() {
        worksheet
            .write_with_format(0, column as u16, *header, &header_format)
            .map_err(|error| error.to_string())?;
    }

    for (row_index, item) in items.iter().enumerate() {
        let row = (row_index + 1) as u32;
        let values = [
            item.id.clone(),
            item.kpi_number.clone(),
            item.category.clone(),
            item.assignee.clone(),
            item.created_at.clone(),
            item.updated_at.clone(),
            item.status.clone(),
            item.rank.clone(),
            item.deal_size.clone(),
            item.external_stakeholders.clone(),
            item.internal_departments.clone(),
            item.customer.clone(),
            item.content.clone(),
            item.next_action.clone(),
            item.report_memo.clone(),
            item.updated_by.clone(),
            item.version.to_string(),
        ];

        for (column, value) in values.iter().enumerate() {
            worksheet
                .write_string(row, column as u16, value)
                .map_err(|error| error.to_string())?;
        }
    }

    workbook.save(path).map_err(|error| error.to_string())
}

pub fn ensure_workbook(path: &Path) -> Result<(), String> {
    if !path.exists() {
        write_workbook(path, &[])?;
    }

    let _ = read_items(path)?;
    Ok(())
}

pub fn read_items(path: &Path) -> Result<Vec<ProgressItem>, String> {
    if !path.exists() {
        write_workbook(path, &[])?;
    }

    let mut workbook = open_workbook_auto(path).map_err(|error| error.to_string())?;
    let sheet_name = if workbook.sheet_names().iter().any(|name| name == SHEET_NAME) {
        SHEET_NAME.to_string()
    } else {
        workbook
            .sheet_names()
            .first()
            .cloned()
            .ok_or_else(|| "Excel シートが見つかりません。".to_string())?
    };

    let range = workbook
        .worksheet_range(&sheet_name)
        .map_err(|error| error.to_string())?;

    let mut rows = range.rows();
    let header_row = rows
        .next()
        .ok_or_else(|| "Excel ヘッダーが見つかりません。".to_string())?;
    let actual_headers: Vec<String> = header_row.iter().map(ToString::to_string).collect();
    let header_map: HashMap<&str, usize> = actual_headers
        .iter()
        .enumerate()
        .map(|(index, header)| (header.as_str(), index))
        .collect();

    if !LEGACY_HEADERS.iter().all(|header| header_map.contains_key(header)) {
        return Err("Excel ファイルのヘッダーが想定と一致しません。README のテンプレート列を使用してください。".to_string());
    }

    let mut items = Vec::new();

    for row in rows {
        let id = row
            .get(*header_map.get("RowID").unwrap_or(&0))
            .map(ToString::to_string)
            .unwrap_or_default();

        if id.trim().is_empty() {
            continue;
        }

        let get = |header: &str| {
            header_map
                .get(header)
                .and_then(|index| row.get(*index))
                .map(ToString::to_string)
                .unwrap_or_default()
        };
        let version = get("Version").parse::<u32>().unwrap_or(1);

        items.push(ProgressItem {
            id,
            kpi_number: get("KPI番号"),
            category: get("カテゴリー"),
            assignee: get("担当者名"),
            created_at: get("登録日"),
            updated_at: get("更新日"),
            status: get("ステータス"),
            rank: get("ランク"),
            deal_size: get("ディールサイズ"),
            external_stakeholders: get("社外関係者"),
            internal_departments: get("社内関連部署"),
            customer: get("顧客名"),
            content: get("内容"),
            next_action: get("NextAction"),
            report_memo: get("報告メモ"),
            updated_by: get("更新者"),
            version,
        });
    }

    items.sort_by(|left, right| right.updated_at.cmp(&left.updated_at));
    Ok(items)
}

fn validate_payload(payload: &ProgressPayload, settings: &AppSettings) -> Result<(), String> {
    if payload.category.trim().is_empty() {
        return Err("カテゴリー は必須です。".to_string());
    }

    if !settings
        .category_options
        .iter()
        .any(|value| value == payload.category.trim())
    {
        return Err("カテゴリー は設定済みの候補から選択してください。".to_string());
    }

    if payload.assignee.trim().is_empty() {
        return Err("担当者名 は必須です。".to_string());
    }

    if !settings
        .assignee_options
        .iter()
        .any(|value| value == payload.assignee.trim())
    {
        return Err("担当者名 は指定された候補から選択してください。".to_string());
    }

    if payload.status.trim().is_empty() {
        return Err("ステータス は必須です。".to_string());
    }

    if !settings
        .status_options
        .iter()
        .any(|value| value == payload.status.trim())
    {
        return Err("ステータス は設定済みの候補から選択してください。".to_string());
    }

    if payload.category == "営業"
        && !payload.rank.trim().is_empty()
        && !settings
            .rank_options
            .iter()
            .any(|value| value == payload.rank.trim())
    {
        return Err("ランク は設定済みの候補から選択してください。".to_string());
    }

    if payload.category == "営業" {
        normalize_deal_size(&payload.deal_size)?;
    }

    Ok(())
}

pub fn create_item(path: &Path, settings: &AppSettings, payload: ProgressPayload) -> Result<ProgressItem, String> {
    validate_payload(&payload, settings)?;
    let mut items = read_items(path)?;
    let now = Utc::now().to_rfc3339();
    let (rank, deal_size) = normalize_sales_fields(&payload.category, payload.rank, payload.deal_size)?;
    let content = normalize_content(&payload.content);
    let updated_by = payload.assignee.trim().to_string();

    let item = ProgressItem {
        id: Uuid::new_v4().to_string(),
        kpi_number: payload.kpi_number,
        category: payload.category,
        assignee: payload.assignee,
        created_at: now.clone(),
        updated_at: now,
        status: payload.status,
        rank,
        deal_size,
        external_stakeholders: payload.external_stakeholders,
        internal_departments: payload.internal_departments,
        customer: payload.customer,
        content,
        next_action: payload.next_action,
        report_memo: payload.report_memo,
        updated_by,
        version: 1,
    };

    items.push(item.clone());
    write_workbook(path, &items)?;
    Ok(item)
}

pub fn update_item(path: &Path, settings: &AppSettings, id: &str, payload: ProgressPayload) -> Result<ProgressItem, String> {
    validate_payload(&payload, settings)?;
    let mut items = read_items(path)?;
    let target = items
        .iter_mut()
        .find(|item| item.id == id)
        .ok_or_else(|| "対象データが見つかりません。".to_string())?;

    let expected_version = payload.version.unwrap_or(0);

    if expected_version != target.version {
        return Err("他の更新が先に保存されています。最新データを再読み込みしてください。".to_string());
    }

    let (rank, deal_size) = normalize_sales_fields(&payload.category, payload.rank, payload.deal_size)?;
    let content = normalize_content(&payload.content);

    target.kpi_number = payload.kpi_number;
    target.category = payload.category;
    target.assignee = payload.assignee;
    target.status = payload.status;
    target.rank = rank;
    target.deal_size = deal_size;
    target.external_stakeholders = payload.external_stakeholders;
    target.internal_departments = payload.internal_departments;
    target.customer = payload.customer;
    target.content = content;
    target.next_action = payload.next_action;
    target.report_memo = payload.report_memo;
    target.updated_by = target.assignee.clone();
    target.updated_at = Utc::now().to_rfc3339();
    target.version += 1;

    let item = target.clone();
    write_workbook(path, &items)?;
    Ok(item)
}

pub fn delete_item(path: &Path, id: &str) -> Result<(), String> {
    let mut items = read_items(path)?;
    let original_len = items.len();
    items.retain(|item| item.id != id);

    if items.len() == original_len {
        return Err("対象データが見つかりません。".to_string());
    }

    write_workbook(path, &items)
}

pub fn build_summary(items: &[ProgressItem]) -> Summary {
    let mut by_status = BTreeMap::new();

    for item in items {
        *by_status.entry(item.status.clone()).or_insert(0) += 1;
    }

    Summary {
        total: items.len(),
        by_status,
    }
}
