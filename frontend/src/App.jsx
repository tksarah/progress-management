import { useEffect, useMemo, useRef, useState } from "react"
import { invoke } from "@tauri-apps/api/core"
import { open, save } from "@tauri-apps/plugin-dialog"

const pageSize = 10

const emptyForm = {
  id: "",
  title: "",
  updatedAtInput: "",
  kpiNumber: "",
  category: "",
  assignee: "",
  status: "",
  rank: "",
  dealSize: "",
  leadSource: "",
  externalStakeholders: "",
  internalDepartments: "",
  customer: "",
  content: "",
  nextAction: "",
  reportMemo: "",
  updatedBy: "",
  version: 0
}

const defaultCategoryOptions = ["営業", "マーケティング"]
const defaultAssigneeOptions = ["山﨑", "倉持", "西田"]
const defaultStatusOptions = ["進捗中", "計画中", "クローズ", "保留"]
const defaultRankOptions = ["A", "B", "C", "D", "X1", "X2", "1"]
const kpiOptions = ["①", "②", "③", "④", "⑤"]
const reportHistoryStorageKey = "progress-tracker-report-history"
const reportPresetOptions = [
  { value: "7", label: "直近1週間", days: 7 },
  { value: "14", label: "直近2週間", days: 14 },
  { value: "custom", label: "範囲指定" }
]

const rankChartPalette = [
  "#0d5ea6",
  "#2c7be5",
  "#15aabf",
  "#2f9e77",
  "#f08c00",
  "#d9485f",
  "#7950f2",
  "#5f6b7a",
  "#c2255c"
]

const defaultLeadSourceOptions = ["TDW", "主催・共催イベント", "オフラインイベント", "アウトバウンド", "社内", "個別ネットワーキング", "ウェビナー"]

const tableColumnDefinitions = [
  { key: "title", label: "タイトル", render: (item) => truncateText(item.title) },
  { key: "status", label: "ステータス", render: (item) => item.status || "-" },
  { key: "kpiNumber", label: "KPI", render: (item) => item.kpiNumber || "-" },
  { key: "category", label: "カテゴリ", render: (item) => item.category || "-" },
  { key: "assignee", label: "担当者", render: (item) => item.assignee || "-" },
  { key: "updatedAt", label: "更新日", render: (item) => formatDate(item.updatedAt) },
  { key: "customer", label: "顧客名 / Project 名", render: (item) => truncateText(item.customer || "-") },
  { key: "rank", label: "ランク", render: (item) => item.rank || "-" },
  { key: "dealSize", label: "ディールサイズ", render: (item) => truncateText(formatDealSizeDisplay(item.dealSize)) },
  { key: "leadSource", label: "リード元", render: (item) => item.leadSource || "-" },
  { key: "content", label: "内容", render: (item) => truncateText(item.content) },
  { key: "nextAction", label: "Next Action", render: (item) => truncateText(item.nextAction) },
  { key: "reportMemo", label: "報告メモ", render: (item) => truncateText(item.reportMemo) }
]

const defaultVisibleColumns = ["status", "kpiNumber", "category", "assignee", "updatedAt", "content"]

const defaultAppSettings = {
  excelFilePath: "",
  categoryOptions: defaultCategoryOptions,
  assigneeOptions: defaultAssigneeOptions,
  statusOptions: defaultStatusOptions,
  rankOptions: defaultRankOptions,
  visibleColumns: defaultVisibleColumns,
  leadSourceOptions: defaultLeadSourceOptions
}

function reorderList(items, fromIndex, toIndex) {
  if (fromIndex === toIndex || fromIndex < 0 || toIndex < 0 || fromIndex >= items.length || toIndex >= items.length) {
    return items
  }

  const nextItems = [...items]
  const [movedItem] = nextItems.splice(fromIndex, 1)
  nextItems.splice(toIndex, 0, movedItem)
  return nextItems
}

function insertListItem(items, fromIndex, toIndex) {
  if (fromIndex < 0 || fromIndex >= items.length || toIndex < 0 || toIndex > items.length) {
    return items
  }

  const nextItems = [...items]
  const [movedItem] = nextItems.splice(fromIndex, 1)
  const adjustedIndex = fromIndex < toIndex ? toIndex - 1 : toIndex

  if (adjustedIndex === fromIndex) {
    return items
  }

  nextItems.splice(adjustedIndex, 0, movedItem)
  return nextItems
}

function sanitizeOptionList(values) {
  return Array.from(new Set(values.map((value) => value.trim()).filter(Boolean)))
}

function normalizeVisibleColumns(columns) {
  const allowedKeys = new Set(tableColumnDefinitions.map((column) => column.key))
  const normalized = Array.from(new Set((columns || []).filter((key) => allowedKeys.has(key))))

  return normalized.length > 0 ? normalized : defaultVisibleColumns
}

function normalizeAppSettings(settings) {
  return {
    excelFilePath: String(settings?.excelFilePath || "").trim(),
    categoryOptions: sanitizeOptionList(settings?.categoryOptions || defaultCategoryOptions),
    assigneeOptions: sanitizeOptionList(settings?.assigneeOptions || defaultAssigneeOptions),
    statusOptions: sanitizeOptionList(settings?.statusOptions || defaultStatusOptions),
    rankOptions: sanitizeOptionList(settings?.rankOptions || defaultRankOptions),
    leadSourceOptions: sanitizeOptionList((settings?.leadSourceOptions && settings.leadSourceOptions.length) ? settings.leadSourceOptions : defaultLeadSourceOptions),
    visibleColumns: normalizeVisibleColumns(settings?.visibleColumns || defaultVisibleColumns)
  }
}

function withCurrentOption(options, currentValue) {
  if (!currentValue || options.includes(currentValue)) {
    return options
  }

  return [...options, currentValue]
}

function truncateText(value, maxLength = 18) {
  const text = String(value || "").trim()

  if (!text) {
    return "-"
  }

  return text.length > maxLength ? `${text.slice(0, maxLength)}...` : text
}

function parseDealSizeUnits(value) {
  const text = String(value || "").trim()

  if (!text) {
    return null
  }

  const normalized = text.replace(/,/g, "").replace(/\s+/g, "")

  if (/^\d+$/.test(normalized)) {
    return Number(normalized)
  }

  if (/^\d+万(円)?$/.test(normalized)) {
    return Number(normalized.replace(/万円?|円/g, ""))
  }

  if (/^\d+円$/.test(normalized)) {
    const yenValue = Number(normalized.replace(/円/g, ""))
    return yenValue % 10000 === 0 ? yenValue / 10000 : null
  }

  return null
}

function normalizeDealSizeInputValue(value) {
  const units = parseDealSizeUnits(value)
  return units === null ? "" : String(units)
}

function formatDealSizeDisplay(value) {
  const units = parseDealSizeUnits(value)

  if (units === null) {
    const text = String(value || "").trim()
    return text || "-"
  }

  return `${units.toLocaleString("ja-JP")}万円`
}

function formatKpiDisplay(value) {
  return String(value || "").trim() || "-"
}

function formatDealSizeUnitsLabel(value) {
  return `${value.toLocaleString("ja-JP")}万円`
}

function describeRankShare(value, total) {
  if (!total) {
    return "0.0"
  }

  return ((value / total) * 100).toFixed(1)
}

function buildRankDonutGradient(rows, totalUnits) {
  if (!totalUnits || rows.length === 0) {
    return "conic-gradient(from -105deg, transparent 0deg 360deg)"
  }

  let currentAngle = 0
  const segments = []

  rows.forEach((row) => {
    const sweepAngle = (row.totalUnits / totalUnits) * 360
    const gapAngle = Math.min(3.2, sweepAngle / 5)
    const startAngle = currentAngle
    const endAngle = currentAngle + sweepAngle
    const visibleStart = sweepAngle > gapAngle ? startAngle + gapAngle / 2 : startAngle
    const visibleEnd = sweepAngle > gapAngle ? endAngle - gapAngle / 2 : endAngle

    if (visibleStart > startAngle) {
      segments.push(`transparent ${startAngle.toFixed(2)}deg ${visibleStart.toFixed(2)}deg`)
    }

    segments.push(`${row.color} ${visibleStart.toFixed(2)}deg ${visibleEnd.toFixed(2)}deg`)

    if (visibleEnd < endAngle) {
      segments.push(`transparent ${visibleEnd.toFixed(2)}deg ${endAngle.toFixed(2)}deg`)
    }

    currentAngle = endAngle
  })

  if (currentAngle < 360) {
    segments.push(`transparent ${currentAngle.toFixed(2)}deg 360deg`)
  }

  return `conic-gradient(from -105deg, ${segments.join(", ")})`
}

// 履歴保存は削除されました

function toReportSnapshot(items) {
  return items.map((item) => ({
    id: item.id,
    title: item.title,
    kpiNumber: item.kpiNumber,
    assignee: item.assignee,
    status: item.status,
    updatedAt: item.updatedAt,
    content: item.content,
    nextAction: item.nextAction,
    reportMemo: item.reportMemo
  }))
}

function hasItemChanged(previousItem, currentItem) {
  if (!previousItem) {
    return true
  }

  return previousItem.updatedAt !== currentItem.updatedAt
    || previousItem.title !== currentItem.title
    || previousItem.status !== currentItem.status
    || previousItem.content !== currentItem.content
    || previousItem.nextAction !== currentItem.nextAction
    || previousItem.reportMemo !== currentItem.reportMemo
}

// 差分生成は削除されました

async function call(command, payload) {
  try {
    return await invoke(command, payload)
  } catch (error) {
    throw new Error(String(error))
  }
}

function normalizeExcelPath(path) {
  if (!path) {
    return path
  }

  return path.toLowerCase().endsWith(".xlsx") ? path : `${path}.xlsx`
}

function formatDate(value) {
  if (!value) {
    return "-"
  }

  return new Date(value).toLocaleString("ja-JP")
}

function toValidDate(value) {
  if (!value) {
    return null
  }

  const parsed = new Date(value)
  return Number.isNaN(parsed.getTime()) ? null : parsed
}

function formatDateOnly(value) {
  const parsed = toValidDate(value)

  if (!parsed) {
    return "-"
  }

  return parsed.toLocaleDateString("ja-JP")
}

function shiftDate(baseDate, days) {
  const next = new Date(baseDate)
  next.setDate(next.getDate() + days)
  return next
}

function formatDateInputValue(value) {
  const parsed = toValidDate(value)

  if (!parsed) {
    return ""
  }

  const year = parsed.getFullYear()
  const month = String(parsed.getMonth() + 1).padStart(2, "0")
  const day = String(parsed.getDate()).padStart(2, "0")

  return `${year}-${month}-${day}`
}

function parseDateInputValue(value) {
  if (!value) {
    return null
  }

  const [year, month, day] = value.split("-").map(Number)

  if (!year || !month || !day) {
    return null
  }

  const parsed = new Date(year, month - 1, day)
  return Number.isNaN(parsed.getTime()) ? null : parsed
}

function formatPeriodLabel(start, end) {
  return `${start.toLocaleDateString("ja-JP")} - ${end.toLocaleDateString("ja-JP")}`
}

function resolveReportRange({ preset, startDateInput, endDateInput }) {
  if (preset === "custom") {
    const start = parseDateInputValue(startDateInput)
    const end = parseDateInputValue(endDateInput)

    if (!start || !end) {
      return {
        isValid: false,
        errorMessage: "開始日と終了日を入力してください。"
      }
    }

    start.setHours(0, 0, 0, 0)
    end.setHours(23, 59, 59, 999)

    if (start > end) {
      return {
        isValid: false,
        errorMessage: "開始日は終了日以前にしてください。"
      }
    }

    return {
      isValid: true,
      start,
      end,
      label: formatPeriodLabel(start, end)
    }
  }

  const presetOption = reportPresetOptions.find((item) => item.value === preset && item.days) || reportPresetOptions[1]
  const end = new Date()
  const start = shiftDate(end, -(presetOption.days - 1))

  start.setHours(0, 0, 0, 0)

  return {
    isValid: true,
    start,
    end,
    label: formatPeriodLabel(start, end)
  }
}

function formatReportLine(item) {
  const memo = item.reportMemo?.trim()
  const content = item.content?.trim()
  const nextAction = item.nextAction?.trim()

  if (memo) {
    return `・${formatKpiDisplay(item.kpiNumber)} ${item.assignee}: ${memo}`
  }

  if (content) {
    return `・${formatKpiDisplay(item.kpiNumber)} ${item.assignee}: ${content}`
  }

  if (nextAction) {
    return `・${formatKpiDisplay(item.kpiNumber)} ${item.assignee}: 次回対応 ${nextAction}`
  }

  return `・${formatKpiDisplay(item.kpiNumber)} ${item.assignee}: 更新あり`
}

function RankDealDonutChart({ rows, totalUnits, totalCount }) {
  const donutGradient = buildRankDonutGradient(rows, totalUnits)

  return (
    <div className="rank-chart-panel">
      <div className="rank-chart-head">
        <span className="rank-chart-label">ランク別</span>
        <p className="rank-chart-copy">ディールサイズの構成比</p>
      </div>

      <div className="rank-donut-stage">
        <div className="rank-donut-shell" role="img" aria-label="ランク別ディールサイズ構成">
          <div className="rank-donut-aura" aria-hidden="true" />
          <div className="rank-donut-track" aria-hidden="true" />
          <div className="rank-donut" style={{ "--rank-donut-gradient": donutGradient }} aria-hidden="true" />
          <div className="rank-donut-center">
            <span>総ディールサイズ</span>
            <strong>{formatDealSizeUnitsLabel(totalUnits)}</strong>
            <small>{totalCount.toLocaleString("ja-JP")}件 / {rows.length}ランク</small>
          </div>
        </div>
      </div>

      <div className="rank-legend" role="list" aria-label="ランク別凡例">
        {rows.map((row) => (
          <article className="rank-legend-item" key={row.rank} role="listitem">
            <span className="rank-legend-swatch" style={{ "--legend-color": row.color }} aria-hidden="true" />
            <div>
              <strong>{row.rank}</strong>
              <span>{formatDealSizeUnitsLabel(row.totalUnits)}</span>
              <small>{describeRankShare(row.totalUnits, totalUnits)}%</small>
            </div>
          </article>
        ))}
      </div>
    </div>
  )
}

function normalizeProgressPayload(payload) {
  const content = payload.content.trim()

  return {
    ...payload,
    title: String(payload.title || "").trim(),
    kpiNumber: payload.kpiNumber.trim(),
    category: payload.category.trim(),
    assignee: payload.assignee.trim(),
    status: payload.status.trim(),
    rank: payload.rank.trim(),
    dealSize: normalizeDealSizeInputValue(payload.dealSize),
    leadSource: String(payload.leadSource || "").trim(),
    externalStakeholders: payload.externalStakeholders.trim(),
    internalDepartments: payload.internalDepartments.trim(),
    customer: payload.customer.trim(),
    content: content || "空です",
    nextAction: payload.nextAction.trim(),
    reportMemo: payload.reportMemo.trim(),
    updatedAt: payload.updatedAtInput ? payload.updatedAtInput.trim() : null,
    updatedBy: payload.assignee.trim()
  }
}

function toFormState(item) {
  return {
    ...item,
    kpiNumber: String(item.kpiNumber || "").trim(),
    dealSize: normalizeDealSizeInputValue(item.dealSize),
    updatedAtInput: formatDateInputValue(item.updatedAt),
    updatedBy: item.assignee || item.updatedBy
  }
}

function toDuplicateFormState(item) {
  return {
    ...toFormState(item),
    id: "",
    createdAt: "",
    updatedAt: "",
    updatedAtInput: "",
    version: 0
  }
}

function buildReportDraft({
  items,
  range,
  statusFilter,
  kpiFilter,
  categoryFilter,
  query
}) {
  const { start, end } = range

  const updatedItems = items.filter((item) => {
    const updatedAt = toValidDate(item.updatedAt)
    return updatedAt && updatedAt >= start && updatedAt <= end
  })
  const completedItems = updatedItems.filter((item) => item.status === "クローズ")
  const onHoldItems = updatedItems.filter((item) => item.status === "保留")
  const noNextActionItems = updatedItems.filter((item) => !item.nextAction?.trim())
  const noReportMemoItems = updatedItems.filter((item) => !item.reportMemo?.trim())
  const staleItems = items.filter((item) => {
    const updatedAt = toValidDate(item.updatedAt)
    return !updatedAt || updatedAt < start
  })

  const filterLabels = [
    statusFilter ? `ステータス: ${statusFilter}` : null,
    kpiFilter ? `KPI: ${kpiFilter}` : null,
    categoryFilter ? `カテゴリ: ${categoryFilter}` : null,
    query ? `検索語: ${query}` : null
  ].filter(Boolean)

  const progressLines = updatedItems.slice(0, 5).map(formatReportLine)
  const riskLines = [
    ...onHoldItems.slice(0, 3).map((item) => `・${formatKpiDisplay(item.kpiNumber)} ${item.assignee}: 保留中。${item.reportMemo?.trim() || item.content?.trim() || "状況確認が必要"}`),
    ...staleItems.slice(0, 3).map((item) => `・${formatKpiDisplay(item.kpiNumber)} ${item.assignee}: ${formatDateOnly(item.updatedAt)} 以降更新なし。`)
  ].slice(0, 5)
  const actionLines = updatedItems
    .filter((item) => item.nextAction?.trim())
    .slice(0, 5)
    .map((item) => `・${formatKpiDisplay(item.kpiNumber)} ${item.assignee}: ${item.nextAction.trim()}`)
  const meetingPoints = [
    ...(onHoldItems.slice(0, 2).map((item) => `・確認事項: ${formatKpiDisplay(item.kpiNumber)} ${item.assignee} / ${item.reportMemo?.trim() || item.content?.trim() || "判断が必要"}`))
  ].slice(0, 5)

  const text = [
    // 新しい要求仕様: 各更新案件をMarkdown形式で列挙し、項目ごとに --- で区切る
    // ヘッダ
    `定例会議サマリ（${formatPeriodLabel(start, end)}）`,
    "",
    "報告対象の範囲に更新された内容をすべて報告。",
    "報告する項目は以下のとおりです。",
    "",
    // 各案件をMarkdownで出力
    ...(updatedItems.length > 0 ? updatedItems.map((item) => {
      const lines = []
      lines.push(`- タイトル: ${item.title?.trim() || "-"}`)
      lines.push(`- 担当者: ${item.assignee || "-"}`)
      lines.push(`- カテゴリー: ${item.category || "-"}`)

      if (item.category === "営業") {
        const customer = item.customer?.trim() || "-"
        const rank = item.rank?.trim() || "-"
        const dealSize = formatDealSizeDisplay(item.dealSize)
        lines.push(`- 顧客名/ランク/ディールサイズ: ${customer} / ${rank} / ${dealSize}`)
      }

      lines.push(`- 内容: ${item.content?.trim() || "-"}`)
      lines.push(`- NextAction: ${item.nextAction?.trim() || "-"}`)
      lines.push(`- 報告メモ: ${item.reportMemo?.trim() || "-"}`)

      // 各報告を '---' で区切る
      return lines.join("\n") + "\n\n---"
    }) : ["・期間内に更新された案件はありません。"])
  ].join("\n")

  return {
    text,
    metrics: {
      total: items.length,
      updated: updatedItems.length,
      completed: completedItems.length,
      onHold: onHoldItems.length,
      stale: staleItems.length,
      noNextAction: noNextActionItems.length,
      noReportMemo: noReportMemoItems.length,
      changed: 0,
      newItems: 0,
      newlyCompleted: 0
    },
    periodLabel: formatPeriodLabel(start, end),
      diffLines: [],
    snapshot: toReportSnapshot(items),
    isValid: true,
    errorMessage: ""
  }
}

function buildInvalidReportDraft(items, errorMessage) {
  return {
    text: `定例会議サマリ\n\n${errorMessage}`,
    metrics: {
      total: items.length,
      updated: 0,
      completed: 0,
      onHold: 0,
      stale: 0,
      noNextAction: 0,
      noReportMemo: 0,
      changed: 0,
      newItems: 0,
      newlyCompleted: 0
    },
    periodLabel: "期間未設定",
    diffLines: [errorMessage],
    snapshot: toReportSnapshot(items),
    isValid: false,
    errorMessage
  }
}

function SortableTagList({ items, onRemove, onMove, onReorder, showRemove = true }) {
  const listRef = useRef(null)
  const [dragState, setDragState] = useState({
    sourceIndex: -1,
    targetIndex: -1,
    position: "before",
    active: false,
    pointerX: 0,
    pointerY: 0
  })

  function resetDragState() {
    setDragState({
      sourceIndex: -1,
      targetIndex: -1,
      position: "before",
      active: false,
      pointerX: 0,
      pointerY: 0
    })
  }

  function handlePointerDown(event, index) {
    event.preventDefault()
    setDragState({
      sourceIndex: index,
      targetIndex: index,
      position: "before",
      active: true,
      pointerX: event.clientX,
      pointerY: event.clientY
    })
  }

  useEffect(() => {
    if (!dragState.active) {
      return undefined
    }

    function handlePointerMove(event) {
      const listElement = listRef.current

      if (!listElement) {
        return
      }

      const targetElement = document.elementFromPoint(event.clientX, event.clientY)?.closest("[data-sort-index]")

      if (!(targetElement instanceof HTMLElement) || !listElement.contains(targetElement)) {
        return
      }

      const nextIndex = Number(targetElement.dataset.sortIndex)

      if (Number.isNaN(nextIndex)) {
        return
      }

      const bounds = targetElement.getBoundingClientRect()
      const nextPosition = event.clientY >= bounds.top + bounds.height / 2 ? "after" : "before"

      setDragState((current) => {
        if (!current.active) {
          return current
        }

        if (current.targetIndex === nextIndex && current.position === nextPosition) {
          return current
        }

        return {
          ...current,
          targetIndex: nextIndex,
          position: nextPosition,
          pointerX: event.clientX,
          pointerY: event.clientY
        }
      })
    }

    function handlePointerUp() {
      if (onReorder && dragState.sourceIndex >= 0 && dragState.targetIndex >= 0) {
        const targetIndex = dragState.targetIndex + (dragState.position === "after" ? 1 : 0)
        onReorder(dragState.sourceIndex, targetIndex)
      }

      resetDragState()
    }

    window.addEventListener("pointermove", handlePointerMove)
    window.addEventListener("pointerup", handlePointerUp)

    return () => {
      window.removeEventListener("pointermove", handlePointerMove)
      window.removeEventListener("pointerup", handlePointerUp)
    }
  }, [dragState.active, dragState.position, dragState.sourceIndex, dragState.targetIndex, onReorder])

  function getDropClass(index) {
    if (dragState.sourceIndex < 0 || dragState.targetIndex !== index || dragState.sourceIndex === index) {
      return ""
    }

    return dragState.position === "after" ? "drop-after" : "drop-before"
  }

  return (
    <>
      <div className={`settings-sortable-list ${dragState.active ? "drag-active" : ""}`.trim()} ref={listRef}>
        {items.map((item, index) => (
          <div
            className={`settings-sortable-item ${dragState.sourceIndex === index ? "dragging" : ""} ${getDropClass(index)}`.trim()}
            key={item}
            data-sort-index={index}
          >
            <div className="settings-item-main">
              <button
                type="button"
                className="settings-drag-handle"
                aria-label={`${item} をドラッグして並び替え`}
                onPointerDown={(event) => handlePointerDown(event, index)}
              >
                ::
              </button>
              <strong>{item}</strong>
            </div>
            <div className="settings-item-actions">
              <button type="button" className="secondary-button table-action" onClick={() => onMove(index, -1)} disabled={index === 0}>上へ</button>
              <button type="button" className="secondary-button table-action" onClick={() => onMove(index, 1)} disabled={index === items.length - 1}>下へ</button>
              {showRemove ? (
                <button type="button" className="secondary-button table-action danger-action" onClick={() => onRemove(item)}>削除</button>
              ) : null}
            </div>
          </div>
        ))}
      </div>
      {dragState.active && dragState.sourceIndex >= 0 ? (
        <div
          className="settings-drag-preview"
          aria-hidden="true"
          style={{
            left: `${dragState.pointerX + 18}px`,
            top: `${dragState.pointerY + 18}px`
          }}
        >
          <span className="settings-drag-preview-handle">::</span>
          <strong>{items[dragState.sourceIndex]}</strong>
          <span className="settings-drag-preview-badge">移動中</span>
        </div>
      ) : null}
    </>
  )
}

function OptionEditorSection({
  title,
  items,
  inputValue,
  onInputChange,
  onAdd,
  onRemove,
  onMove,
  onReorder,
  placeholder
}) {
  return (
    <section className="settings-editor-section">
      <div className="settings-editor-head">
        <h3>{title}</h3>
        <span>{items.length}件</span>
      </div>
      <div className="settings-editor-inputs">
        <input
          value={inputValue}
          onChange={(event) => onInputChange(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault()
              onAdd()
            }
          }}
          placeholder={placeholder}
        />
        <button type="button" className="secondary-button" onClick={onAdd}>追加</button>
      </div>
      <p className="hint settings-sort-hint">ドラッグアンドドロップ、または上下ボタンで並び順を変更できます。</p>
      <div className="settings-tag-list">
        {items.length === 0 ? <p className="hint">まだ登録がありません。</p> : null}
        {items.length > 0 ? (
          <SortableTagList
            items={items}
            onRemove={onRemove}
            onMove={onMove}
            onReorder={onReorder}
          />
        ) : null}
      </div>
    </section>
  )
}

function SettingsModal({
  settingsDraft,
  optionDrafts,
  onOptionDraftChange,
  onAddOption,
  onRemoveOption,
  onMoveOption,
  onReorderOption,
  onToggleColumn,
  onMoveColumn,
  onReorderColumn,
  onClose,
  error,
  saving
}) {
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <section
        className="settings-modal"
        onClick={(event) => event.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-label="表示設定と選択肢の編集"
      >
        <div className="drawer-head">
          <div>
            <p className="eyebrow">DISPLAY SETTINGS</p>
            <h2>設定</h2>
            <p className="hint section-copy">カテゴリー、担当者名、ステータス、ランク、一覧列の表示内容と順序を編集します。</p>
          </div>
          <button type="button" className="secondary-button" onClick={onClose} disabled={saving}>閉じる</button>
        </div>

        <div className="settings-modal-grid">
          <OptionEditorSection
            title="カテゴリー"
            items={settingsDraft.categoryOptions}
            inputValue={optionDrafts.categoryOptions}
            onInputChange={(value) => onOptionDraftChange("categoryOptions", value)}
            onAdd={() => onAddOption("categoryOptions")}
            onRemove={(value) => onRemoveOption("categoryOptions", value)}
            onMove={(index, direction) => onMoveOption("categoryOptions", index, direction)}
            onReorder={(fromIndex, toIndex) => onReorderOption("categoryOptions", fromIndex, toIndex)}
            placeholder="カテゴリー名を入力"
          />
          <OptionEditorSection
            title="担当者名"
            items={settingsDraft.assigneeOptions}
            inputValue={optionDrafts.assigneeOptions}
            onInputChange={(value) => onOptionDraftChange("assigneeOptions", value)}
            onAdd={() => onAddOption("assigneeOptions")}
            onRemove={(value) => onRemoveOption("assigneeOptions", value)}
            onMove={(index, direction) => onMoveOption("assigneeOptions", index, direction)}
            onReorder={(fromIndex, toIndex) => onReorderOption("assigneeOptions", fromIndex, toIndex)}
            placeholder="担当者名を入力"
          />
          <OptionEditorSection
            title="ステータス"
            items={settingsDraft.statusOptions}
            inputValue={optionDrafts.statusOptions}
            onInputChange={(value) => onOptionDraftChange("statusOptions", value)}
            onAdd={() => onAddOption("statusOptions")}
            onRemove={(value) => onRemoveOption("statusOptions", value)}
            onMove={(index, direction) => onMoveOption("statusOptions", index, direction)}
            onReorder={(fromIndex, toIndex) => onReorderOption("statusOptions", fromIndex, toIndex)}
            placeholder="ステータス名を入力"
          />
          <OptionEditorSection
            title="ランク"
            items={settingsDraft.rankOptions}
            inputValue={optionDrafts.rankOptions}
            onInputChange={(value) => onOptionDraftChange("rankOptions", value)}
            onAdd={() => onAddOption("rankOptions")}
            onRemove={(value) => onRemoveOption("rankOptions", value)}
            onMove={(index, direction) => onMoveOption("rankOptions", index, direction)}
            onReorder={(fromIndex, toIndex) => onReorderOption("rankOptions", fromIndex, toIndex)}
            placeholder="ランク名を入力"
          />

          <OptionEditorSection
            title="リード元"
            items={settingsDraft.leadSourceOptions}
            inputValue={optionDrafts.leadSourceOptions}
            onInputChange={(value) => onOptionDraftChange("leadSourceOptions", value)}
            onAdd={() => onAddOption("leadSourceOptions")}
            onRemove={(value) => onRemoveOption("leadSourceOptions", value)}
            onMove={(index, direction) => onMoveOption("leadSourceOptions", index, direction)}
            onReorder={(fromIndex, toIndex) => onReorderOption("leadSourceOptions", fromIndex, toIndex)}
            placeholder="リード元を入力"
          />

          <section className="settings-editor-section settings-column-section">
            <div className="settings-editor-head">
              <h3>進捗一覧の列</h3>
              <span>{settingsDraft.visibleColumns.length}列を表示</span>
            </div>
            <p className="hint">No. と 操作 は固定表示です。その他の表示列を切り替えられます。</p>
            <div className="settings-column-grid">
              {tableColumnDefinitions.map((column) => (
                <label className="settings-checkbox" key={column.key}>
                  <input
                    type="checkbox"
                    checked={settingsDraft.visibleColumns.includes(column.key)}
                    onChange={() => onToggleColumn(column.key)}
                  />
                  <span>{column.label}</span>
                </label>
              ))}
            </div>
            <p className="hint settings-sort-hint">表示中の列はドラッグアンドドロップ、または上下ボタンで順序変更できます。</p>
            <SortableTagList
              items={settingsDraft.visibleColumns.map((key) => tableColumnDefinitions.find((column) => column.key === key)?.label || key)}
              onRemove={(label) => {
                const column = tableColumnDefinitions.find((item) => item.label === label)
                if (column) {
                  onToggleColumn(column.key)
                }
              }}
              onMove={onMoveColumn}
              onReorder={onReorderColumn}
              showRemove={false}
            />
            <p className="hint">最低1列は選択してください。</p>
          </section>
        </div>

        <p className="hint settings-current-path">現在の保存先: <span className="path-label">{settingsDraft.excelFilePath || "未設定"}</span></p>

        <div className="confirm-modal-actions">
          <button type="button" className="secondary-button" onClick={onClose}>閉じる</button>
        </div>
        {error ? <p className="message error compact-message">{error}</p> : null}
      </section>
    </div>
  )
}

function StartupWizard({ startupState, saving, error, onOpenExisting, onCreateNew }) {
  return (
    <section className="startup-shell">
      <div className="startup-card">
        <div className="startup-head">
          <div>
            <p className="eyebrow">STARTUP WIZARD</p>
            <h2>起動方法を選択してください</h2>
            <p className="hero-copy startup-copy">
              初回起動時、または前回指定した Excel が見つからない場合は、利用する Excel ファイルをここで選択します。
            </p>
          </div>
          <span className="pill alt">Excel 未選択</span>
        </div>

        {startupState.hasMissingConfiguredExcel ? (
          <div className="message error compact-message">
            前回指定された Excel が見つかりません: <span className="path-label">{startupState.configuredExcelPath}</span>
          </div>
        ) : null}

        <div className="startup-choice-grid">
          <article className="startup-choice-card">
            <p className="eyebrow">CREATE NEW</p>
            <h3>新規で Excel 名を指定して起動</h3>
            <p className="hint startup-choice-copy">
              保存先とファイル名を指定して、新しい Excel を作成してからアプリを開きます。
            </p>
            <p className="hint startup-path-hint">推奨パス: <span className="path-label">{startupState.suggestedNewExcelPath}</span></p>
            <button type="button" onClick={onCreateNew} disabled={saving}>新規 Excel を指定</button>
          </article>

          <article className="startup-choice-card">
            <p className="eyebrow">OPEN EXISTING</p>
            <h3>既存の Excel で起動</h3>
            <p className="hint startup-choice-copy">
              既に使っている Excel ファイルを選択して、その内容を読み込んで開きます。
            </p>
            <p className="hint startup-path-hint">共有フォルダ上の .xlsx も選択できます。</p>
            <button type="button" className="secondary-button" onClick={onOpenExisting} disabled={saving}>既存 Excel を選択</button>
          </article>
        </div>

        {error ? <p className="message error compact-message">{error}</p> : null}
      </div>
    </section>
  )
}

function ProgressForm({
  form,
  onChange,
  onSubmit,
  onCancel,
  saving,
  submitLabel,
  cancelLabel,
  categoryOptions,
  assigneeOptions,
  statusOptions,
  rankOptions
  ,leadSourceOptions
}) {
  const categorySelectOptions = withCurrentOption(categoryOptions, form.category)
  const assigneeSelectOptions = withCurrentOption(assigneeOptions, form.assignee)
  const statusSelectOptions = withCurrentOption(statusOptions, form.status)
  const rankSelectOptions = withCurrentOption(rankOptions, form.rank)
  const isEditMode = Boolean(form.id)

  const leadSourceSelectOptions = withCurrentOption(
    (leadSourceOptions && leadSourceOptions.length) ? leadSourceOptions : defaultLeadSourceOptions,
    form.leadSource
  )

  const salesContentTemplate = "* 概要：\n* Budget / Authority / Need / Timeline：\n* コンペリングイベント：\n"

  return (
    <form className="detail-form" onSubmit={onSubmit}>
      <label>
        <span>タイトル</span>
        <input value={form.title} onChange={(event) => onChange({ ...form, title: event.target.value })} />
      </label>
      <label>
        <span>対象KPI番号</span>
        <select value={form.kpiNumber} onChange={(event) => onChange({ ...form, kpiNumber: event.target.value })}>
          <option value="">選択してください</option>
          {kpiOptions.map((kpi) => (
            <option key={kpi} value={kpi}>{kpi}</option>
          ))}
        </select>
      </label>
      <label>
        <span className="field-label">カテゴリー <span className="required-mark" aria-label="必須">必須</span></span>
        <select
          value={form.category}
          onChange={(event) => {
            const nextCategory = event.target.value
            const nextContent =
              nextCategory === "営業" && !String(form.content ?? "").trim()
                ? salesContentTemplate
                : form.content
            onChange({
              ...form,
              category: nextCategory,
              rank: nextCategory === "営業" ? form.rank : "",
              dealSize: nextCategory === "営業" ? form.dealSize : "",
              leadSource: nextCategory === "営業" ? form.leadSource : "",
              content: nextContent
            })
          }}
          required
        >
          <option value="">選択してください</option>
          {categorySelectOptions.map((category) => (
            <option key={category} value={category}>{category}</option>
          ))}
        </select>
      </label>
      <label>
        <span className="field-label">担当者名 <span className="required-mark" aria-label="必須">必須</span></span>
        <select value={form.assignee} onChange={(event) => onChange({ ...form, assignee: event.target.value })} required>
          <option value="">選択してください</option>
          {assigneeSelectOptions.map((assignee) => (
            <option key={assignee} value={assignee}>{assignee}</option>
          ))}
        </select>
      </label>
      <label>
        <span className="field-label">ステータス <span className="required-mark" aria-label="必須">必須</span></span>
        <select value={form.status} onChange={(event) => onChange({ ...form, status: event.target.value })} required>
          <option value="">選択してください</option>
          {statusSelectOptions.map((status) => (
            <option key={status} value={status}>{status}</option>
          ))}
        </select>
      </label>
      {form.category === "営業" ? (
        <>
          <label>
            <span>ランク</span>
            <select value={form.rank} onChange={(event) => onChange({ ...form, rank: event.target.value })}>
              <option value="">選択してください</option>
              {rankSelectOptions.map((rank) => (
                <option key={rank} value={rank}>{rank}</option>
              ))}
            </select>
          </label>
          <label>
            <span>リード元</span>
            <select value={form.leadSource} onChange={(event) => onChange({ ...form, leadSource: event.target.value })}>
              <option value="">選択してください</option>
              {leadSourceSelectOptions.map((source) => (
                <option key={source} value={source}>{source}</option>
              ))}
            </select>
          </label>
          <label>
            <span>ディールサイズ（万円）</span>
            <input
              type="number"
              min="0"
              step="1"
              inputMode="numeric"
              value={form.dealSize}
              onChange={(event) => onChange({ ...form, dealSize: event.target.value })}
              placeholder="例: 300"
            />
          </label>
        </>
      ) : null}
      <label>
        <span>社外関係者</span>
        <input value={form.externalStakeholders} onChange={(event) => onChange({ ...form, externalStakeholders: event.target.value })} />
      </label>
      <label>
        <span>社内関連部署</span>
        <input value={form.internalDepartments} onChange={(event) => onChange({ ...form, internalDepartments: event.target.value })} />
      </label>
      <label>
        <span>顧客名 / Project 名</span>
        <input value={form.customer} onChange={(event) => onChange({ ...form, customer: event.target.value })} />
      </label>
      <label className="full-width">
        <span>内容</span>
        <textarea value={form.content} onChange={(event) => onChange({ ...form, content: event.target.value })} rows="4" placeholder="未入力のまま保存すると「空です」で登録されます。" />
      </label>
      <label className="full-width">
        <span>Next Action</span>
        <textarea value={form.nextAction} onChange={(event) => onChange({ ...form, nextAction: event.target.value })} rows="3" />
      </label>
      <label className="full-width">
        <span>報告メモ</span>
        <textarea value={form.reportMemo} onChange={(event) => onChange({ ...form, reportMemo: event.target.value })} rows="3" />
      </label>
      {isEditMode ? (
        <label>
          <span>更新日</span>
          <input
            type="date"
            value={form.updatedAtInput || ""}
            onChange={(event) => onChange({ ...form, updatedAtInput: event.target.value })}
            required
          />
        </label>
      ) : null}
      <div className="form-actions">
        {onCancel ? <button type="button" className="secondary-button" onClick={onCancel} disabled={saving}>{cancelLabel}</button> : null}
        <button type="submit" disabled={saving}>{submitLabel}</button>
      </div>
    </form>
  )
}

export default function App() {
  const defaultReportEndDate = formatDateInputValue(new Date())
  const defaultReportStartDate = formatDateInputValue(shiftDate(new Date(), -13))
  const [items, setItems] = useState([])
  const [summary, setSummary] = useState({ total: 0, byStatus: {} })
  const [appSettings, setAppSettings] = useState(defaultAppSettings)
  const [query, setQuery] = useState("")
  const [statusFilter, setStatusFilter] = useState("")
  const [kpiFilter, setKpiFilter] = useState("")
  const [categoryFilter, setCategoryFilter] = useState("")
  const [settingsPath, setSettingsPath] = useState("")
  const [draftPath, setDraftPath] = useState("")
  const [settingsDraft, setSettingsDraft] = useState(null)
  const [settingsOptionDrafts, setSettingsOptionDrafts] = useState({
    categoryOptions: "",
    assigneeOptions: "",
    statusOptions: "",
    rankOptions: "",
    leadSourceOptions: ""
  })
  const [settingsError, setSettingsError] = useState("")
  const [settingsSaving, setSettingsSaving] = useState(false)
  const [form, setForm] = useState(emptyForm)
  const [selectedId, setSelectedId] = useState("")
  const [panelMode, setPanelMode] = useState("detail")
  const [isDrawerOpen, setIsDrawerOpen] = useState(false)
  const [message, setMessage] = useState("")
  const [error, setError] = useState("")
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [pathDirty, setPathDirty] = useState(false)
  const [reportPreset, setReportPreset] = useState("14")
  const [reportStartDate, setReportStartDate] = useState(defaultReportStartDate)
  const [reportEndDate, setReportEndDate] = useState(defaultReportEndDate)
  const [reportText, setReportText] = useState("")
  const [reportNotice, setReportNotice] = useState("")
  const [currentPage, setCurrentPage] = useState(1)
  const [duplicateForm, setDuplicateForm] = useState(null)
  const [duplicateSourceLabel, setDuplicateSourceLabel] = useState("")
  const [duplicateSaving, setDuplicateSaving] = useState(false)
  const [duplicateError, setDuplicateError] = useState("")
  const [deleteTarget, setDeleteTarget] = useState(null)
  const [startupState, setStartupState] = useState(null)
  const [startupError, setStartupError] = useState("")

  async function loadData(nextQuery = query, nextStatus = statusFilter, manageLoading = true) {
    if (manageLoading) {
      setLoading(true)
    }

    setError("")

    try {
      const response = await call("get_dashboard", {
        query: nextQuery || null,
        status: nextStatus || null
      })
      const normalizedSettings = normalizeAppSettings(response.settings)

      setItems(response.items)
      setSummary(response.summary)
      setAppSettings(normalizedSettings)
      setSettingsPath(normalizedSettings.excelFilePath)
      setDraftPath(normalizedSettings.excelFilePath)
      setPathDirty(false)
      setStartupState(null)

      if (selectedId) {
        const selected = response.items.find((item) => item.id === selectedId)
        if (selected && panelMode === "form") {
          setForm(toFormState(selected))
        }

        if (!selected && panelMode === "detail") {
          setSelectedId("")
        }
      }
    } catch (loadError) {
      setError(loadError.message)
    } finally {
      if (manageLoading) {
        setLoading(false)
      }
    }
  }

  async function initializeApp() {
    setLoading(true)
    setError("")
    setStartupError("")

    try {
      const startup = await call("get_startup_state")
      const normalizedSettings = normalizeAppSettings(startup.settings)
      const fallbackPath = normalizeExcelPath(startup.suggestedNewExcelPath || "progress.xlsx")

      setAppSettings(normalizedSettings)
      setSettingsPath(normalizedSettings.excelFilePath)
      setDraftPath(normalizedSettings.excelFilePath || fallbackPath)
      setPathDirty(false)

      if (startup.needsOnboarding) {
        setStartupState({
          configuredExcelPath: normalizedSettings.excelFilePath,
          hasMissingConfiguredExcel: startup.hasConfiguredExcel && !startup.configuredExcelExists,
          suggestedNewExcelPath: fallbackPath
        })
        return
      }

      await loadData("", "", false)
    } catch (loadError) {
      setError(loadError.message)
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    void initializeApp()
  }, [])

  // 履歴保存機能を削除したため、初期化は不要

  useEffect(() => {
    if (!isDrawerOpen) {
      return undefined
    }

    function handleKeyDown(event) {
      if (event.key === "Escape" && !saving) {
        setIsDrawerOpen(false)
      }
    }

    window.addEventListener("keydown", handleKeyDown)

    return () => {
      window.removeEventListener("keydown", handleKeyDown)
    }
  }, [isDrawerOpen, saving])

  useEffect(() => {
    if (statusFilter && !appSettings.statusOptions.includes(statusFilter)) {
      setStatusFilter("")
      void loadData(query, "")
    }
  }, [appSettings.statusOptions, query, statusFilter])

  useEffect(() => {
    if (categoryFilter && !appSettings.categoryOptions.includes(categoryFilter)) {
      setCategoryFilter("")
    }
  }, [appSettings.categoryOptions, categoryFilter])

  const visibleTableColumns = useMemo(() => {
    const columnsByKey = new Map(tableColumnDefinitions.map((column) => [column.key, column]))
    return appSettings.visibleColumns.map((key) => columnsByKey.get(key)).filter(Boolean)
  }, [appSettings.visibleColumns])

  const filteredItems = useMemo(() => {
    return items.filter((item) => {
      const matchesKpi = kpiFilter ? item.kpiNumber === kpiFilter : true
      const matchesCategory = categoryFilter ? item.category === categoryFilter : true

      return matchesKpi && matchesCategory
    })
  }, [items, kpiFilter, categoryFilter])

  const summaryCards = useMemo(() => {
    const counts = new Map()

    filteredItems.forEach((item) => {
      const status = item.status?.trim() || "未設定"
      counts.set(status, (counts.get(status) || 0) + 1)
    })

    const orderedStatuses = [
      ...appSettings.statusOptions.filter((status) => counts.has(status)),
      ...Array.from(counts.keys()).filter((status) => !appSettings.statusOptions.includes(status)).sort()
    ]

    return orderedStatuses.map((status) => ({ status, count: counts.get(status) || 0 }))
  }, [appSettings.statusOptions, filteredItems])

  const rankDealSummary = useMemo(() => {
    const summaryByRank = new Map()

    filteredItems.forEach((item) => {
      const rank = item.rank?.trim()
      const dealSizeUnits = parseDealSizeUnits(item.dealSize)

      if (!rank || dealSizeUnits === null) {
        return
      }

      const current = summaryByRank.get(rank) || { count: 0, totalUnits: 0 }
      summaryByRank.set(rank, {
        count: current.count + 1,
        totalUnits: current.totalUnits + dealSizeUnits
      })
    })

    const orderedRanks = [
      ...appSettings.rankOptions.filter((rank) => summaryByRank.has(rank)),
      ...Array.from(summaryByRank.keys()).filter((rank) => !appSettings.rankOptions.includes(rank)).sort()
    ]

    const rows = orderedRanks.map((rank, index) => ({
      rank,
      count: summaryByRank.get(rank)?.count || 0,
      totalUnits: summaryByRank.get(rank)?.totalUnits || 0,
      color: rankChartPalette[index % rankChartPalette.length]
    }))

    return {
      rows,
      totalCount: rows.reduce((total, row) => total + row.count, 0),
      totalUnits: rows.reduce((total, row) => total + row.totalUnits, 0)
    }
  }, [appSettings.rankOptions, filteredItems])

  const totalPages = useMemo(() => {
    return Math.max(1, Math.ceil(filteredItems.length / pageSize))
  }, [filteredItems.length])

  const paginatedItems = useMemo(() => {
    const start = (currentPage - 1) * pageSize
    return filteredItems.slice(start, start + pageSize)
  }, [currentPage, filteredItems])

  const selectedItem = useMemo(() => {
    return items.find((item) => item.id === selectedId) || null
  }, [items, selectedId])

  const selectedVisibleIndex = useMemo(() => {
    return filteredItems.findIndex((item) => item.id === selectedId)
  }, [filteredItems, selectedId])

  const reportRange = useMemo(() => {
    return resolveReportRange({
      preset: reportPreset,
      startDateInput: reportStartDate,
      endDateInput: reportEndDate
    })
  }, [reportEndDate, reportPreset, reportStartDate])

  const reportPreview = useMemo(() => {
    if (!reportRange.isValid) {
      return buildInvalidReportDraft(filteredItems, reportRange.errorMessage)
    }

    return buildReportDraft({
      items: filteredItems,
      range: reportRange,
      statusFilter,
      kpiFilter,
      categoryFilter,
      query
    })
  }, [categoryFilter, filteredItems, kpiFilter, query, reportRange, statusFilter])

  function handleShowDetail(item) {
    setSelectedId(item.id)
    setPanelMode("detail")
    setIsDrawerOpen(true)
    setMessage("")
    setError("")
  }

  function handleNew() {
    setSelectedId("")
    setForm(emptyForm)
    setPanelMode("form")
    setIsDrawerOpen(true)
    setMessage("")
    setError("")
  }

  function handleEdit(item) {
    setSelectedId(item.id)
    setForm(toFormState(item))
    setPanelMode("form")
    setIsDrawerOpen(true)
    setMessage("")
    setError("")
  }

  function handleDuplicate(item) {
    setDuplicateForm(toDuplicateFormState(item))
    setDuplicateSourceLabel(`${formatKpiDisplay(item.kpiNumber)} / ${item.assignee}`)
    setDuplicateSaving(false)
    setDuplicateError("")
  }

  function handleCloseDuplicateModal() {
    if (duplicateSaving) {
      return
    }

    setDuplicateForm(null)
    setDuplicateSourceLabel("")
    setDuplicateError("")
  }

  function handleRequestDelete(item) {
    setDeleteTarget(item)
    setMessage("")
    setError("")
  }

  function handleCloseDeleteModal() {
    if (saving) {
      return
    }

    setDeleteTarget(null)
  }

  function handleOpenReport() {
    setPanelMode("report")
    setReportText(reportPreview.text)
    setReportNotice("")
    setIsDrawerOpen(true)
    setMessage("")
    setError("")
  }

  function handleOpenSettings() {
    setSettingsDraft(normalizeAppSettings({
      ...appSettings,
      excelFilePath: settingsPath || appSettings.excelFilePath
    }))
    setSettingsOptionDrafts({
      categoryOptions: "",
      assigneeOptions: "",
      statusOptions: "",
      rankOptions: "",
      leadSourceOptions: ""
    })
    setSettingsError("")
  }

  function handleCloseSettings() {
    if (settingsSaving) {
      return
    }

    setSettingsDraft(null)
    setSettingsError("")
  }

  function handleCloseDrawer() {
    if (saving) {
      return
    }

    setIsDrawerOpen(false)
  }

  useEffect(() => {
    setCurrentPage(1)
  }, [query, statusFilter, kpiFilter, categoryFilter, items])

  useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(totalPages)
    }
  }, [currentPage, totalPages])

  function handleNavigateDetail(offset) {
    const nextItem = filteredItems[selectedVisibleIndex + offset]

    if (!nextItem) {
      return
    }

    setSelectedId(nextItem.id)
    setPanelMode("detail")
    setIsDrawerOpen(true)
  }

  async function handleResetFilters() {
    setQuery("")
    setStatusFilter("")
    setKpiFilter("")
    setCategoryFilter("")
    await loadData("", "")
  }

  async function persistExcelPath(nextPath, successMessage) {
    const normalizedPath = normalizeExcelPath(nextPath)
    const response = await call("set_excel_file_path", {
      excelFilePath: normalizedPath
    })
    const normalizedSettings = normalizeAppSettings(response)

    setAppSettings(normalizedSettings)
    setSettingsPath(normalizedSettings.excelFilePath)
    setDraftPath(normalizedSettings.excelFilePath)
    setPathDirty(false)
    setMessage(successMessage)
    await loadData()
  }

  async function commitDraftPath() {
    const trimmedPath = draftPath.trim()

    if (!pathDirty) {
      return
    }

    if (!trimmedPath) {
      setDraftPath(settingsPath)
      setPathDirty(false)
      return
    }

    if (normalizeExcelPath(trimmedPath) === normalizeExcelPath(settingsPath)) {
      setPathDirty(false)
      return
    }

    setSaving(true)
    setMessage("")
    setError("")

    try {
      await persistExcelPath(trimmedPath, "Excel 保存先を更新しました。")
    } catch (saveError) {
      setError(saveError.message)
    } finally {
      setSaving(false)
    }
  }

  async function handleBrowseExcel() {
    setSaving(true)
    setError("")
    setMessage("")

    try {
      const selected = await open({
        multiple: false,
        directory: false,
        filters: [
          {
            name: "Excel Workbook",
            extensions: ["xlsx"]
          }
        ]
      })

      if (typeof selected === "string") {
        await persistExcelPath(selected, "既存の Excel 保存先へ切り替えました。")
      }
    } catch (dialogError) {
      setError(dialogError.message)
    } finally {
      setSaving(false)
    }
  }

  async function handleCreateExcel() {
    setSaving(true)
    setError("")
    setMessage("")

    try {
      const selected = await save({
        filters: [
          {
            name: "Excel Workbook",
            extensions: ["xlsx"]
          }
        ],
        defaultPath: normalizeExcelPath(draftPath || "progress.xlsx")
      })

      if (typeof selected === "string") {
        await persistExcelPath(selected, "新しい Excel 保存先へ切り替えました。")
      }
    } catch (dialogError) {
      setError(dialogError.message)
    } finally {
      setSaving(false)
    }
  }

  async function handleStartupOpenExisting() {
    setSaving(true)
    setStartupError("")
    setError("")
    setMessage("")

    try {
      const selected = await open({
        multiple: false,
        directory: false,
        filters: [
          {
            name: "Excel Workbook",
            extensions: ["xlsx"]
          }
        ]
      })

      if (typeof selected === "string") {
        await persistExcelPath(selected, "既存の Excel を読み込んで起動しました。")
        setStartupState(null)
      }
    } catch (dialogError) {
      setStartupError(dialogError.message)
    } finally {
      setSaving(false)
    }
  }

  async function handleStartupCreateNew() {
    setSaving(true)
    setStartupError("")
    setError("")
    setMessage("")

    try {
      const selected = await save({
        filters: [
          {
            name: "Excel Workbook",
            extensions: ["xlsx"]
          }
        ],
        defaultPath: startupState?.suggestedNewExcelPath || normalizeExcelPath(draftPath || "progress.xlsx")
      })

      if (typeof selected === "string") {
        await persistExcelPath(selected, "新しい Excel を作成して起動しました。")
        setStartupState(null)
      }
    } catch (dialogError) {
      setStartupError(dialogError.message)
    } finally {
      setSaving(false)
    }
  }

  async function handleExportExcel() {
    setSaving(true)
    setError("")
    setMessage("")

    try {
      const selected = await save({
        filters: [
          {
            name: "Excel Workbook",
            extensions: ["xlsx"]
          }
        ],
        defaultPath: normalizeExcelPath("progress-export.xlsx")
      })

      if (typeof selected === "string") {
        const exportPath = normalizeExcelPath(selected)
        const savedPath = await call("export_current_excel", {
          exportFilePath: exportPath
        })
        setMessage(`現在の状態を Excel にエクスポートしました: ${savedPath}`)
      }
    } catch (dialogError) {
      setError(dialogError.message)
    } finally {
      setSaving(false)
    }
  }

  function handlePathInputChange(event) {
    setDraftPath(event.target.value)
    setPathDirty(true)
  }

  function handlePathInputBlur(event) {
    const nextTarget = event.relatedTarget

    if (nextTarget?.dataset?.pathAction === "true") {
      return
    }

    void commitDraftPath()
  }

  function handlePathInputKeyDown(event) {
    if (event.key !== "Enter") {
      return
    }

    event.preventDefault()
    void commitDraftPath()
  }

  function handleSettingsOptionDraftChange(field, value) {
    setSettingsOptionDrafts((current) => ({
      ...current,
      [field]: value
    }))
  }

  function handleAddOption(field) {
    if (!settingsDraft) {
      return
    }

    const draftValue = String(settingsOptionDrafts[field] || "").trim()

    if (!draftValue) {
      return
    }

    const nextDraft = settingsDraft ? ({
      ...settingsDraft,
      [field]: sanitizeOptionList([...settingsDraft[field], draftValue])
    }) : null
    if (!nextDraft) return
    setSettingsDraft(nextDraft)
    setSettingsOptionDrafts((current) => ({
      ...current,
      [field]: ""
    }))
    setSettingsError("")
    void persistSettings(nextDraft)
  }

  function handleRemoveOption(field, value) {
    if (!settingsDraft) return
    const nextDraft = {
      ...settingsDraft,
      [field]: settingsDraft[field].filter((item) => item !== value)
    }
    setSettingsDraft(nextDraft)
    void persistSettings(nextDraft)
  }

  function handleMoveOption(field, index, direction) {
    if (!settingsDraft) return
    const nextDraft = {
      ...settingsDraft,
      [field]: reorderList(settingsDraft[field], index, index + direction)
    }
    setSettingsDraft(nextDraft)
    void persistSettings(nextDraft)
  }

  function handleReorderOption(field, fromIndex, toIndex) {
    if (!settingsDraft) return
    const nextItems = insertListItem(settingsDraft[field], fromIndex, toIndex)

    if (nextItems === settingsDraft[field]) {
      return
    }

    const nextDraft = {
      ...settingsDraft,
      [field]: nextItems
    }
    setSettingsDraft(nextDraft)
    void persistSettings(nextDraft)
  }

  function handleToggleColumn(columnKey) {
    if (!settingsDraft) return
    const isActive = settingsDraft.visibleColumns.includes(columnKey)
    const visibleColumns = isActive
      ? settingsDraft.visibleColumns.filter((item) => item !== columnKey)
      : [...settingsDraft.visibleColumns, columnKey]

    const nextDraft = { ...settingsDraft, visibleColumns }
    setSettingsDraft(nextDraft)
    void persistSettings(nextDraft)
  }

  function handleMoveColumn(index, direction) {
    if (!settingsDraft) return
    const nextDraft = {
      ...settingsDraft,
      visibleColumns: reorderList(settingsDraft.visibleColumns, index, index + direction)
    }
    setSettingsDraft(nextDraft)
    void persistSettings(nextDraft)
  }

  function handleReorderColumn(fromIndex, toIndex) {
    if (!settingsDraft) return
    const nextColumns = insertListItem(settingsDraft.visibleColumns, fromIndex, toIndex)

    if (nextColumns === settingsDraft.visibleColumns) {
      return
    }

    const nextDraft = {
      ...settingsDraft,
      visibleColumns: nextColumns
    }
    setSettingsDraft(nextDraft)
    void persistSettings(nextDraft)
  }

  async function handleSaveSettings() {
    if (!settingsDraft) {
      return
    }

    const nextSettings = normalizeAppSettings({
      ...settingsDraft,
      excelFilePath: settingsPath || appSettings.excelFilePath
    })

    if (nextSettings.visibleColumns.length === 0) {
      setSettingsError("進捗一覧に表示する列を1つ以上選択してください。")
      return
    }

    // keep legacy save available but delegate to persistSettings
    await persistSettings(nextSettings)
  }

  async function persistSettings(nextDraft) {
    if (!nextDraft) return

    const nextSettings = normalizeAppSettings({
      ...nextDraft,
      excelFilePath: settingsPath || appSettings.excelFilePath
    })

    if (nextSettings.visibleColumns.length === 0) {
      setSettingsError("進捗一覧に表示する列を1つ以上選択してください。")
      return
    }

    setSettingsSaving(true)
    setSettingsError("")
    setMessage("")
    setError("")

    try {
      const response = await call("update_app_settings", {
        settings: nextSettings
      })

      const normalizedSettings = normalizeAppSettings(response)
      setAppSettings(normalizedSettings)
      setSettingsDraft(normalizedSettings)
      setMessage("設定を更新しました。")
    } catch (saveError) {
      setSettingsError(saveError.message)
    } finally {
      setSettingsSaving(false)
    }
  }

  async function handleSubmit(event) {
    event.preventDefault()
    setSaving(true)
    setMessage("")
    setError("")

    try {
      const payload = normalizeProgressPayload(form)
      const savedItem = selectedId
        ? await call("update_progress", {
            id: selectedId,
            payload
          })
        : await call("create_progress", {
            payload
          })

      setSelectedId(savedItem.id)
          setForm(toFormState(savedItem))
      setPanelMode("detail")
      setIsDrawerOpen(false)
      setMessage(selectedId ? "進捗を更新しました。" : "進捗を登録しました。")
      await loadData()
    } catch (submitError) {
      setError(submitError.message)
    } finally {
      setSaving(false)
    }
  }

  async function handleDelete(item) {
    if (!item) {
      return
    }

    setSaving(true)
    setMessage("")
    setError("")

    try {
      await call("delete_progress", { id: item.id })

      if (selectedId === item.id) {
        setSelectedId("")
        setIsDrawerOpen(false)
      }

      setDeleteTarget(null)
      setMessage("進捗を削除しました。")
      await loadData()
    } catch (deleteError) {
      setError(deleteError.message)
    } finally {
      setSaving(false)
    }
  }

  async function handleDuplicateSubmit(event) {
    event.preventDefault()
    setDuplicateSaving(true)
    setDuplicateError("")
    setMessage("")
    setError("")

    try {
      const savedItem = await call("create_progress", {
        payload: normalizeProgressPayload(duplicateForm)
      })

      setSelectedId(savedItem.id)
      setForm(toFormState(savedItem))
      setDuplicateForm(null)
      setDuplicateSourceLabel("")
      setMessage("進捗を複製して登録しました。")
      await loadData()
    } catch (submitError) {
      setDuplicateError(submitError.message)
    } finally {
      setDuplicateSaving(false)
    }
  }

  function handleGenerateReport() {
    setReportText(reportPreview.text)
    setReportNotice(reportPreview.isValid ? "最新の条件で会議向け報告文を再生成しました。" : reportPreview.errorMessage)
  }

  async function handleCopyReport() {
    try {
      await navigator.clipboard.writeText(reportText)
      setReportNotice("報告文をクリップボードにコピーしました。")
    } catch {
      setReportNotice("コピーに失敗しました。プレビューから手動でコピーしてください。")
    }
  }

  // 履歴保存は削除されています

  function handleReportPresetChange(event) {
    const nextPreset = event.target.value

    if (nextPreset === "custom" && reportPreset !== "custom" && reportRange.isValid) {
      setReportStartDate(formatDateInputValue(reportRange.start))
      setReportEndDate(formatDateInputValue(reportRange.end))
    }

    setReportPreset(nextPreset)
    setReportNotice("")
  }

  return (
    <div className="app-shell">
      <header className="hero">
        <div>
          <p className="eyebrow">LOCAL EXCEL PROTOTYPE</p>
          <h1>進捗管理 App</h1>
          <p className="hero-copy">
            共有フォルダまたはローカルの Excel を直接データソースにして、サーバーなしで一覧確認・登録・更新を行う最小構成です。
          </p>
        </div>
      </header>

      <div className="app-version">Version v0.4</div>

      {startupState ? (
        <StartupWizard
          startupState={startupState}
          saving={saving}
          error={startupError}
          onOpenExisting={handleStartupOpenExisting}
          onCreateNew={handleStartupCreateNew}
        />
      ) : (
        <>
      <main className="content-grid">
        

        <section className="summary-section">
          <div className="summary-export-action">
            <button type="button" onClick={handleExportExcel} disabled={saving}>エクスポート</button>
            <button type="button" className="settings-button" onClick={handleOpenSettings} disabled={saving || loading}>設定</button>
          </div>
          <div className="card summary-card">
            <h2>ステータス集計</h2>
            <div className="summary-list">
              <article className="summary-chip total-chip">
                <span>総件数</span>
                <strong>{filteredItems.length}</strong>
              </article>
              {summaryCards.length === 0 ? <p className="hint">データ未登録です。</p> : null}
              {summaryCards.map((item) => (
                <article className="summary-chip" key={item.status}>
                  <span>{item.status}</span>
                  <strong>{item.count}</strong>
                </article>
              ))}
            </div>
            <div className="summary-breakdown">
              {rankDealSummary.rows.length === 0 ? (
                <p className="hint">ランクとディールサイズが入力された営業案件はありません。</p>
              ) : (
                <RankDealDonutChart
                  rows={rankDealSummary.rows}
                  totalUnits={rankDealSummary.totalUnits}
                  totalCount={rankDealSummary.totalCount}
                />
              )}
            </div>
          </div>
        </section>

        <section className="card list-card">
          <div className="section-head">
            <div>
              <h2>進捗一覧</h2>
              <p className="hint section-copy">主要項目だけを一覧に残し、詳細は右からスライド表示する構成です。</p>
            </div>
            <div className="section-actions">
              <button type="button" className="report-button" onClick={handleOpenReport}>定例報告書作成</button>
              <button type="button" onClick={handleNew}>新規入力</button>
            </div>
          </div>

          <div className="filters filters-extended">
            <input
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              placeholder="KPI番号・カテゴリ・担当者・内容・報告メモで検索"
            />
            <button type="button" onClick={() => void loadData(query, statusFilter)}>検索</button>
            <select value={statusFilter} onChange={(event) => {
              const value = event.target.value
              setStatusFilter(value)
              setCurrentPage(1)
              void loadData(query, value)
            }}>
              <option value="">全ステータス</option>
              {appSettings.statusOptions.map((status) => (
                <option key={status} value={status}>{status}</option>
              ))}
            </select>
            <select value={kpiFilter} onChange={(event) => {
              setKpiFilter(event.target.value)
              setCurrentPage(1)
            }}>
              <option value="">全KPI</option>
              {kpiOptions.map((kpi) => (
                <option key={kpi} value={kpi}>{kpi}</option>
              ))}
            </select>
            <select value={categoryFilter} onChange={(event) => {
              setCategoryFilter(event.target.value)
              setCurrentPage(1)
            }}>
              <option value="">全カテゴリ</option>
              {appSettings.categoryOptions.map((category) => (
                <option key={category} value={category}>{category}</option>
              ))}
            </select>
            <button type="button" className="secondary-button" onClick={() => void handleResetFilters()}>条件クリア</button>
          </div>

          {loading ? <p className="hint">読み込み中...</p> : null}
          {!loading && filteredItems.length === 0 ? <p className="hint">条件に一致するデータがありません。</p> : null}

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>No.</th>
                  {visibleTableColumns.map((column) => (
                    <th key={column.key} className={`col-${column.key}`}>{column.label}</th>
                  ))}
                  <th>操作</th>
                </tr>
              </thead>
              <tbody>
                {paginatedItems.map((item, index) => (
                  <tr
                    key={item.id}
                    className={selectedId === item.id ? "selected" : ""}
                    onClick={() => handleShowDetail(item)}
                  >
                    <td>{(currentPage - 1) * pageSize + index + 1}</td>
                    {visibleTableColumns.map((column) => (
                      <td key={`${item.id}-${column.key}`} className={`col-${column.key}`}>{column.render(item)}</td>
                    ))}
                    <td>
                      <div className="row-actions">
                        <button
                          type="button"
                          className="secondary-button table-action"
                          onClick={(event) => {
                            event.stopPropagation()
                            handleShowDetail(item)
                          }}
                        >
                          詳細
                        </button>
                        <button
                          type="button"
                          className="secondary-button table-action"
                          onClick={(event) => {
                            event.stopPropagation()
                            handleEdit(item)
                          }}
                        >
                          編集
                        </button>
                        <button
                          type="button"
                          className="secondary-button table-action"
                          onClick={(event) => {
                            event.stopPropagation()
                            handleDuplicate(item)
                          }}
                        >
                          複製
                        </button>
                        <button
                          type="button"
                          className="secondary-button table-action danger-action"
                          onClick={(event) => {
                            event.stopPropagation()
                            handleRequestDelete(item)
                          }}
                        >
                          削除
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {filteredItems.length > pageSize ? (
            <div className="pagination">
              <button type="button" className="secondary-button" onClick={() => setCurrentPage((page) => Math.max(1, page - 1))} disabled={currentPage === 1}>前の10件</button>
              <span className="pagination-label">{currentPage} / {totalPages} ページ</span>
              <button type="button" className="secondary-button" onClick={() => setCurrentPage((page) => Math.min(totalPages, page + 1))} disabled={currentPage === totalPages}>次の10件</button>
            </div>
          ) : null}
          <p className="hint table-caption">詳細、内容、関係者、Next Action、報告メモは右側のスライドパネルで確認できます。KPI とカテゴリは上部の絞り込みで絞れます。</p>
          {message ? <p className="message success">{message}</p> : null}
          {error ? <p className="message error">{error}</p> : null}
        </section>
          <section className="card settings-card full-width">
            <h2>Excel 保存先</h2>
            <div className="inline-form">
              <input
                value={draftPath}
                onBlur={handlePathInputBlur}
                onChange={handlePathInputChange}
                onKeyDown={handlePathInputKeyDown}
                placeholder="C:\\Users\\...\\progress.xlsx または \\\\server\\share\\progress.xlsx"
              />
              <button type="button" className="secondary-button" data-path-action="true" onClick={handleBrowseExcel} disabled={saving}>参照</button>
              <button type="button" className="secondary-button" data-path-action="true" onClick={handleCreateExcel} disabled={saving}>新規作成</button>
            </div>
            <p className="hint settings-current-path">現在の保存先: <span className="path-label">{settingsPath || "未設定"}</span></p>
            <p className="hint">手入力は Enter またはフォーカス移動で反映されます。既存ファイルは参照、新規ファイルは新規作成を選んだ時点で切り替わります。</p>
          </section>
      </main>
      <div className={`drawer-backdrop ${isDrawerOpen ? "open" : ""}`} onClick={handleCloseDrawer}>
        <aside
          className={`detail-drawer ${isDrawerOpen ? "open" : ""}`}
          onClick={(event) => event.stopPropagation()}
          role="dialog"
          aria-modal="true"
          aria-label={panelMode === "form" ? "進捗登録・編集" : panelMode === "report" ? "定例報告書作成" : "進捗詳細"}
        >
          <div className="drawer-head">
            <div>
              <p className="eyebrow">{panelMode === "form" ? "ENTRY PANEL" : panelMode === "report" ? "REPORT PANEL" : "DETAIL PANEL"}</p>
              <h2>{panelMode === "form" ? (selectedId ? "進捗更新" : "進捗登録") : panelMode === "report" ? "定例報告書作成" : "進捗詳細"}</h2>
            </div>
            <div className="drawer-head-actions">
              {panelMode === "detail" && selectedItem ? (
                <div className="drawer-nav">
                  <button type="button" className="secondary-button" onClick={() => handleNavigateDetail(-1)} disabled={selectedVisibleIndex <= 0}>前へ</button>
                  <button type="button" className="secondary-button" onClick={() => handleNavigateDetail(1)} disabled={selectedVisibleIndex < 0 || selectedVisibleIndex >= filteredItems.length - 1}>次へ</button>
                </div>
              ) : null}
              <button type="button" className="secondary-button" onClick={handleCloseDrawer} disabled={saving}>閉じる</button>
            </div>
          </div>

          {panelMode === "report" ? (
            <div className="report-panel">
              <section className="detail-section report-controls">
                <div className="report-toolbar">
                  <label>
                    <span>報告期間</span>
                    <select value={reportPreset} onChange={handleReportPresetChange}>
                      {reportPresetOptions.map((option) => (
                        <option key={option.value} value={option.value}>{option.label}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    <span>対象範囲</span>
                    <input value={reportPreview.periodLabel} readOnly />
                  </label>
                </div>
                {reportPreset === "custom" ? (
                  <div className="report-date-range">
                    <label>
                      <span>開始日</span>
                      <input type="date" value={reportStartDate} onChange={(event) => setReportStartDate(event.target.value)} />
                    </label>
                    <label>
                      <span>終了日</span>
                      <input type="date" value={reportEndDate} onChange={(event) => setReportEndDate(event.target.value)} />
                    </label>
                  </div>
                ) : null}
                {!reportPreview.isValid ? <p className="message error compact-message">{reportPreview.errorMessage}</p> : null}
                <p className="hint report-filter-note">
                  現在の一覧条件をそのまま使います。検索語、ステータス、KPI、カテゴリで絞った結果から報告文を作成します。
                </p>
                {/* 履歴保存・差分機能は削除されています */}
                <div className="report-actions">
                  <button type="button" className="secondary-button" onClick={handleGenerateReport}>再生成</button>
                  <button type="button" onClick={handleCopyReport}>コピー</button>
                </div>
              </section>

              <section className="report-metrics">
                <article className="summary-chip">
                  <span>対象案件</span>
                  <strong>{reportPreview.metrics.total}</strong>
                </article>
                <article className="summary-chip">
                  <span>期間内更新</span>
                  <strong>{reportPreview.metrics.updated}</strong>
                </article>
                <article className="summary-chip">
                  <span>完了</span>
                  <strong>{reportPreview.metrics.completed}</strong>
                </article>
                <article className="summary-chip">
                  <span>保留</span>
                  <strong>{reportPreview.metrics.onHold}</strong>
                </article>
                <article className="summary-chip">
                  <span>更新なし</span>
                  <strong>{reportPreview.metrics.stale}</strong>
                </article>
                <article className="summary-chip">
                  <span>報告メモ未入力</span>
                  <strong>{reportPreview.metrics.noReportMemo}</strong>
                </article>
                {/* 差分関連の統計は削除 */}
              </section>

              {/* 履歴と差分表示は削除されました */}

              <section className="detail-section report-preview">
                <h3>会議向け報告プレビュー</h3>
                <textarea value={reportText} onChange={(event) => setReportText(event.target.value)} rows="18" />
                {reportNotice ? <p className="message success compact-message">{reportNotice}</p> : null}
              </section>
            </div>
          ) : panelMode === "detail" ? (
            selectedItem ? (
              <div className="detail-panel">
                <div className="detail-hero">
                  <div>
                    <p className="eyebrow">{selectedItem.category || "未設定カテゴリ"}</p>
                    <h3>{selectedItem.title?.trim() || `${formatKpiDisplay(selectedItem.kpiNumber)} / ${selectedItem.assignee}`}</h3>
                    <p className="hint">{formatKpiDisplay(selectedItem.kpiNumber)} / {selectedItem.assignee}</p>
                  </div>
                  <div className="detail-hero-actions">
                    <span className="pill">{selectedItem.status}</span>
                    <button type="button" onClick={() => handleEdit(selectedItem)}>編集する</button>
                  </div>
                </div>

                <div className="detail-grid">
                  <article className="detail-block">
                    <span>タイトル</span>
                    <strong>{selectedItem.title || "-"}</strong>
                  </article>
                  <article className="detail-block">
                    <span>顧客名 / Project 名</span>
                    <strong>{truncateText(selectedItem.customer || "-")}</strong>
                  </article>
                  <article className="detail-block">
                    <span>更新日</span>
                    <strong>{formatDate(selectedItem.updatedAt)}</strong>
                  </article>
                  {selectedItem.category === "営業" ? (
                    <>
                      <article className="detail-block">
                        <span>ランク</span>
                        <strong>{selectedItem.rank || "-"}</strong>
                      </article>
                      <article className="detail-block">
                        <span>ディールサイズ</span>
                        <strong>{formatDealSizeDisplay(selectedItem.dealSize)}</strong>
                      </article>
                    </>
                  ) : null}
                </div>

                <div className="detail-sections">
                  <section className="detail-section">
                    <h3>内容</h3>
                    <p>{selectedItem.content || "空です"}</p>
                  </section>
                  <section className="detail-section">
                    <h3>Next Action</h3>
                    <p>{selectedItem.nextAction || "-"}</p>
                  </section>
                  <section className="detail-section">
                    <h3>報告メモ</h3>
                    <p>{selectedItem.reportMemo || "-"}</p>
                  </section>
                  <section className="detail-section">
                    <h3>関係者・関連部署</h3>
                    <dl className="detail-list">
                      <div>
                        <dt>社外関係者</dt>
                        <dd>{selectedItem.externalStakeholders || "-"}</dd>
                      </div>
                      <div>
                        <dt>社内関連部署</dt>
                        <dd>{selectedItem.internalDepartments || "-"}</dd>
                      </div>
                    </dl>
                  </section>
                </div>
              </div>
            ) : (
              <div className="empty-detail-state">
                <p className="eyebrow">DETAIL VIEW</p>
                <h3>一覧から詳細を選択してください</h3>
                <p className="hint">一覧では KPI、カテゴリ、担当者、更新日に絞り、詳細はこのドロワーで確認する構成です。</p>
              </div>
            )
          ) : (
            <ProgressForm
              form={form}
              onChange={setForm}
              onSubmit={handleSubmit}
              onCancel={selectedId ? () => setPanelMode("detail") : null}
              saving={saving}
              submitLabel={selectedId ? "更新する" : "登録する"}
              cancelLabel="詳細に戻る"
              categoryOptions={appSettings.categoryOptions}
              assigneeOptions={appSettings.assigneeOptions}
              statusOptions={appSettings.statusOptions}
              rankOptions={appSettings.rankOptions}
              leadSourceOptions={appSettings.leadSourceOptions}
            />
          )}
        </aside>
      </div>

      {settingsDraft ? (
        <SettingsModal
          settingsDraft={settingsDraft}
          optionDrafts={settingsOptionDrafts}
          onOptionDraftChange={handleSettingsOptionDraftChange}
          onAddOption={handleAddOption}
          onRemoveOption={handleRemoveOption}
          onMoveOption={handleMoveOption}
          onReorderOption={handleReorderOption}
          onToggleColumn={handleToggleColumn}
          onMoveColumn={handleMoveColumn}
          onReorderColumn={handleReorderColumn}
          onClose={handleCloseSettings}
          error={settingsError}
          saving={settingsSaving}
        />
      ) : null}

      {duplicateForm ? (
        <div className="modal-backdrop" onClick={handleCloseDuplicateModal}>
          <section
            className="duplicate-modal"
            onClick={(event) => event.stopPropagation()}
            role="dialog"
            aria-modal="true"
            aria-label="進捗を複製して新規入力"
          >
            <div className="drawer-head">
              <div>
                <p className="eyebrow">DUPLICATE ENTRY</p>
                <h2>複製して新規入力</h2>
                <p className="hint section-copy">元データ: {duplicateSourceLabel}</p>
              </div>
              <button type="button" className="secondary-button" onClick={handleCloseDuplicateModal} disabled={duplicateSaving}>閉じる</button>
            </div>
            <ProgressForm
              form={duplicateForm}
              onChange={setDuplicateForm}
              onSubmit={handleDuplicateSubmit}
              onCancel={handleCloseDuplicateModal}
              saving={duplicateSaving}
              submitLabel="複製して登録する"
              cancelLabel="キャンセル"
              categoryOptions={appSettings.categoryOptions}
              assigneeOptions={appSettings.assigneeOptions}
              statusOptions={appSettings.statusOptions}
              rankOptions={appSettings.rankOptions}
              leadSourceOptions={appSettings.leadSourceOptions}
            />
            {duplicateError ? <p className="message error compact-message">{duplicateError}</p> : null}
          </section>
        </div>
      ) : null}

      {deleteTarget ? (
        <div className="modal-backdrop" onClick={handleCloseDeleteModal}>
          <section
            className="confirm-modal"
            onClick={(event) => event.stopPropagation()}
            role="dialog"
            aria-modal="true"
            aria-label="進捗削除の確認"
          >
            <div className="confirm-modal-copy">
              <p className="eyebrow">DELETE CONFIRMATION</p>
              <h2>この進捗を削除しますか</h2>
              <p>
                対象: {formatKpiDisplay(deleteTarget.kpiNumber)} / {deleteTarget.assignee}
              </p>
              <p className="hint">
                削除すると元に戻せません。内容、Next Action、報告メモを含めて一覧から完全に削除されます。
              </p>
            </div>
            <div className="confirm-modal-actions">
              <button type="button" className="secondary-button" onClick={handleCloseDeleteModal} disabled={saving}>キャンセル</button>
              <button type="button" className="danger-button" onClick={() => void handleDelete(deleteTarget)} disabled={saving}>削除する</button>
            </div>
            {error ? <p className="message error compact-message">{error}</p> : null}
          </section>
        </div>
      ) : null}
        </>
      )}
    </div>
  )
}
