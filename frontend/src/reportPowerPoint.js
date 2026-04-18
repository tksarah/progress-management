import PptxGenJS from "pptxgenjs"

const STATUS_PRIORITY = {
  "保留": 0,
  "進捗中": 1,
  "計画中": 2,
  "クローズ": 3
}

const STATUS_COLORS = {
  "保留": "D9485F",
  "進捗中": "2E89C9",
  "計画中": "63B3ED",
  "クローズ": "2F855A"
}

const RANK_CHART_PALETTE = [
  "0D5EA6",
  "2C7BE5",
  "15AABF",
  "2F9E77",
  "F08C00",
  "D9485F",
  "7950F2",
  "5F6B7A",
  "C2255C"
]

const BRAND = {
  navy: "14324A",
  blue: "2E89C9",
  cyan: "63B3ED",
  ink: "1F2933",
  muted: "5F6C7B",
  border: "D9E2EC",
  soft: "EEF6FB",
  softAlt: "F7FAFC",
  success: "2F855A",
  warning: "D69E2E",
  danger: "C53030"
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

function formatDateTime(value) {
  const parsed = toValidDate(value)

  if (!parsed) {
    return "-"
  }

  return parsed.toLocaleString("ja-JP")
}

function formatPeriodLabel(start, end) {
  return `${start.toLocaleDateString("ja-JP")} - ${end.toLocaleDateString("ja-JP")}`
}

function formatKpiDisplay(value) {
  return String(value || "").trim() || "-"
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

function formatDealSizeDisplay(value) {
  const units = parseDealSizeUnits(value)

  if (units === null) {
    const text = String(value || "").trim()
    return text || "-"
  }

  return `${units.toLocaleString("ja-JP")}万円`
}

function formatDealSizeUnitsLabel(value) {
  return `${value.toLocaleString("ja-JP")}万円`
}

function sanitizeText(value) {
  const text = String(value || "").trim()
  return text || "-"
}

function summarizeStatus(items) {
  return items.reduce((counts, item) => {
    const status = sanitizeText(item.status)
    counts[status] = (counts[status] || 0) + 1
    return counts
  }, {})
}

function statusRank(status) {
  return STATUS_PRIORITY[status] ?? 99
}

function sortItems(items) {
  return [...items].sort((left, right) => {
    const statusDifference = statusRank(left.status) - statusRank(right.status)

    if (statusDifference !== 0) {
      return statusDifference
    }

    const leftDate = toValidDate(left.updatedAt)?.getTime() || 0
    const rightDate = toValidDate(right.updatedAt)?.getTime() || 0

    if (leftDate !== rightDate) {
      return rightDate - leftDate
    }

    return sanitizeText(left.title).localeCompare(sanitizeText(right.title), "ja")
  })
}

function groupByCategory(items, categoryOrder) {
  const groups = new Map()

  items.forEach((item) => {
    const category = sanitizeText(item.category)

    if (!groups.has(category)) {
      groups.set(category, [])
    }

    groups.get(category).push(item)
  })

  const orderMap = new Map((categoryOrder || []).map((category, index) => [category, index]))

  return Array.from(groups.entries())
    .map(([category, categoryItems]) => {
      const sortedItems = sortItems(categoryItems)

      return {
        category,
        items: sortedItems,
        total: sortedItems.length,
        statusCounts: summarizeStatus(sortedItems),
        highlightItems: sortedItems
          .filter((item) => item.status === "保留" || !String(item.reportMemo || "").trim() || !String(item.nextAction || "").trim())
          .slice(0, 3)
      }
    })
    .sort((left, right) => {
      const leftOrder = orderMap.has(left.category) ? orderMap.get(left.category) : Number.MAX_SAFE_INTEGER
      const rightOrder = orderMap.has(right.category) ? orderMap.get(right.category) : Number.MAX_SAFE_INTEGER

      if (leftOrder !== rightOrder) {
        return leftOrder - rightOrder
      }

      return left.category.localeCompare(right.category, "ja")
    })
}

function chunkItems(items, size) {
  const chunks = []

  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size))
  }

  return chunks
}

function buildFilterLabels(filters) {
  return [
    filters.statusFilter ? `ステータス: ${filters.statusFilter}` : null,
    filters.kpiFilter ? `KPI: ${filters.kpiFilter}` : null,
    filters.categoryFilter ? `カテゴリ: ${filters.categoryFilter}` : null,
    filters.query ? `検索語: ${filters.query}` : null
  ].filter(Boolean)
}

function buildStatusRows(items) {
  const statusCounts = summarizeStatus(items)

  return Object.entries(statusCounts)
    .map(([status, count]) => ({
      label: status,
      value: count,
      color: STATUS_COLORS[status] || BRAND.blue
    }))
    .sort((left, right) => {
      const statusDifference = statusRank(left.label) - statusRank(right.label)

      if (statusDifference !== 0) {
        return statusDifference
      }

      return left.label.localeCompare(right.label, "ja")
    })
}

function buildRankRows(items, rankOrder) {
  const summaryByRank = new Map()

  items.forEach((item) => {
    const rank = String(item.rank || "").trim()
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
    ...(rankOrder || []).filter((rank) => summaryByRank.has(rank)),
    ...Array.from(summaryByRank.keys()).filter((rank) => !(rankOrder || []).includes(rank)).sort((left, right) => left.localeCompare(right, "ja"))
  ]

  return orderedRanks.map((rank, index) => ({
    label: rank,
    value: summaryByRank.get(rank)?.totalUnits || 0,
    count: summaryByRank.get(rank)?.count || 0,
    color: RANK_CHART_PALETTE[index % RANK_CHART_PALETTE.length]
  }))
}

export function buildPowerPointReportData({ items, range, metrics, filters, categoryOrder = [], rankOrder = [] }) {
  const updatedItems = items.filter((item) => {
    const updatedAt = toValidDate(item.updatedAt)
    return updatedAt && updatedAt >= range.start && updatedAt <= range.end
  })

  const statusRows = buildStatusRows(updatedItems)
  const rankRows = buildRankRows(updatedItems, rankOrder)

  return {
    title: "定例報告書",
    periodLabel: formatPeriodLabel(range.start, range.end),
    generatedAtLabel: formatDateTime(new Date()),
    filters: buildFilterLabels(filters),
    metrics: {
      total: metrics.total,
      updated: metrics.updated,
      completed: metrics.completed,
      onHold: metrics.onHold,
      stale: metrics.stale,
      noReportMemo: metrics.noReportMemo
    },
    statusRows,
    rankRows,
    categories: groupByCategory(updatedItems, categoryOrder),
    updatedItems
  }
}

function addHeaderBand(slide, subtitle) {
  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: 13.33,
    h: 0.5,
    line: { color: BRAND.navy, transparency: 100 },
    fill: { color: BRAND.navy }
  })

  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.75,
      y: 0.62,
      w: 4.4,
      h: 0.24,
      fontFace: "Meiryo",
      fontSize: 11,
      color: BRAND.blue,
      bold: true,
      margin: 0
    })
  }
}

function addSlideTitle(slide, title, subtitle) {
  addHeaderBand(slide, subtitle)

  slide.addText(title, {
    x: 0.75,
    y: subtitle ? 0.92 : 0.78,
    w: 8.8,
    h: 0.48,
    fontFace: "Meiryo",
    fontSize: 24,
    bold: true,
    color: BRAND.ink,
    margin: 0
  })
}

function addFooter(slide, text) {
  slide.addText(text, {
    x: 0.75,
    y: 7.0,
    w: 11.83,
    h: 0.2,
    fontFace: "Meiryo",
    fontSize: 9,
    color: BRAND.muted,
    align: "right",
    margin: 0
  })
}

function addMetricCard(slide, options) {
  const width = options.w || 2.8
  const height = options.h || 1.2

  slide.addShape("roundRect", {
    x: options.x,
    y: options.y,
    w: width,
    h: height,
    rectRadius: 0.08,
    line: { color: BRAND.border, pt: 1 },
    fill: { color: options.fillColor || "FFFFFF" }
  })

  slide.addText(options.label, {
    x: options.x + 0.2,
    y: options.y + 0.16,
    w: width - 0.4,
    h: 0.18,
    fontFace: "Meiryo",
    fontSize: options.labelFontSize || 10,
    color: BRAND.muted,
    margin: 0
  })

  slide.addText(String(options.value), {
    x: options.x + 0.2,
    y: options.y + 0.42,
    w: width - 0.4,
    h: 0.4,
    fontFace: "Meiryo",
    fontSize: options.valueFontSize || 24,
    bold: true,
    color: options.valueColor || BRAND.ink,
    margin: 0,
    fit: "shrink"
  })
}

function addBulletList(slide, title, lines, x, y, w, h) {
  slide.addShape("roundRect", {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    line: { color: BRAND.border, pt: 1 },
    fill: { color: BRAND.softAlt }
  })

  slide.addText(title, {
    x: x + 0.18,
    y: y + 0.14,
    w: w - 0.36,
    h: 0.2,
    fontFace: "Meiryo",
    fontSize: 12,
    bold: true,
    color: BRAND.ink,
    margin: 0
  })

  slide.addText(lines.length > 0 ? lines.join("\n") : "対象なし", {
    x: x + 0.18,
    y: y + 0.45,
    w: w - 0.36,
    h: h - 0.58,
    fontFace: "Meiryo",
    fontSize: 11,
    color: BRAND.ink,
    valign: "top",
    breakLine: false,
    margin: 0
  })
}

function addLegendCard(slide, options) {
  slide.addShape("roundRect", {
    x: options.x,
    y: options.y,
    w: options.w,
    h: options.h,
    rectRadius: 0.08,
    line: { color: BRAND.border, pt: 1 },
    fill: { color: "FFFFFF" }
  })

  slide.addShape("ellipse", {
    x: options.x + 0.16,
    y: options.y + 0.16,
    w: 0.16,
    h: 0.16,
    line: { color: options.color, transparency: 100 },
    fill: { color: options.color }
  })

  slide.addText(options.label, {
    x: options.x + 0.4,
    y: options.y + 0.11,
    w: options.w - 0.52,
    h: 0.18,
    fontFace: "Meiryo",
    fontSize: 10,
    color: BRAND.ink,
    bold: true,
    margin: 0
  })

  slide.addText(options.valueLabel, {
    x: options.x + 0.4,
    y: options.y + 0.42,
    w: options.w - 0.52,
    h: 0.28,
    fontFace: "Meiryo",
    fontSize: 16,
    color: BRAND.ink,
    margin: 0,
    fit: "shrink"
  })

  if (options.caption) {
    slide.addText(options.caption, {
      x: options.x + 0.4,
      y: options.y + 0.84,
      w: options.w - 0.52,
      h: 0.18,
      fontFace: "Meiryo",
      fontSize: 9,
      color: BRAND.muted,
      margin: 0
    })
  }
}

function addDonutChartPanel(slide, options) {
  slide.addShape("roundRect", {
    x: options.x,
    y: options.y,
    w: options.w,
    h: options.h,
    rectRadius: 0.08,
    line: { color: BRAND.border, pt: 1 },
    fill: { color: BRAND.softAlt }
  })

  slide.addText(options.title, {
    x: options.x + 0.22,
    y: options.y + 0.18,
    w: 2.8,
    h: 0.22,
    fontFace: "Meiryo",
    fontSize: 12,
    bold: true,
    color: BRAND.ink,
    margin: 0
  })

  slide.addText(options.subtitle, {
    x: options.x + 0.22,
    y: options.y + 0.46,
    w: 3.2,
    h: 0.18,
    fontFace: "Meiryo",
    fontSize: 9,
    color: BRAND.muted,
    margin: 0
  })

  if (options.rows.length > 0) {
    const chartX = options.x + 0.12
    const chartY = options.y + 0.78
    const chartSize = 2.3

    slide.addChart("doughnut", [{
      name: options.title,
      labels: options.rows.map((row) => row.label),
      values: options.rows.map((row) => row.value)
    }], {
      x: chartX,
      y: chartY,
      w: chartSize,
      h: chartSize,
      holeSize: 68,
      showLegend: false,
      showTitle: false,
      showValue: false,
      showPercent: false,
      chartColors: options.rows.map((row) => row.color),
      dataBorder: { pt: 1.25, color: "F7FAFC" },
      chartArea: { color: BRAND.softAlt, transparency: 100 },
      plotArea: { color: BRAND.softAlt, transparency: 100, border: { color: BRAND.softAlt, transparency: 100 } },
      layout: { x: 0, y: 0, w: 1, h: 1 }
    })
  } else {
    slide.addText("対象データなし", {
      x: options.x + 0.3,
      y: options.y + 1.85,
      w: 2.4,
      h: 0.2,
      fontFace: "Meiryo",
      fontSize: 11,
      color: BRAND.muted,
      align: "center",
      margin: 0
    })
  }

  slide.addText(options.centerTitle, {
    x: options.x + 0.52,
    y: options.y + 1.7,
    w: 1.55,
    h: 0.16,
    fontFace: "Meiryo",
    fontSize: 8,
    color: BRAND.muted,
    bold: true,
    align: "center",
    margin: 0
  })

  slide.addText(options.centerValue, {
    x: options.x + 0.34,
    y: options.y + 1.9,
    w: 2.0,
    h: 0.3,
    fontFace: "Meiryo",
    fontSize: 16,
    color: BRAND.ink,
    bold: true,
    align: "center",
    margin: 0,
    fit: "shrink"
  })

  if (options.centerCaption) {
    slide.addText(options.centerCaption, {
      x: options.x + 0.38,
      y: options.y + 2.26,
      w: 1.95,
      h: 0.16,
      fontFace: "Meiryo",
      fontSize: 8,
      color: BRAND.muted,
      align: "center",
      margin: 0
    })
  }

  const legendItems = options.rows.slice(0, options.maxLegendItems || 4)

  legendItems.forEach((row, index) => {
    const column = index % 2
    const rowIndex = Math.floor(index / 2)

    addLegendCard(slide, {
      x: options.x + 2.72 + (column * 1.46),
      y: options.y + 0.9 + (rowIndex * 1.08),
      w: 1.34,
      h: 0.98,
      label: row.label,
      valueLabel: options.legendValue(row),
      caption: options.legendCaption ? options.legendCaption(row) : "",
      color: row.color
    })
  })
}

function buildCategoryStatusText(statusCounts) {
  return Object.entries(statusCounts)
    .sort((left, right) => statusRank(left[0]) - statusRank(right[0]))
    .map(([status, count]) => `${status}: ${count}件`)
    .join(" / ")
}

function addCategorySummarySlide(pptx, report, categoryReport) {
  const slide = pptx.addSlide()
  addSlideTitle(slide, `${categoryReport.category} カテゴリ`, "カテゴリサマリ")

  slide.addText(report.periodLabel, {
    x: 8.9,
    y: 0.96,
    w: 3.68,
    h: 0.24,
    fontFace: "Meiryo",
    fontSize: 10,
    color: BRAND.muted,
    align: "right",
    margin: 0
  })

  addMetricCard(slide, { x: 0.75, y: 1.55, label: "期間内更新", value: categoryReport.total, fillColor: BRAND.soft })
  addMetricCard(slide, { x: 3.75, y: 1.55, label: "保留", value: categoryReport.statusCounts["保留"] || 0, valueColor: BRAND.danger })
  addMetricCard(slide, { x: 6.75, y: 1.55, label: "クローズ", value: categoryReport.statusCounts["クローズ"] || 0, valueColor: BRAND.success })
  addMetricCard(slide, { x: 9.75, y: 1.55, label: "報告メモ未入力", value: categoryReport.items.filter((item) => !String(item.reportMemo || "").trim()).length, valueColor: BRAND.warning })

  addBulletList(
    slide,
    "ステータス内訳",
    [buildCategoryStatusText(categoryReport.statusCounts) || "データなし"],
    0.75,
    3.15,
    12.0,
    1.25
  )

  addFooter(slide, `カテゴリ別に整理した定例報告書 | ${report.generatedAtLabel}`)
}

function addItemField(slide, label, value, x, y, w, h) {
  slide.addText(`${label}: ${value}`, {
    x,
    y,
    w,
    h,
    fontFace: "Meiryo",
    fontSize: 10,
    color: BRAND.ink,
    valign: "top",
    margin: 0
  })
}

function addItemCard(slide, item, index) {
  const column = index % 2
  const row = Math.floor(index / 2)
  const x = 0.75 + column * 6.05
  const y = 1.55 + row * 2.55
  const w = 5.4
  const h = 2.15

  slide.addShape("roundRect", {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    line: { color: BRAND.border, pt: 1 },
    fill: { color: "FFFFFF" }
  })

  slide.addText(sanitizeText(item.title), {
    x: x + 0.18,
    y: y + 0.14,
    w: w - 0.36,
    h: 0.24,
    fontFace: "Meiryo",
    fontSize: 13,
    bold: true,
    color: BRAND.ink,
    margin: 0,
    fit: "shrink"
  })

  slide.addText(
    `${formatKpiDisplay(item.kpiNumber)} / ${sanitizeText(item.assignee)} / ${sanitizeText(item.status)} / 更新 ${formatDateOnly(item.updatedAt)}`,
    {
      x: x + 0.18,
      y: y + 0.42,
      w: w - 0.36,
      h: 0.18,
      fontFace: "Meiryo",
      fontSize: 9,
      color: BRAND.blue,
      margin: 0,
      fit: "shrink"
    }
  )

  addItemField(slide, "内容", sanitizeText(item.content), x + 0.18, y + 0.7, w - 0.36, 0.42)
  addItemField(slide, "NextAction", sanitizeText(item.nextAction), x + 0.18, y + 1.12, w - 0.36, 0.34)
  addItemField(slide, "報告メモ", sanitizeText(item.reportMemo), x + 0.18, y + 1.46, w - 0.36, 0.42)

  if (item.category === "営業") {
    addItemField(
      slide,
      "顧客/ランク/ディールサイズ",
      `${sanitizeText(item.customer)} / ${sanitizeText(item.rank)} / ${formatDealSizeDisplay(item.dealSize)}`,
      x + 0.18,
      y + 1.8,
      w - 0.36,
      0.24
    )
  }
}

function addCategoryDetailSlides(pptx, report, categoryReport) {
  const itemPages = chunkItems(categoryReport.items, 4)

  itemPages.forEach((items, pageIndex) => {
    const slide = pptx.addSlide()
    const suffix = itemPages.length > 1 ? ` ${pageIndex + 1}/${itemPages.length}` : ""
    addSlideTitle(slide, `${categoryReport.category} 案件詳細${suffix}`, "カテゴリ詳細")

    slide.addText(
      `並び順: ステータス優先 / 更新日が新しい順`,
      {
        x: 8.2,
        y: 0.96,
        w: 4.4,
        h: 0.2,
        fontFace: "Meiryo",
        fontSize: 9,
        color: BRAND.muted,
        align: "right",
        margin: 0
      }
    )

    items.forEach((item, index) => addItemCard(slide, item, index))
    addFooter(slide, `${categoryReport.category} | ${report.periodLabel}`)
  })
}

function addEmptyStateSlide(pptx, report) {
  const slide = pptx.addSlide()
  addSlideTitle(slide, "期間内の更新案件はありません", "定例報告書")

  slide.addShape("roundRect", {
    x: 0.9,
    y: 1.8,
    w: 11.5,
    h: 2.0,
    rectRadius: 0.08,
    line: { color: BRAND.border, pt: 1 },
    fill: { color: BRAND.softAlt }
  })

  slide.addText(`対象期間: ${report.periodLabel}\n現在の一覧条件に一致する更新案件はありません。`, {
    x: 1.2,
    y: 2.35,
    w: 10.9,
    h: 0.8,
    fontFace: "Meiryo",
    fontSize: 18,
    color: BRAND.ink,
    bold: true,
    align: "center",
    valign: "mid",
    margin: 0
  })

  addFooter(slide, `カテゴリ別に整理した定例報告書 | ${report.generatedAtLabel}`)
}

export async function buildPowerPointArrayBuffer(report) {
  const pptx = new PptxGenJS()
  pptx.layout = "LAYOUT_WIDE"
  pptx.author = "GitHub Copilot"
  pptx.company = "Progress Tracker"
  pptx.subject = "定例報告書"
  pptx.title = `${report.title} ${report.periodLabel}`
  pptx.lang = "ja-JP"

  const titleSlide = pptx.addSlide()
  titleSlide.background = { color: "F7FBFD" }
  titleSlide.addShape("rect", {
    x: 0,
    y: 0,
    w: 13.33,
    h: 0.78,
    line: { color: BRAND.navy, transparency: 100 },
    fill: { color: BRAND.navy }
  })
  titleSlide.addText("定例報告書", {
    x: 0.75,
    y: 1.1,
    w: 5.8,
    h: 0.55,
    fontFace: "Meiryo",
    fontSize: 27,
    bold: true,
    color: BRAND.ink,
    margin: 0
  })
  titleSlide.addText(report.periodLabel, {
    x: 0.75,
    y: 1.72,
    w: 4.8,
    h: 0.24,
    fontFace: "Meiryo",
    fontSize: 13,
    color: BRAND.blue,
    bold: true,
    margin: 0
  })
  titleSlide.addShape("roundRect", {
    x: 0.75,
    y: 2.3,
    w: 5.95,
    h: 1.3,
    rectRadius: 0.08,
    line: { color: BRAND.border, pt: 1 },
    fill: { color: "FFFFFF" }
  })
  titleSlide.addText(`出力日時: ${report.generatedAtLabel}\n対象件数: ${report.metrics.updated}件`, {
    x: 1.0,
    y: 2.68,
    w: 5.4,
    h: 0.56,
    fontFace: "Meiryo",
    fontSize: 13,
    color: BRAND.ink,
    margin: 0
  })
  addBulletList(titleSlide, "現在の絞り込み条件", report.filters, 7.15, 1.2, 5.4, 2.4)
  addBulletList(
    titleSlide,
    "構成",
    ["・全体サマリ", "・カテゴリ別サマリ", "・カテゴリ別の案件詳細"],
    0.75,
    4.15,
    11.8,
    1.4
  )
  addFooter(titleSlide, "Progress Management | PowerPoint Export")

  const overviewSlide = pptx.addSlide()
  addSlideTitle(overviewSlide, "全体サマリ", "定例報告書")
  addMetricCard(overviewSlide, { x: 0.75, y: 1.6, w: 2.75, h: 1.02, label: "対象案件", value: report.metrics.total, fillColor: BRAND.softAlt, valueFontSize: 22 })
  addMetricCard(overviewSlide, { x: 3.65, y: 1.6, w: 2.75, h: 1.02, label: "期間内更新", value: report.metrics.updated, fillColor: BRAND.soft, valueFontSize: 22 })
  addMetricCard(overviewSlide, { x: 6.55, y: 1.6, w: 2.75, h: 1.02, label: "完了", value: report.metrics.completed, valueColor: BRAND.success, valueFontSize: 22 })
  addMetricCard(overviewSlide, { x: 9.45, y: 1.6, w: 2.75, h: 1.02, label: "保留", value: report.metrics.onHold, valueColor: BRAND.danger, valueFontSize: 22 })

  overviewSlide.addText(
    `更新なし ${report.metrics.stale}件 / 報告メモ未入力 ${report.metrics.noReportMemo}件 / カテゴリ ${report.categories.length}件 / 詳細対象 ${report.updatedItems.length}件`,
    {
      x: 0.9,
      y: 2.88,
      w: 11.45,
      h: 0.2,
      fontFace: "Meiryo",
      fontSize: 10,
      color: BRAND.muted,
      align: "center",
      margin: 0
    }
  )

  const totalRankUnits = report.rankRows.reduce((total, row) => total + row.value, 0)
  const totalRankCount = report.rankRows.reduce((total, row) => total + row.count, 0)

  addDonutChartPanel(overviewSlide, {
    x: 0.75,
    y: 3.22,
    w: 5.9,
    h: 3.45,
    title: "ステータス集計",
    subtitle: "期間内更新案件の件数構成",
    rows: report.statusRows,
    centerTitle: "更新案件",
    centerValue: `${report.metrics.updated}件`,
    centerCaption: `${report.statusRows.length}ステータス`,
    legendValue: (row) => `${row.value}件`,
    legendCaption: (row) => report.metrics.updated > 0 ? `${((row.value / report.metrics.updated) * 100).toFixed(1)}%` : "0.0%"
  })

  addDonutChartPanel(overviewSlide, {
    x: 6.7,
    y: 3.22,
    w: 5.85,
    h: 3.45,
    title: "ランク別",
    subtitle: "ディールサイズの構成比",
    rows: report.rankRows,
    centerTitle: "総ディールサイズ",
    centerValue: formatDealSizeUnitsLabel(totalRankUnits),
    centerCaption: `${totalRankCount}件 / ${report.rankRows.length}ランク`,
    legendValue: (row) => formatDealSizeUnitsLabel(row.value),
    legendCaption: (row) => totalRankUnits > 0 ? `${((row.value / totalRankUnits) * 100).toFixed(1)}%` : "0.0%"
  })
  addFooter(overviewSlide, `${report.periodLabel} | ${report.generatedAtLabel}`)

  if (report.updatedItems.length === 0) {
    addEmptyStateSlide(pptx, report)
  }

  report.categories.forEach((categoryReport) => {
    addCategorySummarySlide(pptx, report, categoryReport)
    addCategoryDetailSlides(pptx, report, categoryReport)
  })

  return pptx.write({ outputType: "arraybuffer" })
}

export function buildPowerPointFileName(report) {
  const periodToken = report.periodLabel.replace(/[\\/:*?"<>|\s]+/g, "_")
  return `定例報告書_${periodToken}.pptx`
}