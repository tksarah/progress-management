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

const FONT_JA = "Meiryo"
const FONT_LATIN = "Segoe UI"

const FRAME = {
  width: 13.33,
  marginX: 0.72,
  headerTop: 0.22,
  headerTitleY: 0.26,
  headerSubtitleY: 0.54,
  headerDividerY: 0.94,
  footerDividerY: 6.78,
  footerTextY: 6.92,
  footerPageX: 11.12,
  footerPageW: 1.46
}

const DETAIL_ITEMS_PER_PAGE = 2
const TOC_ENTRIES_PER_PAGE = 8

function isAsciiCharacter(character) {
  return character.charCodeAt(0) <= 0x7f
}

function buildMixedFontRuns(text) {
  const value = String(text ?? "")

  if (value.length === 0) {
    return [{ text: "", options: { fontFace: FONT_JA } }]
  }

  const runs = []
  let currentKind = null
  let buffer = ""

  for (const character of value) {
    const kind = isAsciiCharacter(character) ? "latin" : "ja"

    if (currentKind === null) {
      currentKind = kind
      buffer = character
      continue
    }

    if (kind === currentKind) {
      buffer += character
      continue
    }

    runs.push({
      text: buffer,
      options: { fontFace: currentKind === "latin" ? FONT_LATIN : FONT_JA }
    })

    buffer = character
    currentKind = kind
  }

  if (buffer.length > 0) {
    runs.push({
      text: buffer,
      options: { fontFace: currentKind === "latin" ? FONT_LATIN : FONT_JA }
    })
  }

  return runs
}

function addText(slide, text, options = {}) {
  if (Array.isArray(text)) {
    slide.addText(text, options)
    return
  }

  const value = String(text ?? "")
  const textOptions = { ...options }
  delete textOptions.fontFace

  const runs = buildMixedFontRuns(value)

  if (runs.length === 1) {
    slide.addText(runs[0].text, {
      ...textOptions,
      fontFace: runs[0].options.fontFace
    })
    return
  }

  slide.addText(runs, textOptions)
}

function formatPageNumber(pageNumber, totalPages) {
  const width = Math.max(String(totalPages).length, 2)
  return `${String(pageNumber).padStart(width, "0")} / ${String(totalPages).padStart(width, "0")}`
}

function addPanel(slide, options) {
  const accentColor = options.accentColor || BRAND.blue

  slide.addShape("roundRect", {
    x: options.x,
    y: options.y,
    w: options.w,
    h: options.h,
    rectRadius: options.rectRadius ?? 0.08,
    line: { color: options.borderColor || BRAND.border, pt: 1 },
    fill: { color: options.fillColor || "FFFFFF" }
  })

  slide.addShape("rect", {
    x: options.x,
    y: options.y,
    w: 0.12,
    h: options.h,
    line: { color: accentColor, transparency: 100 },
    fill: { color: accentColor }
  })
}

function applySlideBackground(slide, isCover = false) {
  slide.background = { color: BRAND.softAlt }

  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: FRAME.width,
    h: 0.12,
    line: { color: BRAND.navy, transparency: 100 },
    fill: { color: BRAND.navy }
  })

  slide.addShape("rect", {
    x: 0,
    y: 7.42,
    w: FRAME.width,
    h: 0.08,
    line: { color: BRAND.cyan, transparency: 100 },
    fill: { color: BRAND.cyan }
  })

  if (isCover) {
    slide.addShape("ellipse", {
      x: 10.1,
      y: 0.52,
      w: 3.1,
      h: 3.1,
      line: { color: BRAND.cyan, transparency: 100 },
      fill: { color: BRAND.cyan, transparency: 84 }
    })

    slide.addShape("ellipse", {
      x: -0.7,
      y: 5.65,
      w: 2.1,
      h: 2.1,
      line: { color: BRAND.blue, transparency: 100 },
      fill: { color: BRAND.blue, transparency: 88 }
    })
  }
}

function renderSlideFrame(slide, { title, subtitle, meta, footerLeft, pageNumber, totalPages, isCover = false }) {
  applySlideBackground(slide, isCover)

  slide.addShape("line", {
    x: FRAME.marginX,
    y: FRAME.headerDividerY,
    w: FRAME.width - (FRAME.marginX * 2),
    h: 0,
    line: { color: BRAND.border, pt: 1 }
  })

  slide.addShape("line", {
    x: FRAME.marginX,
    y: FRAME.footerDividerY,
    w: FRAME.width - (FRAME.marginX * 2),
    h: 0,
    line: { color: BRAND.border, pt: 1 }
  })

  slide.addShape("roundRect", {
    x: FRAME.marginX,
    y: FRAME.headerTop,
    w: 0.15,
    h: 0.54,
    rectRadius: 0.04,
    line: { color: BRAND.navy, transparency: 100 },
    fill: { color: BRAND.navy }
  })

  addText(slide, title, {
    x: FRAME.marginX + 0.3,
    y: FRAME.headerTitleY,
    w: 8.55,
    h: 0.34,
    fontSize: 24,
    bold: true,
    color: BRAND.ink,
    margin: 0,
    fit: "shrink"
  })

  if (meta) {
    addText(slide, meta, {
      x: 9.55,
      y: 0.3,
      w: 2.98,
      h: 0.18,
      fontSize: 10,
      color: BRAND.blue,
      align: "right",
      margin: 0,
      fit: "shrink"
    })
  }

  if (footerLeft) {
    addText(slide, footerLeft, {
      x: FRAME.marginX,
      y: FRAME.footerTextY,
      w: 9.5,
      h: 0.16,
      fontSize: 8.5,
      color: BRAND.muted,
      margin: 0,
      fit: "shrink"
    })
  }

  if (pageNumber) {
    addText(slide, String(pageNumber), {
      x: FRAME.footerPageX,
      y: 6.93,
      w: FRAME.footerPageW,
      h: 0.12,
      fontSize: 10,
      bold: true,
      color: BRAND.muted,
      align: "right",
      margin: 0,
      fit: "shrink"
    })
  }
}

function addMetricCard(slide, options) {
  const width = options.w || 2.8
  const height = options.h || 1.2

  addPanel(slide, {
    x: options.x,
    y: options.y,
    w: width,
    h: height,
    fillColor: options.fillColor || "FFFFFF",
    borderColor: options.borderColor || BRAND.border,
    accentColor: options.accentColor || options.valueColor || BRAND.blue
  })

  addText(slide, options.label, {
    x: options.x + 0.22,
    y: options.y + 0.14,
    w: width - 0.44,
    h: 0.18,
    fontSize: options.labelFontSize || 10,
    color: BRAND.muted,
    margin: 0,
    fit: "shrink"
  })

  addText(slide, String(options.value), {
    x: options.x + 0.22,
    y: options.y + 0.4,
    w: width - 0.44,
    h: 0.44,
    fontSize: options.valueFontSize || 24,
    bold: true,
    color: options.valueColor || BRAND.ink,
    margin: 0,
    fit: "shrink"
  })
}

function addBulletList(slide, title, lines, x, y, w, h, accentColor = BRAND.blue) {
  addPanel(slide, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fillColor: BRAND.softAlt,
    accentColor
  })

  addText(slide, title, {
    x: x + 0.22,
    y: y + 0.14,
    w: w - 0.44,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: BRAND.ink,
    margin: 0,
    fit: "shrink"
  })

  addText(slide, lines.length > 0 ? lines.join("\n") : "対象なし", {
    x: x + 0.22,
    y: y + 0.42,
    w: w - 0.44,
    h: h - 0.5,
    fontSize: 10.5,
    color: BRAND.ink,
    valign: "top",
    breakLine: false,
    margin: 0,
    fit: "shrink"
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

  addText(slide, options.label, {
    x: options.x + 0.4,
    y: options.y + 0.11,
    w: options.w - 0.52,
    h: 0.18,
    fontSize: 10,
    color: BRAND.ink,
    bold: true,
    margin: 0,
    fit: "shrink"
  })

  addText(slide, options.valueLabel, {
    x: options.x + 0.4,
    y: options.y + 0.42,
    w: options.w - 0.52,
    h: 0.28,
    fontSize: 16,
    color: BRAND.ink,
    margin: 0,
    fit: "shrink"
  })

  if (options.caption) {
    addText(slide, options.caption, {
      x: options.x + 0.4,
      y: options.y + 0.78,
      w: options.w - 0.52,
      h: 0.18,
      fontSize: 9,
      color: BRAND.muted,
      margin: 0,
      fit: "shrink"
    })
  }
}

function addDonutChartPanel(slide, options) {
  addPanel(slide, {
    x: options.x,
    y: options.y,
    w: options.w,
    h: options.h,
    rectRadius: 0.08,
    fillColor: BRAND.softAlt,
    accentColor: options.rows[0]?.color || BRAND.blue
  })

  addText(slide, options.title, {
    x: options.x + 0.22,
    y: options.y + 0.16,
    w: 2.8,
    h: 0.22,
    fontSize: 12,
    bold: true,
    color: BRAND.ink,
    margin: 0,
    fit: "shrink"
  })

  addText(slide, options.subtitle, {
    x: options.x + 0.22,
    y: options.y + 0.42,
    w: 3.2,
    h: 0.18,
    fontSize: 9,
    color: BRAND.muted,
    margin: 0,
    fit: "shrink"
  })

  if (options.rows.length > 0) {
    const chartX = options.x + 0.12
    const chartY = options.y + 0.74
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
    addText(slide, "対象データなし", {
      x: options.x + 0.3,
      y: options.y + 1.86,
      w: 2.4,
      h: 0.2,
      fontSize: 11,
      color: BRAND.muted,
      align: "center",
      margin: 0,
      fit: "shrink"
    })
  }

  addText(slide, options.centerTitle, {
    x: options.x + 0.52,
    y: options.y + 1.68,
    w: 1.55,
    h: 0.16,
    fontSize: 8,
    color: BRAND.muted,
    bold: true,
    align: "center",
    margin: 0,
    fit: "shrink"
  })

  addText(slide, options.centerValue, {
    x: options.x + 0.34,
    y: options.y + 1.88,
    w: 2.0,
    h: 0.3,
    fontSize: 16,
    color: BRAND.ink,
    bold: true,
    align: "center",
    margin: 0,
    fit: "shrink"
  })

  if (options.centerCaption) {
    addText(slide, options.centerCaption, {
      x: options.x + 0.38,
      y: options.y + 2.22,
      w: 1.95,
      h: 0.16,
      fontSize: 8,
      color: BRAND.muted,
      align: "center",
      margin: 0,
      fit: "shrink"
    })
  }

  const legendItems = options.rows.slice(0, options.maxLegendItems || 4)

  legendItems.forEach((row, index) => {
    const column = index % 2
    const rowIndex = Math.floor(index / 2)

    addLegendCard(slide, {
      x: options.x + 2.72 + (column * 1.46),
      y: options.y + 0.88 + (rowIndex * 1.08),
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

function buildContentSections(report) {
  const sections = [
    {
      type: "overview",
      title: "全体サマリ",
      subtitle: "期間内の更新状況",
      tocLabel: "全体サマリ",
      kindLabel: "概要",
      accentColor: BRAND.navy
    }
  ]

  if (report.updatedItems.length === 0) {
    sections.push({
      type: "empty",
      title: "更新案件なし",
      subtitle: "対象期間内に該当データがありません",
      tocLabel: "更新案件なし",
      kindLabel: "案内",
      accentColor: BRAND.warning
    })
    return sections
  }

  report.categories.forEach((categoryReport) => {
    const itemPages = chunkItems(categoryReport.items, DETAIL_ITEMS_PER_PAGE)

    sections.push({
      type: "categorySummary",
      title: `${categoryReport.category} カテゴリ`,
      subtitle: "カテゴリサマリ",
      tocLabel: `${categoryReport.category} サマリ`,
      kindLabel: "カテゴリ",
      accentColor: BRAND.blue,
      categoryReport
    })

    itemPages.forEach((items, pageIndex) => {
      sections.push({
        type: "categoryDetail",
        title: `${categoryReport.category} 案件詳細`,
        subtitle: "ステータス優先 / 更新日が新しい順",
        tocLabel: `${categoryReport.category} 詳細 ${pageIndex + 1}/${itemPages.length}`,
        kindLabel: "詳細",
        accentColor: BRAND.cyan,
        categoryReport,
        items,
        pageIndex,
        pageCount: itemPages.length
      })
    })
  })

  return sections
}

function buildSlidePlan(report) {
  const contentSections = buildContentSections(report)
  const tocPages = Math.max(1, Math.ceil(contentSections.length / TOC_ENTRIES_PER_PAGE))
  const slidePlan = [{
    type: "cover",
    title: report.title,
    subtitle: "表紙",
    accentColor: BRAND.navy
  }]

  for (let index = 0; index < tocPages; index += 1) {
    slidePlan.push({
      type: "toc",
      title: "目次",
      subtitle: tocPages > 1 ? `掲載内容 ${index + 1}/${tocPages}` : "掲載内容",
      tocPageIndex: index,
      tocPages,
      accentColor: BRAND.blue
    })
  }

  slidePlan.push(...contentSections)

  slidePlan.forEach((slide, index) => {
    slide.pageNumber = index + 1
  })

  const totalPages = slidePlan.length
  const tocEntries = contentSections.map((section) => ({
    label: section.tocLabel,
    pageNumber: section.pageNumber,
    kindLabel: section.kindLabel,
    accentColor: section.accentColor
  }))
  const tocChunks = chunkItems(tocEntries, TOC_ENTRIES_PER_PAGE)

  slidePlan.forEach((slide) => {
    slide.totalPages = totalPages

    if (slide.type === "toc") {
      slide.entries = tocChunks[slide.tocPageIndex] || []
    }
  })

  return slidePlan
}

function renderCoverSlide(slide, report, slideInfo) {
  renderSlideFrame(slide, {
    title: "定例報告書",
    meta: report.periodLabel,
    pageNumber: slideInfo.pageNumber,
    totalPages: slideInfo.totalPages,
    isCover: true
  })

  addPanel(slide, {
    x: 0.72,
    y: 1.18,
    w: 6.1,
    h: 5.0,
    fillColor: "FFFFFF",
    accentColor: BRAND.navy
  })

  addText(slide, "定例報告書", {
    x: 1.06,
    y: 1.52,
    w: 4.8,
    h: 0.42,
    fontSize: 26,
    bold: true,
    color: BRAND.ink,
    margin: 0,
    fit: "shrink"
  })

  addText(slide, report.periodLabel, {
    x: 1.06,
    y: 2.12,
    w: 4.8,
    h: 0.22,
    fontSize: 12,
    bold: true,
    color: BRAND.blue,
    margin: 0,
    fit: "shrink"
  })

  addMetricCard(slide, {
    x: 7.15,
    y: 1.28,
    w: 2.55,
    h: 1.0,
    label: "対象案件",
    value: report.metrics.total,
    fillColor: BRAND.softAlt,
    valueFontSize: 20,
    accentColor: BRAND.navy
  })

  addMetricCard(slide, {
    x: 9.96,
    y: 1.28,
    w: 2.55,
    h: 1.0,
    label: "期間内更新",
    value: report.metrics.updated,
    fillColor: BRAND.soft,
    valueFontSize: 20,
    accentColor: BRAND.blue
  })

  addMetricCard(slide, {
    x: 7.15,
    y: 2.5,
    w: 2.55,
    h: 1.0,
    label: "カテゴリ数",
    value: report.categories.length,
    fillColor: "FFFFFF",
    valueFontSize: 20,
    accentColor: BRAND.cyan
  })

  addMetricCard(slide, {
    x: 9.96,
    y: 2.5,
    w: 2.55,
    h: 1.0,
    label: "詳細対象",
    value: report.updatedItems.length,
    fillColor: "FFFFFF",
    valueFontSize: 20,
    accentColor: BRAND.success
  })

  addBulletList(
    slide,
    "出力日時",
    [report.generatedAtLabel],
    7.15,
    3.75,
    5.36,
    1.0,
    BRAND.warning
  )
}

function renderTocSlide(slide, report, slideInfo) {
  renderSlideFrame(slide, {
    title: "目次",
    meta: report.periodLabel,
    pageNumber: slideInfo.pageNumber,
    totalPages: slideInfo.totalPages
  })

  addText(slide, "掲載内容とページを一覧できます。", {
    x: FRAME.marginX,
    y: 1.1,
    w: 8.0,
    h: 0.18,
    fontSize: 10,
    color: BRAND.muted,
    margin: 0,
    fit: "shrink"
  })

  if (slideInfo.entries.length === 0) {
    addBulletList(slide, "目次", ["項目がありません"], 0.75, 1.42, 11.84, 1.0, BRAND.blue)
    return
  }

  slideInfo.entries.forEach((entry, index) => {
    const column = index % 2
    const row = Math.floor(index / 2)
    const x = 0.75 + (column * 6.06)
    const y = 1.4 + (row * 0.84)

    addPanel(slide, {
      x,
      y,
      w: 5.7,
      h: 0.68,
      fillColor: "FFFFFF",
      accentColor: entry.accentColor
    })

    addText(slide, entry.label, {
      x: x + 0.22,
      y: y + 0.11,
      w: 4.5,
      h: 0.18,
      fontSize: 11,
      bold: true,
      color: BRAND.ink,
      margin: 0,
      fit: "shrink"
    })

    addText(slide, entry.kindLabel, {
      x: x + 0.22,
      y: y + 0.37,
      w: 1.2,
      h: 0.12,
      fontSize: 8,
      color: BRAND.muted,
      margin: 0,
      fit: "shrink"
    })

    addText(slide, String(entry.pageNumber), {
      x: x + 5.04,
      y: y + 0.11,
      w: 0.44,
      h: 0.16,
      fontSize: 10,
      bold: true,
      color: entry.accentColor,
      align: "right",
      margin: 0,
      fit: "shrink"
    })
  })
}

function renderOverviewSlide(slide, report, slideInfo) {
  renderSlideFrame(slide, {
    title: "全体サマリ",
    meta: report.periodLabel,
    pageNumber: slideInfo.pageNumber,
    totalPages: slideInfo.totalPages
  })

  addMetricCard(slide, { x: 0.75, y: 1.56, w: 2.75, h: 1.02, label: "対象案件", value: report.metrics.total, fillColor: BRAND.softAlt, valueFontSize: 22, accentColor: BRAND.navy })
  addMetricCard(slide, { x: 3.65, y: 1.56, w: 2.75, h: 1.02, label: "期間内更新", value: report.metrics.updated, fillColor: BRAND.soft, valueFontSize: 22, accentColor: BRAND.blue })
  addMetricCard(slide, { x: 6.55, y: 1.56, w: 2.75, h: 1.02, label: "完了", value: report.metrics.completed, valueColor: BRAND.success, valueFontSize: 22, accentColor: BRAND.success })
  addMetricCard(slide, { x: 9.45, y: 1.56, w: 2.75, h: 1.02, label: "保留", value: report.metrics.onHold, valueColor: BRAND.danger, valueFontSize: 22, accentColor: BRAND.danger })

  addText(slide, `更新なし ${report.metrics.stale}件 / 報告メモ未入力 ${report.metrics.noReportMemo}件 / カテゴリ ${report.categories.length}件 / 詳細対象 ${report.updatedItems.length}件`, {
    x: 0.9,
    y: 2.82,
    w: 11.45,
    h: 0.2,
    fontSize: 10,
    color: BRAND.muted,
    align: "center",
    margin: 0,
    fit: "shrink"
  })

  const totalRankUnits = report.rankRows.reduce((total, row) => total + row.value, 0)
  const totalRankCount = report.rankRows.reduce((total, row) => total + row.count, 0)

  addDonutChartPanel(slide, {
    x: 0.75,
    y: 3.18,
    w: 5.9,
    h: 3.42,
    title: "ステータス集計",
    subtitle: "期間内更新案件の件数構成",
    rows: report.statusRows,
    centerTitle: "更新案件",
    centerValue: `${report.metrics.updated}件`,
    centerCaption: `${report.statusRows.length}ステータス`,
    legendValue: (row) => `${row.value}件`,
    legendCaption: (row) => report.metrics.updated > 0 ? `${((row.value / report.metrics.updated) * 100).toFixed(1)}%` : "0.0%"
  })

  addDonutChartPanel(slide, {
    x: 6.7,
    y: 3.18,
    w: 5.85,
    h: 3.42,
    title: "ランク別",
    subtitle: "ディールサイズの構成比",
    rows: report.rankRows,
    centerTitle: "総ディールサイズ",
    centerValue: formatDealSizeUnitsLabel(totalRankUnits),
    centerCaption: `${totalRankCount}件 / ${report.rankRows.length}ランク`,
    legendValue: (row) => formatDealSizeUnitsLabel(row.value),
    legendCaption: (row) => totalRankUnits > 0 ? `${((row.value / totalRankUnits) * 100).toFixed(1)}%` : "0.0%"
  })
}

function renderEmptyStateSlide(slide, report, slideInfo) {
  renderSlideFrame(slide, {
    title: "更新案件なし",
    meta: report.periodLabel,
    pageNumber: slideInfo.pageNumber,
    totalPages: slideInfo.totalPages
  })

  addPanel(slide, {
    x: 0.9,
    y: 1.8,
    w: 11.5,
    h: 2.0,
    rectRadius: 0.08,
    fillColor: BRAND.softAlt,
    accentColor: BRAND.warning
  })

  addText(slide, `対象期間: ${report.periodLabel}\n現在の一覧条件に一致する更新案件はありません。`, {
    x: 1.2,
    y: 2.34,
    w: 10.9,
    h: 0.8,
    fontSize: 18,
    color: BRAND.ink,
    bold: true,
    align: "center",
    valign: "mid",
    margin: 0,
    fit: "shrink"
  })
}

function renderCategorySummarySlide(slide, report, categoryReport, slideInfo) {
  renderSlideFrame(slide, {
    title: `${categoryReport.category} カテゴリ`,
    meta: report.periodLabel,
    pageNumber: slideInfo.pageNumber,
    totalPages: slideInfo.totalPages
  })

  addMetricCard(slide, { x: 0.75, y: 1.55, label: "期間内更新", value: categoryReport.total, fillColor: BRAND.soft, accentColor: BRAND.blue })
  addMetricCard(slide, { x: 3.75, y: 1.55, label: "保留", value: categoryReport.statusCounts["保留"] || 0, valueColor: BRAND.danger, accentColor: BRAND.danger })
  addMetricCard(slide, { x: 6.75, y: 1.55, label: "クローズ", value: categoryReport.statusCounts["クローズ"] || 0, valueColor: BRAND.success, accentColor: BRAND.success })
  addMetricCard(slide, { x: 9.75, y: 1.55, label: "報告メモ未入力", value: categoryReport.items.filter((item) => !String(item.reportMemo || "").trim()).length, valueColor: BRAND.warning, accentColor: BRAND.warning })

  addBulletList(
    slide,
    "ステータス内訳",
    [buildCategoryStatusText(categoryReport.statusCounts) || "データなし"],
    0.75,
    3.15,
    12.0,
    1.25,
    BRAND.blue
  )
}

function addItemField(slide, label, value, x, y, w, h, fontSize = 9) {
  addText(slide, `${label}: ${value}`, {
    x,
    y,
    w,
    h,
    fontSize,
    color: BRAND.ink,
    valign: "top",
    margin: 0,
    fit: "shrink"
  })
}

function addItemContentBlock(slide, item, x, y, w, h) {
  addPanel(slide, {
    x,
    y,
    w,
    h,
    rectRadius: 0.06,
    fillColor: BRAND.softAlt,
    accentColor: BRAND.cyan
  })

  addText(slide, "内容", {
    x: x + 0.16,
    y: y + 0.08,
    w: 0.5,
    h: 0.14,
    fontSize: 9,
    bold: true,
    color: BRAND.muted,
    margin: 0,
    fit: "shrink"
  })

  addText(slide, sanitizeText(item.content), {
    x: x + 0.16,
    y: y + 0.24,
    w: w - 0.32,
    h: h - 0.3,
    fontSize: 11,
    color: BRAND.ink,
    valign: "top",
    align: "left",
    margin: 0,
    fit: "shrink"
  })
}

function addItemCard(slide, item, index, layout = "grid") {
  const isStackedLayout = layout === "stacked"
  const column = index % 2
  const row = Math.floor(index / 2)
  const x = isStackedLayout ? 0.75 : 0.75 + column * 6.05
  const y = isStackedLayout ? 1.56 + index * 2.46 : 1.55 + row * 2.55
  const w = isStackedLayout ? 11.83 : 5.4
  const h = isStackedLayout ? 2.32 : 2.15

  addPanel(slide, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fillColor: "FFFFFF",
    accentColor: item.status === "保留" ? BRAND.danger : item.status === "クローズ" ? BRAND.success : BRAND.blue
  })

  addText(slide, sanitizeText(item.title), {
    x: x + 0.2,
    y: y + 0.14,
    w: w - 0.4,
    h: 0.28,
    fontSize: isStackedLayout ? 12 : 13,
    bold: true,
    color: BRAND.ink,
    margin: 0,
    fit: "shrink"
  })

  addText(slide, `${formatKpiDisplay(item.kpiNumber)} / ${sanitizeText(item.assignee)} / ${sanitizeText(item.status)} / 更新 ${formatDateOnly(item.updatedAt)}`, {
    x: x + 0.2,
    y: y + 0.46,
    w: w - 0.4,
    h: 0.18,
    fontSize: isStackedLayout ? 8.5 : 9,
    color: BRAND.blue,
    margin: 0,
    fit: "shrink"
  })

  const contentY = isStackedLayout ? y + 0.74 : y + 0.7
  const nextActionY = isStackedLayout ? y + 1.12 : y + 1.12
  const reportMemoY = isStackedLayout ? y + 1.44 : y + 1.46
  const customerY = isStackedLayout ? y + 1.78 : y + 1.8
  const contentH = isStackedLayout ? 0.34 : 0.42
  const nextActionH = isStackedLayout ? 0.3 : 0.34
  const reportMemoH = isStackedLayout ? 0.34 : 0.42
  const customerH = isStackedLayout ? 0.22 : 0.24

  if (isStackedLayout) {
    addItemField(slide, "NextAction", sanitizeText(item.nextAction), x + 0.2, nextActionY, 3.55, nextActionH)
    addItemField(slide, "報告メモ", sanitizeText(item.reportMemo), x + 0.2, reportMemoY, 3.55, reportMemoH)

    if (item.category === "営業") {
      addItemField(
        slide,
        "顧客/ランク/ディールサイズ",
        `${sanitizeText(item.customer)} / ${sanitizeText(item.rank)} / ${formatDealSizeDisplay(item.dealSize)}`,
        x + 0.2,
        customerY,
        3.55,
        customerH
      )
    }

    addItemContentBlock(slide, item, x + 3.95, y + 0.22, 7.66, 1.92)
    return
  }

  addItemField(slide, "内容", sanitizeText(item.content), x + 0.2, contentY, w - 0.4, contentH, 11)
  addItemField(slide, "NextAction", sanitizeText(item.nextAction), x + 0.2, nextActionY, w - 0.4, nextActionH)
  addItemField(slide, "報告メモ", sanitizeText(item.reportMemo), x + 0.2, reportMemoY, w - 0.4, reportMemoH)

  if (item.category === "営業") {
    addItemField(
      slide,
      "顧客/ランク/ディールサイズ",
      `${sanitizeText(item.customer)} / ${sanitizeText(item.rank)} / ${formatDealSizeDisplay(item.dealSize)}`,
      x + 0.2,
      customerY,
      w - 0.4,
      customerH
    )
  }
}

function renderCategoryDetailSlide(slide, report, categoryReport, slideInfo) {
  renderSlideFrame(slide, {
    title: `${categoryReport.category} 案件詳細`,
    meta: report.periodLabel,
    pageNumber: slideInfo.pageNumber,
    totalPages: slideInfo.totalPages
  })

  slideInfo.items.forEach((item, index) => addItemCard(slide, item, index, "stacked"))
}

export async function buildPowerPointArrayBuffer(report) {
  const pptx = new PptxGenJS()
  pptx.layout = "LAYOUT_WIDE"
  pptx.author = "GitHub Copilot"
  pptx.company = "Progress Tracker"
  pptx.subject = "定例報告書"
  pptx.title = `${report.title} ${report.periodLabel}`
  pptx.lang = "ja-JP"

  const slidePlan = buildSlidePlan(report)

  slidePlan.forEach((slideInfo) => {
    const slide = pptx.addSlide()

    if (slideInfo.type === "cover") {
      renderCoverSlide(slide, report, slideInfo)
      return
    }

    if (slideInfo.type === "toc") {
      renderTocSlide(slide, report, slideInfo)
      return
    }

    if (slideInfo.type === "overview") {
      renderOverviewSlide(slide, report, slideInfo)
      return
    }

    if (slideInfo.type === "empty") {
      renderEmptyStateSlide(slide, report, slideInfo)
      return
    }

    if (slideInfo.type === "categorySummary") {
      renderCategorySummarySlide(slide, report, slideInfo.categoryReport, slideInfo)
      return
    }

    if (slideInfo.type === "categoryDetail") {
      renderCategoryDetailSlide(slide, report, slideInfo.categoryReport, slideInfo)
    }
  })

  return pptx.write({ outputType: "arraybuffer" })
}

export function buildPowerPointFileName(report) {
  const periodToken = report.periodLabel.replace(/[\\/:*?"<>|\s]+/g, "_")
  return `定例報告書_${periodToken}.pptx`
}