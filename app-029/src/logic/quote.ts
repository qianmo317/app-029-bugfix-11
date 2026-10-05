/**
 * 报价单导出（规格书第 4.6 节）：
 * - PDF：走浏览器打印（打印样式见各页面 @media print），不引入需要编译的依赖；
 * - Excel：纯前端生成 .xls（HTML 表格 + Excel MIME），无需第三方库；
 * - 工艺卡：CSV（RFC 4180 转义），Excel/WPS 可直接打开；
 * - 金额单位：整数「分」。
 *
 * 所有写入单据的文本一律先过转义：
 * - XLS 是 HTML 文档，文本走 escapeHtml（& < >），否则材料名里的尖括号会被当成标签，整行挤进一格、后续列串位；
 * - CSV 走 csvCell（含逗号 / 引号 / 换行时加双引号，内部双引号翻倍）。
 */

import type { BomResult, CompareRow } from './materials'
import { yuan } from './materials'
import type { LayoutResult } from './layout'
import { alignLabel, mountingLabel } from './layout'
import { ledRows } from './led'
import type { Project } from './types'

export function bomGroupLabel(kind: string): string {
  switch (kind) {
    case 'acrylic':
      return '面板材料'
    case 'led_module':
      return 'LED 模组'
    case 'psu':
      return '电源'
    case 'glue':
      return '胶与配件'
    default:
      return '加工费'
  }
}

export interface QuoteDoc {
  title: string
  projectName: string
  date: string
  validUntil: string
  panelText: string
  fontText: string
  layoutText: string
  rows: Array<{ group: string; spec: string; qty: string; unit: string; unitPrice: string; amount: string }>
  total: string
  notes: string[]
  footer: string
}

export function buildQuoteDoc(project: Project, layout: LayoutResult, bom: BomResult, fontLabel: string): QuoteDoc {
  const now = new Date()
  const valid = new Date(now.getTime() + 30 * 24 * 3600 * 1000)
  const fmt = (d: Date): string => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
  const rows = bom.materials.map((m) => ({
    group: bomGroupLabel(m.kind),
    spec: m.spec,
    qty: String(m.qty),
    unit: m.unit,
    unitPrice: yuan(m.unitPriceCents),
    amount: yuan(m.amountCents)
  }))
  return {
    title: '招牌字制作报价单',
    projectName: project.name,
    date: fmt(now),
    validUntil: fmt(valid),
    panelText: `${project.layout.panel.wMm}×${project.layout.panel.hMm}mm（边框 ${project.layout.panel.frameMm}mm，${mountingLabel(
      project.layout.panel.mounting
    )}）`,
    fontText: `${fontLabel}　字重 ${project.layout.settings.weight}　字号 ${layout.sizeMm}mm（${alignLabel(project.layout.settings.align)}）`,
    layoutText: `占宽 ${layout.occupiedW}mm × 占高 ${layout.occupiedH}mm；左右留边 ${layout.margins.left}/${layout.margins.right}mm；视觉间距极差 ${layout.gapSpread}mm`,
    rows,
    total: yuan(bom.totalCents),
    notes: [
      `面板材料：${bom.panelMaterial.name}（${bom.panelMaterial.desc}）`,
      `亚克力拼版：${bom.nesting.sheetCount} 张 ${bom.sheet.spec}，利用率 ${(bom.nesting.utilization * 100).toFixed(1)}%`,
      `LED：布点长度 ${bom.led.perimeterTotalMm}mm，模组 ${bom.led.modules} 只，额定功率 ${bom.led.ratedW}W，建议电源 ${bom.led.suggestedPsu}`,
      bom.led.note
    ].filter((s) => !!s),
    footer: '本报价基于当前材料单价，有效期 30 天；含材料与加工费，不含安装与运输。'
  }
}

function download(filename: string, blob: Blob): void {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  setTimeout(() => URL.revokeObjectURL(url), 2000)
}

/** XLS（HTML）文本转义：& 必须最先替换 */
function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function td(value: string | number, tag: 'td' | 'th' = 'td', colspan?: number): string {
  const span = colspan ? ` colspan="${colspan}"` : ''
  return `<${tag}${span}>${escapeHtml(String(value))}</${tag}>`
}

/** CSV 单元格转义（RFC 4180）：含逗号、双引号或换行时整段加双引号，内部双引号翻倍 */
function csvCell(value: string | number): string {
  const s = String(value)
  if (/[",\r\n]/.test(s)) {
    return `"${s.replace(/"/g, '""')}"`
  }
  return s
}

function csvRow(cells: Array<string | number>): string {
  return cells.map(csvCell).join(',')
}

/** 导出 Excel（.xls，Excel/WPS 可直接打开） */
export function exportQuoteXls(project: Project, layout: LayoutResult, bom: BomResult, fontLabel: string, compare: CompareRow[]): void {
  const doc = buildQuoteDoc(project, layout, bom, fontLabel)
  const table = `
  <table border="1">
    <tr>${td(doc.title, 'th', 6)}</tr>
    <tr>${td('项目')}${td(doc.projectName, 'td', 5)}</tr>
    <tr>${td('门头尺寸')}${td(doc.panelText, 'td', 5)}</tr>
    <tr>${td('字体/排版')}${td(doc.fontText, 'td', 5)}</tr>
    <tr>${td('排版结果')}${td(doc.layoutText, 'td', 5)}</tr>
    <tr>${td('类别', 'th')}${td('规格/说明', 'th')}${td('数量', 'th')}${td('单位', 'th')}${td('单价(元)', 'th')}${td(
      '金额(元)',
      'th'
    )}</tr>
    ${doc.rows.map((r) => `<tr>${td(r.group)}${td(r.spec)}${td(r.qty)}${td(r.unit)}${td(r.unitPrice)}${td(r.amount)}</tr>`).join('\n')}
    <tr>${td('合计', 'td', 5)}${td(doc.total)}</tr>
    <tr>${td('工艺说明', 'th', 6)}</tr>
    ${doc.notes.map((n) => `<tr>${td(n, 'td', 6)}</tr>`).join('\n')}
    <tr>${td(`多材质对照（共 ${compare.length} 种材质，单位：元）`, 'th', 6)}</tr>
    <tr>${td('材质', 'th')}${td('面板', 'th')}${td('LED+电源', 'th')}${td('配件+加工', 'th')}${td('合计', 'th')}</tr>
    ${compare
      .map(
        (c) =>
          `<tr>${td(c.name)}${td(yuan(c.panelCents))}${td(yuan(c.ledCents + c.psuCents))}${td(
            yuan(c.accessoryCents + c.laborCents)
          )}${td(yuan(c.totalCents))}</tr>`
      )
      .join('\n')}
    <tr>${td(doc.footer, 'td', 6)}</tr>
  </table>`
  const html = `<html><head><meta charset="utf-8"></head><body>${table}</body></html>`
  download(`${project.name || '招牌'}报价单.xls`, new Blob(['﻿' + html], { type: 'application/vnd.ms-excel;charset=utf-8' }))
}

/** 导出工艺卡（CSV，供车间流转；PDF 走浏览器打印）。四段固定顺序，缺段也不跳过。 */
export function exportProcessCardCsv(project: Project, layout: LayoutResult, bom: BomResult, fontLabel: string): void {
  const lines: string[] = []
  const st = project.layout.settings
  lines.push(csvRow(['招牌字工艺卡']))
  lines.push(csvRow(['项目', project.name]))
  lines.push(csvRow(['门头', `${project.layout.panel.wMm}×${project.layout.panel.hMm}mm 边框${project.layout.panel.frameMm}mm`]))
  lines.push(csvRow(['安装方式', mountingLabel(project.layout.panel.mounting)]))
  lines.push(csvRow(['字体', `${fontLabel} 字重${st.weight} 字号${layout.sizeMm}mm`]))
  lines.push(
    csvRow([
      '排版',
      `${alignLabel(st.align)} 占宽${layout.occupiedW}mm 占高${layout.occupiedH}mm 视觉间距极差${layout.gapSpread}mm`
    ])
  )
  lines.push('')

  // 第 1 段：字形工艺分析（逐字表，车间据此分件）
  lines.push(csvRow(['字形工艺分析']))
  lines.push(csvRow(['字符', '字号 mm', '笔画块', '外轮廓周长 mm', '最细笔画 mm', '警告']))
  for (const g of layout.glyphs) {
    const outerPerimeter = g.contours
      .filter((c) => !c.isHole)
      .reduce((sum, c) => sum + c.perimeterMm, 0)
      .toFixed(1)
    lines.push(
      csvRow([
        g.char,
        g.sizeMm,
        g.strokeBlocks,
        outerPerimeter,
        g.minStrokeMm,
        g.warnings.join('；') || '—'
      ])
    )
  }
  lines.push('')

  // 第 2 段：裁切清单（同尺寸合并计数）
  lines.push(csvRow(['裁切清单']))
  lines.push(csvRow(['料件', '宽 mm', '高 mm', '数量']))
  if (bom.cutList.length === 0) {
    lines.push(csvRow(['（无）', '', '', '']))
  } else {
    for (const c of bom.cutList) {
      lines.push(csvRow([c.label, c.wMm, c.hMm, c.count]))
    }
  }
  lines.push('')

  // 第 3 段：灯与电源（每字灯数 + 总电源台数，车间据此分区领料）
  lines.push(csvRow(['灯与电源']))
  lines.push(csvRow(['布点总长度 mm', bom.led.perimeterTotalMm]))
  lines.push(csvRow(['模组总数（只）', bom.led.modules]))
  lines.push(csvRow(['额定功率 W', bom.led.ratedW]))
  lines.push(csvRow(['建议电源', bom.led.suggestedPsu]))
  lines.push(csvRow(['单台电源功率 W', bom.led.psuUnitW]))
  lines.push(csvRow(['电源数量（台）', bom.led.psuCount]))
  if (bom.led.note) lines.push(csvRow(['说明', bom.led.note]))
  lines.push(csvRow(['字符', '笔画块', '外轮廓周长 mm', '灯数（只）', '额定功率 W']))
  for (const r of ledRows(layout.chars, project.led)) {
    lines.push(csvRow([r.char, r.blocks, r.outerPerimeterMm, r.modules, r.ratedW]))
  }
  lines.push('')

  // 第 4 段：拼版（板材、板数、利用率，固定在最后）
  lines.push(csvRow(['亚克力拼版']))
  lines.push(csvRow(['板材', bom.sheet.spec]))
  lines.push(csvRow(['板数（张）', bom.nesting.sheetCount]))
  lines.push(csvRow(['利用率', `${(bom.nesting.utilization * 100).toFixed(1)}%`]))
  lines.push(csvRow(['料件总数', bom.nesting.pieceCount]))
  if (bom.nesting.oversize.length > 0) {
    lines.push(csvRow(['超板料件', bom.nesting.oversize.map((p) => `${p.label} ${p.wMm}×${p.hMm}mm`).join('；')]))
  }

  download(`${project.name || '招牌'}工艺卡.csv`, new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }))
}
