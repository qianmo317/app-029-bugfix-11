/**
 * 报价单导出（规格书第 4.6 节）：
 * - PDF：走浏览器打印（打印样式见各页面 @media print），不引入需要编译的依赖；
 * - Excel：纯前端生成 .xls（HTML 表格 + Excel MIME），无需第三方库；
 * - 金额单位：整数「分」。
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

/**
 * HTML 转义：导出的 .xls 是 HTML 表格，所有文本写入前统一走这一套规则，
 * 否则材料名里的 & < > " 会拆散表格结构（整行挤进一格、后续列串位）。
 */
function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

/** CSV 单元格转义：工艺卡所有字段统一走这一套规则（含逗号/引号/换行时加引号，引号翻倍） */
function csvCell(s: string): string {
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s
}

function csvRow(cells: Array<string | number>): string {
  return cells.map((c) => csvCell(String(c))).join(',')
}

/** 导出 Excel（.xls，Excel/WPS 可直接打开） */
export function exportQuoteXls(project: Project, layout: LayoutResult, bom: BomResult, fontLabel: string, compare: CompareRow[]): void {
  const doc = buildQuoteDoc(project, layout, bom, fontLabel)
  const table = `
  <table border="1">
    <tr><th colspan="6">${esc(doc.title)}</th></tr>
    <tr><td>项目</td><td colspan="5">${esc(doc.projectName)}</td></tr>
    <tr><td>门头尺寸</td><td colspan="5">${esc(doc.panelText)}</td></tr>
    <tr><td>字体/排版</td><td colspan="5">${esc(doc.fontText)}</td></tr>
    <tr><td>排版结果</td><td colspan="5">${esc(doc.layoutText)}</td></tr>
    <tr><th>类别</th><th>规格/说明</th><th>数量</th><th>单位</th><th>单价(元)</th><th>金额(元)</th></tr>
    ${doc.rows
      .map(
        (r) =>
          `<tr><td>${esc(r.group)}</td><td>${esc(r.spec)}</td><td>${esc(r.qty)}</td><td>${esc(r.unit)}</td><td>${esc(
            r.unitPrice
          )}</td><td>${esc(r.amount)}</td></tr>`
      )
      .join('\n')}
    <tr><td colspan="5">合计</td><td>${esc(doc.total)}</td></tr>
    <tr><th colspan="6">工艺说明</th></tr>
    ${doc.notes.map((n) => `<tr><td colspan="6">${esc(n)}</td></tr>`).join('\n    ')}
    <tr><th colspan="6">多材质对照（元）</th></tr>
    <tr><th>材质</th><th>面板</th><th>LED+电源</th><th>配件+加工</th><th colspan="2">合计</th></tr>
    ${compare
      .map(
        (c) =>
          `<tr><td>${esc(c.name)}</td><td>${esc(yuan(c.panelCents))}</td><td>${esc(yuan(c.ledCents + c.psuCents))}</td><td>${esc(
            yuan(c.accessoryCents + c.laborCents)
          )}</td><td colspan="2">${esc(yuan(c.totalCents))}</td></tr>`
      )
      .join('\n    ')}
    <tr><td colspan="6">${esc(doc.footer)}</td></tr>
  </table>`
  const html = `<html><head><meta charset="utf-8"></head><body>${table}</body></html>`
  download(`${project.name || '招牌'}报价单.xls`, new Blob([`\ufeff${html}`], { type: 'application/vnd.ms-excel;charset=utf-8' }))
}

/** 导出工艺卡（CSV，供车间流转；PDF 走浏览器打印） */
export function exportProcessCardCsv(project: Project, layout: LayoutResult, bom: BomResult, fontLabel: string): void {
  const lines: string[] = []
  lines.push(csvRow(['招牌字工艺卡']))
  lines.push(csvRow(['项目', project.name]))
  lines.push(
    csvRow([
      '门头',
      `${project.layout.panel.wMm}×${project.layout.panel.hMm}mm 边框${project.layout.panel.frameMm}mm ${mountingLabel(project.layout.panel.mounting)}`
    ])
  )
  lines.push(csvRow(['字体', `${fontLabel} 字重${project.layout.settings.weight} 字号${layout.sizeMm}mm`]))
  lines.push(csvRow(['排版', `${alignLabel(project.layout.settings.align)} 占宽${layout.occupiedW}mm 占高${layout.occupiedH}mm`]))
  lines.push('')

  // 四段固定顺序：字形工艺分析 → 裁切清单 → 灯与电源 → 拼版；段内无数据也保留段标题与表头，不跳过
  lines.push(csvRow(['字形工艺分析']))
  lines.push(csvRow(['字符', '字号mm', '笔画块', '外轮廓周长mm', '最细笔画mm', '警告']))
  for (const g of layout.glyphs) {
    const outerMm = g.contours.filter((c) => !c.isHole).reduce((s, c) => s + c.perimeterMm, 0)
    lines.push(csvRow([g.char, g.sizeMm, g.strokeBlocks, outerMm.toFixed(1), g.minStrokeMm, g.warnings.join('；') || '—']))
  }
  lines.push('')

  lines.push(csvRow(['裁切清单']))
  lines.push(csvRow(['料件', '宽mm', '高mm', '数量']))
  for (const c of bom.cutList) {
    lines.push(csvRow([c.label, c.wMm, c.hMm, c.count]))
  }
  lines.push('')

  lines.push(csvRow(['灯与电源']))
  lines.push(csvRow(['模组规格', bom.module.spec]))
  lines.push(csvRow(['布点长度mm', bom.led.perimeterTotalMm]))
  lines.push(csvRow(['模组数量只', bom.led.modules]))
  lines.push(csvRow(['额定功率W', bom.led.ratedW]))
  lines.push(csvRow(['建议电源', bom.led.suggestedPsu]))
  if (bom.led.note) lines.push(csvRow(['备注', bom.led.note]))
  lines.push(csvRow(['逐字布点', '笔画块', '外轮廓周长mm', '模组数', '功率W']))
  for (const r of ledRows(layout.chars, project.led)) {
    lines.push(csvRow([r.char, r.blocks, r.outerPerimeterMm, r.modules, r.ratedW]))
  }
  lines.push('')

  lines.push(csvRow(['拼版']))
  lines.push(csvRow(['板材', bom.sheet.spec]))
  lines.push(csvRow(['板数', bom.nesting.sheetCount]))
  lines.push(csvRow(['利用率', `${(bom.nesting.utilization * 100).toFixed(1)}%`]))
  download(`${project.name || '招牌'}工艺卡.csv`, new Blob([`\ufeff${lines.join('\r\n')}`], { type: 'text/csv;charset=utf-8' }))
}