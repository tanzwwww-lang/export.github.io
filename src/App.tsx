import { useMemo, useState, useEffect } from 'react'
import { bitable, FieldType } from '@lark-base-open/js-sdk'

async function getActiveTable() {
  return await bitable.base.getActiveTable()
}

async function getActiveView(table: any) {
  try {
    if (typeof table.getActiveView === 'function') return await table.getActiveView()
  } catch {}
  return null
}

async function getFieldIds(table: any, view?: any) {
  try {
    if (view && typeof view.getVisibleFieldIdList === 'function') {
      return await view.getVisibleFieldIdList()
    }
  } catch {}
  return await table.getFieldIdList()
}

async function getRecordIds(table: any, view?: any) {
  try {
    if (view && typeof view.getVisibleRecordIdList === 'function') {
      return await view.getVisibleRecordIdList()
    }
  } catch {}
  return await table.getRecordIdList()
}

async function getFieldName(table: any, fieldId: string) {
  const field = await table.getField(fieldId)
  if (typeof field.getName === 'function') return await field.getName()
  return fieldId
}

function normalize(val: any): any {
  if (val === undefined || val === null) return ''
  if (Array.isArray(val)) return val.map(v => normalize(v)).join(',')
  if (typeof val === 'object') {
    if ('text' in val && typeof (val as any).text === 'string') return (val as any).text
    try {
      return JSON.stringify(val)
    } catch {
      return String(val)
    }
  }
  return val
}

async function exportExcel(filename: string, onStatus?: (msg: string) => void, onProgress?: (done: number, total: number) => void, tableId?: string, viewId?: string, insertImages?: boolean) {
  let table: any
  if (tableId) {
    table = await bitable.base.getTableById(tableId)
  } else {
    try {
      table = await getActiveTable()
    } catch {
      const list = await bitable.base.getTableList()
      table = list && list[0]
      if (!table) throw new Error('无法获取数据表')
    }
  }
  let view: any = null
  if (viewId) {
    try {
      const vlist = await table.getViewList()
      view = (vlist || []).find((v: any) => v.id === viewId) || null
    } catch {}
  }
  if (!view) {
    try {
      view = await getActiveView(table)
    } catch {}
  }
  const fieldIds: string[] = await getFieldIds(table, view)
  const recordIds: string[] = await getRecordIds(table, view)
  onProgress?.(0, recordIds.length)
  onStatus?.('收集字段与记录')

  const fieldNameMap = new Map<string, string>()
  for (const fid of fieldIds) {
    const name = await getFieldName(table, fid)
    fieldNameMap.set(fid, name)
  }

  const attachMeta = await table.getFieldMetaListByType(FieldType.Attachment)
  const attachmentFieldIds = new Set<string>((attachMeta || []).map((m: any) => m.id))

  const fieldObjMap = new Map<string, any>()
  for (const fid of fieldIds) {
    const field = await table.getField(fid)
    fieldObjMap.set(fid, field)
  }

  type ColPlan = { fid: string, header: string, attachment: boolean, slot?: number }
  const colPlan: ColPlan[] = []
  for (const fid of fieldIds) {
    const headerBase = fieldNameMap.get(fid) || fid
    if (!attachmentFieldIds.has(fid)) {
      colPlan.push({ fid, header: headerBase, attachment: false })
    }
  }

  onStatus?.('加载 Excel 库')
  let ExcelJS: any = (globalThis as any).ExcelJS
  if (!ExcelJS) {
    try {
      await new Promise<void>((resolve, reject) => {
        const s = document.createElement('script')
        s.src = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js'
        s.async = true
        s.onload = () => resolve()
        s.onerror = () => reject(new Error('CDN 加载失败'))
        document.head.appendChild(s)
      })
      ExcelJS = (globalThis as any).ExcelJS
    } catch {}
  }
  if (!ExcelJS) {
    try {
      const ExcelMod: any = await import('exceljs/dist/exceljs.min.js')
      ExcelJS = ExcelMod?.ExcelJS || ExcelMod?.default || ExcelMod
    } catch {}
  }
  if (!ExcelJS || !ExcelJS.Workbook) {
    onStatus?.('Excel 库无效：缺少 Workbook')
    throw new Error('ExcelJS.Workbook 未找到')
  }
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Sheet1')
  ws.addRow(colPlan.map(p => p.header))
  ws.getRow(1).font = { bold: true }

  const attachCount: Record<string, number> = {}
  function findColIndex(fid: string, slot?: number): number {
    for (let i = 0; i < colPlan.length; i++) {
      const p = colPlan[i]
      if (p.fid === fid && (!!p.attachment === !!(typeof slot === 'number')) && (typeof slot !== 'number' || p.slot === slot)) return i + 1
    }
    return -1
  }
  function ensureAttachmentSlot(fid: string, slotIndex: number, headerBase: string): number {
    const current = attachCount[fid] || 0
    if (slotIndex >= current) {
      for (let i = current; i <= slotIndex; i++) {
        const header = `${headerBase}(${i + 1})`
        colPlan.push({ fid, header, attachment: true, slot: i })
        const colIdx = colPlan.length
        ws.getRow(1).getCell(colIdx).value = header
        ws.getRow(1).font = { bold: true }
      }
      attachCount[fid] = slotIndex + 1
    }
    return findColIndex(fid, slotIndex)
  }
  async function runBatches(tasks: (() => Promise<void>)[], size: number) {
    for (let i = 0; i < tasks.length; i += size) {
      const slice = tasks.slice(i, i + size)
      await Promise.all(slice.map(fn => fn()))
    }
  }

  for (let r = 0; r < recordIds.length; r++) {
    const rid = recordIds[r]
    const rowIndex = r + 2
    const rowValues: any[] = new Array(colPlan.length).fill('')
    let targetRowHeight = 20
    if ((r + 1) % 10 === 0 || r === recordIds.length - 1) {
      onProgress?.(r + 1, recordIds.length)
      onStatus?.(`写入 ${r + 1}/${recordIds.length}`)
    }

    const valMap: Record<string, any> = {}
    await Promise.all(fieldIds.map(async fid => {
      if (!attachmentFieldIds.has(fid)) return
      const field = fieldObjMap.get(fid)
      valMap[fid] = await field.getValue(rid)
    }))

    const textTasks: Promise<void>[] = []
    for (const fid of fieldIds) {
      if (attachmentFieldIds.has(fid)) continue
      const idx = findColIndex(fid)
      const c = idx - 1
      textTasks.push((async () => {
        try {
          const s = await fieldObjMap.get(fid).getCellString(rid)
          rowValues[c] = s ?? ''
        } catch {
          rowValues[c] = ''
        }
      })())
    }
    await Promise.all(textTasks)

    const imgTasks: (() => Promise<void>)[] = []
    for (const fid of fieldIds) {
      if (!attachmentFieldIds.has(fid)) continue
      const headerBase = fieldNameMap.get(fid) || fid
      const val = valMap[fid]
      const arr = Array.isArray(val) ? val : (val ? [val] : [])
      if (arr.length > 0) {
        ensureAttachmentSlot(fid, arr.length - 1, headerBase)
        if (rowValues.length < colPlan.length) {
          const extra = colPlan.length - rowValues.length
          for (let x = 0; x < extra; x++) rowValues.push('')
        }
      }
      if (!insertImages) {
        for (let i = 0; i < arr.length; i++) {
          const item = arr[i]
          const colIndex = findColIndex(fid, i)
          const c = colIndex - 1
          if (colIndex < 1) continue
          rowValues[c] = String(item?.name || '')
        }
        continue
      }
      if (arr.length > 0) {
        const tokenIdxs: number[] = []
        const tokenList: string[] = []
        for (let i = 0; i < arr.length; i++) {
          const t = arr[i]?.token
          if (t) { tokenIdxs.push(i); tokenList.push(t) }
        }
        imgTasks.push(async () => {
          let urls: (string | undefined)[] = []
          try {
            const u = await table.getCellAttachmentUrls(tokenList, fid, rid)
            urls = Array.isArray(u) ? u : []
          } catch {}
          const subs: (() => Promise<void>)[] = arr.map((item, i) => async () => {
            const colIndex = findColIndex(fid, i)
            const c = colIndex - 1
            const fileName = String(item?.name || '')
            if (colIndex < 1) return
            const idxInTokens = tokenIdxs.indexOf(i)
            let url = idxInTokens >= 0 ? urls[idxInTokens] : undefined
            if (!url) {
              const tk = item?.token
              if (tk) {
                try {
                  const single = await table.getCellAttachmentUrls([tk], fid, rid)
                  url = single && single[0]
                } catch {}
              }
            }
            if (!url) {
              ws.getRow(rowIndex).getCell(colIndex).value = fileName
              return
            }
            try {
              const res = await fetch(url)
              const blob = await res.blob()
              if (!blob.type.startsWith('image/')) {
                ws.getRow(rowIndex).getCell(colIndex).value = fileName
                return
              }
              const initialDataUrl = await new Promise<string>((resolve, reject) => {
                const fr = new FileReader()
                fr.onload = () => resolve(String(fr.result))
                fr.onerror = reject
                fr.readAsDataURL(blob)
              })
              let base64 = initialDataUrl.includes(',') ? initialDataUrl.split(',')[1] : initialDataUrl
              let ext = 'png'
              const contentType = blob.type || ''
              if (contentType.includes('png')) {
                ext = 'png'
              } else if (contentType.includes('jpeg') || contentType.includes('jpg')) {
                ext = 'jpeg'
              } else {
                try {
                  const img = await new Promise<HTMLImageElement>((resolve) => {
                    const im = new Image()
                    im.onload = () => resolve(im)
                    im.src = initialDataUrl
                  })
                  const canvas = document.createElement('canvas')
                  canvas.width = img.naturalWidth || img.width || 1
                  canvas.height = img.naturalHeight || img.height || 1
                  const ctx = canvas.getContext('2d')
                  if (ctx) ctx.drawImage(img, 0, 0)
                  const pngUrl = canvas.toDataURL('image/png')
                  base64 = pngUrl.includes(',') ? pngUrl.split(',')[1] : pngUrl
                  ext = 'png'
                } catch {}
              }
              const imgId = wb.addImage({ base64, extension: ext as any })
              const imgWidth = 120
              const imgHeight = 90
              ws.getColumn(colIndex).width = Math.max(ws.getColumn(colIndex).width || 15, Math.ceil(imgWidth / 7))
              targetRowHeight = Math.max(targetRowHeight, imgHeight)
              ws.addImage(imgId, { tl: { col: colIndex - 1, row: rowIndex - 1 }, ext: { width: imgWidth, height: imgHeight } })
              ws.getRow(rowIndex).getCell(colIndex).value = ''
            } catch {
              ws.getRow(rowIndex).getCell(colIndex).value = fileName
            }
          })
          const innerSize = subs.length <= 4 ? 4 : subs.length <= 12 ? 3 : 2
          await runBatches(subs, innerSize)
        })
      }
    }
    // 先插入行，确保图片锚点行已存在，避免偏移到上一行/表头
    ws.addRow(rowValues)
    if (imgTasks.length) {
      const totalAttachments = Object.values(valMap).reduce((acc, v: any) => acc + (Array.isArray(v) ? v.length : v ? 1 : 0), 0)
      const outerSize = totalAttachments <= 6 ? 6 : totalAttachments <= 12 ? 4 : totalAttachments <= 30 ? 3 : 2
      await runBatches(imgTasks, outerSize)
    }
    ws.getRow(rowIndex).height = targetRowHeight
  }

  const safe = filename && filename.trim().length ? filename.trim() : '导出.xlsx'
  const finalName = safe.endsWith('.xlsx') ? safe : `${safe}.xlsx`
  onStatus?.('生成文件')
  const buffer = await wb.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = finalName
  document.body.appendChild(a)
  a.click()
  a.remove()
  URL.revokeObjectURL(url)
}

async function loadDocx(): Promise<any> {
  let docx: any = (globalThis as any).docx
  if (docx) return docx
  try {
    const mod: any = await import(/* @vite-ignore */ 'docx')
    if (mod) return mod
  } catch {}
  const urls = [
    'https://cdn.jsdelivr.net/npm/docx@9.5.1/build/index.min.js',
    'https://cdn.jsdelivr.net/npm/docx@8.6.0/build/index.min.js',
    'https://unpkg.com/docx@9.5.1/build/index.min.js',
    'https://unpkg.com/docx@8.6.0/build/index.min.js',
  ]
  for (const u of urls) {
    try {
      await new Promise<void>((resolve, reject) => {
        const s = document.createElement('script')
        let done = false
        const timer = setTimeout(() => {
          if (done) return
          done = true
          reject(new Error('docx 加载超时'))
        }, 8000)
        s.src = u
        s.async = true
        s.onload = () => {
          if (done) return
          done = true
          clearTimeout(timer)
          resolve()
        }
        s.onerror = () => {
          if (done) return
          done = true
          clearTimeout(timer)
          reject(new Error('docx 加载失败'))
        }
        document.head.appendChild(s)
      })
      docx = (globalThis as any).docx
      if (docx) return docx
    } catch {}
  }
  return (globalThis as any).docx
}

async function exportWord(filename: string, onStatus?: (msg: string) => void, onProgress?: (done: number, total: number) => void, tableId?: string, viewId?: string, insertImages?: boolean, pageOrientation?: 'portrait' | 'landscape', showIndex?: boolean) {
  let table: any
  if (tableId) {
    table = await bitable.base.getTableById(tableId)
  } else {
    try {
      table = await getActiveTable()
    } catch {
      const list = await bitable.base.getTableList()
      table = list && list[0]
      if (!table) throw new Error('无法获取数据表')
    }
  }
  let view: any = null
  if (viewId) {
    try {
      const vlist = await table.getViewList()
      view = (vlist || []).find((v: any) => v.id === viewId) || null
    } catch {}
  }
  if (!view) {
    try {
      view = await getActiveView(table)
    } catch {}
  }
  const fieldIds: string[] = await getFieldIds(table, view)
  const recordIds: string[] = await getRecordIds(table, view)
  onProgress?.(0, recordIds.length)
  onStatus?.('收集字段与记录')
  const fieldNameMap = new Map<string, string>()
  const fieldObjMap = new Map<string, any>()
  for (const fid of fieldIds) {
    const name = await getFieldName(table, fid)
    fieldNameMap.set(fid, name)
    const field = await table.getField(fid)
    fieldObjMap.set(fid, field)
  }
  const attachMeta = await table.getFieldMetaListByType(FieldType.Attachment)
  const attachmentFieldIds = new Set<string>((attachMeta || []).map((m: any) => m.id))
  onStatus?.('加载 Word 库')
  const docx = await loadDocx()
  if (docx) {
    const displayFids: string[] = (showIndex === false) ? fieldIds : ['__INDEX__', ...fieldIds]
    const headers = displayFids.map(fid => fid === '__INDEX__' ? '序号' : String(fieldNameMap.get(fid) || fid))
    const canvas = document.createElement('canvas')
    const ctx = canvas.getContext('2d')
    if (ctx) {
      ctx.font = '16px "Noto Sans SC","Microsoft YaHei",Arial,sans-serif'
    }
    const minColPx = 60
    function measureTextPx(t: string): number {
      if (!ctx) return Math.max(minColPx, (t || '').length * 8 + 12)
      const w = ctx.measureText(t || '').width
      return Math.max(minColPx, Math.ceil(w + 12))
    }
    const desired: number[] = headers.map(h => measureTextPx(h))
    const sample = Math.min(recordIds.length, 30)
    for (let r = 0; r < sample; r++) {
      const rid = recordIds[r]
      for (let i = 0; i < displayFids.length; i++) {
        const dfid = displayFids[i]
        if (dfid === '__INDEX__') {
          desired[i] = Math.max(desired[i], measureTextPx(String(recordIds.length)))
        } else {
          try {
            const s = await fieldObjMap.get(dfid).getCellString(rid)
            desired[i] = Math.max(desired[i], measureTextPx(String(s ?? '')))
          } catch {}
        }
      }
    }
    let sumDesired = desired.reduce((a, b) => a + b, 0)
    if (!sumDesired || !isFinite(sumDesired)) sumDesired = minColPx * displayFids.length
    let colPercents = desired.map(d => (d / sumDesired) * 100)
    const sumPerc = colPercents.reduce((a, b) => a + b, 0)
    if (Math.abs(sumPerc - 100) > 0.001) {
      const adjust = 100 - sumPerc
      colPercents[colPercents.length - 1] += adjust
    }
    const rows: any[] = []
  rows.push(new docx.TableRow({
    children: displayFids.map((fid, i) => new docx.TableCell({
      width: { size: colPercents[i], type: docx.WidthType.PERCENTAGE },
      children: [new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        children: [new docx.TextRun({ text: fid === '__INDEX__' ? '序号' : String(fieldNameMap.get(fid) || fid), bold: true, size: 24, font: 'Microsoft YaHei' })]
      })]
    }))
  }))
    for (let r = 0; r < recordIds.length; r++) {
      const rid = recordIds[r]
    const cells = await Promise.all(displayFids.map(async (fid, i) => {
      const children: any[] = []
      if (fid === '__INDEX__') {
        const textVal = String(r + 1)
        children.push(new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, children: [new docx.TextRun({ text: textVal, size: 24, font: 'Microsoft YaHei' })] }))
      } else {
        let textVal = ''
        try {
          const s = await fieldObjMap.get(fid).getCellString(rid)
          textVal = String(s ?? '')
        } catch {
          textVal = ''
        }
        if (insertImages && attachmentFieldIds.has(fid)) {
          textVal = ''
        }
        if (textVal) children.push(new docx.Paragraph({ children: [new docx.TextRun({ text: textVal, size: 24, font: 'Microsoft YaHei' })] }))
        if (insertImages && attachmentFieldIds.has(fid)) {
          try {
            const val = await fieldObjMap.get(fid).getValue(rid)
            const arr = Array.isArray(val) ? val : (val ? [val] : [])
            const tokens = arr.map(a => a?.token).filter(Boolean) as string[]
            if (tokens.length) {
              const urls = await table.getCellAttachmentUrls(tokens, fid, rid)
              const tasks = (urls || []).map(async (url: string | undefined, idx: number) => {
                const fileName = String(arr[idx]?.name || '未知文件')
                if (!url) return { type: 'text', content: fileName } as const
                const res = await fetch(url)
                const blob = await res.blob()
                if (!blob.type.startsWith('image/')) return { type: 'text', content: fileName } as const
                const buf = await blob.arrayBuffer()
                const data = new Uint8Array(buf)
                let iw = 1, ih = 1
                try {
                  const bmp = await (globalThis as any).createImageBitmap ? await createImageBitmap(blob) : null
                  if (bmp) { iw = bmp.width || 1; ih = bmp.height || 1; bmp.close?.() }
                } catch {
                  const dataUrlForSize = await new Promise<string>((resolve, reject) => {
                    const fr = new FileReader()
                    fr.onload = () => resolve(String(fr.result))
                    fr.onerror = reject
                    fr.readAsDataURL(blob)
                  })
                  await new Promise<void>((resolve) => {
                    const img = new Image()
                    img.onload = () => { iw = img.naturalWidth || 1; ih = img.naturalHeight || 1; resolve() }
                    img.src = dataUrlForSize
                  })
                }
                const pagePx = pageOrientation === 'landscape' ? 1120 : 794
                const innerPx = pagePx - 192
                const colMaxW = Math.max(180, Math.floor(innerPx * colPercents[i] / 100) - 12)
                const maxW = Math.max(240, colMaxW)
                const maxH = 400
                const minW = 300
                const minH = 225
                const ratioMax = Math.min(maxW / iw, maxH / ih)
                const ratioMin = Math.max(minW / iw, minH / ih)
                const ratio = Math.min(Math.max(ratioMin, 0), ratioMax)
                return { type: 'image', data, w: Math.max(1, Math.floor(iw * ratio)), h: Math.max(1, Math.floor(ih * ratio)) } as const
              })
              const sizeTasks = tasks.length <= 4 ? 4 : tasks.length <= 12 ? 3 : 2
              const results: Array<{ type: 'text'; content: string } | { type: 'image'; data: Uint8Array; w: number; h: number }> = []
              for (let k = 0; k < tasks.length; k += sizeTasks) {
                const part = await Promise.all(tasks.slice(k, k + sizeTasks))
                for (const it of part) if (it) results.push(it)
              }
              for (const res of results) {
                if (res.type === 'text') {
                  children.push(new docx.Paragraph({ children: [new docx.TextRun({ text: res.content, size: 24, font: 'Microsoft YaHei' })] }))
                } else {
                  children.push(new docx.Paragraph({ children: [new docx.ImageRun({ data: res.data, transformation: { width: res.w, height: res.h } })] }))
                }
              }
            }
          } catch {}
        }
      }
      return new docx.TableCell({
        width: { size: colPercents[i], type: docx.WidthType.PERCENTAGE },
        children
      })
    }))
      rows.push(new docx.TableRow({ children: cells }))
      if ((r + 1) % 10 === 0 || r === recordIds.length - 1) {
        onProgress?.(r + 1, recordIds.length)
        onStatus?.(`写入 ${r + 1}/${recordIds.length}`)
      }
    }
    const doc = new docx.Document({
      sections: [{
        properties: {
          page: {
            size: { orientation: pageOrientation === 'landscape' ? docx.PageOrientation.LANDSCAPE : docx.PageOrientation.PORTRAIT }
          }
        },
        children: [new docx.Table({ rows })]
      }]
    })
    const safe = filename && filename.trim().length ? filename.trim() : '导出.docx'
    const finalName = safe.endsWith('.docx') ? safe : `${safe}.docx`
    const blob = await docx.Packer.toBlob(doc)
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = finalName
    document.body.appendChild(a)
    a.click()
    a.remove()
    URL.revokeObjectURL(url)
    return
  }
  onStatus?.('docx 加载失败，尝试 HTML DOCX')
  const htmlDocx = await loadHtmlDocx()
  if (!htmlDocx || !htmlDocx.asBlob) throw new Error('html-docx-js 未找到')
  function esc(s: string) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  }
  let html = '<!doctype html><html><head><meta charset="utf-8"><style>table{border-collapse:collapse;width:100%}th,td{border:1px solid #000;padding:4px;font-size:12pt;font-family:\"Noto Sans SC\",\"Microsoft YaHei\",sans-serif;}th{text-align:center;font-weight:bold}</style></head><body><table>'
  const displayFidsHtml: string[] = (showIndex === false) ? fieldIds : ['__INDEX__', ...fieldIds]
  html += '<tr>' + displayFidsHtml.map(fid => `<th>${esc(fid === '__INDEX__' ? '序号' : String(fieldNameMap.get(fid) || fid))}</th>`).join('') + '</tr>'
  for (let r = 0; r < recordIds.length; r++) {
    const rid = recordIds[r]
    const cells = await Promise.all(displayFidsHtml.map(async fid => {
      if (fid === '__INDEX__') return `<td style="text-align:center">${esc(String(r + 1))}</td>`
      try {
        const s = await fieldObjMap.get(fid).getCellString(rid)
        return `<td>${esc(String(s ?? ''))}</td>`
      } catch {
        return '<td></td>'
      }
    }))
    html += `<tr>${cells.join('')}</tr>`
    if ((r + 1) % 10 === 0 || r === recordIds.length - 1) {
      onProgress?.(r + 1, recordIds.length)
      onStatus?.(`写入 ${r + 1}/${recordIds.length}`)
    }
  }
  html += '</table></body></html>'
  const safe = filename && filename.trim().length ? filename.trim() : '导出.docx'
  const finalName = safe.endsWith('.docx') ? safe : `${safe}.docx`
  const blob = htmlDocx.asBlob(html, { orientation: pageOrientation === 'landscape' ? 'landscape' : 'portrait' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = finalName
  document.body.appendChild(a)
  a.click()
  a.remove()
  URL.revokeObjectURL(url)
}

async function loadHtmlDocx(): Promise<any> {
  let htmlDocx: any = (globalThis as any).htmlDocx
  if (htmlDocx) return htmlDocx
  const urls = [
    'https://cdn.jsdelivr.net/npm/html-docx-js@0.3.1/dist/html-docx.min.js',
    'https://unpkg.com/html-docx-js@0.3.1/dist/html-docx.min.js',
  ]
  for (const u of urls) {
    try {
      await new Promise<void>((resolve, reject) => {
        const s = document.createElement('script')
        let done = false
        const timer = setTimeout(() => {
          if (done) return
          done = true
          reject(new Error('html-docx-js 加载超时'))
        }, 8000)
        s.src = u
        s.async = true
        s.onload = () => {
          if (done) return
          done = true
          clearTimeout(timer)
          resolve()
        }
        s.onerror = () => {
          if (done) return
          done = true
          clearTimeout(timer)
          reject(new Error('html-docx-js 加载失败'))
        }
        document.head.appendChild(s)
      })
      htmlDocx = (globalThis as any).htmlDocx
      if (htmlDocx) return htmlDocx
    } catch {}
  }
  return (globalThis as any).htmlDocx
}

async function loadJsPDF(): Promise<any> {
  let jspdf: any = (globalThis as any).jspdf
  if (!jspdf) {
    await new Promise<void>((resolve, reject) => {
      const s = document.createElement('script')
      s.src = 'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js'
      s.async = true
      s.onload = () => resolve()
      s.onerror = () => reject(new Error('jsPDF 加载失败'))
      document.head.appendChild(s)
    })
    jspdf = (globalThis as any).jspdf
  }
  return jspdf
}

async function exportPDF(filename: string, onStatus?: (msg: string) => void, onProgress?: (done: number, total: number) => void, tableId?: string, viewId?: string, insertImages?: boolean, pageOrientation?: 'portrait' | 'landscape', showIndex?: boolean, showPageNumber?: boolean) {
  let table: any
  if (tableId) {
    table = await bitable.base.getTableById(tableId)
  } else {
    try {
      table = await getActiveTable()
    } catch {
      const list = await bitable.base.getTableList()
      table = list && list[0]
      if (!table) throw new Error('无法获取数据表')
    }
  }
  let view: any = null
  if (viewId) {
    try {
      const vlist = await table.getViewList()
      view = (vlist || []).find((v: any) => v.id === viewId) || null
    } catch {}
  }
  if (!view) {
    try {
      view = await getActiveView(table)
    } catch {}
  }
  const fieldIds: string[] = await getFieldIds(table, view)
  const recordIds: string[] = await getRecordIds(table, view)
  onProgress?.(0, recordIds.length)
  onStatus?.('收集字段与记录')
  const fieldNameMap = new Map<string, string>()
  const fieldObjMap = new Map<string, any>()
  for (const fid of fieldIds) {
    const name = await getFieldName(table, fid)
    fieldNameMap.set(fid, name)
    const field = await table.getField(fid)
    fieldObjMap.set(fid, field)
  }
  const attachMeta = await table.getFieldMetaListByType(FieldType.Attachment)
  const attachmentFieldIds = new Set<string>((attachMeta || []).map((m: any) => m.id))
  onStatus?.('加载 PDF 库')
  const jspdf = await loadJsPDF()
  if (!jspdf || !jspdf.jsPDF) throw new Error('jsPDF 未找到')
  const doc = new jspdf.jsPDF({ orientation: pageOrientation || 'portrait', unit: 'pt', format: 'a4' })
  async function toB64(ab: ArrayBuffer) {
    let s = ''
    const bytes = new Uint8Array(ab)
    const chunk = 0x8000
    for (let i = 0; i < bytes.length; i += chunk) {
      s += String.fromCharCode(...bytes.subarray(i, i + chunk))
    }
    return btoa(s)
  }
  async function ensureCJKFont() {
    const urls = [
      'https://cdn.jsdelivr.net/gh/jsntn/webfonts/NotoSansSC-Regular.ttf',
      'https://raw.githubusercontent.com/jsntn/webfonts/master/NotoSansSC-Regular.ttf',
    ]
    for (const u of urls) {
      try {
        const res = await fetch(u, { cache: 'force-cache' })
        if (!res.ok) continue
        const buf = await res.arrayBuffer()
        const b64 = await toB64(buf)
        const fname = 'NotoSansSC-Regular.ttf'
        doc.addFileToVFS(fname, b64)
        doc.addFont(fname, 'NotoSansSC', 'normal')
        doc.setFont('NotoSansSC', 'normal')
        return
      } catch {}
    }
  }
  await ensureCJKFont()
  const pageW = doc.internal.pageSize.getWidth()
  const pageH = doc.internal.pageSize.getHeight()
  const margin = 40
  const startX = margin
  let y = margin
  const innerW = pageW - margin * 2
  const lh = 14
  doc.setFontSize(12)
  doc.setLineWidth(0.5)
  doc.setDrawColor(160, 160, 160)
  doc.setTextColor(0, 0, 0)
  const displayFids: string[] = (showIndex === false) ? fieldIds : ['__INDEX__', ...fieldIds]
  const headers = displayFids.map(fid => fid === '__INDEX__' ? '序号' : String(fieldNameMap.get(fid) || fid))
  let headerH = lh + 8
  function textBold(t: string, x: number, yy: number) {
    doc.text(t, x, yy)
    doc.text(t, x + 0.25, yy)
    doc.text(t, x, yy + 0.25)
  }
  function wrapCJK(t: string, maxW: number): string[] {
    const out: string[] = []
    const parts = String(t || '').split(/\r?\n/)
    for (const part of parts) {
      let buf = ''
      for (const ch of part) {
        const w = doc.getTextWidth(buf + ch)
        if (w <= maxW) {
          buf += ch
        } else {
          if (buf.length) out.push(buf)
          buf = ch
        }
      }
      out.push(buf)
    }
    return out.length ? out : ['']
  }
  const minColW = 60
  const desired: number[] = headers.map(h => Math.max(doc.getTextWidth(h) + 12, minColW))
    const sample = Math.min(recordIds.length, 30)
  for (let r = 0; r < sample; r++) {
    const rid = recordIds[r]
    for (let i = 0; i < displayFids.length; i++) {
      const dfid = displayFids[i]
      if (dfid === '__INDEX__') {
        const w = doc.getTextWidth(String(recordIds.length))
        desired[i] = Math.max(desired[i], w + 12)
      } else {
        try {
          const s = await fieldObjMap.get(dfid).getCellString(rid)
          const w = doc.getTextWidth(String(s ?? ''))
          desired[i] = Math.max(desired[i], w + 12)
        } catch {}
      }
    }
  }
  let colWs: number[] = []
  if (minColW * displayFids.length > innerW) {
    const eq = innerW / displayFids.length
    colWs = new Array(displayFids.length).fill(eq)
  } else {
    const sumDesired = desired.reduce((a, b) => a + b, 0) || innerW
    colWs = desired.map(d => innerW * (d / sumDesired))
    for (let i = 0; i < colWs.length; i++) {
      if (colWs[i] < minColW) colWs[i] = minColW
    }
    let sumWs = colWs.reduce((a, b) => a + b, 0)
    let delta = innerW - sumWs
    if (Math.abs(delta) > 0.5) {
      if (delta > 0) {
        const weightSum = colWs.reduce((a, b) => a + b, 0)
        for (let i = 0; i < colWs.length; i++) {
          colWs[i] += delta * (colWs[i] / weightSum)
        }
      } else {
        const flex = colWs.map(w => Math.max(0, w - minColW))
        const flexSum = flex.reduce((a, b) => a + b, 0)
        if (flexSum > 0) {
          for (let i = 0; i < colWs.length; i++) {
            const reduce = (-delta) * (flex[i] / flexSum)
            colWs[i] = Math.max(minColW, colWs[i] - reduce)
          }
        } else {
          const eq = innerW / displayFids.length
          colWs = new Array(displayFids.length).fill(eq)
        }
      }
      const adjust = innerW - colWs.reduce((a, b) => a + b, 0)
      colWs[colWs.length - 1] += adjust
    }
  }
  function colX(i: number) {
    let x = startX
    for (let k = 0; k < i; k++) x += colWs[k]
    return x
  }
  const headerLinesArr: string[][] = headers.map((h, i) => {
    const maxW = Math.max(4, colWs[i] - 8)
    let lines = doc.splitTextToSize(h, maxW)
    if (doc.getTextWidth(String(h || '')) > maxW) {
      if (lines.length <= 1) {
        lines = wrapCJK(String(h || ''), maxW)
      } else {
        const fixed: string[] = []
        for (const l of lines) {
          if (doc.getTextWidth(l) > maxW) {
            fixed.push(...wrapCJK(l, maxW))
          } else {
            fixed.push(l)
          }
        }
        lines = fixed
      }
    }
    return lines
  })
  const headerHeights = headerLinesArr.map(ls => Math.max(lh, ls.length * lh))
  headerH = Math.max(headerH, Math.max(...headerHeights) + 6)
  function drawHeader() {
    for (let i = 0; i < headers.length; i++) {
      const x = colX(i)
      const w = colWs[i]
      doc.rect(x, y, w, headerH, 'S')
      const lines = headerLinesArr[i]
      for (let idx = 0; idx < lines.length; idx++) {
        const line = lines[idx]
        const tw = doc.getTextWidth(line)
        const cx = x + w / 2 - tw / 2
        textBold(line, cx, y + lh + idx * lh)
      }
    }
    y += headerH
  }
  drawHeader()
  for (let r = 0; r < recordIds.length; r++) {
    const rid = recordIds[r]
    const texts: string[] = await Promise.all(displayFids.map(async fid => {
      if (fid === '__INDEX__') return String(r + 1)
      try {
        if (insertImages && attachmentFieldIds.has(fid)) return ''
        const s = await fieldObjMap.get(fid).getCellString(rid)
        return String(s ?? '')
      } catch {
        return ''
      }
    }))
  const linesArr: string[][] = []
  const textHeights: number[] = []
  let rowH = lh
  for (let i = 0; i < displayFids.length; i++) {
    const maxW = Math.max(4, colWs[i] - 8)
    let lines = doc.splitTextToSize(texts[i], maxW)
    if (doc.getTextWidth(String(texts[i] || '')) > maxW) {
      if (lines.length <= 1) {
        lines = wrapCJK(String(texts[i] || ''), maxW)
      } else {
        const fixed: string[] = []
        for (const l of lines) {
          if (doc.getTextWidth(l) > maxW) {
            fixed.push(...wrapCJK(l, maxW))
          } else {
            fixed.push(l)
          }
        }
        lines = fixed
      }
    }
    linesArr[i] = lines
    const th = Math.max(lh, lines.length * lh)
    textHeights[i] = th
    rowH = Math.max(rowH, th + 6)
  }
    const imageInfos: Array<Array<{ dataUrl: string; ext: string; w: number; h: number }>> = Array.from({ length: displayFids.length }, () => [])
    const extraTextLines: Array<Array<string>> = Array.from({ length: displayFids.length }, () => [])
    if (insertImages) {
      const tasks: Promise<void>[] = []
      for (let i = 0; i < displayFids.length; i++) {
        const fid = displayFids[i]
        if (fid === '__INDEX__') continue
        if (!attachmentFieldIds.has(fid)) continue
        tasks.push((async () => {
          try {
            const val = await fieldObjMap.get(fid).getValue(rid)
            const arr = Array.isArray(val) ? val : (val ? [val] : [])
            const tokens = arr.map(a => a?.token).filter(Boolean) as string[]
            if (!tokens.length) return
            const urls = await table.getCellAttachmentUrls(tokens, fid, rid)
            const subtasks = (urls || []).map(async (url: string | undefined, idx: number) => {
              const fileName = String(arr[idx]?.name || '未知文件')
              if (!url) {
                const lines = doc.splitTextToSize(fileName, colWs[i] - 8)
                extraTextLines[i].push(...lines)
                return null
              }
              const res = await fetch(url)
              const blob = await res.blob()
              if (!blob.type.startsWith('image/')) {
                const lines = doc.splitTextToSize(fileName, colWs[i] - 8)
                extraTextLines[i].push(...lines)
                return null
              }
              const dataUrl = await new Promise<string>((resolve, reject) => {
                const fr = new FileReader()
                fr.onload = () => resolve(String(fr.result))
                fr.onerror = reject
                fr.readAsDataURL(blob)
              })
              const ext = blob.type.includes('jpeg') || blob.type.includes('jpg') ? 'JPEG' : blob.type.includes('png') ? 'PNG' : blob.type.includes('gif') ? 'GIF' : 'PNG'
              const maxW = colWs[i] - 8
              const maxH = 300
              const minW = Math.min(maxW, 240)
              const minH = 180
              const dims = await new Promise<{ w: number; h: number }>((resolve) => {
                const img = new Image()
                img.onload = () => {
                  const iw = img.naturalWidth || 1
                  const ih = img.naturalHeight || 1
                  const ratioMax = Math.min(maxW / iw, maxH / ih)
                  const ratioMin = Math.max(minW / iw, minH / ih)
                  const ratio = Math.min(Math.max(ratioMin, 0), ratioMax)
                  resolve({ w: Math.max(1, Math.floor(iw * ratio)), h: Math.max(1, Math.floor(ih * ratio)) })
                }
                img.src = dataUrl
              })
              return { dataUrl, ext, w: dims.w, h: dims.h }
            })
            const sizeSubs = subtasks.length <= 4 ? 4 : subtasks.length <= 12 ? 3 : 2
            const infosParts: Array<{ dataUrl: string; ext: string; w: number; h: number } | null> = []
            for (let k = 0; k < subtasks.length; k += sizeSubs) {
              const part = await Promise.all(subtasks.slice(k, k + sizeSubs))
              infosParts.push(...part)
            }
            const infos = infosParts.filter(Boolean) as Array<{ dataUrl: string; ext: string; w: number; h: number }>
            imageInfos[i] = infos
          } catch {}
        })())
      }
      await Promise.all(tasks)
      for (let i = 0; i < displayFids.length; i++) {
        if (extraTextLines[i].length) {
          linesArr[i].push(...extraTextLines[i])
          textHeights[i] = (linesArr[i].length) * lh
          rowH = Math.max(rowH, textHeights[i] + 6)
        }
      }
      for (let i = 0; i < displayFids.length; i++) {
        const sumH = imageInfos[i].reduce((a, b) => a + b.h + 2, 0)
        rowH = Math.max(rowH, (textHeights[i] || 0) + 6 + sumH)
      }
    }
    if (y + rowH + margin > pageH) {
      doc.addPage()
      y = margin
      drawHeader()
    }
    for (let i = 0; i < displayFids.length; i++) {
      const x = colX(i)
      doc.rect(x, y, colWs[i], rowH, 'S')
      if (displayFids[i] === '__INDEX__') {
        doc.text(String(texts[i] || ''), x + colWs[i] / 2, y + lh, { align: 'center' })
      } else {
        doc.text(linesArr[i], x + 4, y + lh)
      }
      if (imageInfos[i].length) {
        let iy = y + (textHeights[i] || 0) + 6
        for (const info of imageInfos[i]) {
          doc.addImage(info.dataUrl, info.ext, x + 4, iy, info.w, info.h)
          iy += info.h + 2
        }
      }
    }
    y += rowH
    if ((r + 1) % 10 === 0 || r === recordIds.length - 1) {
      onProgress?.(r + 1, recordIds.length)
      onStatus?.(`写入 ${r + 1}/${recordIds.length}`)
    }
  }
  if (showPageNumber) {
    const count = doc.internal.getNumberOfPages()
    const size = 10
    for (let i = 1; i <= count; i++) {
      doc.setPage(i)
      const pw = doc.internal.pageSize.getWidth()
      const ph = doc.internal.pageSize.getHeight()
      doc.setFontSize(size)
      doc.text(`${i}/${count}`, pw / 2, ph - 20, { align: 'center' })
    }
    doc.setFontSize(12)
  }
  const safe = filename && filename.trim().length ? filename.trim() : '导出.pdf'
  const finalName = safe.endsWith('.pdf') ? safe : `${safe}.pdf`
  doc.save(finalName)
}

export default function App() {
  const [name, setName] = useState('')
  const [nameDirty, setNameDirty] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [status, setStatus] = useState('')
  const [progress, setProgress] = useState<{ done: number; total: number }>({ done: 0, total: 0 })
  const [tables, setTables] = useState<{ label: string; value: string }[]>([])
  const [views, setViews] = useState<{ label: string; value: string }[]>([])
  const [tableId, setTableId] = useState<string>()
  const [viewId, setViewId] = useState<string>()
  const [format, setFormat] = useState<'xlsx' | 'docx' | 'pdf'>('xlsx')
  const [insertImages, setInsertImages] = useState<boolean>(false)
  const [pageOrientation, setPageOrientation] = useState<'portrait' | 'landscape'>('portrait')
  const [showIndex, setShowIndex] = useState<boolean>(true)
  const [showPageNumber, setShowPageNumber] = useState<boolean>(true)
  const [startTs, setStartTs] = useState<number | null>(null)
  const disabled = useMemo(() => !name.trim().length, [name])
  const percent = useMemo(() => (progress.total ? Math.round((progress.done * 100) / progress.total) : 0), [progress])
  

  const onExport = async () => {
    setExporting(true)
    setStatus('准备中')
    setProgress({ done: 0, total: 0 })
    setStartTs(null)
    try {
      if (format === 'xlsx') {
        await exportExcel(name, (msg: string) => setStatus(msg), (d: number, t: number) => { setProgress({ done: d, total: t }); setStartTs(s => s ?? Date.now()) }, tableId, viewId, insertImages)
      } else if (format === 'docx') {
        await exportWord(name, (msg: string) => setStatus(msg), (d: number, t: number) => { setProgress({ done: d, total: t }); setStartTs(s => s ?? Date.now()) }, tableId, viewId, insertImages, pageOrientation, showIndex)
      } else {
        await exportPDF(name, (msg: string) => setStatus(msg), (d: number, t: number) => { setProgress({ done: d, total: t }); setStartTs(s => s ?? Date.now()) }, tableId, viewId, insertImages, pageOrientation, showIndex, showPageNumber)
      }
      setStatus('导出完成')
    } catch (e: any) {
      setStatus(`导出失败: ${e?.message || String(e)}`)
    } finally {
      setExporting(false)
    }
  }
  useEffect(() => {
    const init = async () => {
      const list = await bitable.base.getTableList()
      const opts = await Promise.all(list.map(async t => ({ label: await t.getName(), value: t.id })))
      setTables(opts)
      let active: any = null
      try { active = await bitable.base.getActiveTable() } catch {}
      const defId = active?.id || opts[0]?.value
      if (!defId) return
      setTableId(defId)
      const table = await bitable.base.getTableById(defId)
      const vlist = await table.getViewList()
      const vopts = await Promise.all(vlist.map(async v => ({ label: await v.getName(), value: v.id })))
      setViews(vopts)
      let aview: any = null
      try { aview = await table.getActiveView() } catch {}
      setViewId(aview?.id || vopts[0]?.value)
    }
    init()
  }, [])
  useEffect(() => {
    const loadViews = async () => {
      if (!tableId) return
      const table = await bitable.base.getTableById(tableId)
      const vlist = await table.getViewList()
      const vopts = await Promise.all(vlist.map(async v => ({ label: await v.getName(), value: v.id })))
      setViews(vopts)
      let aview: any = null
      try { aview = await table.getActiveView() } catch {}
      setViewId(aview?.id || vopts[0]?.value)
    }
    loadViews()
  }, [tableId])
  useEffect(() => {
    if (nameDirty) return
    const tName = tables.find(o => o.value === tableId)?.label
    const vName = views.find(o => o.value === viewId)?.label
    const auto = tName && vName ? `${tName}-${vName}` : (tName || '')
    if (auto && auto !== name) setName(auto)
  }, [tableId, viewId, tables, views, nameDirty])
  return (
    <div className="plugin-container">
      <div className="card">
        <div className="field">
          <label className="label">数据表</label>
          <select className="select" value={tableId} onChange={e => setTableId(e.target.value)}>
            {tables.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
          </select>
        </div>
        <div className="field">
          <label className="label">视图</label>
          <select className="select" value={viewId} onChange={e => setViewId(e.target.value)}>
            {views.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
          </select>
        </div>
        <div className="field">
          <label className="label">文件名</label>
          <input className="input" value={name} onChange={e => { setName(e.target.value); setNameDirty(true) }} placeholder="导出文件名" />
        </div>
        <div className="field">
          <label className="label">导出格式</label>
          <select className="select" value={format} onChange={e => setFormat(e.target.value as any)}>
            <option value="xlsx">Excel</option>
            <option value="docx">Word</option>
            <option value="pdf">PDF</option>
          </select>
        </div>
        {(format === 'docx' || format === 'pdf' || format === 'xlsx') ? (
          <div className="field">
            <label className="label">插入图片</label>
            <div className="radio-group">
              <label style={{ marginRight: 16 }}>
                <input type="radio" checked={insertImages === true} onChange={() => setInsertImages(true)} /> 是
              </label>
              <label>
                <input type="radio" checked={insertImages === false} onChange={() => setInsertImages(false)} /> 否
              </label>
            </div>
          </div>
        ) : null}
        {(format === 'docx' || format === 'pdf') ? (
          <div className="field">
            <label className="label">页面方向</label>
            <div className="radio-group">
              <label style={{ marginRight: 16 }}>
                <input type="radio" checked={pageOrientation === 'portrait'} onChange={() => setPageOrientation('portrait')} /> 纵向
              </label>
              <label>
                <input type="radio" checked={pageOrientation === 'landscape'} onChange={() => setPageOrientation('landscape')} /> 横向
              </label>
            </div>
          </div>
        ) : null}
        {(format === 'docx' || format === 'pdf') ? (
          <div className="field">
            <label className="label">显示序号</label>
            <div className="radio-group">
              <label style={{ marginRight: 16 }}>
                <input type="radio" checked={showIndex === true} onChange={() => setShowIndex(true)} /> 是
              </label>
              <label>
                <input type="radio" checked={showIndex === false} onChange={() => setShowIndex(false)} /> 否
              </label>
            </div>
          </div>
        ) : null}
        {format === 'pdf' ? (
          <div className="field">
            <label className="label">显示页码</label>
            <div className="radio-group">
              <label style={{ marginRight: 16 }}>
                <input type="radio" checked={showPageNumber === true} onChange={() => setShowPageNumber(true)} /> 是
              </label>
              <label>
                <input type="radio" checked={showPageNumber === false} onChange={() => setShowPageNumber(false)} /> 否
              </label>
            </div>
          </div>
        ) : null}
        <button className="btn btn-primary" onClick={onExport} disabled={disabled || exporting}>导出</button>
        <div className="status">状态：{status}</div>
        {exporting || progress.total > 0 ? (
          <>
            <div className="progress">
              <div className="progress-bar" style={{ width: `${percent}%` }} />
            </div>
            <div className="progress-text">{progress.done}/{progress.total}（{percent}%）</div>
          </>
        ) : null}
      </div>
    </div>
  )
}
