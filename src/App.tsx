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

async function exportExcel(filename: string, onStatus?: (msg: string) => void, onProgress?: (done: number, total: number) => void, tableId?: string, viewId?: string) {
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
      for (let i = 0; i < arr.length; i++) {
        const item = arr[i]
        const token: string | undefined = item?.token
        const colIndex = findColIndex(fid, i)
        const c = colIndex - 1
        if (!token || colIndex < 1) {
          rowValues[c] = ''
          continue
        }
        imgTasks.push(async () => {
          try {
            const urls = await table.getCellAttachmentUrls([token], fid, rid)
            const url = urls && urls[0]
            if (!url) {
              rowValues[c] = ''
              return
            }
            const res = await fetch(url)
            const blob = await res.blob()
            if (!blob.type.startsWith('image/')) {
              rowValues[c] = normalize(item)
              return
            }
            const dataUrl = await new Promise<string>((resolve, reject) => {
              const fr = new FileReader()
              fr.onload = () => resolve(String(fr.result))
              fr.onerror = reject
              fr.readAsDataURL(blob)
            })
            const base64 = dataUrl.includes(',') ? dataUrl.split(',')[1] : dataUrl
            const contentType = blob.type || 'image/png'
            const ext = contentType.includes('jpeg') || contentType.includes('jpg') ? 'jpeg' : contentType.includes('png') ? 'png' : contentType.includes('gif') ? 'gif' : 'png'
            const imgId = wb.addImage({ base64, extension: ext as any })
            const imgWidth = 120
            const imgHeight = 90
            ws.getColumn(colIndex).width = Math.max(ws.getColumn(colIndex).width || 15, Math.ceil(imgWidth / 7))
            targetRowHeight = Math.max(targetRowHeight, imgHeight)
            ws.addImage(imgId, { tl: { col: colIndex - 1, row: rowIndex - 1 }, ext: { width: imgWidth, height: imgHeight } })
            rowValues[c] = ''
          } catch {
            rowValues[c] = normalize(item)
          }
        })
      }
    }
    if (imgTasks.length) await runBatches(imgTasks, 4)

    ws.addRow(rowValues)
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
  const [startTs, setStartTs] = useState<number | null>(null)
  const disabled = useMemo(() => !name.trim().length, [name])
  const percent = useMemo(() => (progress.total ? Math.round((progress.done * 100) / progress.total) : 0), [progress])
  const remainSec = useMemo(() => {
    if (!startTs || !progress.total || !progress.done) return null
    const elapsed = (Date.now() - startTs) / 1000
    if (elapsed <= 0) return null
    const rate = progress.done / elapsed
    if (rate <= 0) return null
    const left = (progress.total - progress.done) / rate
    if (!isFinite(left) || left <= 0) return null
    return Math.ceil(left)
  }, [progress, startTs])
  const remainText = useMemo(() => {
    if (!remainSec || remainSec <= 10) return ''
    const m = Math.floor(remainSec / 60)
    const s = remainSec % 60
    const ss = String(s).padStart(2, '0')
    return m ? `${m}分${ss}秒` : `${ss}秒`
  }, [remainSec])

  const onExport = async () => {
    setExporting(true)
    setStatus('准备中')
    setProgress({ done: 0, total: 0 })
    setStartTs(null)
    try {
      await exportExcel(name, (msg: string) => setStatus(msg), (d: number, t: number) => { setProgress({ done: d, total: t }); setStartTs(s => s ?? Date.now()) }, tableId, viewId)
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
        <button className="btn btn-primary" onClick={onExport} disabled={disabled || exporting}>导出为 Excel</button>
        <div className="status">状态：{status}</div>
        {exporting || progress.total > 0 ? (
          <>
            <div className="progress">
              <div className="progress-bar" style={{ width: `${percent}%` }} />
            </div>
            <div className="progress-text">{progress.done}/{progress.total}（{percent}%）</div>
            {remainText ? <div className="progress-text">预计剩余：{remainText}</div> : null}
          </>
        ) : null}
      </div>
    </div>
  )
}
