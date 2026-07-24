import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { Upload, Database, Factory, FlaskConical, CalendarDays, Search, CheckCircle2, AlertTriangle, Clock3, X, BarChart3 } from 'lucide-react'
import './styles.css'

const TARGETS = {
  '1519': 80000, '1521': 60000, '1523': 40000, '1524': 6000,
  '1525': 130000, '1528': 80000, '1540': 18000, '1541': 210000,
  '1542': 60000, '1543': 60000,
}
const FACILITIES = Object.keys(TARGETS)

const normalize = (v) => String(v ?? '').trim()
const normKey = (v) => normalize(v).toLowerCase().replace(/\s+/g, ' ')
const num = (v) => {
  const n = Number(String(v ?? '').replace(/,/g, ''))
  return Number.isFinite(n) ? n : 0
}
const excelDate = (v) => {
  if (!v) return null
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v)
    return d ? new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0) : null
  }
  const d = new Date(v)
  return Number.isNaN(d.getTime()) ? null : d
}
const iso = (d) => d ? new Date(d).toISOString().slice(0, 10) : ''
const fmt = (n) => Math.round(n || 0).toLocaleString('he-IL')
const getField = (row, names) => {
  const map = new Map(Object.keys(row || {}).map(k => [normKey(k), row[k]]))
  for (const name of names) {
    const value = map.get(normKey(name))
    if (value !== undefined) return value
  }
  return ''
}

async function readWorkbook(file) {
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf, { type: 'array', cellDates: true })
  let rows = []
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName]
    const sheetRows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true })
    if (sheetRows.length) rows = rows.concat(sheetRows)
  }
  return rows
}

function classifyFile(rows) {
  const keys = Object.keys(rows[0] || {}).map(normKey)
  const has = (...terms) => terms.some(term => keys.some(k => k.includes(normKey(term))))
  if (has('actual finish time', 'delivered quantity', 'storage location')) return 'production'
  if (has('rejected characteristics', 'qa status', 'ud remarks')) return 'deviations'
  if (has('master insp charactristic', 'result status', 'inspection lot')) return 'quality'
  return 'unknown'
}

export default function App() {
  const [production, setProduction] = useState([])
  const [quality, setQuality] = useState([])
  const [deviations, setDeviations] = useState([])
  const [status, setStatus] = useState('ממתין לטעינת קבצים')
  const [busy, setBusy] = useState(false)
  const [query, setQuery] = useState('')
  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')
  const [selectedFacility, setSelectedFacility] = useState('')
  const [activeTab, setActiveTab] = useState('production')

  const handleFiles = async (files) => {
    if (!files.length) return
    setBusy(true)
    try {
      const loaded = []
      let prodRows = null, qualityRows = null, deviationRows = null
      for (const file of files) {
        setStatus(`קורא את ${file.name}...`)
        const rows = await readWorkbook(file)
        const kind = classifyFile(rows)
        if (kind === 'production') prodRows = rows
        if (kind === 'quality') qualityRows = rows
        if (kind === 'deviations') deviationRows = rows
        loaded.push(`${file.name}: ${fmt(rows.length)} רשומות${kind === 'unknown' ? ' (סוג לא זוהה)' : ''}`)
      }
      if (prodRows) setProduction(prodRows)
      if (qualityRows) setQuality(qualityRows)
      if (deviationRows) setDeviations(deviationRows)
      setStatus(`נטען בהצלחה — ${loaded.join(' | ')}`)
    } catch (e) {
      console.error(e)
      setStatus(`שגיאה בקריאת קובץ: ${e.message}`)
    } finally { setBusy(false) }
  }

  const prod = useMemo(() => production.map(r => {
    const finish = excelDate(getField(r, ['Actual Finish Time', 'Actual finish date', 'Release date (actual)']))
    return {
      facility: normalize(getField(r, ['Storage Location', 'Storage location'])),
      date: finish,
      qty: num(getField(r, ['Delivered quantity (GMEIN)', 'Confirmed Yield Quantity (GMEIN)', 'Delivered quantity'])),
      order: normalize(getField(r, ['Order', 'Process Order', 'Work Order'])),
      batch: normalize(getField(r, ['Batch'])),
      material: normalize(getField(r, ['Material'])),
      desc: normalize(getField(r, ['Material description', 'Material Description'])),
      orderType: normalize(getField(r, ['Order Type'])),
      hour: finish ? finish.getHours() : null,
    }
  }).filter(r => r.facility), [production])

  const qualityRows = useMemo(() => quality.map(r => ({
    facility: normalize(getField(r, ['Production Line', 'Facility', 'Storage Location'])),
    batch: normalize(getField(r, ['Batch'])),
    material: normalize(getField(r, ['Material'])),
    order: normalize(getField(r, ['Process Order', 'Order'])),
    status: normalize(getField(r, ['Result Status', 'QA Approval', 'Status'])),
    inspectionLot: normalize(getField(r, ['Inspection Lot'])),
  })), [quality])

  const deviationRows = useMemo(() => deviations.map(r => ({
    facility: normalize(getField(r, ['Facility', 'Production Line', 'Storage Location'])),
    batch: normalize(getField(r, ['Batch'])),
    material: normalize(getField(r, ['Material'])),
    status: normalize(getField(r, ['QA Status', 'Status'])),
    rejected: normalize(getField(r, ['Rejected characteristics', 'Rejected characteristics '])),
    remarks: normalize(getField(r, ['UD Remarks', 'Remarks'])),
  })), [deviations])

  const dateBounds = useMemo(() => {
    const ds = prod.map(r => r.date).filter(Boolean).sort((a, b) => a - b)
    return { min: iso(ds[0]), max: iso(ds.at(-1)) }
  }, [prod])

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase()
    return prod.filter(r => {
      const d = iso(r.date)
      const okDate = (!from || d >= from) && (!to || d <= to)
      const okFacility = !selectedFacility || r.facility === selectedFacility
      const okQ = !q || [r.facility, r.order, r.batch, r.material, r.desc].some(v => v.toLowerCase().includes(q))
      return okDate && okFacility && okQ
    })
  }, [prod, from, to, query, selectedFacility])

  const days = useMemo(() => new Set(filtered.map(r => iso(r.date)).filter(Boolean)).size || 1, [filtered])
  const facilityStats = useMemo(() => FACILITIES.map(id => {
    let rows = filtered.filter(r => r.facility === id)
    if (id === '1542') rows = rows.filter(r => r.orderType.toUpperCase().includes('ZFIN'))
    const actual = rows.reduce((s, r) => s + r.qty, 0)
    const target = TARGETS[id] * days
    return { id, actual, target, pct: target ? Math.round(actual / target * 100) : 0, orders: new Set(rows.map(r => r.order).filter(Boolean)).size }
  }), [filtered, days])

  const total = facilityStats.reduce((s, x) => s + x.actual, 0)
  const activeFacilities = facilityStats.filter(x => x.actual > 0).length
  const morningQty = filtered.filter(r => r.hour !== null && r.hour >= 7 && r.hour < 19).reduce((s, r) => s + r.qty, 0)
  const nightQty = filtered.filter(r => r.hour !== null && (r.hour >= 19 || r.hour < 7)).reduce((s, r) => s + r.qty, 0)
  const qualityBad = qualityRows.filter(r => {
    const st = r.status.toLowerCase()
    return st && !['accepted', 'תקין', 'pass', 'approved', 'מאושר'].some(x => st.includes(x))
  })
  const openDeviations = deviationRows.filter(r => {
    const st = r.status.toLowerCase()
    return !st || !['approved', 'closed', 'מאושר', 'סגור'].some(x => st.includes(x))
  })

  const setQuickRange = (daysBack) => {
    if (!dateBounds.max) return
    const end = new Date(`${dateBounds.max}T12:00:00`)
    const start = new Date(end)
    start.setDate(end.getDate() - (daysBack - 1))
    setFrom(iso(start)); setTo(dateBounds.max)
  }

  return <div className="dashboard" dir="rtl">
    <aside className="side">
      <div className="brand">IML<span>CONTROL</span></div>
      <div className="side-stat"><Database/><div><b>{fmt(production.length)}</b><small>רשומות תפוקה</small></div></div>
      <div className="side-stat"><FlaskConical/><div><b>{fmt(quality.length + deviations.length)}</b><small>רשומות איכות</small></div></div>
      <div className="side-stat"><Clock3/><div><b>{fmt(morningQty)}</b><small>משמרת בוקר</small></div></div>
      <div className="side-note">הנתונים מעובדים בדפדפן בלבד ואינם נשלחים לשרת.</div>
    </aside>

    <main className="main">
      <header className="header">
        <div><h1>מרכז שליטה למתקני אריזה</h1><p>טעינת Excel, KPI, איכות והשוואת משמרות</p></div>
        <label className={`upload ${busy ? 'disabled' : ''}`}><Upload size={19}/>{busy ? 'טוען...' : 'טעינת קובצי Excel'}<input type="file" multiple accept=".xlsx,.xls" disabled={busy} onChange={e => handleFiles([...e.target.files])}/></label>
      </header>

      <div className="load-status"><CheckCircle2 size={18}/>{status}</div>

      <section className="filters">
        <label><Search size={17}/><input value={query} onChange={e => setQuery(e.target.value)} placeholder="חיפוש מתקן, Order, Batch או חומר"/></label>
        <label><CalendarDays size={17}/>מתאריך<input type="date" min={dateBounds.min} max={dateBounds.max} value={from} onChange={e => setFrom(e.target.value)}/></label>
        <label><CalendarDays size={17}/>עד תאריך<input type="date" min={dateBounds.min} max={dateBounds.max} value={to} onChange={e => setTo(e.target.value)}/></label>
        <button onClick={() => setQuickRange(1)}>יום אחרון</button>
        <button onClick={() => setQuickRange(2)}>יומיים</button>
        <button onClick={() => setQuickRange(30)}>30 יום</button>
        <button onClick={() => { setFrom(''); setTo(''); setQuery(''); setSelectedFacility('') }}>נקה</button>
      </section>

      {selectedFacility && <div className="selection"><span>מסנן לפי מתקן <b>{selectedFacility}</b></span><button onClick={() => setSelectedFacility('')}><X size={16}/> הסר סינון</button></div>}

      <section className="summary-grid">
        <Summary title="סה״כ תפוקה" value={fmt(total)} sub="ליטר / ק״ג לפי הקובץ"/>
        <Summary title="מתקנים פעילים" value={activeFacilities} sub={`מתוך ${FACILITIES.length} מתקנים`}/>
        <Summary title="משמרת בוקר" value={fmt(morningQty)} sub="07:00–19:00"/>
        <Summary title="משמרת לילה" value={fmt(nightQty)} sub="19:00–07:00"/>
        <Summary title="חריגות פתוחות" value={openDeviations.length} sub="לפי קובץ המנות החריגות" warn/>
        <Summary title="תוצאות איכות לא תקינות" value={qualityBad.length} sub="לפי קובץ האיכות" warn/>
      </section>

      <div className="section-title"><Factory/><div><h2>ביצועים לפי מתקן</h2><p>לחיצה על כרטיס מסננת את כל הנתונים למתקן הנבחר</p></div></div>
      <section className="facility-grid">
        {facilityStats.map(x => <Facility key={x.id} {...x} onClick={() => setSelectedFacility(x.id)}/>) }
      </section>

      <section className="tabs">
        <button className={activeTab === 'production' ? 'active' : ''} onClick={() => setActiveTab('production')}><BarChart3 size={16}/> תפוקה</button>
        <button className={activeTab === 'quality' ? 'active' : ''} onClick={() => setActiveTab('quality')}><FlaskConical size={16}/> איכות ({qualityBad.length})</button>
        <button className={activeTab === 'deviations' ? 'active' : ''} onClick={() => setActiveTab('deviations')}><AlertTriangle size={16}/> מנות חריגות ({openDeviations.length})</button>
      </section>

      {activeTab === 'production' && <section className="details"><h2>רשומות תפוקה אחרונות</h2><div className="table-wrap"><table><thead><tr><th>תאריך</th><th>שעה</th><th>מתקן</th><th>הזמנה</th><th>Batch</th><th>חומר</th><th>כמות</th></tr></thead><tbody>
        {filtered.slice(-200).reverse().map((r, i) => <tr key={i}><td>{iso(r.date)}</td><td>{r.date ? r.date.toLocaleTimeString('he-IL', {hour:'2-digit',minute:'2-digit'}) : ''}</td><td>{r.facility}</td><td>{r.order}</td><td>{r.batch}</td><td>{r.desc || r.material}</td><td>{fmt(r.qty)}</td></tr>)}
        {!filtered.length && <tr><td colSpan="7" className="empty">טען את קובץ התפוקות כדי להציג נתונים</td></tr>}
      </tbody></table></div></section>}

      {activeTab === 'quality' && <section className="details"><h2>תוצאות איכות לא תקינות</h2><div className="table-wrap"><table><thead><tr><th>מתקן</th><th>Inspection Lot</th><th>Order</th><th>Batch</th><th>Material</th><th>סטטוס</th></tr></thead><tbody>
        {qualityBad.slice(0, 300).map((r, i) => <tr key={i}><td>{r.facility}</td><td>{r.inspectionLot}</td><td>{r.order}</td><td>{r.batch}</td><td>{r.material}</td><td><span className="status-bad">{r.status || 'ללא סטטוס'}</span></td></tr>)}
        {!qualityBad.length && <tr><td colSpan="6" className="empty">לא נמצאו תוצאות איכות לא תקינות</td></tr>}
      </tbody></table></div></section>}

      {activeTab === 'deviations' && <section className="details"><h2>מנות חריגות פתוחות</h2><div className="table-wrap"><table><thead><tr><th>מתקן</th><th>Batch</th><th>Material</th><th>סטטוס</th><th>מאפיינים שנדחו</th><th>הערות</th></tr></thead><tbody>
        {openDeviations.slice(0, 300).map((r, i) => <tr key={i}><td>{r.facility}</td><td>{r.batch}</td><td>{r.material}</td><td><span className="status-bad">{r.status || 'פתוח'}</span></td><td>{r.rejected}</td><td>{r.remarks}</td></tr>)}
        {!openDeviations.length && <tr><td colSpan="6" className="empty">לא נמצאו מנות חריגות פתוחות</td></tr>}
      </tbody></table></div></section>}
    </main>
  </div>
}

function Summary({ title, value, sub, warn }) { return <div className={`summary ${warn ? 'warn' : ''}`}><span>{title}</span><b>{value}</b><small>{sub}</small></div> }
function Facility({ id, actual, target, pct, orders, onClick }) {
  const capped = Math.min(pct, 100)
  const state = pct >= 100 ? 'good' : pct >= 75 ? 'mid' : 'bad'
  return <article className={`facility ${state}`} onClick={onClick} role="button" tabIndex="0">
    <div className="facility-top"><div><small>מתקן</small><h3>{id}</h3></div><b>{pct}%</b></div>
    <div className="bar"><i style={{ width: `${capped}%` }}/></div>
    <div className="facility-numbers"><span>בפועל<strong>{fmt(actual)}</strong></span><span>יעד<strong>{fmt(target)}</strong></span><span>הזמנות<strong>{orders}</strong></span></div>
  </article>
}
