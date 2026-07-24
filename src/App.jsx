import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { Upload, Database, Factory, FlaskConical, CalendarDays, Search, CheckCircle2, AlertTriangle } from 'lucide-react'
import './styles.css'

const TARGETS = {
  '1519': 80000, '1521': 60000, '1523': 40000, '1524': 6000,
  '1525': 130000, '1528': 80000, '1540': 18000, '1541': 210000,
  '1542': 60000, '1543': 60000,
}
const FACILITIES = Object.keys(TARGETS)

const normalize = (v) => String(v ?? '').trim()
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
const iso = (d) => d ? new Date(d).toISOString().slice(0,10) : ''
const fmt = (n) => Math.round(n || 0).toLocaleString('he-IL')

async function readWorkbook(file) {
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf, { type: 'array', cellDates: true })
  const ws = wb.Sheets[wb.SheetNames[0]]
  return XLSX.utils.sheet_to_json(ws, { defval: '', raw: true })
}

function classifyFile(rows) {
  const keys = new Set(Object.keys(rows[0] || {}))
  if (keys.has('Actual Finish Time') || keys.has('Delivered quantity (GMEIN)')) return 'production'
  if (keys.has('Rejected characteristics ') || keys.has('QA Status')) return 'deviations'
  if (keys.has('Master Insp Charactristic') || keys.has('Result Status')) return 'quality'
  return 'unknown'
}

export default function App(){
  const [production, setProduction] = useState([])
  const [quality, setQuality] = useState([])
  const [deviations, setDeviations] = useState([])
  const [status, setStatus] = useState('ממתין לטעינת קבצים')
  const [busy, setBusy] = useState(false)
  const [query, setQuery] = useState('')
  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')

  const handleFiles = async (files) => {
    setBusy(true)
    try {
      let loaded = []
      for (const file of files) {
        setStatus(`קורא את ${file.name}...`)
        const rows = await readWorkbook(file)
        const kind = classifyFile(rows)
        if (kind === 'production') setProduction(rows)
        if (kind === 'quality') setQuality(rows)
        if (kind === 'deviations') setDeviations(rows)
        loaded.push(`${file.name}: ${rows.length.toLocaleString('he-IL')} רשומות`)
      }
      setStatus(`נטען בהצלחה — ${loaded.join(' | ')}`)
    } catch (e) {
      console.error(e)
      setStatus(`שגיאה בקריאת קובץ: ${e.message}`)
    } finally { setBusy(false) }
  }

  const prod = useMemo(() => production.map(r => ({
    facility: normalize(r['Storage Location']),
    date: excelDate(r['Actual finish date'] || r['Release date (actual)']),
    qty: num(r['Delivered quantity (GMEIN)'] || r['Confirmed Yield Quantity (GMEIN)']),
    order: normalize(r['Order']), batch: normalize(r['Batch']), material: normalize(r['Material']),
    desc: normalize(r['Material description']), orderType: normalize(r['Order Type']),
  })).filter(r => r.facility), [production])

  const dateBounds = useMemo(() => {
    const ds = prod.map(r=>r.date).filter(Boolean).sort((a,b)=>a-b)
    return { min: iso(ds[0]), max: iso(ds.at(-1)) }
  }, [prod])

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase()
    return prod.filter(r => {
      const d = iso(r.date)
      const okDate = (!from || d >= from) && (!to || d <= to)
      const okQ = !q || [r.facility,r.order,r.batch,r.material,r.desc].some(v=>v.toLowerCase().includes(q))
      return okDate && okQ
    })
  }, [prod,from,to,query])

  const days = useMemo(() => new Set(filtered.map(r=>iso(r.date)).filter(Boolean)).size || 1, [filtered])
  const facilityStats = useMemo(() => FACILITIES.map(id => {
    let rows = filtered.filter(r=>r.facility === id)
    if (id === '1542') rows = rows.filter(r=>r.orderType.includes('ZFIN'))
    const actual = rows.reduce((s,r)=>s+r.qty,0)
    const target = TARGETS[id] * days
    return { id, actual, target, pct: target ? Math.round(actual/target*100) : 0, orders: new Set(rows.map(r=>r.order).filter(Boolean)).size }
  }), [filtered,days])

  const total = facilityStats.reduce((s,x)=>s+x.actual,0)
  const activeFacilities = facilityStats.filter(x=>x.actual>0).length
  const deviationRows = deviations.filter(r => {
    const st = normalize(r['QA Status']).toLowerCase()
    return st && !st.includes('מאושר')
  })
  const qualityBad = quality.filter(r => {
    const st = normalize(r['Result Status']).toLowerCase()
    return st && !st.includes('accepted') && !st.includes('תקין') && !st.includes('pass')
  }).length

  return <div className="dashboard" dir="rtl">
    <aside className="side">
      <div className="brand">IML<span>CONTROL</span></div>
      <div className="side-stat"><Database/><div><b>{fmt(production.length)}</b><small>רשומות תפוקה</small></div></div>
      <div className="side-stat"><FlaskConical/><div><b>{fmt(quality.length + deviations.length)}</b><small>רשומות איכות</small></div></div>
      <div className="side-note">הנתונים נשארים במחשב שלך בלבד ואינם נשלחים לשרת.</div>
    </aside>

    <main className="main">
      <header className="header">
        <div><h1>מרכז שליטה למתקני אריזה</h1><p>טעינת Excel, חישוב KPI וסינון נתונים בזמן אמת</p></div>
        <label className={`upload ${busy?'disabled':''}`}><Upload size={19}/>{busy?'טוען...':'טעינת קובצי Excel'}<input type="file" multiple accept=".xlsx,.xls" disabled={busy} onChange={e=>handleFiles([...e.target.files])}/></label>
      </header>

      <div className="load-status"><CheckCircle2 size={18}/>{status}</div>

      <section className="filters">
        <label><Search size={17}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="חיפוש מתקן, Order, Batch או חומר"/></label>
        <label><CalendarDays size={17}/>מתאריך<input type="date" min={dateBounds.min} max={dateBounds.max} value={from} onChange={e=>setFrom(e.target.value)}/></label>
        <label><CalendarDays size={17}/>עד תאריך<input type="date" min={dateBounds.min} max={dateBounds.max} value={to} onChange={e=>setTo(e.target.value)}/></label>
        <button onClick={()=>{setFrom('');setTo('');setQuery('')}}>נקה סינון</button>
      </section>

      <section className="summary-grid">
        <Summary title="סה״כ תפוקה" value={fmt(total)} sub="ליטר / ק״ג לפי הקובץ"/>
        <Summary title="מתקנים פעילים" value={activeFacilities} sub={`מתוך ${FACILITIES.length} מתקנים`}/>
        <Summary title="ימי פעילות" value={days} sub={dateBounds.min ? `${dateBounds.min} עד ${dateBounds.max}` : 'אין נתונים'}/>
        <Summary title="חריגות איכות" value={deviationRows.length + qualityBad} sub="לפי קובצי האיכות" warn/>
      </section>

      <div className="section-title"><Factory/><div><h2>ביצועים לפי מתקן</h2><p>יעד מחושב לפי מספר הימים בטווח הנבחר</p></div></div>
      <section className="facility-grid">
        {facilityStats.map(x=><Facility key={x.id} {...x}/>) }
      </section>

      <section className="details">
        <h2>רשומות תפוקה אחרונות</h2>
        <div className="table-wrap"><table><thead><tr><th>תאריך</th><th>מתקן</th><th>הזמנה</th><th>Batch</th><th>חומר</th><th>כמות</th></tr></thead><tbody>
          {filtered.slice(-100).reverse().map((r,i)=><tr key={i}><td>{iso(r.date)}</td><td>{r.facility}</td><td>{r.order}</td><td>{r.batch}</td><td>{r.desc || r.material}</td><td>{fmt(r.qty)}</td></tr>)}
          {!filtered.length && <tr><td colSpan="6" className="empty">טען את קובץ התפוקות כדי להציג נתונים</td></tr>}
        </tbody></table></div>
      </section>
    </main>
  </div>
}

function Summary({title,value,sub,warn}) { return <div className={`summary ${warn?'warn':''}`}><span>{title}</span><b>{value}</b><small>{sub}</small></div> }
function Facility({id,actual,target,pct,orders}) {
  const capped=Math.min(pct,100)
  const state=pct>=100?'good':pct>=75?'mid':'bad'
  return <article className={`facility ${state}`}>
    <div className="facility-top"><div><small>מתקן</small><h3>{id}</h3></div><b>{pct}%</b></div>
    <div className="bar"><i style={{width:`${capped}%`}}/></div>
    <div className="facility-numbers"><span>בפועל<strong>{fmt(actual)}</strong></span><span>יעד<strong>{fmt(target)}</strong></span><span>הזמנות<strong>{orders}</strong></span></div>
  </article>
}
