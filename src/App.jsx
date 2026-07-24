import { useState } from 'react'
import { Bell, Search, CalendarDays, FlaskConical, ArrowLeft } from 'lucide-react'
import Sidebar from './components/Sidebar'
import KpiCard from './components/KpiCard'
import FacilityCard from './components/FacilityCard'
import BatchCard from './components/BatchCard'
import { kpis, facilities } from './data/demo'
import './styles.css'

export default function App(){
  const [active,setActive]=useState('סקירה ראשית')
  const [batchOpen,setBatchOpen]=useState(true)
  return <div className="app-shell">
    <Sidebar active={active} onSelect={setActive}/>
    <main>
      <header className="topbar"><div><h1>IML CONTROL CENTER</h1><p>מרכז שליטה לייצור, אריזה ואיכות</p></div><div className="top-actions"><label className="global-search"><Search size={18}/><input placeholder="חיפוש Batch, חומר או הזמנה"/></label><button><CalendarDays/></button><button><Bell/></button><div className="avatar">LB</div></div></header>
      <section className="content">
        <div className="page-title"><div><h2>סקירה ראשית</h2><p>תמונת מצב עדכנית של מתקני האריזה</p></div><span>עודכן היום · 14:30</span></div>
        <div className="kpi-grid">{kpis.map(item=><KpiCard key={item.label} item={item}/>)}</div>
        <div className="section-head"><div><h3>מצב מתקנים</h3><p>ביצועים מול יעד יומי</p></div><button>לכל המתקנים <ArrowLeft size={17}/></button></div>
        <div className="facility-grid">{facilities.map(item=><FacilityCard key={item.id} item={item}/>)}</div>
        <div className="quality-banner"><div className="quality-icon"><FlaskConical/></div><div><h3>התראת איכות פעילה</h3><p>Batch 352 · FLUROXYPR-MEPT · תוצאה 11,742 PPM מול גבול עליון 1 PPM</p></div><button onClick={()=>setBatchOpen(true)}>פתח כרטיס מנה</button></div>
      </section>
    </main>
    {batchOpen && <BatchCard onClose={()=>setBatchOpen(false)}/>} 
  </div>
}
