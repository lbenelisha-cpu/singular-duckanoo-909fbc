export default function FacilityCard({ item }) {
  const pct = Math.min(100, Math.round(item.actual/item.target*100))
  return <div className="facility-card">
    <div className="facility-head"><div><strong>{item.id}</strong><span>{item.name}</span></div><span className={`state ${item.status==='פעיל'?'ok':'wait'}`}>{item.status}</span></div>
    <div className="progress"><i style={{width:`${pct}%`}}/></div>
    <div className="facility-foot"><span>{item.actual.toLocaleString()} בפועל</span><strong>{pct}%</strong><span>{item.target.toLocaleString()} יעד</span></div>
  </div>
}
