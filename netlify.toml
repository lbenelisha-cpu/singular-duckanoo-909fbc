import { X, Search, Download } from 'lucide-react'
import { labResults } from '../data/demo'

export default function BatchCard({ onClose }) {
  return <div className="modal-backdrop"><section className="batch-card">
    <header><div><h2>כרטיס מנה — Batch 352</h2><p>Material 10000000999 · כל תוצאות המעבדה</p></div><button onClick={onClose} className="icon-btn"><X/></button></header>
    <div className="batch-summary">
      {[['כמות מנה','8,800'],['מתקן','1523'],['מק״ט','10000000999'],['Inspection Lot','40000235230'],['Inspection Lot תוצאות','40000232978'],['UD Code','A2-Rejected'],['QA Status','אין בדיקה'],['פקודת ייצור','—']].map(([k,v])=><div key={k}><span>{k}</span><strong>{v}</strong></div>)}
    </div>
    <div className="batch-toolbar"><div className="filter-pills"><button className="selected">הכול 12</button><button>תקין 11</button><button>גבולי 0</button><button className="danger">חריג 1</button></div><label><Search size={17}/><input placeholder="חיפוש בדיקה..."/></label><button className="export"><Download size={17}/>ייצוא</button></div>
    <div className="table-wrap"><table><thead><tr><th>מאפיין</th><th>גבול תחתון</th><th>ערך בפועל</th><th>גבול עליון</th><th>יחידה</th><th>סטטוס</th></tr></thead><tbody>{labResults.map(r=><tr key={r.characteristic} className={r.status==='חריג'?'rejected':''}><td>{r.characteristic}</td><td>{r.lower}</td><td>{r.actual}</td><td>{r.upper}</td><td>{r.unit}</td><td><span className={`result ${r.status==='חריג'?'bad':'good'}`}>{r.status}</span></td></tr>)}</tbody></table></div>
  </section></div>
}
