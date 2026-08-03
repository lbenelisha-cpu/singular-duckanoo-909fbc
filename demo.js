export default function KpiCard({ item }) {
  return <div className="kpi-card">
    <div className="kpi-label">{item.label}</div>
    <div className="kpi-value">{item.value}<small>{item.unit}</small></div>
    <div className="trend">{item.trend} לעומת התקופה הקודמת</div>
  </div>
}
