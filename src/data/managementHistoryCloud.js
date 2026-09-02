import { supabase } from '../supabase'

const monthKey = value => { const s=String(value||''); return s.length>=7?s.slice(0,7):'' }
const monthDate = value => { const key=monthKey(value); return key ? `${key}-01` : '' }
const normalizePlanRow = row => ({ month:monthKey(row.month), facility:String(row.facility||''), plan:Number(row.plan||0), source_actual:Number(row.source_actual||0), groups:row.groups||{}, source_label:row.source_label||'' })
const normalizeCostRow = row => ({ month:monthKey(row.month), facility:String(row.facility||'42'), packaged:Number(row.packaged||0), cost:Number(row.cost||0), cost_per_unit:Number(row.cost_per_unit||0), lines:row.lines||{}, shift_qty:Array.isArray(row.shift_qty)?row.shift_qty:[], shift_cost:Array.isArray(row.shift_cost)?row.shift_cost:[], source_label:row.source_label||'' })

const uniquePayload = rows => {
  const seen=new Map()
  ;(rows||[]).forEach(r=>{
    const key=`${monthKey(r.month)}|${String(r.facility||'')}`
    if(!monthKey(r.month)||!String(r.facility||'')) return
    seen.set(key,r)
  })
  return [...seen.values()]
}

async function classifyExisting(table, rows){
  const normalized=uniquePayload(rows)
  const months=[...new Set(normalized.map(r=>monthDate(r.month)).filter(Boolean))]
  if(!months.length) return {created:normalized.length,updated:0,existingKeys:new Set()}
  const {data,error}=await supabase.from(table).select('month,facility').in('month',months)
  if(error) throw error
  const existingKeys=new Set((data||[]).map(r=>`${monthKey(r.month)}|${String(r.facility||'')}`))
  let created=0,updated=0
  normalized.forEach(r=>{ const k=`${monthKey(r.month)}|${String(r.facility||'')}`; existingKeys.has(k)?updated++:created++ })
  return {created,updated,existingKeys}
}

export async function inspectManagementRows(kind, rows){
  if(!supabase) throw new Error('Supabase אינו מחובר')
  const table=kind==='contractor'?'iml_management_contractor_costs':'iml_management_plan_actual'
  return classifyExisting(table,rows)
}

export async function loadManagementHistoryFromCloud(fallbackHistory) {
  if (!supabase) return { history:fallbackHistory, source:'embedded', error:null }
  try {
    const [planRes,contractorRes]=await Promise.all([
      supabase.from('iml_management_plan_actual').select('month,facility,plan,source_actual,groups,source_label,updated_at').order('month',{ascending:true}),
      supabase.from('iml_management_contractor_costs').select('month,facility,packaged,cost,cost_per_unit,lines,shift_qty,shift_cost,source_label,updated_at').order('month',{ascending:true}),
    ])
    if(planRes.error) throw planRes.error; if(contractorRes.error) throw contractorRes.error
    if(!(planRes.data||[]).length&&!(contractorRes.data||[]).length) return {history:fallbackHistory,source:'embedded',error:null}
    const planActual={}; (planRes.data||[]).forEach(row=>{const r=normalizePlanRow(row);if(!r.month)return;if(!planActual[r.month])planActual[r.month]={};planActual[r.month][r.facility]={plan:r.plan,actual:r.source_actual,groups:r.groups}})
    const contractor42={}; (contractorRes.data||[]).forEach(row=>{const r=normalizeCostRow(row);if(!r.month||r.facility!=='42')return;contractor42[r.month]={facility:'42',packaged:r.packaged,cost:r.cost,costPerUnit:r.cost_per_unit,lines:r.lines,shiftQty:r.shift_qty,shiftCost:r.shift_cost}})
    const mergedPlanActual={...(fallbackHistory?.planActual||{})}
    Object.entries(planActual).forEach(([month,facilities])=>{
      mergedPlanActual[month]={...(mergedPlanActual[month]||{}),...facilities}
    })
    // Cloud rows override matching months, while embedded audited months remain
    // available until they are explicitly uploaded to Supabase.
    const mergedContractor42={...(fallbackHistory?.contractor42||{}),...contractor42}
    const contractorKeys=Object.keys(mergedContractor42).sort()
    return {history:{planActual:mergedPlanActual,contractor42:mergedContractor42,meta:{...(fallbackHistory?.meta||{}),source:'supabase+embedded',contractor2026Through:contractorKeys.filter(k=>k.startsWith('2026-')).slice(-1)[0]||fallbackHistory?.meta?.contractor2026Through||''}},source:'supabase+embedded',error:null}
  } catch(error){return {history:fallbackHistory,source:'embedded',error:error?.message||String(error)}}
}

export async function getManagementCloudStatus(){
  if(!supabase) return {planRows:0,contractorRows:0,lastUpdated:'',error:'Supabase אינו מחובר'}
  try{
    const [p,c]=await Promise.all([
      supabase.from('iml_management_plan_actual').select('updated_at',{count:'exact',head:false}).order('updated_at',{ascending:false}).limit(1),
      supabase.from('iml_management_contractor_costs').select('updated_at',{count:'exact',head:false}).order('updated_at',{ascending:false}).limit(1),
    ])
    if(p.error)throw p.error;if(c.error)throw c.error
    const last=[p.data?.[0]?.updated_at,c.data?.[0]?.updated_at].filter(Boolean).sort().slice(-1)[0]||''
    return {planRows:p.count||0,contractorRows:c.count||0,lastUpdated:last,error:''}
  }catch(error){return {planRows:0,contractorRows:0,lastUpdated:'',error:error?.message||String(error)}}
}

export async function upsertManagementPlanRows(rows){
  if(!supabase) throw new Error('Supabase אינו מחובר')
  const normalized=uniquePayload((rows||[]).map(normalizePlanRow).filter(r=>r.month&&r.facility))
  const payload=normalized.map(r=>({...r,month:monthDate(r.month)}))
  if(!payload.length) throw new Error('לא נמצאו רשומות Plan Vs Actual תקינות')
  const stats=await classifyExisting('iml_management_plan_actual',normalized)
  const {error}=await supabase.from('iml_management_plan_actual').upsert(payload,{onConflict:'month,facility'})
  if(error)throw error
  return {written:payload.length,created:stats.created,updated:stats.updated}
}
export async function upsertManagementContractorRows(rows){
  if(!supabase) throw new Error('Supabase אינו מחובר')
  const normalized=uniquePayload((rows||[]).map(normalizeCostRow).filter(r=>r.month&&r.facility))
  const payload=normalized.map(r=>({...r,month:monthDate(r.month)}))
  if(!payload.length) throw new Error('לא נמצאו רשומות עלות קבלן תקינות')
  const stats=await classifyExisting('iml_management_contractor_costs',normalized)
  const {error}=await supabase.from('iml_management_contractor_costs').upsert(payload,{onConflict:'month,facility'})
  if(error)throw error
  return {written:payload.length,created:stats.created,updated:stats.updated}
}



export async function getManagementUploadHistory(limit=30){
  if(!supabase) return []
  const {data,error}=await supabase.from('iml_management_upload_log').select('id,file_name,data_kind,status,rows_written,periods,uploaded_at,error_message').order('uploaded_at',{ascending:false}).limit(limit)
  if(error){ if(String(error.message||'').includes('iml_management_upload_log')) return []; throw error }
  return data||[]
}
export async function logManagementUpload(entry){
  if(!supabase) return
  const payload={file_name:String(entry.file_name||''),data_kind:String(entry.data_kind||''),status:String(entry.status||'success'),rows_written:Number(entry.rows_written||0),periods:Array.isArray(entry.periods)?entry.periods:[],error_message:String(entry.error_message||'')}
  const {error}=await supabase.from('iml_management_upload_log').insert(payload)
  if(error && !String(error.message||'').includes('iml_management_upload_log')) throw error
}
