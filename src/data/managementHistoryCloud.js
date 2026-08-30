import { supabase } from '../supabase'

const monthKey = value => {
  const s = String(value || '')
  return s.length >= 7 ? s.slice(0, 7) : ''
}

export async function loadManagementHistoryFromCloud(fallbackHistory) {
  if (!supabase) return { history: fallbackHistory, source: 'embedded', error: null }
  try {
    const [planRes, contractorRes] = await Promise.all([
      supabase.from('iml_management_plan_actual').select('month,facility,plan,source_actual,groups,source_label,updated_at').order('month', { ascending: true }),
      supabase.from('iml_management_contractor_costs').select('month,facility,packaged,cost,cost_per_unit,lines,shift_qty,shift_cost,source_label,updated_at').order('month', { ascending: true }),
    ])
    if (planRes.error) throw planRes.error
    if (contractorRes.error) throw contractorRes.error
    if (!(planRes.data || []).length && !(contractorRes.data || []).length) {
      return { history: fallbackHistory, source: 'embedded', error: null }
    }
    const planActual = {}
    ;(planRes.data || []).forEach(row => {
      const key = monthKey(row.month)
      if (!key) return
      if (!planActual[key]) planActual[key] = {}
      planActual[key][String(row.facility)] = {
        plan: Number(row.plan || 0),
        actual: Number(row.source_actual || 0),
        groups: row.groups || {},
      }
    })
    const contractor42 = {}
    ;(contractorRes.data || []).forEach(row => {
      const key = monthKey(row.month)
      if (!key || String(row.facility) !== '42') return
      contractor42[key] = {
        facility: '42',
        packaged: Number(row.packaged || 0),
        cost: Number(row.cost || 0),
        costPerUnit: Number(row.cost_per_unit || 0),
        lines: row.lines || {},
        shiftQty: Array.isArray(row.shift_qty) ? row.shift_qty : [],
        shiftCost: Array.isArray(row.shift_cost) ? row.shift_cost : [],
      }
    })
    const contractorKeys = Object.keys(contractor42).sort()
    return {
      history: {
        planActual,
        contractor42,
        meta: {
          ...(fallbackHistory?.meta || {}),
          source: 'supabase',
          contractor2026Through: contractorKeys.filter(k => k.startsWith('2026-')).slice(-1)[0] || fallbackHistory?.meta?.contractor2026Through || '',
        },
      },
      source: 'supabase',
      error: null,
    }
  } catch (error) {
    return { history: fallbackHistory, source: 'embedded', error: error?.message || String(error) }
  }
}
