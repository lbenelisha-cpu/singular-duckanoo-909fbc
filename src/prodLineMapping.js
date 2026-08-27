const clean = value => String(value ?? '').trim().toUpperCase()

// IML CONTROL — approved PROD LINE master mapping (27/08/2026).
// PROD LINE is the fastest/most precise first-pass assignment. Existing SAP
// Storage/Material rules remain as fallback for resources not listed here.
const EXACT = {
  // Facility 42 packaging
  '42-P-02': { facility:'1542', resource:'LQ 1LT (42)', tool:'1 liter' },
  '42-P-03': { facility:'1542', resource:'LQ 5 LT (42)', tool:'5 liter' },
  '42-P-04': { facility:'1542', resource:'LQ 10/20 LT (42)', tool:'10/20 liter' },

  // Facility 43 packaging
  '43-P-A': { facility:'1543', resource:'LQ 43', tool:'43-A-1/5/10' },
  '43-P-B': { facility:'1543', resource:'LQ 43', tool:'43-B-5/10' },

  // Facility 19
  '19-F-02': { facility:'1519', resource:'WG (19)', tool:'south - 19' },
  '19-F-12': { facility:'1519', resource:'WG (19)', tool:'north - 19' },
  '19-F-13': { facility:'1519', resource:'WG (19)', tool:'new - 19' },
  '19-F-14': { facility:'1519', resource:'WG (19)', tool:'north - 19 - pilot' },
  '19-P-03': { facility:'1519', resource:'WG SMALL PACKS (19)', tool:'284' },
  '19-P-04': { facility:'1519', resource:'WG SMALL PACKS (19)', tool:'255' },
  '19-P-05': { facility:'1519', resource:'WG SMALL PACKS (19)', tool:'brevis' },
  'FCLT-19': { facility:'1519', resource:'WG SMALL PACKS (19)', tool:'Facility 19' },

  // Facility 23 — one dashboard facility, Excel keeps vessel detail
  '23-F-01': { facility:'1523', resource:'EC (23)', tool:'R-120', displayFacility:'23' },
  '23-F-02': { facility:'1523', resource:'EC (23)', tool:'R-125', displayFacility:'23' },
  '23-F-03': { facility:'1523', resource:'EC (23)', tool:'R-243', displayFacility:'23' },
  '23-F-04': { facility:'1523', resource:'EC (23)', tool:'R-246', displayFacility:'23' },
  '23-F-05': { facility:'1523', resource:'EC (23)', tool:'R247', displayFacility:'23' },
  '23-F-06': { facility:'1523', resource:'EC (23)', tool:'R-220', displayFacility:'23' },
  '23-F-07': { facility:'1523', resource:'EC (23)', tool:'R-902', displayFacility:'23' },
  '23-F-08': { facility:'1523', resource:'EC (23)', tool:'R-934', displayFacility:'23' },
  '23-F-09': { facility:'1523', resource:'EC (23)', tool:'R-202', displayFacility:'23' },
  '23-P-247': { facility:'1523', resource:'EC (23)', tool:'R247 ISOTANK PACK', displayFacility:'23' },
  'FCLT-23': { facility:'1523', resource:'EC (23)', tool:'Facility 23', displayFacility:'23' },

  // Facility 24
  '24-F-02': { facility:'1524', resource:'24F', tool:'R-129' },
  '24-P-03': { facility:'1524', resource:'24F', tool:'R-127' },
  '24-F-12': { facility:'1524', resource:'24F128', tool:'R-128' },

  // Facility 28 — one dashboard facility, Excel keeps area detail
  '28-F-01': { facility:'1528', resource:'SC (28)', tool:'מזרחי', displayFacility:'28' },
  '28-F-02': { facility:'1528', resource:'SC (28)', tool:'חדש', displayFacility:'28' },
  '28-F-03': { facility:'1528', resource:'SC (28)', tool:'מרכזי', displayFacility:'28' },
  '28-F-04': { facility:'1528', resource:'SC (28)', tool:'מערבי', displayFacility:'28' },
  '28-F-05': { facility:'1528', resource:'SC (28)', tool:'KELZAN', displayFacility:'28' },

  // Facility 40 approved cards
  '40-T-12': { facility:'1540', resource:'DIURON (40)', tool:'diuron - B' },
  '40-T-13': { facility:'1540', resource:'TOLUREX (40)', tool:'cotoran/tolurex/protugan' },
  '40-T-14': { facility:'1540', resource:'DIURON (40)', tool:'diuron - D' },
}

const FACILITY_25_TOOLS = {
  '25-F-01':'Sodium 45% Tank farm', '25-F-13':'flurochloridone cs pendimethlin cs',
  '25-F-14':'propaquizafop', '25-F-15':'galigan',
  '25-S-02':'uracil', '25-S-03':'bromacil', '25-S-04':'krovar', '25-S-05':'propaquizafop',
  '25-S-07':'flurochloridone', '25-S-08':'Koh Solu In Ethn 22.5%', '25-S-09':'Ethoxy (Galigan)',
  '25-S-10':'Galigan in Toluene', '25-S-11':'galigan', '25-S-12':'imazamox', '25-S-16':'Aganile crude',
  '25-S-17':'aganil', '25-S-18':'AP FLASH ester', '25-S-19':'AP FLASH', '25-S-20':'ap 200',
  '25-S-21':'pomelone', '25-S-22':'saflufenacil', '25-S-28':'synth', '25-S-29':'cottoran',
  '25-S-30':'tolurex', '25-S-31':'protugan', '25-S-38':'synth', '25-S-48':'synth',
}

export const normalizeProdLine = clean

export function prodLineInfo(value) {
  const code = clean(value)
  if (!code) return null
  if (EXACT[code]) return { code, ...EXACT[code] }
  // User-approved rule: all listed 25-F-* and 25-S-* resources belong to EC 25.
  if (/^25-(F|S)-/.test(code)) {
    return { code, facility:'1525', resource:'EC (25)', tool:FACILITY_25_TOOLS[code] || code }
  }
  return null
}

export function isExcludedProdLine(value) {
  const code = clean(value)
  if (!code) return false
  // FCLT rows are not production by default. Only FCLT-19 and FCLT-23 were
  // explicitly approved above and therefore return false here.
  return /^FCLT-/.test(code) && !EXACT[code]
}

export function excelFacilityLabel(row) {
  const info = prodLineInfo(row?.prodLine)
  if (info?.displayFacility) return info.displayFacility
  const facility = String(row?.facility || '')
  if (facility === '1523') return '23'
  if (facility === '1528') return '28'
  return facility
}
