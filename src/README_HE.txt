IML CONTROL - Excel style root fix

IMPORTANT:
1. Replace package.json in the PROJECT ROOT (same folder as netlify.toml), NOT src/package.json.
2. Replace src/DashboardApp.jsx.
3. Do NOT replace netlify.toml.

Verification in Netlify log:
Before vite build you MUST see:
IML_STYLE_PATCH_ACTIVE

If that line does not appear, the root package.json was not replaced.
