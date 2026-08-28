IML CONTROL - Netlify build fix
החלף רק את src/DashboardApp.jsx.
התיקון מסיר את התלות xlsx-js-style ומחזיר שימוש בחבילת xlsx שכבר קיימת בפרויקט.
אין צורך לשנות package.json או netlify.toml.
