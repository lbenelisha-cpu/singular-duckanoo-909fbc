IML CONTROL 11.22.2 — תיקון Build למצגת PowerPoint

מה תוקן:
- הוסר import של pptxgenjs מתוך npm שגרם לכשל Build ב-Netlify.
- נוסף מנוע PowerPoint כקובץ browser מקומי בתוך public/pptxgen.bundle.js.
- המנוע נטען רק כאשר לוחצים על "הפק PowerPoint".
- אין צורך לשנות package.json או package-lock.json.
- אין צורך להריץ SQL.

קבצים להחלפה/הוספה:
1. src/DashboardApp.jsx
2. public/pptxgen.bundle.js  (חדש)
3. public/version.json

לא להעתיק את package.json מ-11.22.1. אם כבר הוחלף אצלך, החזר את package.json לגרסה שלפני 11.22.1 או הסר ממנו את שורת pptxgenjs.
