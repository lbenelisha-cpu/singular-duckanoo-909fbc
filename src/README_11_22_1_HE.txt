IML CONTROL 11.22.1 — תיקון PowerPoint

מה תוקן:
- הייצוא הישן יצר קובץ HTML עם סיומת .ppt ולכן PowerPoint סירב לפתוח אותו.
- כעת האפליקציה יוצרת קובץ PowerPoint אמיתי בפורמט .pptx באמצעות PptxGenJS.
- שם הקובץ מסתיים ב-.pptx ונפתח ישירות ב-Microsoft PowerPoint.

קבצים להחלפה/להוסיף:
1. package.json
2. src/DashboardApp.jsx
3. public/version.json

חשוב: package.json כולל תלות חדשה pptxgenjs. ב-Netlify npm install יתקין אותה אוטומטית בזמן Build.
אין צורך ב-SQL ואין שינוי ב-Supabase.
