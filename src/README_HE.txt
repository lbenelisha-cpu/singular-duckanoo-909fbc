IML CONTROL — New ADAMA IML Brand Icon

להחליף:
- index.html
- public/manifest.webmanifest
- public/apple-touch-icon.png
- public/favicon-32.png
- public/icon-192.png
- public/icon-512.png
- כל הקבצים המצורפים תחת public/icons/

מה משתנה:
- אייקון חדש ADAMA + IML בתוך האפליקציה.
- אייקון חדש לקיצור דרך באייפון.
- אייקון חדש ב-PWA / מסך הבית.
- favicon חדש בדפדפן.
- נוספה הפניית manifest תקינה ו-cache-busting כדי שהאייקון החדש ייטען.

לא נוגעים ב:
- DashboardApp.jsx
- cloudData.js
- styles.css
- Supabase
- package.json
- netlify.toml

באייפון:
לאחר Deploy, אם קיצור הדרך הישן עדיין מציג את האייקון הקודם,
מחק את הקיצור ממסך הבית והוסף אותו מחדש דרך Safari > שיתוף > הוסף למסך הבית.
