# הגדרת GitHub ו-Netlify — Sprint 11.3.2

## מבנה המאגר
יש להעלות את תוכן התיקייה הזאת ישירות לשורש מאגר GitHub. בשורש חייבים להופיע:

- package.json
- index.html
- netlify.toml
- vite.config.js
- src
- public

אין להעלות תיקייה חיצונית שמכילה את כל הפרויקט.

## הגדרות Netlify
באתר Netlify, תחת Site configuration > Build & deploy:

- Base directory: ריק
- Build command: npm run build
- Publish directory: dist
- Production branch: main

לאחר העדכון יש לבחור:
Deploys > Trigger deploy > Clear cache and deploy site

## בדיקת הצלחה
בלוג הבנייה צריכים להופיע `vite build` ויצירת תיקיית `dist`.
בדף שפורסם קבצי JavaScript צריכים להיטען מכתובת `/assets/index-....js` ולא מ-`/app.js`.

## מדוע הופיע מסך שחור
האתר הישן ביקש `app.js` שלא היה קיים. כלל ההפניה של Netlify החזיר את `index.html` במקום JavaScript, ולכן הדפדפן דיווח `Unexpected token '<'`.
