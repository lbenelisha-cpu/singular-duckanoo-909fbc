# Sprint 9.1 — כניסה והרשאות

הגרסה מוסיפה מסך כניסה באמצעות Supabase Auth ושלושה תפקידים:
- admin: טעינה, מחיקה, צפייה וייצוא
- manager: צפייה וייצוא
- viewer: צפייה וייצוא

## הקמה
1. צור פרויקט Supabase.
2. הרץ את `supabase_schema.sql` ב-SQL Editor.
3. צור משתמש ראשון ב-Authentication > Users.
4. עדכן אותו ל-admin לפי הפקודה שבתחתית קובץ ה-SQL.
5. ב-Netlify הוסף:
   - VITE_SUPABASE_URL
   - VITE_SUPABASE_ANON_KEY
6. בצע Deploy מחדש.

הערה: Build זה מוסיף אימות והרשאות. הנתונים עצמם עדיין נשמרים מקומית בדפדפן; העברתם למסד המשותף תתבצע ב-Build 9.2.
