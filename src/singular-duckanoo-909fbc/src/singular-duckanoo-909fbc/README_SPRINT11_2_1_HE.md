# IML Control — Sprint 11.2.1

## מה חדש

- כניסת Viewer קבועה באמצעות `VITE_VIEWER_EMAIL` ו־`VITE_VIEWER_PASSWORD`.
- אין שימוש ב־Supabase Anonymous Sign-In.
- הרשאות: Admin, Manager, Viewer.
- Admin: טעינה, מחיקה וניהול מלא.
- Manager: טעינת נתונים ללא מחיקת כלל הנתונים.
- Viewer: צפייה, סינון, חיפוש וייצוא בלבד.
- סרגל סינון קבוע בצד ימין: Quick Search, חודש, תאריכים ומתקן.

## הגדרת חשבון צפייה

1. צור ב־Supabase Auth משתמש קבוע לצפייה.
2. בטבלת `profiles` הגדר לאותו משתמש `role = viewer` ו־`is_active = true`.
3. ב־Netlify הוסף:
   - `VITE_VIEWER_EMAIL`
   - `VITE_VIEWER_PASSWORD`
4. בצע Deploy מחדש.

> הערה: משתני Vite נכללים בבניית צד הלקוח. ההגנה האמיתית חייבת להישען על RLS והרשאות Supabase, ולא על הסתרת הסיסמה בלבד.
