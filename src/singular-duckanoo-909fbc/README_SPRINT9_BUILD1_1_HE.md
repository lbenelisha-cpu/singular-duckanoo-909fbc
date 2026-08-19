# Sprint 9.1.1 – תיקון חיבור Supabase

בגרסה זו נוספו:

- ניקוי אוטומטי של `/rest/v1/` או `/auth/v1/` מתוך `VITE_SUPABASE_URL`.
- תמיכה גם בשם המשתנה הרשמי החדש `VITE_SUPABASE_PUBLISHABLE_KEY`.
- בדיקת תקשורת מוקדמת מול Supabase Auth.
- הודעת שגיאה ברורה כאשר שם המתחם אינו נגיש או הוקלד באופן שגוי.
- טיפול בטוח בשגיאות רשת בזמן הכניסה.

## משתנים ב-Netlify

```text
VITE_SUPABASE_URL=https://PROJECT-REF.supabase.co
VITE_SUPABASE_ANON_KEY=sb_publishable_...
```

אפשר להשתמש במקום המשתנה השני גם ב:

```text
VITE_SUPABASE_PUBLISHABLE_KEY=sb_publishable_...
```

לאחר שינוי משתנים יש לבצע:

`Deploys → Trigger deploy → Clear cache and deploy site`
