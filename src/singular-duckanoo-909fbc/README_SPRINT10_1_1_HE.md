# IML Control — Sprint 10.1.1

גרסה זו מתקנת את חוסר ההתאמה בין קוד האתר לסכמת Supabase.

## סדר התקנה

1. ב־Supabase פתח SQL Editor.
2. פתח את `SPRINT_10_1_1_MIGRATION.sql`, העתק את כל הקובץ ולחץ Run.
3. בשורת התוצאה ודא שמופיע `true` בשלושת השדות:
   - `versions_table_ready`
   - `chunks_table_ready`
   - `active_version_column_ready`
4. העלה את הפרויקט ל־GitHub.
5. ב־Netlify בחר `Clear cache and deploy site`.
6. לאחר הכניסה ודא שמופיע `Sprint 10.1.1 — Schema Recovery & Data Engine`.

## התנהגות בטוחה

אם ההגירה עדיין לא הותקנה, האתר ימשיך לקרוא את הנתונים הישנים מ־Supabase ולא יעבור בטעות ל־Browser Cache. טעינה חדשה תיחסם עם הוראה ברורה להריץ את קובץ ההגירה.
