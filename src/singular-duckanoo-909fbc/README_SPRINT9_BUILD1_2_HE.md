# Sprint 9.1.2 – תיקון חיבור והרשאות

## מה תוקן

- נוסף מסך תיקון כתובת Supabase מתוך מסך הכניסה.
- הכתובת המתוקנת נשמרת בדפדפן ומחליפה את כתובת Netlify רק במחשב שבו בוצע התיקון.
- ניתן לחזור בכל עת לכתובת Netlify.
- לפני כניסה מתבצעת בדיקת חיבור אמיתית ל־Supabase.
- האפליקציה ממשיכה לעבוד במצב Viewer גם אם טבלת `profiles` עדיין לא הותקנה.
- קובץ `supabase_schema.sql` יוצר פרופילים גם למשתמשים שכבר קיימים ומגדיר את lbenelisha@gmail.com כמנהל.

## התקנה

1. העלה את הפרויקט ל־GitHub/Netlify ובצע Deploy חדש.
2. השאר ב־Netlify את המשתנים:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
3. ב־Supabase פתח SQL Editor והרץ פעם אחת את `supabase_schema.sql`.
4. במקרה של שגיאת DNS, לחץ במסך הכניסה על "הגדרות חיבור", הדבק את Project URL מחלון Connect של Supabase ולחץ "שמור ובדוק מחדש".
