# Sprint 11.2.5 — תיקון מלא להעברת API Key

## מה תוקן
- שכבת ה-fetch של Supabase מזהה כעת גם URL objects ולא רק מחרוזות ו-Request.
- כל בקשה לדומיין Supabase מקבלת במפורש apikey.
- נשמר Authorization של המשתמש המחובר; נוסף fallback למפתח הציבורי.
- תוקנה שגיאת 400: No API key found in request.

לא נדרש SQL נוסף. לאחר העלאה ל-GitHub יש להמתין ל-Published ולבצע Ctrl+Shift+R.
