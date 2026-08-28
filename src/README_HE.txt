תיקון Build עבור xlsx-js-style

להחליף בשורש הפרויקט:
1. package.json
2. netlify.toml

אין צורך לשנות שוב את DashboardApp.jsx אם כבר הוחלף בעדכון RTL האחרון.

הסיבה: package-lock.json הישן לא הכיל את xlsx-js-style, ולכן Netlify בנה עם dependencies ישנים וה-import נכשל.
הפקודה החדשה מריצה npm install לפני build ומסנכרנת את התלות החדשה.
