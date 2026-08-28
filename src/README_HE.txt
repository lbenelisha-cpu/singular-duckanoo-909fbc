IML CONTROL – Excel RTL & Style Update

להחליף בפרויקט:
1. src/DashboardApp.jsx
2. package.json

מה תוקן:
- קובץ XLSX אמיתי.
- גיליון דיווח יומי פורמולציות מימין לשמאל.
- כותרת כהה ומודגשת כמו בדוחות הקודמים.
- גבולות מלאים לכל תאי הטבלה.
- הדגשת מתקן וסה״כ תפוקה.
- שמירה על Short Description בעמודת קו ייצור.
- עיצוב RTL גם בגיליונות הנוספים.

הערה: package.json כולל xlsx-js-style כדי שהעיצוב יישמר בתוך XLSX אמיתי. Netlify יתקין את התלות בזמן Build.
