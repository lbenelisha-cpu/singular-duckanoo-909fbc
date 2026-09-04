# IML CONTROL — Sprint 11.9.25

עדכון טעינת יעדים חודשיים:

- ערכי `PLAN` מוכפלים ב־1,000 פעם אחת בלבד.
- כאשר `PLAN` ריק או אפס, היעד נלקח מ־`CAPACITY` ומוכפל ב־1,000.
- כרטיסיות התכנון נוצרות רק ממשאבי היעד העסקיים המאושרים בקובץ החודשי.
- תפוקה ללא יעד אינה יוצרת כרטיסייה חדשה ואינה מקבלת יעד של מוצר אחר.
- נוספה בדיקת רגרסיה מול קובץ ספטמבר עבור CS, Saflufenacil Tech, Galigan, D. Damascone והחרגות Galigan ISO, LINURON ו־Phenol Oxime.

בדיקות שבוצעו:

- `npm run verify:september-targets -- <targets.xlsx>`
- `npm run build`
