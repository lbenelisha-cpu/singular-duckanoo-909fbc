# תיקון Netlify — Sprint 11.4.3 Build 3

קובץ `netlify.toml` חייב להישמר בשורש המאגר בדיוק בתוכן הבא:

```toml
[build]
  command = "npm run build"
  publish = "dist"

[build.environment]
  NODE_VERSION = "20"

[[redirects]]
  from = "/*"
  to = "/index.html"
  status = 200
```

לאחר החלפת הקובץ יש לבצע ב-Netlify:
Deploys → Trigger deploy → Clear cache and deploy site.
