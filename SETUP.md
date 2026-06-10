# SDR.qualify v3 — Setup Guide

## Environment Variables (Vercel Dashboard → Settings → Environment Variables)

| Variable | Value |
|----------|-------|
| `PIPEDRIVE_CLIENT_ID` | מ-Pipedrive Developer Hub |
| `PIPEDRIVE_CLIENT_SECRET` | מ-Pipedrive Developer Hub |
| `PIPEDRIVE_REDIRECT_URI` | `https://YOUR-APP.vercel.app/api/auth/callback` |
| `SESSION_SECRET` | מחרוזת רנדומלית ארוכה (לפחות 32 תווים) |
| `ADMIN_EMAIL` | `roye@contel.co.il` |

## Pipedrive Developer Hub Setup

1. כנס ל: https://developers.pipedrive.com → "My Apps" → Create App
2. Callback URL (חובה מדויק): `https://YOUR-APP.vercel.app/api/auth/callback`
3. Scopes: leads, organizations, persons, notes, users
4. העתק Client ID + Client Secret → הכנס ל-Vercel

## Deploy Steps

1. **GitHub**: העלה את כל הקבצים ל-repo
2. **Vercel**: Import repo → Deploy
3. **Env Vars**: הגדר את כל המשתנים למעלה → Redeploy
4. **Pipedrive App**: הגדר Callback URL → שמור

## Test Checklist

- [ ] פתח את האפליקציה → רואים מסך Login
- [ ] לחץ "התחבר באמצעות Pipedrive" → מועבר ל-Pipedrive
- [ ] אחרי אישור → חוזר לאפליקציה עם שם המשתמש בכותרת
- [ ] צור ליד חדש → הקלד שם חברה → רואים חיפוש ב-Pipedrive
- [ ] לחץ "צור ב-Pipedrive" → מוצג Lead ID בכרטיסייה
- [ ] בדוק ב-Pipedrive Leads Inbox → הליד מופיע
- [ ] כנס ל: /admin → רואים את כל הלידים

## Architecture

```
Browser → /api/auth/login → Pipedrive OAuth → /api/auth/callback
                                                      ↓
                                              Session Cookie (encrypted)
                                                      ↓
Browser → /api/pd/search-orgs   → Pipedrive API (user token)
Browser → /api/pd/search-persons → Pipedrive API (user token)
Browser → /api/pd/create-lead   → Pipedrive API (user token)
                                    1. Create/find Org
                                    2. Create/find Person
                                    3. Create Lead in Inbox
                                    4. Add Note with BANT data
Browser → /api/admin/leads      → Pipedrive API (admin token)
```
