# Anniks Sales Tracker

## Project
Google Apps Script project for a sales tracking sheet with a mobile-friendly web form for a beauty/wellness business.

## Files
- `setup_sheet.gs` — main Apps Script: sheet setup, dropdowns, custom menu, web app entry point, dialog data/submit functions, sales person management
- `dialog.html` — full-page mobile web app form (dark luxury design, Cormorant Garamond + Outfit fonts)
- `appsscript.json` — Apps Script manifest (timezone: Asia/Kuala_Lumpur)
- `.clasp.json` — links this folder to the Google Apps Script project
- `.claude/settings.json` — auto-push hook: editing `.gs` or `.html` triggers `clasp push --force` automatically

## Sheet Columns
Date | No | Redeem Type | Package | Trial | Product | Amount | Payment Method | Sales Person | Remark

## Dropdowns
- **Redeem Type** (C): New, Existing
- **Package** (D): P6880 脸部塑型, P6880 开肩, P6880 体态, P6880 祈龄, P6880 局部, P4880 高级波肽, Gold 脸部塑型, Gold 开肩, T2388 小腿
- **Trial** (E): Yes, No
- **Product** (F): T388 脸部塑型, T388 祈龄魔法, T298 体态, Firming Cream
- **Payment Method** (H): Cash, Card, Online Transfer, Debit Card, Credit Card, QR, Transfer
- **Sales Person** (I): Florence, Annika, Celine, Jane, KitKit, Tracy — managed via ⚙️ Manage menu, stored in Script Properties, `allowInvalid: true` so deleted persons don't error existing rows

## Custom Menu (desktop only)
`⚙️ Manage` menu appears on sheet open via `onOpen()`:
- ➕ New Redeem Entry — opens modal dialog
- Add Sales Person — prompt dialog, saves to Script Properties
- Remove Sales Person — prompt dialog, updates Script Properties

## Web App (mobile)
- Accessed via deployed URL (Google-hosted, no external server needed)
- `doGet()` serves `dialog.html` as standalone page
- Full-page dark luxury UI: charcoal bg, gold accents, sticky submit button
- After every `clasp push`, must **manually redeploy**: Extensions → Apps Script → Deploy → Manage deployments → edit → New version → Deploy

## Auto-push
`.claude/settings.json` PostToolUse hook runs `clasp push --force` whenever a `.gs` or `.html` file is edited via Claude Code.

## Script ID
`1jXdznVj23UEHPmRhR2z5-t1HjfmshohGZUDVmP5fcDr31kx8strZCzYR`

## clasp Auth
Logged in as `kokhou.choi@gmail.com`. Apps Script API enabled at script.google.com/home/usersettings.
