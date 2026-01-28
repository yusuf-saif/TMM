# The Market Masters (TMM) — Website (Static)

This is a clean, minimal, modern **multi-page** website built with:

- **HTML / CSS / JavaScript**
- **Bootstrap 5 (CDN)**

## Pages
- `index.html` — Home
- `about.html` — About
- `services.html` — Services
- `events.html` — Events
- `contact.html` — Contact
- `community.html` — Join our Community (Google Form CTA)

## Quick start
Just open `index.html` in a browser.

For best results (so links work like a real site), run a tiny local server:

### Option A: VS Code Live Server
Install the **Live Server** extension, right-click `index.html`, click **Open with Live Server**.

### Option B: Python
```bash
cd tmm-site
python -m http.server 5500
```
Then open: http://localhost:5500

## Things you should update
### 1) Community Google Form link
Open:
- `community.html`

Replace:
`https://forms.gle/REPLACE_WITH_YOUR_FORM_LINK`

### 2) Contact form (currently demo)
Open:
- `contact.html`

Replace the `<form action="#">` with your endpoint:
- Formspree / Getform / custom backend, etc.

### 3) Events list
Open:
- `assets/js/main.js`

Edit the `eventsData` array.

## Assets
- Logo: `assets/img/logo.jpg`
- Custom styles: `assets/css/styles.css`
- Custom JS: `assets/js/main.js`

---
Built for TMM. Keep it minimal. Keep it sharp.
