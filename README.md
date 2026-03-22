# VulnApp Lite + Jules Bot — Presentation

Visual PowerPoint presentation for **VulnApp Lite** (vulnerability tracker) and **Jules** (Telegram automation bot).

## Slides

| # | Title | Visual Type |
|---|-------|-------------|
| 01 | Tytuł i Wizja | Feature cards, tech stack |
| 02 | Architektura Systemu | System architecture diagram |
| 03 | Logika Bota Jules | 3-column flowchart |
| 04 | Mapa Interakcji Użytkownika | Timeline user journey map |
| 05 | Statystyki i Wydajność | KPI dashboard + mock charts |

## Tech Stack (Presented)

- **Backend:** Flask + Gunicorn + SQLite (rafserver)
- **Proxy:** Nginx (port 80) + ZeroTier VPN
- **Mobile:** VulnApp iOS (Swift, read-only mirror via `/api/cves`)
- **Bot:** Jules — `jules-listener.service` + `jules_actions.sh` + `jules_review.sh`
- **Process manager:** systemd

## Generate

```bash
pip install python-pptx
python3 generate_presentation.py
```

Output: `VulnApp_Lite_Jules_Bot_Presentation.pptx`

## Color Palette

| Color | Hex | Usage |
|-------|-----|-------|
| Cyan | `#00E5FF` | Primary accent, listener flow |
| Green | `#00E676` | Success states, architecture |
| Amber | `#FFBF00` | Bot triggers, warnings |
| Red | `#FF4545` | Critical CVEs, GitHub actions |
| Dark BG | `#0D1B2A` | Background |
