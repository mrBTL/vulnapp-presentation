"""
VulnApp Lite + Jules Bot — Prezentacja PowerPoint
Generuje wizualną prezentację projektu.
Kompatybilna z Google Slides (brak add_connector).
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Paleta kolorów ──────────────────────────────────────────────────
BG_DARK      = RGBColor(0x0D, 0x1B, 0x2A)
BG_CARD      = RGBColor(0x11, 0x2D, 0x4A)
ACCENT_CYAN  = RGBColor(0x00, 0xE5, 0xFF)
ACCENT_GREEN = RGBColor(0x00, 0xE6, 0x76)
ACCENT_RED   = RGBColor(0xFF, 0x45, 0x45)
ACCENT_AMBER = RGBColor(0xFF, 0xBF, 0x00)
TEXT_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_GRAY    = RGBColor(0xA0, 0xB4, 0xC8)
TEXT_DARK    = RGBColor(0x0D, 0x1B, 0x2A)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

blank_layout = prs.slide_layouts[6]


# ══════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════

def add_bg(slide, color=BG_DARK):
    bg = slide.shapes.add_shape(1, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()


def add_rect(slide, x, y, w, h, fill_color, line_color=None, line_w=Pt(0)):
    shape = slide.shapes.add_shape(1, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_w
    else:
        shape.line.fill.background()
    return shape


def add_text(slide, text, x, y, w, h,
             font_size=Pt(14), bold=False, color=TEXT_WHITE,
             align=PP_ALIGN.LEFT, italic=False):
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Calibri"
    return txBox


def add_rounded_box(slide, x, y, w, h, fill, text, txt_size=Pt(13),
                    txt_color=TEXT_WHITE, bold=False, line_color=None):
    shape = slide.shapes.add_shape(5, x, y, w, h)  # ROUNDED_RECTANGLE
    shape.adjustments[0] = 0.05
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1.5)
    else:
        shape.line.fill.background()
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = txt_size
    run.font.bold = bold
    run.font.color.rgb = txt_color
    run.font.name = "Calibri"
    return shape


def add_arrow_v(slide, cx, y_top, y_bot, color=ACCENT_CYAN):
    """Pionowa strzałka: cienki prostokąt + mały trójkąt na dole."""
    line_w = Inches(0.04)
    add_rect(slide, cx - line_w/2, y_top, line_w, y_bot - y_top - Inches(0.12), color)
    # Trójkąt (symulacja grotu) — mały kwadrat obrócony
    tip_size = Inches(0.14)
    tip = slide.shapes.add_shape(
        5,  # rounded rect jako grot
        cx - tip_size/2, y_bot - Inches(0.14), tip_size, Inches(0.14)
    )
    tip.fill.solid()
    tip.fill.fore_color.rgb = color
    tip.line.fill.background()


def add_arrow_h(slide, x_left, x_right, cy, color=ACCENT_CYAN):
    """Pozioma strzałka: cienki prostokąt."""
    line_h = Inches(0.04)
    add_rect(slide, x_left, cy - line_h/2, x_right - x_left, line_h, color)
    tip_size = Inches(0.14)
    tip = slide.shapes.add_shape(
        5,
        x_right - Inches(0.14), cy - tip_size/2, Inches(0.14), tip_size
    )
    tip.fill.solid()
    tip.fill.fore_color.rgb = color
    tip.line.fill.background()


def accent_bar(slide, color=ACCENT_CYAN):
    bar = add_rect(slide, Inches(0.5), Inches(1.35), Inches(12.33), Inches(0.05), color)
    return bar


# ══════════════════════════════════════════════════════════════════════
# SLAJD 1 — Tytuł i Wizja
# ══════════════════════════════════════════════════════════════════════

s1 = prs.slides.add_slide(blank_layout)
add_bg(s1)

# Lewy panel
add_rect(s1, 0, 0, Inches(5.8), SLIDE_H, RGBColor(0x07, 0x23, 0x40))

add_text(s1, "VulnApp Lite", Inches(0.55), Inches(1.8), Inches(5), Inches(1.1),
         font_size=Pt(48), bold=True, color=ACCENT_CYAN)
add_text(s1, "+ Jules Bot", Inches(0.55), Inches(2.9), Inches(5), Inches(0.8),
         font_size=Pt(36), color=TEXT_WHITE)
add_rect(s1, Inches(0.55), Inches(3.65), Inches(1.2), Inches(0.05), ACCENT_CYAN)
add_text(s1, "Zarządzanie podatnościami\ni automatyzacja przez Telegram",
         Inches(0.55), Inches(3.75), Inches(5.2), Inches(1.1),
         font_size=Pt(16), color=TEXT_GRAY)

# 4 feature cards
labels = [
    ("CVE\nTracking",   ACCENT_CYAN,  "🔍"),
    ("Mobile\niOS App", ACCENT_GREEN, "📱"),
    ("Telegram\nBot",   ACCENT_AMBER, "🤖"),
    ("Security\nFirst", ACCENT_RED,   "🔒"),
]
for i, (label, col, icon) in enumerate(labels):
    bx = Inches(6.3) + i * Inches(1.7)
    by = Inches(2.2)
    add_rounded_box(s1, bx, by, Inches(1.5), Inches(2.2), BG_CARD, "", line_color=col)
    add_text(s1, icon,  bx, by + Inches(0.25), Inches(1.5), Inches(0.6),
             font_size=Pt(30), align=PP_ALIGN.CENTER)
    add_text(s1, label, bx, by + Inches(0.9),  Inches(1.5), Inches(0.9),
             font_size=Pt(13), bold=True, color=col, align=PP_ALIGN.CENTER)

add_text(s1, "rafserver · Flask · SQLite · ZeroTier · Nginx · systemd",
         Inches(0.5), Inches(6.9), Inches(12.3), Inches(0.45),
         font_size=Pt(11), color=TEXT_GRAY, align=PP_ALIGN.CENTER)
add_text(s1, "01 / 05", Inches(11.8), Inches(6.9), Inches(1.3), Inches(0.4),
         font_size=Pt(11), color=ACCENT_CYAN, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════
# SLAJD 2 — Architektura Systemu
# ══════════════════════════════════════════════════════════════════════

s2 = prs.slides.add_slide(blank_layout)
add_bg(s2)
accent_bar(s2, ACCENT_CYAN)

add_text(s2, "Architektura Systemu", Inches(0.5), Inches(0.3), Inches(9), Inches(0.9),
         font_size=Pt(32), bold=True, color=TEXT_WHITE)
add_text(s2, "Jak działają wszystkie komponenty razem",
         Inches(0.5), Inches(1.1), Inches(9), Inches(0.5),
         font_size=Pt(14), color=TEXT_GRAY)

# Warstwa 1: Klienci
add_text(s2, "KLIENCI", Inches(0.5), Inches(1.55), Inches(12.3), Inches(0.3),
         font_size=Pt(10), bold=True, color=ACCENT_CYAN, align=PP_ALIGN.CENTER)

clients = [
    ("🌐  Przeglądarka Web\n(LAN / port 80)",    Inches(0.5)),
    ("📱  VulnApp Mobile\n(iOS / ZeroTier)",      Inches(3.5)),
    ("🤖  Jules Bot\n(Telegram API)",             Inches(6.5)),
    ("👤  Admin\n(SSH / terminal)",               Inches(9.5)),
]
for label, bx in clients:
    add_rounded_box(s2, bx, Inches(1.9), Inches(2.7), Inches(0.85),
                    BG_CARD, label, txt_size=Pt(12), line_color=ACCENT_CYAN)

# Strzałki klienci → proxy
for bx in [Inches(1.85), Inches(4.85), Inches(7.85), Inches(10.85)]:
    add_arrow_v(s2, bx, Inches(2.75), Inches(3.25), ACCENT_CYAN)

# Warstwa 2: Proxy
add_text(s2, "SIEĆ / PROXY", Inches(0.5), Inches(3.1), Inches(12.3), Inches(0.3),
         font_size=Pt(10), bold=True, color=ACCENT_GREEN, align=PP_ALIGN.CENTER)

add_rounded_box(s2, Inches(0.5), Inches(3.45), Inches(5.8), Inches(0.75),
                RGBColor(0x00, 0x33, 0x22),
                "🔀  Nginx Reverse Proxy  (port 80)",
                txt_size=Pt(13), bold=True, txt_color=ACCENT_GREEN, line_color=ACCENT_GREEN)
add_rounded_box(s2, Inches(7.0), Inches(3.45), Inches(5.8), Inches(0.75),
                RGBColor(0x33, 0x22, 0x00),
                "🌐  ZeroTier VPN  (zdalny dostęp)",
                txt_size=Pt(13), bold=True, txt_color=ACCENT_AMBER, line_color=ACCENT_AMBER)
# łącznik poziomy między proxy
add_arrow_h(s2, Inches(6.3), Inches(7.0), Inches(3.82), ACCENT_AMBER)

# Strzałki proxy → backend
add_arrow_v(s2, Inches(3.4), Inches(4.2), Inches(4.75), ACCENT_GREEN)
add_arrow_v(s2, Inches(9.9), Inches(4.2), Inches(4.75), ACCENT_AMBER)

# Warstwa 3: Backend
add_text(s2, "BACKEND  (rafserver — Ubuntu 24.04, AMD A6, 8GB RAM)",
         Inches(0.5), Inches(4.6), Inches(12.3), Inches(0.3),
         font_size=Pt(10), bold=True, color=ACCENT_RED, align=PP_ALIGN.CENTER)

backend = [
    ("🐍  Flask\n+ Gunicorn",  Inches(0.5),  ACCENT_RED),
    ("📊  SQLite\ntracker.db", Inches(3.0),  ACCENT_AMBER),
    ("📁  CSV Input\nWeekly/", Inches(5.5),  TEXT_GRAY),
    ("⚙️  systemd\nServices", Inches(8.0),  ACCENT_GREEN),
    ("📋  Jules\nScripts",     Inches(10.5), ACCENT_CYAN),
]
for label, bx, col in backend:
    add_rounded_box(s2, bx, Inches(4.95), Inches(2.4), Inches(0.9),
                    BG_CARD, label, txt_size=Pt(12), bold=True,
                    txt_color=col, line_color=col)

add_text(s2, "Dane: Defender CSV → /Input/Weekly/  |  API: /api/cves  |  Logi: /var/log/vulnapp/",
         Inches(0.5), Inches(6.85), Inches(12.3), Inches(0.4),
         font_size=Pt(10), color=TEXT_GRAY, align=PP_ALIGN.CENTER)
add_text(s2, "02 / 05", Inches(11.8), Inches(6.85), Inches(1.3), Inches(0.4),
         font_size=Pt(11), color=ACCENT_CYAN, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════
# SLAJD 3 — Logika Bota Telegram (Flowchart)
# ══════════════════════════════════════════════════════════════════════

s3 = prs.slides.add_slide(blank_layout)
add_bg(s3)
accent_bar(s3, ACCENT_AMBER)

add_text(s3, "Logika Bota Jules", Inches(0.5), Inches(0.3), Inches(9), Inches(0.9),
         font_size=Pt(32), bold=True, color=TEXT_WHITE)
add_text(s3, "Diagram przepływu: od komendy do akcji",
         Inches(0.5), Inches(1.1), Inches(9), Inches(0.5),
         font_size=Pt(14), color=TEXT_GRAY)

bw = Inches(3.7)
bh = Inches(0.72)

cols = [Inches(0.4), Inches(4.55), Inches(8.7)]
col_colors = [ACCENT_AMBER, ACCENT_CYAN, ACCENT_RED]
col_headers = ["⚡  LIVE LISTENER", "🔄  PR REVIEW  (cron 5 min)", "🔧  GITHUB ACTIONS"]

flows = [
    [
        "jules-listener.service\n(systemd long-poll)",
        "Wiadomość od\nużytkownika",
        "jules_actions.sh\nprzetwarza tekst",
        "Wywołanie\nClaude API",
        "Odpowiedź →\nTelegram",
    ],
    [
        "jules_review.sh\n(cron co 5 min)",
        "gh pr list\n— pobierz nowe PR-y",
        "Claude analizuje\ndiff kodu",
        "gh pr review\n— komentarz GitHub",
        "Powiadomienie\nTelegram",
    ],
    [
        "Komenda /deploy\nlub /status",
        "Parsowanie\nkomendy",
        "gh run / gh issue\n/ gh pr",
        "Sukces / Błąd\nzapis logu",
        "Raport\ndo Telegrama",
    ],
]

y_header = Inches(1.55)
y_start  = Inches(1.95)
y_step   = Inches(0.95)

for ci, (col_x, col_col, col_hdr) in enumerate(zip(cols, col_colors, col_headers)):
    add_text(s3, col_hdr, col_x, y_header, bw, Inches(0.35),
             font_size=Pt(12), bold=True, color=col_col, align=PP_ALIGN.CENTER)

    for fi, node_text in enumerate(flows[ci]):
        by = y_start + fi * y_step
        is_last  = fi == len(flows[ci]) - 1
        is_first = fi == 0
        node_col = ACCENT_GREEN if is_last else col_col
        add_rounded_box(s3, col_x, by, bw, bh,
                        BG_CARD, node_text,
                        txt_size=Pt(12), bold=(is_first or is_last),
                        txt_color=node_col, line_color=node_col)
        if not is_last:
            add_arrow_v(s3,
                        col_x + bw/2,
                        by + bh,
                        by + bh + (y_step - bh),
                        col_col)

add_text(s3, "03 / 05", Inches(11.8), Inches(6.9), Inches(1.3), Inches(0.4),
         font_size=Pt(11), color=ACCENT_CYAN, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════
# SLAJD 4 — User Journey Map
# ══════════════════════════════════════════════════════════════════════

s4 = prs.slides.add_slide(blank_layout)
add_bg(s4)
accent_bar(s4, ACCENT_GREEN)

add_text(s4, "Mapa Interakcji Użytkownika", Inches(0.5), Inches(0.3), Inches(10), Inches(0.9),
         font_size=Pt(32), bold=True, color=TEXT_WHITE)
add_text(s4, "User Journey: od importu CSV do zamknięcia ticketu",
         Inches(0.5), Inches(1.1), Inches(10), Inches(0.5),
         font_size=Pt(14), color=TEXT_GRAY)

# Oś czasu
add_rect(s4, Inches(0.5), Inches(3.72), Inches(12.33), Inches(0.08), ACCENT_GREEN)

steps = [
    ("1", "Import\nCSV",             "Defender\nexportuje raport",         ACCENT_AMBER, Inches(0.5)),
    ("2", "Web\nDashboard",          "Przegląd CVE,\nfiltrowanie",          ACCENT_CYAN,  Inches(2.65)),
    ("3", "Aktualizacja\nStatusu",   "action_taken\nticket_number",         ACCENT_GREEN, Inches(4.8)),
    ("4", "Mobile\nPodgląd",         "iOS: /api/cves\nprzez ZeroTier",      ACCENT_CYAN,  Inches(6.95)),
    ("5", "Bot Jules",               "Telegram: pytania,\nPR review",       ACCENT_AMBER, Inches(9.1)),
    ("6", "Raport\nZamknięcia",      "Export / statystyki\ndo zarządu",     TEXT_GRAY,    Inches(11.25)),
]

card_w = Inches(1.85)
card_h = Inches(0.75)

for num, title, desc, col, bx in steps:
    n = int(num)
    # Kółko na osi
    circle = s4.shapes.add_shape(9, bx + Inches(0.22), Inches(3.49), Inches(0.45), Inches(0.45))
    circle.fill.solid()
    circle.fill.fore_color.rgb = col
    circle.line.fill.background()
    add_text(s4, num, bx + Inches(0.22), Inches(3.49), Inches(0.45), Inches(0.45),
             font_size=Pt(13), bold=True, color=TEXT_DARK, align=PP_ALIGN.CENTER)

    if n % 2 == 1:
        # Karta nad osią
        by_card = Inches(2.3)
        add_rounded_box(s4, bx, by_card, card_w, card_h,
                        BG_CARD, title, txt_size=Pt(13), bold=True,
                        txt_color=col, line_color=col)
        add_text(s4, desc, bx, by_card + card_h, card_w, Inches(0.65),
                 font_size=Pt(11), color=TEXT_GRAY, align=PP_ALIGN.CENTER)
        add_arrow_v(s4, bx + card_w/2, by_card + card_h + Inches(0.65), Inches(3.71), col)
    else:
        # Karta pod osią
        by_card = Inches(4.45)
        add_rounded_box(s4, bx, by_card, card_w, card_h,
                        BG_CARD, title, txt_size=Pt(13), bold=True,
                        txt_color=col, line_color=col)
        add_text(s4, desc, bx, by_card + card_h, card_w, Inches(0.65),
                 font_size=Pt(11), color=TEXT_GRAY, align=PP_ALIGN.CENTER)
        add_arrow_v(s4, bx + card_w/2, Inches(3.8), by_card, col)

# Persona labels
add_text(s4, "👔 Admin / Security Team ──────────────────────────────────────────────────────",
         Inches(0.5), Inches(5.55), Inches(12.3), Inches(0.4),
         font_size=Pt(10), color=ACCENT_AMBER)
add_text(s4, "💻 Developer / DevOps ──────────────────────────────────────────────────────────",
         Inches(0.5), Inches(5.95), Inches(12.3), Inches(0.4),
         font_size=Pt(10), color=ACCENT_CYAN)

add_text(s4, "04 / 05", Inches(11.8), Inches(6.9), Inches(1.3), Inches(0.4),
         font_size=Pt(11), color=ACCENT_CYAN, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════
# SLAJD 5 — Dashboard Wydajności
# ══════════════════════════════════════════════════════════════════════

s5 = prs.slides.add_slide(blank_layout)
add_bg(s5)
accent_bar(s5, ACCENT_RED)

add_text(s5, "Statystyki i Wydajność", Inches(0.5), Inches(0.3), Inches(10), Inches(0.9),
         font_size=Pt(32), bold=True, color=TEXT_WHITE)
add_text(s5, "Dashboard — propozycja metryk operacyjnych",
         Inches(0.5), Inches(1.1), Inches(9), Inches(0.5),
         font_size=Pt(14), color=TEXT_GRAY)

# KPI cards
kpis = [
    ("247",  "CVE\nŚledzonych",        ACCENT_CYAN),
    ("89%",  "Zamkniętych\nw terminie", ACCENT_GREEN),
    ("3.2h", "Avg. czas\nreakcji",      ACCENT_AMBER),
    ("12",   "PR Review\n/ tydzień",    ACCENT_RED),
    ("↑5%",  "Poprawa\nMoM",            TEXT_GRAY),
]
for i, (val, label, col) in enumerate(kpis):
    bx = Inches(0.4) + i * Inches(2.55)
    add_rounded_box(s5, bx, Inches(1.65), Inches(2.3), Inches(1.4),
                    BG_CARD, "", line_color=col)
    add_text(s5, val,   bx, Inches(1.75), Inches(2.3), Inches(0.7),
             font_size=Pt(34), bold=True, color=col, align=PP_ALIGN.CENTER)
    add_text(s5, label, bx, Inches(2.55), Inches(2.3), Inches(0.5),
             font_size=Pt(12), color=TEXT_GRAY, align=PP_ALIGN.CENTER)

# Bar chart: CVE severity
add_rounded_box(s5, Inches(0.4), Inches(3.25), Inches(5.8), Inches(2.9),
                BG_CARD, "", line_color=ACCENT_CYAN)
add_text(s5, "CVE wg. Severity",
         Inches(0.6), Inches(3.35), Inches(5.4), Inches(0.4),
         font_size=Pt(13), bold=True, color=ACCENT_CYAN)

bars = [
    ("Critical", 0.85, ACCENT_RED),
    ("High",     0.60, ACCENT_AMBER),
    ("Medium",   0.40, ACCENT_CYAN),
    ("Low",      0.20, ACCENT_GREEN),
]
max_h = Inches(1.5)
for i, (lbl, pct, col) in enumerate(bars):
    bx  = Inches(0.75) + i * Inches(1.3)
    bh2 = Emu(int(max_h * pct))
    by  = Inches(3.85) + (max_h - bh2)
    add_rect(s5, bx, by, Inches(0.9), bh2, col)
    add_text(s5, lbl, bx - Inches(0.1), Inches(5.45), Inches(1.1), Inches(0.35),
             font_size=Pt(10), color=TEXT_GRAY, align=PP_ALIGN.CENTER)
    add_text(s5, f"{int(pct*100)}%", bx, by - Inches(0.3), Inches(0.9), Inches(0.3),
             font_size=Pt(11), bold=True, color=col, align=PP_ALIGN.CENTER)

# Status tickets (lista zamiast donut)
add_rounded_box(s5, Inches(6.5), Inches(3.25), Inches(6.3), Inches(2.9),
                BG_CARD, "", line_color=ACCENT_GREEN)
add_text(s5, "Status Ticketów",
         Inches(6.7), Inches(3.35), Inches(5.8), Inches(0.4),
         font_size=Pt(13), bold=True, color=ACCENT_GREEN)

status_items = [
    ("✅  Zamknięte",   "62%", ACCENT_GREEN,  Inches(3.85)),
    ("🔄  W trakcie",  "24%", ACCENT_CYAN,   Inches(4.5)),
    ("⚠️  Zaległe",   "10%", ACCENT_AMBER,  Inches(5.15)),
    ("🔴  Krytyczne",   " 4%", ACCENT_RED,   Inches(5.8)),
]
for label, pct, col, by in status_items:
    bar_w = Inches(4.5 * float(pct.strip().rstrip('%')) / 100)
    add_rect(s5, Inches(6.7), by, bar_w, Inches(0.38), col)
    add_text(s5, f"{label}  {pct}", Inches(6.7), by, Inches(5.5), Inches(0.38),
             font_size=Pt(12), bold=False, color=TEXT_WHITE)

# Bot activity strip
add_rounded_box(s5, Inches(0.4), Inches(6.25), Inches(12.4), Inches(0.5),
                BG_CARD,
                "🤖  Jules:  47 komend/tydzień  |  12 PR review  |  ∅ 8s latency  |  uptime 99.8%",
                txt_size=Pt(12), bold=False, txt_color=ACCENT_AMBER, line_color=ACCENT_AMBER)

add_text(s5, "* Dane poglądowe — docelowo integracja z logami systemd i GitHub API",
         Inches(0.5), Inches(6.88), Inches(11), Inches(0.4),
         font_size=Pt(10), italic=True, color=TEXT_GRAY)
add_text(s5, "05 / 05", Inches(11.8), Inches(6.88), Inches(1.3), Inches(0.4),
         font_size=Pt(11), color=ACCENT_CYAN, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════
# ZAPIS
# ══════════════════════════════════════════════════════════════════════

out = "VulnApp_Lite_Jules_Bot_Presentation.pptx"
prs.save(out)
print(f"Saved: {out}")
