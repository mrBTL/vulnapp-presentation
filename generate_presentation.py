"""
VulnApp Lite + Jules Bot — Prezentacja PowerPoint
Generuje wizualną prezentację projektu.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pptx.util as util

# ── Paleta kolorów ──────────────────────────────────────────────────
BG_DARK      = RGBColor(0x0D, 0x1B, 0x2A)   # granatowy tło
BG_CARD      = RGBColor(0x11, 0x2D, 0x4A)   # ciemnoniebieski karta
ACCENT_CYAN  = RGBColor(0x00, 0xE5, 0xFF)   # cyjan (akcent)
ACCENT_GREEN = RGBColor(0x00, 0xE6, 0x76)   # zielony (sukces)
ACCENT_RED   = RGBColor(0xFF, 0x45, 0x45)   # czerwony (alert)
ACCENT_AMBER = RGBColor(0xFF, 0xBF, 0x00)   # bursztynowy (uwaga)
TEXT_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_GRAY    = RGBColor(0xA0, 0xB4, 0xC8)
TEXT_DARK    = RGBColor(0x0D, 0x1B, 0x2A)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

blank_layout = prs.slide_layouts[6]   # puste tło


# ══════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════

def add_bg(slide, color=BG_DARK):
    bg = slide.shapes.add_shape(
        1,   # MSO_SHAPE_TYPE.RECTANGLE
        0, 0, SLIDE_W, SLIDE_H
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    return bg


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
    shape = slide.shapes.add_shape(
        5,   # ROUNDED_RECTANGLE
        x, y, w, h
    )
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


def add_arrow(slide, x1, y1, x2, y2, color=ACCENT_CYAN, w=Pt(2)):
    """Prosta strzałka (connector line) — pionowa lub pozioma."""
    from pptx.util import Emu
    connector = slide.shapes.add_connector(
        1,   # STRAIGHT
        x1, y1, x2, y2
    )
    connector.line.color.rgb = color
    connector.line.width = w
    return connector


def accent_bar(slide, color=ACCENT_CYAN, height=Inches(0.06)):
    """Poziomy pasek akcentu pod tytułem."""
    bar = slide.shapes.add_shape(1, Inches(0.5), Inches(1.35), Inches(12.33), height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()
    return bar


# ══════════════════════════════════════════════════════════════════════
# SLAJD 1 — Tytuł i Wizja
# ══════════════════════════════════════════════════════════════════════

s1 = prs.slides.add_slide(blank_layout)
add_bg(s1)

# Lewy gradient-panel (symulacja)
left = add_rect(s1, 0, 0, Inches(5.8), SLIDE_H, RGBColor(0x07, 0x23, 0x40))

# Tytuł
add_text(s1, "VulnApp Lite", Inches(0.55), Inches(1.8), Inches(5), Inches(1.1),
         font_size=Pt(48), bold=True, color=ACCENT_CYAN, align=PP_ALIGN.LEFT)
add_text(s1, "+ Jules Bot", Inches(0.55), Inches(2.9), Inches(5), Inches(0.8),
         font_size=Pt(36), bold=False, color=TEXT_WHITE, align=PP_ALIGN.LEFT)
add_text(s1, "Zarządzanie podatnościami\ni automatyzacja przez Telegram",
         Inches(0.55), Inches(3.75), Inches(5.2), Inches(1.1),
         font_size=Pt(16), color=TEXT_GRAY, align=PP_ALIGN.LEFT)

# Pasek
bar = add_rect(s1, Inches(0.55), Inches(3.65), Inches(1.2), Inches(0.05), ACCENT_CYAN)

# Prawy panel — 4 ikony/pillary
labels = [
    ("🔍", "CVE\nTracking",   ACCENT_CYAN),
    ("📱", "Mobile\niOS App", ACCENT_GREEN),
    ("🤖", "Telegram\nBot",   ACCENT_AMBER),
    ("🔒", "Security\nFirst", ACCENT_RED),
]
for i, (icon, label, col) in enumerate(labels):
    bx = Inches(6.3) + i * Inches(1.7)
    by = Inches(2.2)
    card = add_rounded_box(s1, bx, by, Inches(1.5), Inches(2.2),
                           BG_CARD, "", line_color=col)
    add_text(s1, icon,  bx, by + Inches(0.25), Inches(1.5), Inches(0.6),
             font_size=Pt(30), align=PP_ALIGN.CENTER)
    add_text(s1, label, bx, by + Inches(0.9),  Inches(1.5), Inches(0.9),
             font_size=Pt(13), bold=True, color=col, align=PP_ALIGN.CENTER)

# Stopka
add_text(s1, "rafserver · Flask · SQLite · ZeroTier · Nginx · systemd",
         Inches(0.5), Inches(6.9), Inches(12.3), Inches(0.45),
         font_size=Pt(11), color=TEXT_GRAY, align=PP_ALIGN.CENTER)

# Numer slajdu
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

# ── Warstwa 1: Klienci ─────────────────────────────────────
add_text(s2, "KLIENCI", Inches(0.5), Inches(1.55), Inches(12.3), Inches(0.35),
         font_size=Pt(11), bold=True, color=ACCENT_CYAN, align=PP_ALIGN.CENTER)

clients = [
    ("🌐 Przeglądarka\nWeb (LAN)", Inches(1.0)),
    ("📱 VulnApp\niOS Mobile",    Inches(4.0)),
    ("🤖 Telegram\nJules Bot",    Inches(7.0)),
    ("👤 Użytkownik\nAdmin",      Inches(10.2)),
]
for label, bx in clients:
    add_rounded_box(s2, bx, Inches(1.95), Inches(2.1), Inches(0.85),
                    BG_CARD, label, txt_size=Pt(12), bold=False,
                    line_color=ACCENT_CYAN)

# ── Strzałki klienci → Nginx ────────────────────────────────
for bx in [Inches(2.05), Inches(5.05), Inches(8.05)]:
    add_arrow(s2, bx, Inches(2.8), bx, Inches(3.35), ACCENT_CYAN, Pt(1.5))
add_arrow(s2, Inches(11.25), Inches(2.8), Inches(11.25), Inches(3.35), ACCENT_AMBER, Pt(1.5))

# ── Warstwa 2: Nginx + ZeroTier ────────────────────────────
add_text(s2, "SIEĆ / PROXY", Inches(0.5), Inches(3.2), Inches(12.3), Inches(0.3),
         font_size=Pt(11), bold=True, color=ACCENT_GREEN, align=PP_ALIGN.CENTER)

add_rounded_box(s2, Inches(1.0), Inches(3.5), Inches(5.5), Inches(0.75),
                RGBColor(0x00, 0x4D, 0x33), "🔀  Nginx Reverse Proxy  (port 80)",
                txt_size=Pt(13), bold=True, txt_color=ACCENT_GREEN, line_color=ACCENT_GREEN)
add_rounded_box(s2, Inches(7.2), Inches(3.5), Inches(5.0), Inches(0.75),
                RGBColor(0x33, 0x22, 0x00), "🌐  ZeroTier VPN  (remote access)",
                txt_size=Pt(13), bold=True, txt_color=ACCENT_AMBER, line_color=ACCENT_AMBER)

# strzałki nginx/zerotier → backend
add_arrow(s2, Inches(3.75), Inches(4.25), Inches(3.75), Inches(4.8), ACCENT_GREEN, Pt(1.5))
add_arrow(s2, Inches(9.7),  Inches(4.25), Inches(9.7),  Inches(4.8), ACCENT_AMBER, Pt(1.5))
# crosslink zerotier → nginx
add_arrow(s2, Inches(7.2), Inches(3.875), Inches(6.5), Inches(3.875), ACCENT_AMBER, Pt(1.2))

# ── Warstwa 3: Backend ─────────────────────────────────────
add_text(s2, "BACKEND  (rafserver)", Inches(0.5), Inches(4.65), Inches(12.3), Inches(0.3),
         font_size=Pt(11), bold=True, color=ACCENT_RED, align=PP_ALIGN.CENTER)

backend_items = [
    ("🐍 Flask\n+ Gunicorn", Inches(0.6),  ACCENT_RED),
    ("📊 SQLite\nDB",        Inches(3.2),  ACCENT_AMBER),
    ("📁 CSV\nInput",        Inches(5.8),  TEXT_GRAY),
    ("📜 systemd\nServices", Inches(8.4),  ACCENT_GREEN),
    ("📋 Jules\nScripts",    Inches(11.0), ACCENT_CYAN),
]
for label, bx, col in backend_items:
    add_rounded_box(s2, bx, Inches(4.95), Inches(2.3), Inches(0.9),
                    BG_CARD, label, txt_size=Pt(12), bold=True,
                    txt_color=col, line_color=col)

# ── Legenda ─────────────────────────────────────────────────
add_text(s2, "● LAN  ● ZeroTier  ● Telegram API  ● GitHub API",
         Inches(0.5), Inches(6.85), Inches(12.3), Inches(0.4),
         font_size=Pt(11), color=TEXT_GRAY, align=PP_ALIGN.CENTER)
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

# ── Kolumna A: long-polling listener ──────────────────────
col_a = Inches(0.45)
col_b = Inches(4.55)
col_c = Inches(8.65)
bw    = Inches(3.7)
bh    = Inches(0.75)

nodes_a = [
    ("jules-listener.service\n(systemd long-poll)", Inches(1.5),  BG_CARD,   ACCENT_AMBER, "START"),
    ("Wiadomość\nod użytkownika",                    Inches(2.5),  BG_CARD,   ACCENT_AMBER, None),
    ("jules_actions.sh\nprzetwarza tekst",            Inches(3.5),  BG_CARD,   ACCENT_AMBER, None),
    ("Wywołanie\nClaude API",                         Inches(4.5),  BG_CARD,   ACCENT_CYAN,  None),
    ("Odpowiedź →\nTelegram",                         Inches(5.5),  BG_CARD,   ACCENT_GREEN, "END"),
]

for label, by, fill, col, tag in nodes_a:
    shape_fill = fill
    add_rounded_box(s3, col_a, by, bw, bh, shape_fill, label,
                    txt_size=Pt(12), bold=(tag is not None),
                    txt_color=col, line_color=col)
    if tag == "START":
        add_rounded_box(s3, col_a + Inches(1.3), by - Inches(0.5), Inches(1.1), Inches(0.35),
                        ACCENT_AMBER, "START", txt_size=Pt(10), bold=True, txt_color=TEXT_DARK)
    if tag == "END":
        add_rounded_box(s3, col_a + Inches(1.3), by + bh, Inches(1.1), Inches(0.35),
                        ACCENT_GREEN, "END", txt_size=Pt(10), bold=True, txt_color=TEXT_DARK)

# Strzałki kolumna A
for by_top in [Inches(2.25), Inches(3.25), Inches(4.25), Inches(5.25)]:
    add_arrow(s3, col_a + bw/2, by_top, col_a + bw/2, by_top + Inches(0.25), ACCENT_AMBER)

# ── Kolumna B: PR Review cron ─────────────────────────────
nodes_b = [
    ("jules_review.sh\n(cron co 5 min)",            Inches(1.5),  BG_CARD,   ACCENT_CYAN,  "CRON"),
    ("gh pr list\n— pobierz nowe PR-y",              Inches(2.5),  BG_CARD,   ACCENT_CYAN,  None),
    ("Claude analizuje\ndiff kodu",                  Inches(3.5),  BG_CARD,   ACCENT_CYAN,  None),
    ("gh pr review\n— komentarz na GitHub",          Inches(4.5),  BG_CARD,   ACCENT_CYAN,  None),
    ("Powiadomienie\nTelegram",                       Inches(5.5),  BG_CARD,   ACCENT_GREEN, "END"),
]
for label, by, fill, col, tag in nodes_b:
    add_rounded_box(s3, col_b, by, bw, bh, fill, label,
                    txt_size=Pt(12), bold=(tag is not None),
                    txt_color=col, line_color=col)
    if tag == "CRON":
        add_rounded_box(s3, col_b + Inches(1.3), by - Inches(0.5), Inches(1.1), Inches(0.35),
                        ACCENT_CYAN, "CRON", txt_size=Pt(10), bold=True, txt_color=TEXT_DARK)

for by_top in [Inches(2.25), Inches(3.25), Inches(4.25), Inches(5.25)]:
    add_arrow(s3, col_b + bw/2, by_top, col_b + bw/2, by_top + Inches(0.25), ACCENT_CYAN)

# ── Kolumna C: Akcje GitHub ───────────────────────────────
nodes_c = [
    ("Komenda /deploy\nlub /status",              Inches(1.5),  BG_CARD,   ACCENT_RED,  "TRIGGER"),
    ("Parsowanie\nkomendy",                       Inches(2.5),  BG_CARD,   ACCENT_RED,  None),
    ("gh run\n/ gh issue / gh pr",               Inches(3.5),  BG_CARD,   ACCENT_RED,  None),
    ("Sukces / Błąd\nzapis logu",                 Inches(4.5),  BG_CARD,   ACCENT_RED,  None),
    ("Raport\ndo Telegrama",                      Inches(5.5),  BG_CARD,   ACCENT_GREEN, "END"),
]
for label, by, fill, col, tag in nodes_c:
    add_rounded_box(s3, col_c, by, bw, bh, fill, label,
                    txt_size=Pt(12), bold=(tag is not None),
                    txt_color=col, line_color=col)
    if tag == "TRIGGER":
        add_rounded_box(s3, col_c + Inches(1.3), by - Inches(0.5), Inches(1.1), Inches(0.35),
                        ACCENT_RED, "TRIGGER", txt_size=Pt(10), bold=True, txt_color=TEXT_WHITE)

for by_top in [Inches(2.25), Inches(3.25), Inches(4.25), Inches(5.25)]:
    add_arrow(s3, col_c + bw/2, by_top, col_c + bw/2, by_top + Inches(0.25), ACCENT_RED)

# Nagłówki kolumn
add_text(s3, "⚡ LIVE LISTENER", col_a, Inches(1.2), bw, Inches(0.4),
         font_size=Pt(12), bold=True, color=ACCENT_AMBER, align=PP_ALIGN.CENTER)
add_text(s3, "🔄 PR REVIEW (CRON)", col_b, Inches(1.2), bw, Inches(0.4),
         font_size=Pt(12), bold=True, color=ACCENT_CYAN, align=PP_ALIGN.CENTER)
add_text(s3, "🔧 GITHUB ACTIONS", col_c, Inches(1.2), bw, Inches(0.4),
         font_size=Pt(12), bold=True, color=ACCENT_RED, align=PP_ALIGN.CENTER)

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
add_text(s4, "Punkty styku: Admin · Developer · End User",
         Inches(0.5), Inches(1.1), Inches(9), Inches(0.5),
         font_size=Pt(14), color=TEXT_GRAY)

# Oś czasu (timeline)
add_rect(s4, Inches(0.5), Inches(3.7), Inches(12.33), Inches(0.08), ACCENT_GREEN)

# Kroki journey
steps = [
    ("1", "Import\nCSV",          "Defender\nexportuje raport",   ACCENT_AMBER,  Inches(0.5)),
    ("2", "Web\nDashboard",       "Przegląd CVE,\nfiltrowanie",    ACCENT_CYAN,   Inches(2.8)),
    ("3", "Aktualizacja\nStatusu","action_taken\nticket_number",   ACCENT_GREEN,  Inches(5.1)),
    ("4", "Mobile\nPodgląd",      "iOS: czyta /api/cves\nprzez ZeroTier", ACCENT_CYAN, Inches(7.4)),
    ("5", "Bot\nJules",           "Telegram: pytania,\nraporty, PR review", ACCENT_AMBER, Inches(9.7)),
    ("6", "Raport\nZamknięcia",   "Export / statystyki\ndo zarządu",   TEXT_GRAY, Inches(11.8)),
]

for num, title, desc, col, bx in steps:
    # Kółko na osi
    circle = slide_circle = s4.shapes.add_shape(
        9,   # OVAL
        bx + Inches(0.3), Inches(3.45), Inches(0.5), Inches(0.5)
    )
    circle.fill.solid()
    circle.fill.fore_color.rgb = col
    circle.line.fill.background()
    # Numer
    add_text(s4, num, bx + Inches(0.3), Inches(3.45), Inches(0.5), Inches(0.5),
             font_size=Pt(14), bold=True, color=TEXT_DARK, align=PP_ALIGN.CENTER)

    # Karta góra (tytuł)
    by_card = Inches(1.6) if int(num) % 2 == 1 else Inches(4.4)
    add_rounded_box(s4, bx, by_card, Inches(1.9), Inches(0.7),
                    BG_CARD, title, txt_size=Pt(13), bold=True,
                    txt_color=col, line_color=col)
    # Opis
    by_desc = by_card + Inches(0.75) if int(num) % 2 == 1 else by_card + Inches(0.75)
    add_text(s4, desc, bx, by_desc, Inches(1.9), Inches(0.8),
             font_size=Pt(11), color=TEXT_GRAY, align=PP_ALIGN.CENTER)

    # Linia do osi
    mid_x = bx + Inches(0.55)
    if int(num) % 2 == 1:
        add_arrow(s4, mid_x, by_card + Inches(0.7), mid_x, Inches(3.68), col, Pt(1.5))
    else:
        add_arrow(s4, mid_x, Inches(3.95), mid_x, by_card, col, Pt(1.5))

# Persona-pasy (lewa szpalta)
add_text(s4, "👔 Admin / Security Team",
         Inches(0.0), Inches(1.55), Inches(0.5), Inches(2.5),
         font_size=Pt(9), color=ACCENT_AMBER, align=PP_ALIGN.CENTER)
add_text(s4, "💻 Developer",
         Inches(0.0), Inches(4.2), Inches(0.5), Inches(2.5),
         font_size=Pt(9), color=ACCENT_CYAN, align=PP_ALIGN.CENTER)

add_text(s4, "04 / 05", Inches(11.8), Inches(6.9), Inches(1.3), Inches(0.4),
         font_size=Pt(11), color=ACCENT_CYAN, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════
# SLAJD 5 — Statystyki i Wydajność (Dashboard)
# ══════════════════════════════════════════════════════════════════════

s5 = prs.slides.add_slide(blank_layout)
add_bg(s5)
accent_bar(s5, ACCENT_RED)

add_text(s5, "Statystyki i Wydajność", Inches(0.5), Inches(0.3), Inches(10), Inches(0.9),
         font_size=Pt(32), bold=True, color=TEXT_WHITE)
add_text(s5, "Dashboard — propozycja metryk operacyjnych",
         Inches(0.5), Inches(1.1), Inches(9), Inches(0.5),
         font_size=Pt(14), color=TEXT_GRAY)

# ── Wiersz KPI ────────────────────────────────────────────
kpis = [
    ("247", "CVE\nŚledzonych",       ACCENT_CYAN),
    ("89%", "Zamkniętych\nw terminie", ACCENT_GREEN),
    ("3.2h", "Avg. czas\nreakcji",    ACCENT_AMBER),
    ("12",  "PR\nZrecenzowanych/tyg", ACCENT_RED),
    ("↑5%", "Poprawa MoM",           TEXT_GRAY),
]
for i, (val, label, col) in enumerate(kpis):
    bx = Inches(0.4) + i * Inches(2.55)
    card = add_rounded_box(s5, bx, Inches(1.65), Inches(2.3), Inches(1.4),
                           BG_CARD, "", line_color=col)
    add_text(s5, val,   bx, Inches(1.75), Inches(2.3), Inches(0.7),
             font_size=Pt(34), bold=True, color=col, align=PP_ALIGN.CENTER)
    add_text(s5, label, bx, Inches(2.55), Inches(2.3), Inches(0.5),
             font_size=Pt(12), color=TEXT_GRAY, align=PP_ALIGN.CENTER)

# ── Chart Area 1: Bar Chart (CVE severity) ───────────────
add_rounded_box(s5, Inches(0.4), Inches(3.25), Inches(5.8), Inches(2.9),
                BG_CARD, "", line_color=ACCENT_CYAN)
add_text(s5, "CVE wg. Severity (Mock Bar Chart)",
         Inches(0.6), Inches(3.35), Inches(5.4), Inches(0.4),
         font_size=Pt(13), bold=True, color=ACCENT_CYAN)

# Bars
bars_data = [
    ("Critical", 0.85, ACCENT_RED),
    ("High",     0.60, ACCENT_AMBER),
    ("Medium",   0.40, ACCENT_CYAN),
    ("Low",      0.20, ACCENT_GREEN),
]
for i, (lbl, pct, col) in enumerate(bars_data):
    bx = Inches(0.8) + i * Inches(1.3)
    max_h = Inches(1.5)
    bar_h = Emu(int(max_h * pct))
    bar_y = Inches(3.85) + (max_h - bar_h)
    add_rect(s5, bx, bar_y, Inches(0.9), bar_h, col)
    add_text(s5, lbl, bx - Inches(0.1), Inches(5.45), Inches(1.1), Inches(0.35),
             font_size=Pt(10), color=TEXT_GRAY, align=PP_ALIGN.CENTER)
    add_text(s5, f"{int(pct*100)}%", bx, bar_y - Inches(0.3), Inches(0.9), Inches(0.3),
             font_size=Pt(11), bold=True, color=col, align=PP_ALIGN.CENTER)

# ── Chart Area 2: Donut (status breakdown) ───────────────
add_rounded_box(s5, Inches(6.5), Inches(3.25), Inches(6.3), Inches(2.9),
                BG_CARD, "", line_color=ACCENT_GREEN)
add_text(s5, "Status Ticketów (Mock Donut)",
         Inches(6.7), Inches(3.35), Inches(5.8), Inches(0.4),
         font_size=Pt(13), bold=True, color=ACCENT_GREEN)

donut_items = [
    ("✅ Zamknięte",  "62%", ACCENT_GREEN),
    ("🔄 W trakcie", "24%", ACCENT_CYAN),
    ("⚠️ Zaległe",  "10%", ACCENT_AMBER),
    ("🔴 Krytyczne",  "4%", ACCENT_RED),
]
for i, (label, pct, col) in enumerate(donut_items):
    row = i // 2
    col_pos = i % 2
    bx = Inches(6.7) + col_pos * Inches(3.0)
    by = Inches(3.85) + row * Inches(0.9)
    add_rounded_box(s5, bx, by, Inches(2.7), Inches(0.75),
                    RGBColor(0x16, 0x38, 0x60), f"{label}  {pct}",
                    txt_size=Pt(12), bold=False, txt_color=col, line_color=col)

# ── Bot activity ─────────────────────────────────────────
add_rounded_box(s5, Inches(6.5), Inches(3.25) + Inches(1.9), Inches(6.3), Inches(0.95),
                BG_CARD, "🤖  Jules: 47 komend tygodniowo  |  12 PR review  |  ∅ 8s latency",
                txt_size=Pt(13), bold=False, txt_color=ACCENT_AMBER, line_color=ACCENT_AMBER)

add_text(s5, "05 / 05", Inches(11.8), Inches(6.9), Inches(1.3), Inches(0.4),
         font_size=Pt(11), color=ACCENT_CYAN, align=PP_ALIGN.RIGHT)
add_text(s5, "* Dane poglądowe — docelowo integracja z logami systemd i GitHub API",
         Inches(0.5), Inches(6.9), Inches(11), Inches(0.4),
         font_size=Pt(10), italic=True, color=TEXT_GRAY)


# ══════════════════════════════════════════════════════════════════════
# ZAPIS
# ══════════════════════════════════════════════════════════════════════

out = "VulnApp_Lite_Jules_Bot_Presentation.pptx"
prs.save(out)
print(f"Saved: {out}")
