from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# Color palette
COLOR_PRIMARY   = RGBColor(0x00, 0x47, 0x99)   # dark blue
COLOR_ACCENT    = RGBColor(0x00, 0xA8, 0xE8)   # light blue
COLOR_DARK      = RGBColor(0x1A, 0x1A, 0x2E)   # near black
COLOR_LIGHT     = RGBColor(0xF4, 0xF7, 0xFF)   # off white
COLOR_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_GRAY      = RGBColor(0x66, 0x66, 0x66)
COLOR_GREEN     = RGBColor(0x27, 0xAE, 0x60)
COLOR_ORANGE    = RGBColor(0xE6, 0x7E, 0x22)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

blank_layout = prs.slide_layouts[6]   # completely blank

# ─────────────────────────────────────────────────────────────────────────────
# Helper functions
# ─────────────────────────────────────────────────────────────────────────────

def add_rect(slide, left, top, width, height, fill_color=None, line_color=None, line_width=Pt(0)):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    fill = shape.fill
    if fill_color:
        fill.solid()
        fill.fore_color.rgb = fill_color
    else:
        fill.background()
    line = shape.line
    if line_color:
        line.color.rgb = line_color
        line.width = line_width
    else:
        line.fill.background()
    return shape


def add_textbox(slide, text, left, top, width, height,
                font_size=Pt(14), font_color=COLOR_DARK, bold=False,
                align=PP_ALIGN.LEFT, wrap=True, font_name="Calibri"):
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.color.rgb = font_color
    run.font.bold = bold
    run.font.name = font_name
    return txBox


def add_paragraph(tf, text, font_size=Pt(13), font_color=COLOR_DARK,
                  bold=False, space_before=Pt(4), level=0, font_name="Calibri"):
    p = tf.add_paragraph()
    p.level = level
    p.space_before = space_before
    run = p.add_run()
    run.text = text
    run.font.size = font_size
    run.font.color.rgb = font_color
    run.font.bold = bold
    run.font.name = font_name
    return p


def slide_header(slide, title, subtitle=None):
    """Dark top bar with title."""
    add_rect(slide, 0, 0, 13.33, 1.15, fill_color=COLOR_PRIMARY)
    add_textbox(slide, title, 0.35, 0.15, 12.5, 0.65,
                font_size=Pt(28), font_color=COLOR_WHITE, bold=True,
                align=PP_ALIGN.LEFT)
    if subtitle:
        add_textbox(slide, subtitle, 0.35, 0.78, 12.5, 0.35,
                    font_size=Pt(14), font_color=COLOR_ACCENT,
                    align=PP_ALIGN.LEFT)
    # accent line
    add_rect(slide, 0, 1.15, 13.33, 0.04, fill_color=COLOR_ACCENT)
    # background
    add_rect(slide, 0, 1.19, 13.33, 6.31, fill_color=COLOR_LIGHT)


def slide_footer(slide, page_num, total=10):
    add_rect(slide, 0, 7.2, 13.33, 0.3, fill_color=COLOR_PRIMARY)
    add_textbox(slide, "Marktrecherche: KI-basierte Verwaltungsassistenten | Deutschland 2026",
                0.2, 7.21, 11.5, 0.28,
                font_size=Pt(9), font_color=COLOR_WHITE, align=PP_ALIGN.LEFT)
    add_textbox(slide, f"{page_num} / {total}",
                12.5, 7.21, 0.7, 0.28,
                font_size=Pt(9), font_color=COLOR_WHITE, align=PP_ALIGN.RIGHT)


def card(slide, left, top, width, height, title, bullets,
         title_color=COLOR_PRIMARY, bg_color=COLOR_WHITE,
         bullet_size=Pt(12), title_size=Pt(14)):
    add_rect(slide, left, top, width, height, fill_color=bg_color,
             line_color=COLOR_ACCENT, line_width=Pt(1))
    # title bar inside card
    add_rect(slide, left, top, width, 0.38, fill_color=title_color)
    add_textbox(slide, title, left + 0.1, top + 0.05, width - 0.2, 0.3,
                font_size=title_size, font_color=COLOR_WHITE, bold=True)
    # bullet text
    txBox = slide.shapes.add_textbox(
        Inches(left + 0.12), Inches(top + 0.45),
        Inches(width - 0.24), Inches(height - 0.55)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for b in bullets:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(3)
        run = p.add_run()
        run.text = b
        run.font.size = bullet_size
        run.font.color.rgb = COLOR_DARK
        run.font.name = "Calibri"


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 1 – Titel
# ─────────────────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
# full background
add_rect(slide, 0, 0, 13.33, 7.5, fill_color=COLOR_PRIMARY)
# diagonal accent stripe
add_rect(slide, 0, 4.8, 13.33, 0.08, fill_color=COLOR_ACCENT)
add_rect(slide, 0, 4.88, 13.33, 2.62, fill_color=COLOR_DARK)

add_textbox(slide, "Marktrecherche", 0.8, 1.2, 11.5, 0.7,
            font_size=Pt(20), font_color=COLOR_ACCENT, bold=False, align=PP_ALIGN.CENTER)
add_textbox(slide,
            "KI-basierte Verwaltungsassistenten",
            0.5, 1.85, 12.3, 1.1,
            font_size=Pt(38), font_color=COLOR_WHITE, bold=True, align=PP_ALIGN.CENTER)
add_textbox(slide, "auf dem deutschen Markt", 0.5, 2.9, 12.3, 0.65,
            font_size=Pt(26), font_color=COLOR_ACCENT, bold=False, align=PP_ALIGN.CENTER)

add_textbox(slide, "Stand: April 2026", 0.5, 5.1, 12.3, 0.4,
            font_size=Pt(14), font_color=COLOR_WHITE, align=PP_ALIGN.CENTER)
add_textbox(slide,
            "Oeffentliche Verwaltung  |  Unternehmensbereich  |  Markttrends  |  Empfehlungen",
            0.5, 5.55, 12.3, 0.4,
            font_size=Pt(12), font_color=RGBColor(0xAA, 0xCC, 0xFF), align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 2 – Agenda
# ─────────────────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
slide_header(slide, "Agenda", "Uebersicht der Praesentation")
slide_footer(slide, 2)

agenda_items = [
    ("01", "Marktueberblick & Zahlen",          "Marktgroesse, Wachstum, Adoptionsrate in Deutschland"),
    ("02", "Kategorien KI-Assistenten",          "Typen, Einsatzbereiche und Unterscheidungsmerkmale"),
    ("03", "Loesungen Oeffentliche Verwaltung",  "Bundeslaender-Projekte, Bundesebene, souveraene Cloud"),
    ("04", "Internationale Unternehmensloesungen","Microsoft Copilot, SAP Joule, Notion AI, Slack AI"),
    ("05", "Deutsche Unternehmensloesungen",     "fonio, Vitas, Personio und weitere DSGVO-konforme Anbieter"),
    ("06", "Vergleichsmatrix",                   "Funktionen, Datenschutz, Zielgruppe im Ueberblick"),
    ("07", "Markttrends & Ausblick 2026+",       "Wachstumsprognosen, Konvergenz, Regulierung"),
    ("08", "Fazit & Empfehlungen",               "Entscheidungshilfen fuer den Einsatz"),
]

for i, (num, title, desc) in enumerate(agenda_items):
    col = i % 2
    row = i // 2
    left = 0.4 + col * 6.5
    top  = 1.35 + row * 1.42

    add_rect(slide, left, top, 6.1, 1.25, fill_color=COLOR_WHITE,
             line_color=COLOR_ACCENT, line_width=Pt(1))
    add_rect(slide, left, top, 0.55, 1.25, fill_color=COLOR_PRIMARY)
    add_textbox(slide, num, left + 0.04, top + 0.32, 0.48, 0.55,
                font_size=Pt(18), font_color=COLOR_WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, title, left + 0.65, top + 0.08, 5.3, 0.42,
                font_size=Pt(14), font_color=COLOR_PRIMARY, bold=True)
    add_textbox(slide, desc, left + 0.65, top + 0.52, 5.3, 0.65,
                font_size=Pt(11), font_color=COLOR_GRAY)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 3 – Marktueberblick & Zahlen
# ─────────────────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
slide_header(slide, "Marktueberblick & Zahlen", "KI-Markt Deutschland 2024-2026")
slide_footer(slide, 3)

kpis = [
    ("USD 0,42 Mrd.", "KI-Agenten-Markt\nDeutschland 2024"),
    ("USD 3,30 Mrd.", "Prognose\n2030"),
    ("+45,3 % p.a.", "CAGR\n2026-2033"),
    ("935+", "KI-Startups\nin Deutschland 2025"),
    ("~33 %", "Unternehmen nutzen\nKI produktiv"),
    ("~40 %", "Wissensarbeit\nautomatisierbar"),
]

for idx, (value, label) in enumerate(kpis):
    col = idx % 3
    row = idx // 3
    left = 0.35 + col * 4.3
    top  = 1.4  + row * 2.6
    add_rect(slide, left, top, 4.0, 2.3, fill_color=COLOR_WHITE,
             line_color=COLOR_PRIMARY, line_width=Pt(1.5))
    add_rect(slide, left, top, 4.0, 0.08, fill_color=COLOR_ACCENT)
    add_textbox(slide, value, left + 0.1, top + 0.2, 3.8, 0.9,
                font_size=Pt(28), font_color=COLOR_PRIMARY, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, label, left + 0.1, top + 1.1, 3.8, 1.0,
                font_size=Pt(12), font_color=COLOR_GRAY, align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 4 – Kategorien
# ─────────────────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
slide_header(slide, "Kategorien KI-basierter Verwaltungsassistenten",
             "Typen und Einsatzbereiche im Ueberblick")
slide_footer(slide, 4)

cats = [
    ("Textassistenten",
     ["Texterstellung & Umformulierung",
      "Dokumentenzusammenfassung",
      "Uebersetzungen",
      "Protokollerstellung"]),
    ("Wissens- &\nRechercheassistenten",
     ["Interne Wissensdatenbanken",
      "Dokumentenanalyse",
      "Frage-Antwort-Systeme",
      "Aktenrecherche"]),
    ("Prozess- &\nWorkflow-Agenten",
     ["Formularausfuellung",
      "Aufgabenautomatisierung",
      "Multi-Step-Workflows",
      "ERP/HR-Integration"]),
    ("Kommunikations-\nassistenten",
     ["KI-Telefonassistenten",
      "E-Mail-Verwaltung",
      "Chat-/Ticketsysteme",
      "Buergerservice-Bots"]),
]

for i, (title, bullets) in enumerate(cats):
    left = 0.35 + i * 3.2
    card(slide, left, 1.35, 3.0, 5.65, title, bullets,
         bullet_size=Pt(12), title_size=Pt(13))

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 5 – Oeffentliche Verwaltung
# ─────────────────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
slide_header(slide, "KI-Loesungen fuer die Oeffentliche Verwaltung",
             "Bundeslaender-Projekte und Bundesebene")
slide_footer(slide, 5)

gov_solutions = [
    ("LLMoin\n(Hamburg)",
     ["Entwickelt von Dataport & Hamburg",
      "Rd. 40.000 Beschaeftigte nutzen es",
      "Nachgenutzt von mehreren Bundeslaendern",
      "Texterstellung, Zusammenfassung, Recherche"]),
    ("BaerGPT\n(Berlin)",
     ["Entwickelt vom CityLAB Berlin",
      "Einfuehrung ab November 2025",
      "Open Source, basiert auf Mistral-Modell",
      "Fuer Berliner Verwaltungsmitarbeitende"]),
    ("BayernKI\n(Bayern)",
     ["Federfuehrung: Finanzministerium Bayern",
      "Roll-out ab Oktober 2024",
      "Gesamte Staatsverwaltung + Pilotkommunen",
      "Multifunktionale KI-Anwendung"]),
    ("NRW.Genius\n(NRW)",
     ["Zentrales KI-Projekt in NRW",
      "Fuer alle Verwaltungsebenen entwickelt",
      "Textverarbeitung, Analyse, Assistenz",
      "Landesweiter Rollout geplant"]),
    ("F13 / Aleph Alpha\n(Bund + BaWue)",
     ["Von Aleph Alpha entwickelt",
      "Auf STACKIT souveraener Cloud gehostet",
      "Daten ausschliesslich in Deutschland",
      "Vermerkomat, Textzusammenfassung, Recherche"]),
    ("KIPITZ\n(Bundesebene)",
     ["KI-Plattform des ITZ-Bund",
      "Basisplattform fuer LLMs der Bundesverwaltung",
      "Dokumentzusammenfassung und Uebersetzung",
      "Texterstellung und Umformulierung"]),
]

for i, (title, bullets) in enumerate(gov_solutions):
    col = i % 3
    row = i // 3
    left = 0.3 + col * 4.35
    top  = 1.35 + row * 2.9
    card(slide, left, top, 4.1, 2.7, title, bullets,
         bullet_size=Pt(11), title_size=Pt(12))

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 6 – Internationale Unternehmensloesungen
# ─────────────────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
slide_header(slide, "Internationale Unternehmensloesungen",
             "Microsoft Copilot, SAP Joule, Notion AI, Slack AI")
slide_footer(slide, 6)

intl = [
    ("Microsoft 365\nCopilot",
     COLOR_PRIMARY,
     ["Tief in Word, Teams, Outlook integriert",
      "Zusammenfassungen, E-Mail-Entwuerfe",
      "Integration mit SAP Joule (bidirektional)",
      "Verfuegbar fuer Enterprise-Kunden",
      "Datenhaltung: EU-Option vorhanden"]),
    ("SAP Joule",
     RGBColor(0x00, 0x70, 0xB8),
     ["KI-Copilot in SAP-Anwendungen (S/4HANA etc.)",
      "2.100+ vordefinierte KI-Skills",
      "Routineaufgaben bis 80 % schneller",
      "Nahtlose Verbindung zu Microsoft 365",
      "EU AI Act Konformitaetspfad"]),
    ("Notion AI",
     RGBColor(0x37, 0x35, 0x30),
     ["Meeting-Notizen & Dokumentenzusammenfassung",
      "Integration: Slack, Google Drive, Teams",
      "Ab 11,50 EUR/Monat (Notion Plus)",
      "Wissensmanagement fuer Teams",
      "DSGVO-konform konfigurierbar"]),
    ("Slack AI",
     RGBColor(0x4A, 0x15, 0x4B),
     ["Channel- & Thread-Zusammenfassungen",
      "Tagesrecaps und Highlights",
      "Schnelle Suche im Workspace",
      "Ideal fuer kommunikationsintensive Teams",
      "Datenspeicherung konfigurierbar"]),
]

for i, (title, color, bullets) in enumerate(intl):
    left = 0.3 + i * 3.2
    card(slide, left, 1.35, 3.05, 5.65, title, bullets,
         title_color=color, bullet_size=Pt(11), title_size=Pt(13))

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 7 – Deutsche Unternehmensloesungen
# ─────────────────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
slide_header(slide, "Deutsche Unternehmensloesungen",
             "DSGVO-konforme Anbieter made in Germany")
slide_footer(slide, 7)

de_solutions = [
    ("fonio KI-\nTelefonassistent",
     ["Spezialist fuer Unternehmenstelefonie",
      "Anrufentgegennahme, Terminvereinbarung",
      "Anliegen erfassen & haeufige Fragen beantworten",
      "Server ausschliesslich in Deutschland",
      "DSGVO-konform, Daten in Europa"]),
    ("Vitas\nTelefonassistent",
     ["Made & hosted in Germany",
      "Kein auslaendischer Drittanbieter",
      "Zielmarkt: DACH-Region",
      "Branchenloesungen (Gesundheit, Dienstleistung)",
      "Vollstaendige Datensouveraenitaet"]),
    ("Personio\nAssistant",
     ["KI-Assistent fuer HR-Verwaltung",
      "Auskuenfte zu Urlaub, Gehalt, Recruiting",
      "Automatische Preboarding-Pakete",
      "Datenspeicherung innerhalb der EU",
      "Zielgruppe: mittelstaendische Unternehmen"]),
    ("Aleph Alpha /\nPharia",
     ["Deutsches KI-Unternehmen (Heidelberg)",
      "Souveraenes Sprachmodell fuer Unternehmen",
      "Integration mit STACKIT-Cloud",
      "Einsatz in Verwaltung & Unternehmen",
      "Volle Datensouveraenitaet, On-Premise moeglich"]),
    ("Dataport\nLLMoin",
     ["Oeffentlich-rechtliches IT-Dienstleistungsunternehmen",
      "Loesungen fuer Verwaltung und Unternehmen",
      "KI-Assistent basierend auf Open Source",
      "Nachnutzung fuer mehrere Bundeslaender",
      "Sicherer Betrieb in Deutschland"]),
    ("DeepL\nWrite / API",
     ["Weltweit gefuehrter Uebersetzungsdienst",
      "Sitz: Koeln, Deutschland",
      "DSGVO-konform, EU-Datenhaltung",
      "Text-Optimierung & KI-Uebersetzung",
      "Starke API-Integration fuer Unternehmen"]),
]

for i, (title, bullets) in enumerate(de_solutions):
    col = i % 3
    row = i // 3
    left = 0.3 + col * 4.35
    top  = 1.35 + row * 2.9
    card(slide, left, top, 4.1, 2.7, title, bullets,
         bullet_size=Pt(11), title_size=Pt(12))

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 8 – Vergleichsmatrix
# ─────────────────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
slide_header(slide, "Vergleichsmatrix", "Wichtigste Anbieter im direkten Vergleich")
slide_footer(slide, 8)

headers = ["Anbieter", "Typ", "Zielgruppe", "DSGVO/\nDatenschutz", "Daten-\nsouveraenitaet", "Preis-\nmodell"]
col_w   = [2.2, 2.0, 2.2, 1.7, 1.9, 1.9]

rows = [
    ["Microsoft Copilot",   "Produktivitaet",   "Enterprise",       "EU-Option",    "Konfigurierbar",  "Pro User/Monat"],
    ["SAP Joule",           "ERP-Assistent",    "SAP-Kunden",       "EU AI Act",    "Konfigurierbar",  "Im SAP-Abo"],
    ["LLMoin (Hamburg)",    "Verwaltung",       "Oeffentl. Sektor", "Vollstaendig", "DE-Server",       "Oeffentlich"],
    ["BaerGPT (Berlin)",    "Verwaltung",       "Berliner Verw.",   "Vollstaendig", "DE-Server",       "Open Source"],
    ["F13 / Aleph Alpha",   "Verwaltung",       "Bund + Laender",   "Vollstaendig", "DE/EU only",      "Lizenz"],
    ["fonio",               "Telefon-KI",       "KMU / Unternehm.", "DSGVO",        "DE-Server",       "Abo-Modell"],
    ["Personio Assist.",    "HR-Assistent",     "KMU",              "EU-Daten",     "EU only",         "Im HR-Abo"],
    ["Notion AI",           "Wissen/Doku",      "Teams & KMU",      "Konfigurierbar","Konfigurierbar", "Ab 11,50 EUR/M."],
]

# header row
x = 0.3
y = 1.35
for j, (h, w) in enumerate(zip(headers, col_w)):
    add_rect(slide, x, y, w - 0.04, 0.55, fill_color=COLOR_PRIMARY)
    add_textbox(slide, h, x + 0.05, y + 0.04, w - 0.1, 0.48,
                font_size=Pt(10), font_color=COLOR_WHITE, bold=True)
    x += w

for ri, row in enumerate(rows):
    bg = COLOR_WHITE if ri % 2 == 0 else COLOR_LIGHT
    x = 0.3
    y = 1.9 + ri * 0.62
    for j, (cell, w) in enumerate(zip(row, col_w)):
        add_rect(slide, x, y, w - 0.04, 0.58, fill_color=bg,
                 line_color=RGBColor(0xCC, 0xCC, 0xCC), line_width=Pt(0.5))
        color = COLOR_PRIMARY if j == 0 else COLOR_DARK
        bold  = j == 0
        add_textbox(slide, cell, x + 0.06, y + 0.06, w - 0.14, 0.48,
                    font_size=Pt(10), font_color=color, bold=bold)
        x += w

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 9 – Markttrends & Ausblick
# ─────────────────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
slide_header(slide, "Markttrends & Ausblick 2026+",
             "Wachstumstreiber, Konvergenz und Regulierung")
slide_footer(slide, 9)

trends = [
    ("Markt-\nwachstum",
     ["KI-Agenten-Markt: +45,3 % CAGR",
      "Von USD 8,6 Mrd. (2025) auf",
      "USD 263 Mrd. bis 2035 (global)",
      "Deutschland als fuehrender EU-Markt"]),
    ("Konvergenz\nvon Plattformen",
     ["SAP Joule + Microsoft Copilot\nfusionieren zu einer KI-Einheit",
      "Multi-Agenten-Systeme (Gartner Top 10)",
      "Browser-Agenten automatisieren Online-Aufgaben",
      "Voice-Agenten auf menschlichem Niveau"]),
    ("Datensouver-\naenitatstrend",
     ["Steigende Nachfrage nach DE/EU-Hosting",
      "Open-Source-Loesungen gewinnen an Bedeutung",
      "STACKIT, Aleph Alpha als deutsche Alternativen",
      "EU AI Act treibt Compliance-Anforderungen"]),
    ("Spezialisierung",
     ["Branchenspez. Agenten schlagen Generalisten",
      "Verwaltungs-KI als eigene Kategorie",
      "GovTech-Ecosystem waechst stark",
      "Nachnutzungskonzepte zwischen Laendern"]),
]

for i, (title, bullets) in enumerate(trends):
    col = i % 2
    row = i // 2
    left = 0.3 + col * 6.5
    top  = 1.35 + row * 2.95
    card(slide, left, top, 6.2, 2.8, title, bullets,
         bullet_size=Pt(13), title_size=Pt(15))

# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 10 – Fazit & Empfehlungen
# ─────────────────────────────────────────────────────────────────────────────
slide = prs.slides.add_slide(blank_layout)
slide_header(slide, "Fazit & Empfehlungen",
             "Entscheidungshilfen fuer den Einsatz KI-basierter Verwaltungsassistenten")
slide_footer(slide, 10)

# Left column – key findings
add_rect(slide, 0.3, 1.35, 6.0, 5.85, fill_color=COLOR_WHITE,
         line_color=COLOR_PRIMARY, line_width=Pt(1.5))
add_rect(slide, 0.3, 1.35, 6.0, 0.4, fill_color=COLOR_PRIMARY)
add_textbox(slide, "Kernerkenntnisse", 0.4, 1.37, 5.8, 0.36,
            font_size=Pt(14), font_color=COLOR_WHITE, bold=True)

findings = [
    "Breites Oekosystem: Markt reicht von",
    "  souveraenen DE-Loesungen bis zu globalen",
    "  Enterprise-Plattformen",
    "",
    "Oeffentliche Verwaltung: LLMoin, BaerGPT,",
    "  BayernKI, NRW.Genius, F13 und KIPITZ",
    "  decken fast alle Verwaltungsebenen ab",
    "",
    "Unternehmen: SAP Joule + Microsoft Copilot",
    "  als dominantes Enterprise-Duo; deutsche",
    "  Alternativen bei hohem Datenschutzbedarf",
    "",
    "DSGVO & Datensouveraenitaet sind zentrale",
    "  Kaufentscheidungskriterien in Deutschland",
    "",
    "Markt waechst mit > 45 % CAGR – fruehe",
    "  Einfuehrung verschafft Wettbewerbsvorteil",
]

txBox = slide.shapes.add_textbox(Inches(0.45), Inches(1.85), Inches(5.7), Inches(5.2))
tf = txBox.text_frame
tf.word_wrap = True
first = True
for line in findings:
    if first:
        p = tf.paragraphs[0]; first = False
    else:
        p = tf.add_paragraph()
    p.space_before = Pt(1)
    run = p.add_run()
    run.text = line
    run.font.size = Pt(12)
    run.font.color.rgb = COLOR_DARK
    run.font.name = "Calibri"

# Right column – recommendations
add_rect(slide, 6.7, 1.35, 6.3, 5.85, fill_color=COLOR_WHITE,
         line_color=COLOR_ACCENT, line_width=Pt(1.5))
add_rect(slide, 6.7, 1.35, 6.3, 0.4, fill_color=COLOR_ACCENT)
add_textbox(slide, "Empfehlungen", 6.8, 1.37, 6.1, 0.36,
            font_size=Pt(14), font_color=COLOR_WHITE, bold=True)

recs = [
    ("Oeffentliche Verwaltung:",
     "Nachnutzung bestehender Landesprojekte\n(LLMoin, BayernKI) pruefen; KIPITZ fuer\nBundesbehoerden als Einstieg nutzen"),
    ("Mittelstaendische Unternehmen:",
     "Personio Assistant fuer HR; fonio oder\nVitas fuer Telefonie; Notion AI fuer\nWissensmanagement"),
    ("Grosse Unternehmen:",
     "SAP Joule + Microsoft Copilot als\nintegrierten Stack; Aleph Alpha fuer\nhoehere Datensouveraenitaetsanforderungen"),
    ("Alle Organisationen:",
     "EU AI Act-Konformitaet fruehzeitig\nplanen; Pilotprojekte starten; interne\nKI-Kompetenzen aufbauen"),
]

y = 1.9
for title, text in recs:
    add_rect(slide, 6.75, y, 6.15, 0.28, fill_color=RGBColor(0xE8, 0xF4, 0xFF))
    add_textbox(slide, title, 6.85, y + 0.02, 5.9, 0.24,
                font_size=Pt(11), font_color=COLOR_PRIMARY, bold=True)
    add_textbox(slide, text, 6.85, y + 0.3, 5.9, 0.8,
                font_size=Pt(11), font_color=COLOR_DARK)
    y += 1.3

# ─────────────────────────────────────────────────────────────────────────────
# Save
# ─────────────────────────────────────────────────────────────────────────────
out = "/home/user/TestEva/KI_Verwaltungsassistenten_Marktrecherche.pptx"
prs.save(out)
print(f"Saved: {out}")
