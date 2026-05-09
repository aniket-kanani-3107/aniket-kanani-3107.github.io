"""
gen_german_tax_tool.py
======================
Generates German_Tax_Tool_Plan.docx — a complete, bilingual (DE/EN) plan for an
offline, localhost-based German Tax Return Preparation & Financial Statements tool.

Run:  python gen_german_tax_tool.py
Output: German_Tax_Tool_Plan.docx  (same directory)
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

# ── colour palette ────────────────────────────────────────────────────────────
BLACK   = RGBColor(0x00, 0x00, 0x00)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
GOLD    = RGBColor(0xD4, 0xAF, 0x37)
DARK    = RGBColor(0x1A, 0x1A, 0x2E)
BLUE    = RGBColor(0x16, 0x21, 0x3E)
ACCENT  = RGBColor(0x0F, 0x3C, 0x5F)
SILVER  = RGBColor(0xAA, 0xAA, 0xAA)
LGREY   = RGBColor(0xF5, 0xF5, 0xF5)

# ── helpers ───────────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def add_run(para, text, bold=False, italic=False, size=11,
            color=BLACK, underline=False):
    run = para.add_run(text)
    run.bold      = bold
    run.italic    = italic
    run.underline = underline
    run.font.size  = Pt(size)
    run.font.color.rgb = color
    return run


def heading(doc, text_en, text_de, level=1):
    """Bilingual heading: English / German"""
    p = doc.add_paragraph(style=f"Heading {level}")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run1 = p.add_run(text_en)
    run1.font.color.rgb = GOLD
    run1.bold = True
    run1.font.size = Pt(16 - (level - 1) * 2)
    if text_de:
        run2 = p.add_run(f"  /  {text_de}")
        run2.font.color.rgb = SILVER
        run2.bold = False
        run2.font.size = Pt(14 - (level - 1) * 2)
    return p


def bilingual_bullet(doc, en, de, bullet="•"):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent  = Inches(0.3)
    p.paragraph_format.space_after  = Pt(2)
    add_run(p, f"{bullet} ", bold=True, color=GOLD, size=11)
    add_run(p, en, bold=True, size=11)
    add_run(p, f"  —  {de}", italic=True, color=SILVER, size=10)


def section_divider(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run("─" * 110)
    run.font.color.rgb = ACCENT
    run.font.size = Pt(8)


def info_table(doc, rows_data, col_widths=(3.2, 3.2)):
    """Two-column bilingual info table."""
    table = doc.add_table(rows=1 + len(rows_data), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Table Grid"
    # header
    hdr = table.rows[0].cells
    for i, lbl in enumerate(["🇬🇧  English", "🇩🇪  Deutsch"]):
        hdr[i].text = lbl
        set_cell_bg(hdr[i], "1A1A2E")
        for run in hdr[i].paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = GOLD
            run.font.size = Pt(11)
    # data rows
    for r_idx, (en, de) in enumerate(rows_data):
        row = table.rows[r_idx + 1].cells
        row[0].text = en
        row[1].text = de
        fill = "F5F5F5" if r_idx % 2 == 0 else "FFFFFF"
        set_cell_bg(row[0], fill)
        set_cell_bg(row[1], fill)
        for cell in row:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(10)
    # column widths
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            cell.width = Inches(col_widths[idx])
    return table


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN DOCUMENT BUILDER
# ═══════════════════════════════════════════════════════════════════════════════

def build_document():
    doc = Document()

    # ── page margins ──────────────────────────────────────────────────────────
    for section in doc.sections:
        section.top_margin    = Cm(1.8)
        section.bottom_margin = Cm(1.8)
        section.left_margin   = Cm(2.0)
        section.right_margin  = Cm(2.0)

    # ══════════════════════════════════════════════════════════════════════════
    #  COVER PAGE
    # ══════════════════════════════════════════════════════════════════════════
    for _ in range(3):
        doc.add_paragraph()

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(title_p, "🏛  GERMAN TAX RETURN TOOL", bold=True, size=26, color=GOLD)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(sub_p, "Steuererklärung-Tool für Deutschland", italic=True, size=16, color=SILVER)

    doc.add_paragraph()
    tag_p = doc.add_paragraph()
    tag_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(tag_p,
            "Offline · Localhost · Bilingual (EN / DE) · Excel Export · Financial Statements",
            italic=True, size=12, color=SILVER)

    doc.add_paragraph()
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(date_p, f"Prepared: {datetime.date.today().strftime('%B %d, %Y')}",
            size=10, color=SILVER)

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  TABLE OF CONTENTS  (manual)
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "Table of Contents", "Inhaltsverzeichnis", level=1)
    toc_items = [
        ("1. German Tax System — Overview",          "1. Das deutsche Steuersystem – Überblick"),
        ("2. Tax Forms & Filing Process",             "2. Steuerformulare & Einreichungsprozess"),
        ("3. Project Vision & Goals",                 "3. Projektvision & Ziele"),
        ("4. Application Architecture",               "4. Anwendungsarchitektur"),
        ("5. Technology Stack",                       "5. Technologie-Stack"),
        ("6. File & Folder Structure",                "6. Datei- & Ordnerstruktur"),
        ("7. Module-by-Module Build Plan",            "7. Modul-für-Modul Bauplan"),
        ("   7.1  Personal Info Module",              "   7.1  Persönliche-Daten-Modul"),
        ("   7.2  Income Module",                     "   7.2  Einkommensmodul"),
        ("   7.3  Deductions Module",                 "   7.3  Abzugsmodul"),
        ("   7.4  Tax Calculation Engine",            "   7.4  Steuerberechnungs-Engine"),
        ("   7.5  Financial Statements Module",       "   7.5  Finanzberichte-Modul"),
        ("   7.6  Excel Export Module",               "   7.6  Excel-Export-Modul"),
        ("   7.7  Bilingual UI Module",               "   7.7  Zweisprachiges UI-Modul"),
        ("   7.8  Localhost Server",                  "   7.8  Localhost-Server"),
        ("8. Database Schema",                        "8. Datenbankschema"),
        ("9. Excel Export — Sheets & Layout",         "9. Excel-Export – Blätter & Layout"),
        ("10. Financial Statements Details",          "10. Finanzbericht-Details"),
        ("11. German Tax Law Reference",              "11. Deutsches Steuerrecht – Referenz"),
        ("12. Step-by-Step Development Roadmap",      "12. Schritt-für-Schritt Entwicklungsplan"),
        ("13. Running & Deployment Guide",            "13. Ausführungs- & Bereitstellungsanleitung"),
        ("14. User Workflow Walkthrough",             "14. Benutzer-Workflow"),
        ("15. Security & Data Privacy",               "15. Sicherheit & Datenschutz"),
        ("16. Glossary",                              "16. Glossar"),
    ]
    for en, de in toc_items:
        bilingual_bullet(doc, en, de, bullet="▸")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  1. GERMAN TAX SYSTEM OVERVIEW
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "1. German Tax System — Overview",
            "1. Das deutsche Steuersystem – Überblick")

    p = doc.add_paragraph()
    add_run(p,
        "Germany operates a progressive income tax system governed primarily by the "
        "Einkommensteuergesetz (EStG – Income Tax Act).  Every resident with taxable "
        "income above the basic allowance (Grundfreibetrag) must file an annual tax "
        "return (Einkommensteuererklärung) via ELSTER or paper forms.  The fiscal year "
        "equals the calendar year (1 Jan – 31 Dec); the filing deadline is normally "
        "31 July of the following year (extended to 28/29 Feb if a tax adviser is used).",
        size=11)

    section_divider(doc)
    heading(doc, "1.1  Tax Classes (Steuerklassen)", "", level=2)
    info_table(doc, [
        ("Class I  – Single, divorced, widowed (>2 yrs)",   "Klasse I  – Ledig, geschieden, verwitwet (>2 J.)"),
        ("Class II – Single parent (Alleinerziehend)",        "Klasse II – Alleinerziehend"),
        ("Class III – Married / higher earner",               "Klasse III – Verheiratet / Besserverdiener"),
        ("Class IV – Married / similar income",               "Klasse IV – Verheiratet / ähnl. Einkommen"),
        ("Class V  – Married / lower earner",                 "Klasse V  – Verheiratet / Geringverdiener"),
        ("Class VI – Second / additional employment",         "Klasse VI – Zweiter / weiterer Job"),
    ])

    doc.add_paragraph()
    heading(doc, "1.2  Key Tax Rates 2024", "1.2  Wichtige Steuersätze 2024", level=2)
    info_table(doc, [
        ("Basic allowance (Grundfreibetrag): €11,784",       "Grundfreibetrag: 11.784 €"),
        ("Starting rate: 14 %  (income > €11,784)",          "Eingangssteuersatz: 14 % (ab 11.784 €)"),
        ("Top rate: 42 %  (income > €66,761)",               "Spitzensteuersatz: 42 % (ab 66.761 €)"),
        ("Solidarity surcharge: 5.5 % of income tax",        "Solidaritätszuschlag: 5,5 % der ESt"),
        ("Church tax: 8 – 9 % (Bavaria/BW: 8 %)",           "Kirchensteuer: 8–9 %"),
        ("Capital gains flat tax: 25 % + Soli",              "Abgeltungssteuer: 25 % + Soli"),
        ("Trade tax (GewSt): varies by municipality",         "Gewerbesteuer: je nach Gemeinde"),
        ("VAT standard rate: 19 %",                          "Mehrwertsteuer (MwSt): 19 %"),
        ("VAT reduced rate: 7 %",                            "MwSt ermäßigt: 7 %"),
    ])

    doc.add_paragraph()
    heading(doc, "1.3  Income Types (Einkunftsarten)", "", level=2)
    bilingual_bullet(doc, "Employment income (Einkünfte aus nichtselbstständiger Arbeit)", "§ 19 EStG")
    bilingual_bullet(doc, "Self-employment / Freelance (Selbstständige Arbeit)", "§ 18 EStG")
    bilingual_bullet(doc, "Business income (Gewerbebetrieb)", "§ 15 EStG")
    bilingual_bullet(doc, "Agricultural income (Land- und Forstwirtschaft)", "§ 13 EStG")
    bilingual_bullet(doc, "Rental income (Vermietung & Verpachtung)", "§ 21 EStG")
    bilingual_bullet(doc, "Capital income (Kapitalvermögen)", "§ 20 EStG")
    bilingual_bullet(doc, "Other income (Sonstige Einkünfte)", "§ 22 EStG")

    doc.add_paragraph()
    heading(doc, "1.4  Major Deductions", "1.4  Wichtige Abzüge", level=2)
    info_table(doc, [
        ("Employee flat-rate (Werbungskostenpauschale): €1,230",  "Arbeitnehmer-Pauschbetrag: 1.230 €"),
        ("Special expenses (Sonderausgaben) – various",           "Sonderausgaben – verschiedene"),
        ("Extraordinary burdens (Außergewöhnliche Belastungen)",   "Außergewöhnliche Belastungen"),
        ("Health/care insurance premiums (fully deductible)",      "Kranken-/Pflegeversicherungsbeiträge"),
        ("Pension contributions (Altersvorsorgeaufwendungen)",     "Altersvorsorgeaufwendungen"),
        ("Home office flat rate: €6/day, max €1,260/yr",          "Homeoffice-Pauschale: 6 €/Tag, max. 1.260 €/J."),
        ("Child allowance (Kinderfreibetrag): €6,384/child",      "Kinderfreibetrag: 6.384 € je Kind"),
        ("Donations (Spenden) up to 20 % of income",              "Spenden bis 20 % des Einkommens"),
        ("Commute flat rate: €0.30/km (first 20 km), €0.38 beyond","Pendlerpauschale: 0,30 €/km (ab 21. km: 0,38 €)"),
        ("Professional education / training costs",                "Fortbildungskosten / Weiterbildung"),
        ("Trade union fees (Gewerkschaftsbeiträge)",               "Gewerkschaftsbeiträge"),
    ])

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  2. TAX FORMS & FILING PROCESS
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "2. Tax Forms & Filing Process",
            "2. Steuerformulare & Einreichungsprozess")

    info_table(doc, [
        ("Mantelbogen (ESt 1 A) — Main form for all taxpayers",
         "Mantelbogen (ESt 1 A) — Hauptformular für alle Steuerzahler"),
        ("Anlage N — Employment income",                 "Anlage N — Einkünfte aus nicht-selbst. Arbeit"),
        ("Anlage S — Self-employment / Freelance",       "Anlage S — Selbstständige Arbeit"),
        ("Anlage G — Business income",                   "Anlage G — Einkünfte aus Gewerbebetrieb"),
        ("Anlage V — Rental income",                     "Anlage V — Vermietung & Verpachtung"),
        ("Anlage KAP — Capital income",                  "Anlage KAP — Kapitalvermögen"),
        ("Anlage R — Pension income",                    "Anlage R — Renteneinkünfte"),
        ("Anlage Kind — Child-related deductions",       "Anlage Kind — Kindbezogene Abzüge"),
        ("Anlage Sonderausgaben — Special expenses",     "Anlage Sonderausgaben — Sonderausgaben"),
        ("Anlage AV — Riester pension (§ 10a EStG)",    "Anlage AV — Riester-Rente"),
        ("Anlage Haushaltsnahe — Household services",   "Anlage Haushaltsnahe Dienstleistungen"),
        ("EÜR — Income-surplus calculation (freelance)","EÜR — Einnahmen-Überschuss-Rechnung"),
        ("UStVA — VAT pre-declaration (quarterly)",     "UStVA — Umsatzsteuervoranmeldung"),
        ("GewSt — Trade tax declaration",               "GewSt-Erklärung — Gewerbesteuer"),
    ])

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    add_run(p, "Filing Process: ", bold=True, color=GOLD)
    add_run(p,
        "Gather documents (Lohnsteuerbescheinigung, bank statements, receipts) → "
        "Complete relevant Anlagen → Calculate tax liability → Submit via ELSTER portal "
        "or paper → Receive Steuerbescheid (assessment) → Appeal within 1 month if disputed.",
        size=11)

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  3. PROJECT VISION & GOALS
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "3. Project Vision & Goals",
            "3. Projektvision & Ziele")

    goals = [
        ("Offline-first desktop web app running on localhost — no internet needed",
         "Offline-Desktop-Web-App auf localhost – kein Internet erforderlich"),
        ("Support ALL income categories & every standard German tax form",
         "Alle Einkunftsarten & alle Standard-Steuerformulare unterstützen"),
        ("Automatic tax calculation (income tax, Soli, church tax, trade tax)",
         "Automatische Steuerberechnung (ESt, Soli, KiSt, GewSt)"),
        ("Generate full financial statements: P&L, Balance Sheet, Cash Flow",
         "Vollständige Finanzberichte: GuV, Bilanz, Kapitalflussrechnung"),
        ("One-click export to professionally formatted Excel workbook (.xlsx)",
         "Ein-Klick-Export in professionell formatierte Excel-Arbeitsmappe (.xlsx)"),
        ("Bilingual UI: switch between English and German at any time",
         "Zweisprachige Benutzeroberfläche: jederzeit zwischen EN und DE wechseln"),
        ("Secure local SQLite database — all data stays on your machine",
         "Sichere lokale SQLite-Datenbank – alle Daten verbleiben auf Ihrem Gerät"),
        ("Multi-year support: manage multiple tax years side-by-side",
         "Mehrjährige Unterstützung: mehrere Steuerjahre parallel verwalten"),
        ("Professional PDF print-ready reports",
         "Professionelle, druckfertige PDF-Berichte"),
        ("Dashboard with visual charts & KPIs",
         "Dashboard mit visuellen Diagrammen & KPIs"),
    ]
    for en, de in goals:
        bilingual_bullet(doc, en, de)

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  4. APPLICATION ARCHITECTURE
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "4. Application Architecture",
            "4. Anwendungsarchitektur")

    p = doc.add_paragraph()
    add_run(p,
        "The application follows a classic MVC (Model-View-Controller) pattern "
        "wrapped in a lightweight Python Flask web server.  The browser (Chrome/Firefox) "
        "acts as the display layer; all computation, storage, and export logic runs "
        "server-side in Python.  No cloud services — everything is local.",
        size=11)

    doc.add_paragraph()
    arch_table = info_table(doc, [
        ("Layer",                   "Schicht"),
        ("Browser (HTML/CSS/JS)",   "Browser – Darstellungsschicht"),
        ("Flask Routes (API)",      "Flask-Routen – Controller"),
        ("Business Logic (Python)", "Geschäftslogik – Python-Module"),
        ("SQLite Database",         "SQLite-Datenbank – Modell"),
        ("Excel / PDF Export",      "Excel / PDF Export"),
    ])

    doc.add_paragraph()
    p = doc.add_paragraph()
    add_run(p, "Communication: ", bold=True, color=GOLD)
    add_run(p,
        "The frontend sends JSON requests to Flask REST endpoints.  "
        "Flask processes the request, reads/writes SQLite, and returns JSON.  "
        "Export buttons trigger server-side file generation (openpyxl / reportlab) "
        "and the browser downloads the file.",
        size=11)

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  5. TECHNOLOGY STACK
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "5. Technology Stack",
            "5. Technologie-Stack")

    info_table(doc, [
        ("Python 3.11+",              "Backend-Sprache"),
        ("Flask 3.x",                 "Web-Framework (Localhost-Server)"),
        ("SQLite 3",                  "Lokale Datenbank (keine Installation nötig)"),
        ("SQLAlchemy 2.x",            "ORM – Datenbankabstraktion"),
        ("openpyxl 3.x",              "Excel-Export (.xlsx)"),
        ("reportlab",                 "PDF-Generierung"),
        ("Jinja2",                    "HTML-Template-Engine (in Flask)"),
        ("Bootstrap 5",               "Responsive CSS Framework"),
        ("Chart.js",                  "Dashboard-Diagramme"),
        ("i18next (JS)",              "Mehrsprachigkeit im Browser"),
        ("Flask-Babel",               "Server-seitige Übersetzung"),
        ("Werkzeug",                  "WSGI-Utilities"),
        ("PyInstaller (optional)",    "Standalone-EXE-Paketierung"),
    ])

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  6. FILE & FOLDER STRUCTURE
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "6. File & Folder Structure",
            "6. Datei- & Ordnerstruktur")

    structure = """german_tax_tool/
│
├── app.py                          ← Flask entry point / server start
├── config.py                       ← App-wide settings (DB path, language, etc.)
├── requirements.txt                ← Python dependencies
├── run.bat / run.sh                ← One-click launch scripts
│
├── database/
│   ├── db.py                       ← SQLAlchemy setup & session
│   ├── models.py                   ← All ORM models (TaxReturn, Income, …)
│   └── migrations/                 ← Alembic migration scripts
│
├── modules/
│   ├── personal_info.py            ← Personal data handler
│   ├── income.py                   ← All income types
│   ├── deductions.py               ← All deduction types
│   ├── tax_engine.py               ← Core tax calculation logic
│   ├── financial_statements.py     ← P&L, Balance Sheet, Cash Flow
│   ├── excel_export.py             ← openpyxl workbook builder
│   ├── pdf_export.py               ← reportlab PDF builder
│   └── validators.py               ← Input validation rules
│
├── routes/
│   ├── __init__.py
│   ├── personal.py                 ← /api/personal endpoints
│   ├── income.py                   ← /api/income endpoints
│   ├── deductions.py               ← /api/deductions endpoints
│   ├── tax.py                      ← /api/tax/calculate endpoint
│   ├── statements.py               ← /api/statements endpoints
│   └── export.py                   ← /api/export/excel, /api/export/pdf
│
├── templates/
│   ├── base.html                   ← Base layout (navbar, i18n toggle)
│   ├── index.html                  ← Dashboard / Home
│   ├── personal.html               ← Personal info form
│   ├── income.html                 ← Income entry forms
│   ├── deductions.html             ← Deductions forms
│   ├── tax_result.html             ← Calculation result & summary
│   ├── statements.html             ← Financial statements view
│   └── export.html                 ← Export options page
│
├── static/
│   ├── css/
│   │   ├── custom.css              ← Custom styling (dark/light theme)
│   │   └── print.css               ← Print-optimised stylesheet
│   ├── js/
│   │   ├── main.js                 ← Global JS utilities
│   │   ├── i18n.js                 ← Language switching logic
│   │   └── charts.js               ← Chart.js dashboard charts
│   └── img/
│       └── logo.svg
│
├── translations/
│   ├── en/
│   │   └── messages.json           ← English UI strings
│   └── de/
│       └── messages.json           ← German UI strings
│
├── exports/                        ← Generated Excel/PDF files (auto-created)
│
└── tests/
    ├── test_tax_engine.py
    ├── test_income.py
    ├── test_deductions.py
    └── test_export.py"""

    p = doc.add_paragraph()
    run = p.add_run(structure)
    run.font.name = "Courier New"
    run.font.size = Pt(8.5)
    run.font.color.rgb = ACCENT

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  7. MODULE-BY-MODULE BUILD PLAN
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "7. Module-by-Module Build Plan",
            "7. Modul-für-Modul Bauplan")

    # 7.1
    heading(doc, "7.1  Personal Info Module  (personal_info.py)",
            "Persönliche-Daten-Modul", level=2)
    fields = [
        ("Full name / Vollständiger Name",       "First, last, title"),
        ("Tax ID / Steueridentifikationsnummer", "11-digit TIN"),
        ("Date of birth / Geburtsdatum",         "DD.MM.YYYY"),
        ("Marital status / Familienstand",       "Single/Married/Divorced/Widowed"),
        ("Tax class / Steuerklasse",             "I – VI"),
        ("Church membership / Kirchenmitglied",  "Yes/No + state"),
        ("Address / Adresse",                    "Street, PLZ, City, State (Bundesland)"),
        ("Tax office / Finanzamt",               "Name + number"),
        ("Spouse data / Ehegattendaten",         "If Klasse III/IV/V"),
        ("Children / Kinder",                    "Name, DOB, TIN each"),
        ("Bank account / Bankverbindung",        "IBAN, BIC for refund"),
    ]
    for field, note in fields:
        bilingual_bullet(doc, field, note)

    section_divider(doc)

    # 7.2
    heading(doc, "7.2  Income Module  (income.py)",
            "Einkommensmodul", level=2)
    income_fields = [
        ("Gross salary / Bruttolohn",                "From Lohnsteuerbescheinigung line 3"),
        ("Income tax withheld / Einbehaltene LSt",   "From Lohnsteuerbescheinigung line 4"),
        ("Solidarity withheld / Einbeh. Soli",       "From Lohnsteuerbescheinigung line 5"),
        ("Church tax withheld / Einbeh. KiSt",       "From Lohnsteuerbescheinigung line 6"),
        ("Freelance revenue / Einnahmen §18",         "Gross receipts (Betriebseinnahmen)"),
        ("Business revenue / Einnahmen §15",          "Gross receipts"),
        ("Rental income / Mieteinnahmen",             "Annual rent received"),
        ("Interest & dividends / Zinsen & Dividenden","Gross capital income"),
        ("Pension / Rente",                           "Annual pension received"),
        ("Other income / Sonstige Einkünfte",         "§22 – specify type"),
        ("Foreign income / Ausländische Einkünfte",   "DBA consideration"),
    ]
    for en, de in income_fields:
        bilingual_bullet(doc, en, de)

    section_divider(doc)

    # 7.3
    heading(doc, "7.3  Deductions Module  (deductions.py)",
            "Abzugsmodul", level=2)
    ded_fields = [
        ("Work-related expenses / Werbungskosten",         "Commute, equipment, union fees …"),
        ("Home office / Homeoffice-Pauschale",             "Days × €6, max €1,260"),
        ("Health insurance / Krankenversicherung",         "Beiträge Basis-KV"),
        ("Pension insurance / Rentenversicherung",         "DRV contributions"),
        ("Riester pension / Riester-Rente",                "§10a – with Anlage AV"),
        ("Rürup pension / Rürup-Rente",                   "§10 – up to €27,566 (2024)"),
        ("Donations / Spenden",                            "Receipts required"),
        ("Child care / Kinderbetreuungskosten",            "2/3 of costs, max €4,000/child"),
        ("Household services / Haushaltsnahe Dienste",     "20 % of cost, max €4,000"),
        ("Craft services / Handwerkerleistungen",          "20 % of labour, max €1,200"),
        ("Disability lump sum / Behinderten-Pauschbetrag", "Degree-dependent"),
        ("Extraordinary burdens / Außergewöhnl. Belast.", "Illness, death, catastrophe costs"),
        ("Training / Fortbildungskosten",                  "Courses, books, travel"),
        ("Professional membership / Berufsverbände",       "Annual fees"),
    ]
    for en, de in ded_fields:
        bilingual_bullet(doc, en, de)

    section_divider(doc)

    # 7.4
    heading(doc, "7.4  Tax Calculation Engine  (tax_engine.py)",
            "Steuerberechnungs-Engine", level=2)
    p = doc.add_paragraph()
    add_run(p, "Algorithm steps (§ 32a EStG 2024):", bold=True, color=GOLD)
    steps_calc = [
        ("Step 1", "Sum all income types → Gesamtbetrag der Einkünfte"),
        ("Step 2", "Subtract special expenses (Sonderausgaben)"),
        ("Step 3", "Subtract extraordinary burdens (Außergewöhnliche Belastungen)"),
        ("Step 4", "Apply loss carry-forward / carry-back (Verlustabzug)"),
        ("Step 5", "Apply basic allowance deduction → zu versteuerndes Einkommen (zvE)"),
        ("Step 6", "Apply tax tariff formula (5 tax zones) → Income Tax (ESt)"),
        ("Step 7", "Apply Solidarity Surcharge (Soli): 5.5 % of ESt if applicable"),
        ("Step 8", "Add Church Tax (KiSt): 8 or 9 % of ESt"),
        ("Step 9", "Subtract wage-tax already withheld → Refund or Payment due"),
        ("Step 10","Include capital gains tax (Abgeltungssteuer) if applicable"),
    ]
    for step, desc in steps_calc:
        p2 = doc.add_paragraph()
        p2.paragraph_format.left_indent = Inches(0.3)
        add_run(p2, f"{step}: ", bold=True, color=GOLD, size=10)
        add_run(p2, desc, size=10)

    p3 = doc.add_paragraph()
    add_run(p3, "\nTax zones for zvE (§ 32a Abs. 1 EStG 2024):", bold=True, color=GOLD)
    zones = [
        ("Zone 1", "€0 – €11,784",        "0 % — Grundfreibetrag"),
        ("Zone 2", "€11,785 – €17,005",   "14 % – 23.97 % (linear progression)"),
        ("Zone 3", "€17,006 – €66,760",   "23.97 % – 42 % (linear progression)"),
        ("Zone 4", "€66,761 – €277,825",  "42 % (flat)"),
        ("Zone 5", "> €277,825",          "45 % (Reichensteuer)"),
    ]
    for zone, rng, rate in zones:
        p4 = doc.add_paragraph()
        p4.paragraph_format.left_indent = Inches(0.3)
        add_run(p4, f"{zone} {rng}: ", bold=True, size=10, color=ACCENT)
        add_run(p4, rate, size=10)

    section_divider(doc)

    # 7.5
    heading(doc, "7.5  Financial Statements Module  (financial_statements.py)",
            "Finanzberichte-Modul", level=2)
    fin_items = [
        ("Profit & Loss (GuV / EÜR)",         "Revenue – Expenses = Net Profit/Loss"),
        ("Balance Sheet (Bilanz)",             "Assets = Liabilities + Equity"),
        ("Cash Flow Statement (Kapitalfluss)", "Operating / Investing / Financing activities"),
        ("Net Worth Statement",               "Total assets minus total liabilities"),
        ("Tax Summary Sheet",                 "All taxes computed in one view"),
        ("Year-over-Year Comparison",         "Current vs prior year variance analysis"),
    ]
    for en, de in fin_items:
        bilingual_bullet(doc, en, de)

    section_divider(doc)

    # 7.6
    heading(doc, "7.6  Excel Export Module  (excel_export.py)",
            "Excel-Export-Modul", level=2)
    sheet_names = [
        ("Sheet 1 — Cover",             "Title, taxpayer name, tax year, date"),
        ("Sheet 2 — Personal Info",     "All personal fields"),
        ("Sheet 3 — Income Summary",    "All income categories with subtotals"),
        ("Sheet 4 — Deductions",        "All deduction categories with totals"),
        ("Sheet 5 — Tax Calculation",   "Step-by-step tax computation"),
        ("Sheet 6 — Tax Result",        "Final tax, refund or payment due"),
        ("Sheet 7 — P&L",               "Profit & Loss Statement"),
        ("Sheet 8 — Balance Sheet",     "Assets, Liabilities, Equity"),
        ("Sheet 9 — Cash Flow",         "Operating / Investing / Financing"),
        ("Sheet 10 — Supporting Data",  "Raw entries, receipts log"),
        ("Sheet 11 — Charts (Data)",    "Source data for embedded charts"),
        ("Sheet 12 — Year Comparison",  "Multi-year side-by-side"),
    ]
    for en, de in sheet_names:
        bilingual_bullet(doc, en, de)

    section_divider(doc)

    # 7.7
    heading(doc, "7.7  Bilingual UI Module  (i18n)",
            "Zweisprachiges UI-Modul", level=2)
    bilingual_bullet(doc,
        "translations/en/messages.json  +  translations/de/messages.json",
        "All UI strings stored as JSON key-value pairs")
    bilingual_bullet(doc,
        "Language toggle button in navbar (🇬🇧 / 🇩🇪)",
        "Sofortige Umschaltung ohne Seitenneuladen (JavaScript i18next)")
    bilingual_bullet(doc,
        "Flask-Babel handles server-rendered translations",
        "Flask-Babel verwaltet server-seitige Übersetzungen")
    bilingual_bullet(doc,
        "Date/Number formats auto-adapt: EN uses comma as thousand sep, DE uses period",
        "Datums-/Zahlenformate passen sich automatisch an")

    section_divider(doc)

    # 7.8
    heading(doc, "7.8  Localhost Server  (app.py)",
            "Localhost-Server", level=2)
    bilingual_bullet(doc,
        "flask run --host=127.0.0.1 --port=5000",
        "Startet den Server; im Browser http://localhost:5000 aufrufen")
    bilingual_bullet(doc,
        "run.bat (Windows) / run.sh (Linux/Mac) — double-click to launch",
        "Doppelklick auf Startdatei öffnet automatisch den Browser")
    bilingual_bullet(doc,
        "Optional: package as standalone .exe with PyInstaller (no Python needed)",
        "Optional: als EXE paketieren mit PyInstaller (kein Python erforderlich)")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  8. DATABASE SCHEMA
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "8. Database Schema",
            "8. Datenbankschema")

    tables = [
        ("taxpayers",        "id, first_name, last_name, tax_id, dob, marital_status, tax_class, church_member, address, finanzamt, iban"),
        ("tax_returns",      "id, taxpayer_id, tax_year, status, created_at, updated_at"),
        ("income_entries",   "id, return_id, income_type, description, amount_gross, amount_net, currency"),
        ("deduction_entries","id, return_id, deduction_type, description, amount, receipt_ref"),
        ("tax_results",      "id, return_id, taxable_income, income_tax, soli, church_tax, total_tax, withheld_tax, refund_or_payment"),
        ("financial_data",   "id, return_id, statement_type (PL/BS/CF), line_item, amount, year"),
        ("children",         "id, taxpayer_id, first_name, dob, tax_id, custody_pct"),
        ("spouses",          "id, taxpayer_id, name, tax_id, dob, tax_class"),
        ("settings",         "key, value  (language, theme, currency_format, etc.)"),
    ]
    info_table(doc, [(t, f) for t, f in tables], col_widths=(2.2, 4.2))

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  9. EXCEL EXPORT DETAIL
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "9. Excel Export — Sheets & Layout",
            "9. Excel-Export – Blätter & Layout")

    p = doc.add_paragraph()
    add_run(p,
        "The export engine uses openpyxl to build a multi-sheet workbook.  "
        "Professional formatting: header rows in dark blue (#1A1A2E) with gold (#D4AF37) "
        "text, alternating row colours, bold totals, number formats with € symbol, "
        "borders, and frozen header rows.  Charts are embedded on the Charts sheet.",
        size=11)

    info_table(doc, [
        ("Cover Sheet",        "Logo (optional), taxpayer name, TIN, tax year, generated timestamp"),
        ("Personal Info",      "Label | Value layout, two columns, all personal fields"),
        ("Income Summary",     "Category | Gross | Deductions | Net | % of Total"),
        ("Deductions",         "Type | Description | Amount | Cumulative Total"),
        ("Tax Calculation",    "Step | Description | Amount — mirrors the tax engine logic"),
        ("Tax Result",         "Highlighted result: REFUND (green) or PAYMENT DUE (red)"),
        ("P&L Statement",      "Revenue, COGS, Gross Profit, Operating Expenses, Net Profit"),
        ("Balance Sheet",      "Assets (current/non-current), Liabilities, Equity — balanced"),
        ("Cash Flow",          "Operating, Investing, Financing, Net Change in Cash"),
        ("Supporting Data",    "Raw data rows with sortable/filterable Excel tables"),
        ("Charts (Data)",      "Bar chart (income breakdown), Pie chart (tax distribution)"),
        ("Year Comparison",    "Columns per year; % change highlighted automatically"),
    ], col_widths=(2.5, 3.9))

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  10. FINANCIAL STATEMENTS
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "10. Financial Statements Details",
            "10. Finanzbericht-Details")

    heading(doc, "10.1  Profit & Loss (GuV / EÜR)", "", level=2)
    info_table(doc, [
        ("Revenue (Umsatzerlöse)",             "All income categories subtotal"),
        ("Cost of Goods Sold (Wareneinsatz)",  "Direct costs (if applicable)"),
        ("Gross Profit (Rohertrag)",           "Revenue – COGS"),
        ("Operating Expenses (Betriebsausgaben)", "Rent, salaries, depreciation, utilities …"),
        ("EBIT",                               "Gross Profit – Operating Expenses"),
        ("Financial Income/Expense",           "Interest received / paid"),
        ("EBT (Vorsteuergewinn)",              "EBIT ± Financial items"),
        ("Income Tax Expense",                 "Computed by tax engine"),
        ("Net Profit / Loss (Jahresüberschuss)","EBT – Taxes"),
    ])

    doc.add_paragraph()
    heading(doc, "10.2  Balance Sheet (Bilanz)", "", level=2)
    info_table(doc, [
        ("ASSETS — Current: Cash, Bank, Receivables, Inventory",  "Aktiva – Umlaufvermögen"),
        ("ASSETS — Non-current: Equipment, Vehicles, Goodwill",   "Aktiva – Anlagevermögen"),
        ("LIABILITIES — Current: Payables, Short-term loans",     "Passiva – Kurzfristige Verbindlichkeiten"),
        ("LIABILITIES — Non-current: Long-term loans, Mortgages", "Passiva – Langfristige Verbindlichkeiten"),
        ("EQUITY: Paid-in capital + Retained earnings",           "Eigenkapital"),
        ("Check: Total Assets = Total Liabilities + Equity",      "Prüfung: Bilanzsumme ausgeglichen"),
    ])

    doc.add_paragraph()
    heading(doc, "10.3  Cash Flow Statement (Kapitalflussrechnung)", "", level=2)
    info_table(doc, [
        ("Operating activities (Betriebliche Tätigkeit)", "Net profit ± working capital changes"),
        ("Investing activities (Investitionstätigkeit)",  "Purchase/sale of assets"),
        ("Financing activities (Finanzierungstätigkeit)", "Loans received/repaid, dividends"),
        ("Net change in cash",                            "Sum of the three sections"),
        ("Opening cash balance",                          "Cash at start of year"),
        ("Closing cash balance",                          "Cash at end of year"),
    ])

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  11. GERMAN TAX LAW REFERENCE
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "11. German Tax Law Reference",
            "11. Deutsches Steuerrecht – Referenz")

    info_table(doc, [
        ("§ 1 EStG",   "Unlimited vs. limited tax liability — Unbeschränkte Steuerpflicht"),
        ("§ 2 EStG",   "Taxable income definition — Einkunftsarten"),
        ("§ 9 EStG",   "Work-related expenses — Werbungskosten"),
        ("§ 10 EStG",  "Special expenses — Sonderausgaben"),
        ("§ 13 EStG",  "Agricultural income"),
        ("§ 15 EStG",  "Business income — Gewerbebetrieb"),
        ("§ 18 EStG",  "Freelance income — Selbstständige Arbeit"),
        ("§ 19 EStG",  "Employment income — Nichtselbstständige Arbeit"),
        ("§ 20 EStG",  "Capital income — Kapitalvermögen"),
        ("§ 21 EStG",  "Rental income — Vermietung & Verpachtung"),
        ("§ 22 EStG",  "Other income — Sonstige Einkünfte"),
        ("§ 32a EStG", "Income tax tariff — Einkommensteuertarif 2024"),
        ("§ 32d EStG", "Flat tax on capital gains — Abgeltungssteuer 25 %"),
        ("§ 33 EStG",  "Extraordinary burdens — Außergewöhnliche Belastungen"),
        ("§ 10a EStG", "Riester pension deduction — Altersvorsorgezulage"),
        ("SolZG",      "Solidarity Surcharge Act — Solidaritätszuschlagsgesetz"),
        ("KiStG",      "Church Tax Acts (state level) — Kirchensteuergesetze"),
        ("GewStG",     "Trade Tax Act — Gewerbesteuergesetz"),
        ("UStG",       "VAT Act — Umsatzsteuergesetz"),
        ("AO",         "General Tax Code — Abgabenordnung"),
        ("DSGVO/GDPR", "Data Protection — Datenschutz"),
    ])

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  12. STEP-BY-STEP DEVELOPMENT ROADMAP
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "12. Step-by-Step Development Roadmap",
            "12. Schritt-für-Schritt Entwicklungsplan")

    phases = [
        ("Phase 1 — Foundation (Week 1–2)", "Phase 1 — Grundlagen", [
            ("Install Python 3.11, pip, virtual environment (venv)",
             "Python installieren, virtuelle Umgebung erstellen"),
            ("pip install flask sqlalchemy openpyxl reportlab flask-babel",
             "Abhängigkeiten installieren"),
            ("Create folder structure as shown in Section 6",
             "Ordnerstruktur gemäß Abschnitt 6 erstellen"),
            ("Set up SQLite DB with SQLAlchemy models (models.py)",
             "SQLite-Datenbank mit SQLAlchemy-Modellen einrichten"),
            ("Create app.py with basic Flask server, test http://localhost:5000",
             "Flask-Server testen"),
        ]),
        ("Phase 2 — Personal Info & Income (Week 3–4)", "Phase 2 — Persönl. Daten & Einkommen", [
            ("Build personal info form (personal.html + personal_info.py)",
             "Persönliche-Daten-Formular erstellen"),
            ("Build income entry forms for all 7 income types (income.html)",
             "Einkommensformulare für alle 7 Einkunftsarten erstellen"),
            ("Wire up Flask routes: /api/personal, /api/income",
             "Flask-Routen verdrahten"),
            ("Store data in SQLite via SQLAlchemy",
             "Daten in SQLite speichern"),
            ("Add basic validation (validators.py)",
             "Grundlegende Validierung hinzufügen"),
        ]),
        ("Phase 3 — Deductions & Tax Engine (Week 5–6)", "Phase 3 — Abzüge & Steuer-Engine", [
            ("Build all deduction input forms (deductions.html)",
             "Alle Abzugsformulare erstellen"),
            ("Implement § 32a EStG tax tariff algorithm in tax_engine.py",
             "§ 32a EStG Steuertarif in tax_engine.py implementieren"),
            ("Add Soli, church tax, trade tax calculations",
             "Soli, Kirchensteuer, Gewerbesteuer berechnen"),
            ("Build tax result display page (tax_result.html)",
             "Steuerergebnis-Seite erstellen"),
            ("Write unit tests for all tax engine functions",
             "Unit-Tests für alle Steuerberechnungen schreiben"),
        ]),
        ("Phase 4 — Financial Statements (Week 7–8)", "Phase 4 — Finanzberichte", [
            ("Implement P&L statement builder (financial_statements.py)",
             "GuV-Ersteller implementieren"),
            ("Implement Balance Sheet builder",
             "Bilanz-Ersteller implementieren"),
            ("Implement Cash Flow Statement builder",
             "Kapitalflussrechnung-Ersteller implementieren"),
            ("Build statements display page (statements.html)",
             "Finanzberichte-Seite erstellen"),
            ("Add year-over-year comparison logic",
             "Jahresvergleich-Logik hinzufügen"),
        ]),
        ("Phase 5 — Excel & PDF Export (Week 9–10)", "Phase 5 — Export", [
            ("Build 12-sheet Excel workbook with openpyxl (excel_export.py)",
             "12-Blatt-Excel-Arbeitsmappe mit openpyxl erstellen"),
            ("Apply professional formatting: colours, borders, number formats",
             "Professionelle Formatierung anwenden"),
            ("Embed charts in Excel (bar & pie)",
             "Diagramme in Excel einbetten"),
            ("Build PDF report with reportlab (pdf_export.py)",
             "PDF-Bericht mit reportlab erstellen"),
            ("Add export page and download endpoints (export.py)",
             "Export-Seite und Download-Endpunkte hinzufügen"),
        ]),
        ("Phase 6 — Bilingual UI & Dashboard (Week 11–12)", "Phase 6 — UI & Dashboard", [
            ("Create translations/en/messages.json & translations/de/messages.json",
             "Übersetzungsdateien erstellen"),
            ("Integrate i18next for instant language toggle",
             "i18next für sofortigen Sprachwechsel integrieren"),
            ("Integrate Flask-Babel for server-side translations",
             "Flask-Babel für serverseitige Übersetzungen integrieren"),
            ("Build dashboard with Chart.js (income pie, tax bar, cash flow line)",
             "Dashboard mit Chart.js erstellen"),
            ("Apply Bootstrap 5 responsive layout & custom dark theme",
             "Bootstrap 5 Layout & dunkles Theme anwenden"),
        ]),
        ("Phase 7 — Polish & Packaging (Week 13–14)", "Phase 7 — Finish & Paketierung", [
            ("Full end-to-end user workflow testing",
             "Vollständige End-to-End-Tests"),
            ("Add multi-year selection (dropdown for tax year)",
             "Mehrjährige Auswahl hinzufügen"),
            ("Create run.bat (Windows) and run.sh (Linux/Mac) launcher scripts",
             "Start-Skripte erstellen"),
            ("Optional: package with PyInstaller as standalone .exe",
             "Optional: mit PyInstaller als EXE paketieren"),
            ("Write user documentation (README.md + in-app help)",
             "Benutzerdokumentation schreiben"),
        ]),
    ]

    for phase_en, phase_de, tasks in phases:
        heading(doc, phase_en, phase_de, level=2)
        for task_en, task_de in tasks:
            bilingual_bullet(doc, task_en, task_de, bullet="✔")
        section_divider(doc)

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  13. RUNNING & DEPLOYMENT
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "13. Running & Deployment Guide",
            "13. Ausführungs- & Bereitstellungsanleitung")

    steps_run = [
        ("1. Install Python 3.11+",
         "Download from python.org — add to PATH during install"),
        ("2. Open terminal in project folder",
         "cd german_tax_tool"),
        ("3. Create virtual environment",
         "python -m venv venv"),
        ("4. Activate virtual environment",
         "Windows: venv\\Scripts\\activate  |  Mac/Linux: source venv/bin/activate"),
        ("5. Install dependencies",
         "pip install -r requirements.txt"),
        ("6. Initialise database",
         "python -c \"from database.db import init_db; init_db()\""),
        ("7. Start the server",
         "python app.py   (or double-click run.bat / run.sh)"),
        ("8. Open in browser",
         "Navigate to: http://localhost:5000"),
        ("9. Begin tax return",
         "Follow the wizard: Personal → Income → Deductions → Calculate → Export"),
    ]
    for step, detail in steps_run:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        add_run(p, f"{step}: ", bold=True, color=GOLD, size=11)
        add_run(p, detail, size=11)

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  14. USER WORKFLOW
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "14. User Workflow Walkthrough",
            "14. Benutzer-Workflow")

    workflow = [
        ("Start", "Open http://localhost:5000 → Dashboard shows current tax year"),
        ("Step 1 — Personal Info",     "Enter name, TIN, DOB, tax class, children, spouse, bank details"),
        ("Step 2 — Income",            "Select income type → enter amounts → system maps to correct Anlage"),
        ("Step 3 — Deductions",        "Enter each deduction category; tooltip explains limits & rules"),
        ("Step 4 — Calculate",         "Click 'Calculate Tax' → instant result: refund 🟢 or payment 🔴"),
        ("Step 5 — Review Statements", "View P&L, Balance Sheet, Cash Flow on the Statements page"),
        ("Step 6 — Export Excel",      "Click 'Export Excel' → download 12-sheet workbook"),
        ("Step 7 — Export PDF",        "Click 'Export PDF' → download print-ready report"),
        ("Step 8 — Switch Language",   "Click 🇩🇪/🇬🇧 flag → entire UI instantly switches language"),
        ("Step 9 — Save & Archive",    "Data auto-saved in SQLite; access previous years from dropdown"),
    ]
    for step, desc in workflow:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_after = Pt(3)
        add_run(p, f"  {step}  ", bold=True, color=WHITE, size=10)
        # fake badge colour via text
        add_run(p, f"  {desc}", size=11)

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  15. SECURITY & DATA PRIVACY
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "15. Security & Data Privacy",
            "15. Sicherheit & Datenschutz")

    privacy = [
        ("All data stored locally — SQLite file on your machine only",
         "Alle Daten lokal gespeichert – SQLite nur auf Ihrem Computer"),
        ("No internet connection required or used",
         "Keine Internetverbindung erforderlich oder verwendet"),
        ("Optional: encrypt SQLite with SQLCipher",
         "Optional: SQLite mit SQLCipher verschlüsseln"),
        ("File permissions: restrict exports/ folder to current user",
         "Dateiberechtigungen: Exportordner nur für aktuellen Benutzer"),
        ("Password-protect the app with Flask login (optional)",
         "App mit Flask-Login passwortschützen (optional)"),
        ("Never log sensitive data (TIN, IBAN) to console",
         "Niemals sensible Daten (TIN, IBAN) in Konsole loggen"),
        ("Backup reminder: copy SQLite file to encrypted USB periodically",
         "Backup-Erinnerung: SQLite-Datei regelmäßig auf verschl. USB kopieren"),
        ("GDPR-compliant: data stays with user; no third-party transmission",
         "DSGVO-konform: Daten verbleiben beim Nutzer; keine Drittübermittlung"),
    ]
    for en, de in privacy:
        bilingual_bullet(doc, en, de)

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════════
    #  16. GLOSSARY
    # ══════════════════════════════════════════════════════════════════════════
    heading(doc, "16. Glossary",
            "16. Glossar")

    glossary = [
        ("AO",                    "Abgabenordnung — General Tax Code"),
        ("Anlage",                "Supplementary form attached to the main tax return"),
        ("Arbeitnehmer-Pauschbetrag","Employee lump-sum deduction (€1,230 in 2024)"),
        ("Außergewöhnl. Belastungen","Extraordinary burdens — unusually high expenses"),
        ("ELSTER",               "Electronic tax return portal (elster.de)"),
        ("EStG",                  "Einkommensteuergesetz — Income Tax Act"),
        ("EÜR",                   "Einnahmen-Überschuss-Rechnung — cash-basis P&L for freelancers"),
        ("Finanzamt",             "Tax office (local authority)"),
        ("GewSt",                 "Gewerbesteuer — Trade Tax"),
        ("Grundfreibetrag",       "Basic personal allowance (tax-free threshold)"),
        ("GuV",                   "Gewinn- und Verlustrechnung — Profit & Loss Statement"),
        ("KiSt",                  "Kirchensteuer — Church Tax"),
        ("Lohnsteuerbescheinigung","Annual wage tax certificate issued by employer"),
        ("Mantelbogen",           "Main form of the income tax return (ESt 1 A)"),
        ("Sonderausgaben",        "Special expenses — specific deductible items"),
        ("Solidaritätszuschlag",  "Solidarity surcharge on income tax"),
        ("Steuerbescheid",        "Tax assessment notice from the Finanzamt"),
        ("Steuerklasse",          "Tax class (I – VI) affecting withholding"),
        ("Steueridentifikationsnummer","11-digit permanent tax ID number"),
        ("UStVA",                 "Umsatzsteuervoranmeldung — VAT pre-declaration"),
        ("Werbungskosten",        "Work-related expenses / income-related deductions"),
        ("zvE",                   "zu versteuerndes Einkommen — taxable income"),
    ]
    info_table(doc, [(term, defn) for term, defn in glossary], col_widths=(2.4, 4.0))

    doc.add_paragraph()
    section_divider(doc)

    # footer note
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(p,
        "\n★  German Tax Tool — Complete Plan  ★\n"
        "Generated automatically by gen_german_tax_tool.py\n"
        f"Date: {datetime.date.today().strftime('%B %d, %Y')}",
        italic=True, size=9, color=SILVER)

    return doc


# ── entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    out = "German_Tax_Tool_Plan.docx"
    doc = build_document()
    doc.save(out)
    print(f"✔  Saved: {out}")
