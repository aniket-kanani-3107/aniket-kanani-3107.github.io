# -*- coding: utf-8 -*-
"""Generate German_Learning_Book_C1_Gujarati_Days_271_330.docx (Days 271-330)"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BLUE = RGBColor(0x1A, 0x73, 0xE8)
YELLOW = RGBColor(0xF9, 0xAB, 0x00)
GRAY = RGBColor(0x55, 0x55, 0x55)
GREEN = RGBColor(0x0F, 0x96, 0x60)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x20, 0x20, 0x20)
PURPLE = RGBColor(0x6A, 0x0D, 0xAD)
ORANGE = RGBColor(0xE6, 0x5C, 0x00)


def add_para(doc, text, size_pt=11, bold=False, color=None, align=None, space_before=0, space_after=4):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size_pt)
    if color:
        run.font.color.rgb = color
    return p


def add_bullet(doc, text, size_pt=11):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.size = Pt(size_pt)
    return p


def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def cell_para(cell, text, bold=False, size_pt=10, color=None, align=None):
    p = cell.paragraphs[0]
    if align:
        p.alignment = align
    p.clear()
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size_pt)
    if color:
        run.font.color.rgb = color
    return p


def add_vocab_table(doc, rows):
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    headers = ["🇩🇪 German", "🇬🇧 English", "ગુજરાતી", "🔊 Sound like…"]
    for i, h in enumerate(headers):
        set_cell_bg(hdr[i], "1A73E8")
        cell_para(hdr[i], h, bold=True, size_pt=10, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
    for row_data in rows:
        row = table.add_row().cells
        colors = [DARK, DARK, DARK, GRAY]
        for i, (val, clr) in enumerate(zip(row_data, colors)):
            cell_para(row[i], val, size_pt=10, color=clr)
    doc.add_paragraph()


def add_sentence_table(doc, de, en, gu):
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    cells = table.rows[0].cells
    set_cell_bg(cells[0], "E8F0FE")
    set_cell_bg(cells[1], "E6F4EA")
    set_cell_bg(cells[2], "FFF8E1")
    cell_para(cells[0], de, size_pt=10, color=BLUE)
    cell_para(cells[1], en, size_pt=10, color=GREEN)
    cell_para(cells[2], gu, size_pt=10, color=DARK)
    doc.add_paragraph()


def add_tip_box(doc, emoji, label, text, bg="FFF3CD"):
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.rows[0].cells[0]
    set_cell_bg(cell, bg)
    cell.paragraphs[0].clear()
    run1 = cell.paragraphs[0].add_run(f"{emoji} {label}  ")
    run1.bold = True
    run1.font.size = Pt(10)
    run1.font.color.rgb = DARK
    run2 = cell.paragraphs[0].add_run(text)
    run2.font.size = Pt(10)
    run2.font.color.rgb = DARK
    doc.add_paragraph()


def add_day(doc, day_num, title, topic, eng_exp, guj_exp, vocab, sentences, tips, tasks, mini_test, outcome):
    doc.add_page_break()
    add_para(doc, f"📘  DAY {day_num}", size_pt=21, bold=True, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=2)
    add_para(doc, title, size_pt=14, bold=True, color=YELLOW, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_para(doc, f"📚 Topic: {topic}", size_pt=10.5, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    doc.add_paragraph()

    add_para(doc, "📖  English Explanation", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    add_para(doc, eng_exp, size_pt=10.5, space_after=6)

    add_para(doc, "🇮🇳  ગુજરાતી સ્પષ્ટીકરણ", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    add_para(doc, guj_exp, size_pt=10.5, space_after=6)

    add_para(doc, "📋  Vocabulary List", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    add_vocab_table(doc, vocab)

    add_para(doc, "💬  Sentence Examples — Real Life", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    for de, en, gu in sentences:
        add_sentence_table(doc, de, en, gu)

    for emoji, label, text, bg in tips:
        add_tip_box(doc, emoji, label, text, bg)

    add_para(doc, "✍️  Practice Tasks", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    for task in tasks:
        add_bullet(doc, task)
    doc.add_paragraph()

    add_para(doc, "🧪  Mini Test", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    for i, (q, a) in enumerate(mini_test, 1):
        add_para(doc, f"Q{i}. {q}", size_pt=10.5, space_after=1)
        add_para(doc, f"   ✅ {a}", size_pt=10.5, color=GREEN, space_after=3)
    doc.add_paragraph()

    add_para(doc, "🏆  Final Outcome", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    add_para(doc, outcome, size_pt=10.5, space_after=8)


MODULES = [
    {
        "range": (271, 280),
        "name": "Rhetoric and Advanced Discourse",
        "focus": "rhetorical precision, argument layering, and stylistic impact",
        "vocab": [
            ("die Rhetorik", "rhetoric", "વક્તૃત્વશાસ્ત્ર", "reh-TOH-rik"),
            ("die Zuspitzung", "sharpening", "તીક્ષ્ણતા", "TSOO-shpit-soong"),
            ("der Subtext", "subtext", "આંતરિક અર્થ", "SUB-text"),
            ("die Prämisse", "premise", "મૂળ ધારણા", "PRAY-mis-seh"),
            ("die Folgerichtigkeit", "logical consistency", "તર્કસુસંગતતા", "FOL-ge-rikh-kite"),
            ("die Gegenrede", "counter-speech", "પ્રતિઉક્તિ", "GAY-gen-ray-deh"),
            ("die Akzentuierung", "emphasis", "ઉચ્ચાર ભાર આપવું", "ak-tsen-too-EE-roong"),
            ("die Einordnung", "contextualization", "સંદર્ભીકરણ", "EYN-ord-noong"),
            ("die Konnotation", "connotation", "અંતરઅર્થ", "kon-no-ta-TSY-ohn"),
            ("die Plausibilität", "plausibility", "વિશ્વસનીયતા", "plow-zi-bi-LEE-tet"),
            ("die Verknüpfung", "linkage", "કડી", "fer-KNÜP-foong"),
            ("die Differenzierung", "differentiation", "વિવેચન", "dif-fe-ren-TSY-roong"),
            ("die Übertreibung", "exaggeration", "અતિશયોક્તિ", "ü-ber-TRY-boong"),
            ("die Ironisierung", "ironic framing", "વ્યંગ્યાત્મક રજૂઆત", "ee-ro-nee-ZEE-roong"),
            ("die Antithese", "antithesis", "વિરોધાભાસ", "an-ti-TAY-zeh"),
            ("die Nuancierung", "nuancing", "સૂક્ષ્મતા", "nu-an-SEE-roong"),
            ("die Schlagkraft", "impact", "પ્રભાવશક્તિ", "SHLAHG-kraft"),
            ("die Aussagekraft", "expressive power", "અભિવ્યક્તિ શક્તિ", "OWS-sah-ge-kraft"),
            ("die Dramaturgie", "dramaturgy", "રચનાત્મક નાટ્યગતિ", "dra-ma-toor-GEE"),
            ("die Verdichtung", "condensation", "સારાંશીકરણ", "fer-DIKH-toong"),
        ],
    },
    {
        "range": (281, 290),
        "name": "Academic Research and Critical Reading",
        "focus": "methodology control, evidence evaluation, and synthesis writing",
        "vocab": [
            ("die Methodik", "methodology", "પદ્ધતિશાસ્ત્ર", "meh-TOH-dik"),
            ("die Herleitung", "derivation", "તર્કપૂર્વક ઉત્પન્ન કરવું", "HAIR-ly-toong"),
            ("die Hypothese", "hypothesis", "ધારણા", "hy-po-TAY-zeh"),
            ("die Stichprobe", "sample", "નમૂના", "SHTIKH-pro-beh"),
            ("die Verzerrung", "distortion", "વાંકડુંપણું", "fer-TSEHR-roong"),
            ("die Evidenzlage", "evidence base", "પુરાવાની સ્થિતિ", "eh-vee-DENTS-la-geh"),
            ("die Replikation", "replication", "પુનરાવર્તન", "reh-pli-ka-TSY-ohn"),
            ("die Querverbindung", "cross-link", "આડું જોડાણ", "KVAIR-fer-bin-doong"),
            ("die Schlussfolgerung", "conclusion", "નિષ્કર્ષ", "SHLOOS-fol-ge-roong"),
            ("die Kausalität", "causality", "કારણકાર્ય", "kow-za-li-TET"),
            ("die Korrelation", "correlation", "સહસંબંધ", "ko-re-la-TSY-ohn"),
            ("die Limitation", "limitation", "મર્યાદા", "li-mi-ta-TSY-ohn"),
            ("die Quellenkritik", "source critique", "સ્રોતવિચાર", "KVEL-len-kri-tik"),
            ("die Auslassung", "omission", "કાપ", "OWS-las-soong"),
            ("die Literaturanalyse", "literature analysis", "સાહિત્ય વિશ્લેષણ", "li-te-ra-TOOR-a-na-ly-zeh"),
            ("die Synthese", "synthesis", "સમન્વય", "SIN-teh-zeh"),
            ("das Abstract", "abstract", "સારાંશ", "AP-strakt"),
            ("die Datenbasis", "data base", "ડેટા આધાર", "DA-ten-ba-zis"),
            ("die Validität", "validity", "માન્યતા", "va-li-di-TET"),
            ("die Reliabilität", "reliability", "વિશ્વસનીયતા", "re-li-a-bi-li-TET"),
        ],
    },
    {
        "range": (291, 300),
        "name": "Professional Leadership and Strategy",
        "focus": "strategic communication, executive decisions, and stakeholder alignment",
        "vocab": [
            ("die Strategie", "strategy", "યોજનાત્મક દૃષ્ટિ", "shtra-TEE-gee"),
            ("die Ausrichtung", "direction/alignment", "દિશા નિર્ધારણ", "OWS-rikh-toong"),
            ("die Steuerung", "steering", "નિયંત્રણ", "SHTOY-roong"),
            ("die Skalierung", "scaling", "વિસ્તરણ", "ska-LEE-roong"),
            ("die Wertschöpfung", "value creation", "મૂળ્ય સર્જન", "VERT-sher-foong"),
            ("die Risikoabschätzung", "risk assessment", "જોખમ મૂલ્યાંકન", "REE-zi-ko-AP-she-tsung"),
            ("die Entscheidungsgrundlage", "decision basis", "નિર્ણય આધાર", "ent-SHYE-doongs-groond-la-geh"),
            ("die Ressourcenallokation", "resource allocation", "સંસાધન વહેંચણી", "re-ZOR-sen-a-lo-ka-TSY-ohn"),
            ("die Umsetzungskraft", "execution strength", "અમલ કરવાની શક્તિ", "OOM-zet-soong-kraft"),
            ("die Verantwortungsübernahme", "taking responsibility", "જવાબદારી સ્વીકાર", "fer-ANT-vor-toongs-ü-ber-na-me"),
            ("die Erwartungshaltung", "expectation", "અપેક્ષા રાખવાની ભાવના", "er-VAR-toongs-hal-toong"),
            ("die Delegation", "delegation", "સોંપણી", "de-le-ga-TSY-ohn"),
            ("die Eskalationsstufe", "escalation level", "તણાવ સ્તર", "es-ka-la-TSY-ohns-shtoo-feh"),
            ("die Leistungskennzahl", "performance metric", "પ્રદર્શન માપદંડ", "LY-stoongs-ken-tsahl"),
            ("die Abstimmungsschleife", "alignment loop", "સુમેળ ચક્ર", "AP-shtim-moongs-shly-feh"),
            ("die Handlungsoption", "course of action", "ક્રિયા વિકલ્પ", "HAND-loongs-op-TSY-ohn"),
            ("die Interessenslage", "interest position", "હિત સ્થિતિ", "in-te-RESS-ens-la-geh"),
            ("der Entscheidungsspielraum", "decision leeway", "નિર્ણય લવચીકતા", "ent-SHYE-doongs-shpeel-raum"),
            ("die Prioritätsmatrix", "priority matrix", "પ્રાથમિકતા મેટે્રક્સ", "pri-o-ri-TETS-ma-triks"),
            ("die Wirkungskette", "impact chain", "અસર શ્રેણી", "VIR-koongs-ket-te"),
        ],
    },
    {
        "range": (301, 310),
        "name": "Society, Ethics, and Policy Analysis",
        "focus": "ethical reasoning, policy trade-offs, and societal impact language",
        "vocab": [
            ("die Gerechtigkeitsfrage", "justice question", "ન્યાય પ્રશ્ન", "ge-REKH-tik-kites-fra-geh"),
            ("der Zielkonflikt", "goal conflict", "લક્ષ્ય ટકરાવ", "TSEEL-kon-flikt"),
            ("die Legitimation", "legitimation", "માન્યતા", "le-gee-ti-ma-TSY-ohn"),
            ("die Gemeinwohlorientierung", "common good orientation", "સામૂહિક હિત દિશા", "ge-MYN-vol-o-ri-en-TEE-roong"),
            ("die Rahmensetzung", "framework setting", "ઢાંચા નિર્ધારણ", "RAH-men-zet-soong"),
            ("die Folgenabschätzung", "impact assessment", "પરિણામ મૂલ્યાંકન", "FOL-gen-AP-she-tsung"),
            ("die Interessenvermittlung", "interest mediation", "હિત સમન્વય", "in-te-RESS-en-fer-mit-loong"),
            ("die Verpflichtungsethik", "ethics of duty", "કર્તવ્ય નૈતિકતા", "fer-PFLIKH-toongs-eh-tik"),
            ("die Verantwortungsethik", "ethics of responsibility", "જવાબદારી નૈતિકતા", "fer-ANT-vor-toongs-eh-tik"),
            ("die Verhältnismäßigkeit", "proportionality", "મર્યાદિત અનુપાતતા", "fer-HELTS-nis-may-sig-kite"),
            ("die Interessengruppe", "interest group", "હિતગ્રુપ", "in-te-RESS-en-groo-pe"),
            ("die Zustimmungsrate", "approval rate", "સમર્થન દર", "TSOO-shtim-moongs-ra-te"),
            ("die Polarisierung", "polarization", "ધ્રુવીકરણ", "po-la-ri-ZEE-roong"),
            ("die Durchsetzbarkeit", "enforceability", "અમલક્ષમતા", "DOORKH-zetz-bar-kite"),
            ("die Schutzklausel", "safeguard clause", "સુરક્ષા કલમ", "SHOOTS-klow-zel"),
            ("die Wirksamkeit", "effectiveness", "કાર્યક્ષમતા", "VIRK-zam-kite"),
            ("die Risikoethik", "risk ethics", "જોખમ નૈતિકતા", "REE-zi-ko-eh-tik"),
            ("die Gegenmaßnahme", "countermeasure", "પ્રતિ પગલું", "GAY-gen-mas-na-me"),
            ("die Langzeitwirkung", "long-term effect", "દીર્ઘકાળીન અસર", "LANG-tsyte-vir-koong"),
            ("die Reformagenda", "reform agenda", "સુધારા એજન્ડા", "re-FORM-a-gen-da"),
        ],
    },
    {
        "range": (311, 320),
        "name": "C1 Exam Mastery and High-Impact Writing",
        "focus": "text type control, advanced structure, and evaluation readiness",
        "vocab": [
            ("die Textsortenkompetenz", "text type competence", "લેખ પ્રકાર કુશળતા", "TEKST-zor-ten-kom-pe-TENTS"),
            ("die Aufgabenanalyse", "task analysis", "પ્રશ્ન વિશ્લેષણ", "OWF-ga-ben-a-na-ly-zeh"),
            ("die Bewertungsrubrik", "grading rubric", "મૂલ્યાંકન માપદંડ", "beh-VER-toongs-roo-brik"),
            ("die Kohäsion", "cohesion", "એકતા", "ko-hay-ZY-ohn"),
            ("die Argumentationslinie", "argument line", "દલીલ રેખા", "ar-gu-men-ta-TSY-ohns-LEE-nee"),
            ("die Stilvariation", "style variation", "શૈલી ફેરફાર", "SHTIL-fa-ri-a-TSY-ohn"),
            ("die Überarbeitungsschleife", "revision loop", "પુનઃસંપાદન ચક્ર", "ü-ber-AR-by-toongs-shly-feh"),
            ("die Ausdrucksgenauigkeit", "expressive precision", "અભિવ્યક્તિ ચોકસાઈ", "OWS-drooks-ge-now-ig-kite"),
            ("die Fehleranalyse", "error analysis", "ભૂલ વિશ્લેષણ", "FAY-ler-a-na-ly-zeh"),
            ("die Zeitreserve", "time buffer", "સમય બચત", "TSYTE-re-zer-ve"),
            ("die Schreibökonomie", "writing economy", "લેખન કાર્યક્ષમતા", "SHRYB-ö-ko-no-mee"),
            ("die Präsentationsstruktur", "presentation structure", "પ્રસ્તુતિ બંધારણ", "pray-zen-ta-TSY-ohns-shtruk-toor"),
            ("die Gesprächsführung", "conversation leadership", "વાતચીત માર્ગદર્શન", "ge-SHPREKHS-fü-roong"),
            ("die Replik", "reply/rebuttal", "પ્રતિઉત્તર", "reh-PLIK"),
            ("die Präzisierung", "specification", "સ્પષ્ટીકરણ", "pray-tsi-ZEE-roong"),
            ("die Spannungsführung", "tension management", "રૂચિ જાળવવી", "SHPAN-noongs-fü-roong"),
            ("die Leserführung", "reader guidance", "પાઠક માર્ગદર્શન", "LAY-zer-fü-roong"),
            ("die Gliederungstiefe", "outline depth", "રચનાની ઊંડાણ", "GLI-de-roongs-TEE-feh"),
            ("die Argumentationsschärfe", "argument sharpness", "દલીલ તીવ્રતા", "ar-gu-men-ta-TSY-ohs-sher-fe"),
            ("die Aufgabenpriorisierung", "task prioritization", "પ્રશ્ન પ્રાથમિકતા", "OWF-ga-ben-pri-o-ri-SEE-roong"),
        ],
    },
    {
        "range": (321, 330),
        "name": "C1 Fluency, Style, and C2 Bridge",
        "focus": "near-native fluency, register switching, and stylistic sophistication",
        "vocab": [
            ("die Sprachsouveränität", "language sovereignty", "ભાષા આત્મવિશ્વાસ", "SHPRAKH-zoo-ve-re-nee-TET"),
            ("der Registerwechsel", "register shift", "રજિસ્ટર બદલાવ", "reh-GIS-ter-vek-sel"),
            ("die Idiomatik", "idiomatics", "રૂઢિપ્રયોગ કુશળતા", "ee-dee-oh-MA-tik"),
            ("die Feindifferenzierung", "fine differentiation", "સૂક્ષ્મ ભેદ", "FYN-dif-fe-ren-TSY-roong"),
            ("die Wortschatzdichte", "lexical density", "શબ્દભંડાર ઘનતા", "VORT-shats-dikh-te"),
            ("die Anspielung", "allusion", "ઇશારો", "AN-shpee-loong"),
            ("die Bildlichkeit", "imagery", "છબીમયતા", "BILD-likh-kite"),
            ("die Metapher", "metaphor", "રૂપક", "meh-TA-fer"),
            ("die Klangfarbe", "tone color", "ધ્વનિ રંગ", "KLANG-far-be"),
            ("die Stilbrillanz", "stylistic brilliance", "શૈલી પ્રભા", "SHTIL-bri-LANTS"),
            ("die Anschlussfähigkeit", "connectability", "સંદર્ભ જોડાણક્ષમતા", "AN-shloos-fee-ikh-kite"),
            ("die Sprachökonomie", "language economy", "ભાષા કાર્યક્ષમતા", "SHPRAKH-ö-ko-no-mee"),
            ("die Verdichtung", "condensation", "સારાંશીકરણ", "fer-DIKH-toong"),
            ("die Präzisionsarbeit", "precision work", "ચોકસાઇ કાર્ય", "pray-tsi-ZY-ohns-ar-bite"),
            ("die Selbstkorrektur", "self-correction", "સ્વ-સુધારો", "ZELBST-kor-rek-toor"),
            ("die Anschlussrede", "follow-up speech", "આગળનું ભાષણ", "AN-shloos-ray-deh"),
            ("die Sprachintelligenz", "language intelligence", "ભાષા બુદ્ધિ", "SHPRAKH-in-te-li-GENTS"),
            ("die Ausdrucksvielfalt", "expressive variety", "અભિવ્યક્તિ વિવિધતા", "OWS-drooks-fee-falt"),
            ("die Prägnanzwirkung", "concise impact", "સંક્ષિપ્ત અસર", "PRAYG-nants-vir-koong"),
            ("die Sprechhaltung", "speaking stance", "બોલવાની સ્થિતિ", "SHPREKH-hal-toong"),
        ],
    },
]

SUBTOPICS = [
    "Rhetorical clarity", "Argument layering", "Subtext and implication", "Premise testing", "Logical consistency checks", "Counter-speech tactics", "Emphasis and cadence", "Context framing", "Connotation control", "Impactful endings",
    "Research question focus", "Hypothesis evaluation", "Sample bias analysis", "Evidence grading", "Cross-study comparison", "Causality vs correlation", "Limitation mapping", "Source critique workshop", "Synthesis paragraph craft", "Academic mini-review",
    "Strategic vision talk", "Decision pathway language", "Risk communication", "Stakeholder alignment", "Priority matrix usage", "Execution narrative", "Delegation clarity", "Metric-driven reporting", "Escalation protocol", "Leadership simulation",
    "Ethical dilemma framing", "Policy trade-off language", "Legitimation strategies", "Common-good arguments", "Framework setting", "Impact assessment", "Mediation and consensus", "Proportionality debate", "Polarization repair", "Reform roadmap",
    "Text type selection", "Task analysis under pressure", "C1 cohesion practice", "Argument line development", "Style variation drills", "Revision loop mastery", "Precision and density", "Reader guidance", "Speaking exam lead-in", "Full C1 mock set A",
    "Register switching", "Idiom density", "Imagery and metaphor", "Allusion handling", "Tone color control", "Stylistic brilliance", "Language economy", "Self-correction under speed", "C1 performance polish", "C1 graduation and C2 bridge",
]


def pick_vocab(bank, start, size=15):
    if size > len(bank):
        raise ValueError(f"Vocabulary size {size} cannot exceed vocabulary bank length {len(bank)}.")
    out = []
    for i in range(size):
        out.append(bank[(start + i) % len(bank)])
    return out


def build_day_entry(day_num, module, subtopic, idx):
    vocab = pick_vocab(module["vocab"], idx)
    k1, k2, k3 = vocab[0][0], vocab[1][0], vocab[2][0]

    title = f"{subtopic.upper()} 🚀"
    topic = f"{module['name']} — {subtopic}"

    eng_exp = (
        f"Day {day_num} advances your C1 German with a focused lesson on '{subtopic}'. "
        f"Today you develop {module['focus']}. The goal is to speak and write with near-native control, "
        f"strong logic, and refined style across academic, professional, and public contexts."
    )
    guj_exp = (
        f"Day {day_num} માં તમે C1 સ્તરે '{subtopic}' પર focused અભ્યાસ કરો છો. આજે તમારું ધ્યાન "
        f"{module['focus']} પર છે. હેતુ એ છે કે તમે academic, professional અને public contexts માં "
        f"near-native નિયંત્રણ, મજબૂત તર્ક અને પરિષ્કૃત શૈલી સાથે લખી અને બોલી શકો."
    )

    sentences = [
        (
            f"Im heutigen Training steht {k1} im Zentrum.",
            f"In today's training, {k1} is central.",
            f"આજના અભ્યાસમાં {k1} મુખ્ય કેન્દ્ર છે."
        ),
        (
            f"Ich nutze {k2}, um die Argumentation präziser zu steuern.",
            f"I use {k2} to steer the argumentation more precisely.",
            f"હું {k2} નો ઉપયોગ કરીને દલીલને વધુ ચોક્કસ રીતે દિશા આપું છું."
        ),
        (
            f"Mit {k3} gewinne ich in der Diskussion an Tiefe.",
            f"With {k3}, I gain depth in the discussion.",
            f"{k3} થી ચર્ચામાં ઊંડાણ આવે છે."
        ),
        (
            "Je feiner die Struktur, desto überzeugender wirkt die Aussage.",
            "The finer the structure, the more convincing the statement becomes.",
            "રચના જેટલી સૂક્ષ્મ હોય, નિવેદન તેટલું જ મનાવનાર લાગે છે."
        ),
        (
            "Ich optimiere Wortwahl und Register, bevor ich den Text abgebe.",
            "I optimize word choice and register before submitting the text.",
            "હું લખાણ સબમિટ કરતાં પહેલા શબ્દચયન અને રજિસ્ટર સુધારો છું."
        ),
    ]

    tips = [
        ("📌", "C1 Focus", f"Aim for density + clarity. Integrate {k1} and {k2} in complex sentences.", "E8F0FE"),
        ("💡", "Strategy", "Draft fast, then polish for cohesion, register shifts, and rhetorical impact.", "FFF8E1"),
        ("🎯", "Progress", "C1 growth comes from precision edits. Track recurring patterns and fix them weekly.", "E6F4EA"),
    ]

    tasks = [
        f"Write a 220–250 word analysis on '{subtopic}' using at least 7 target words.",
        "Record a 3-minute spoken response with a clear stance and structured reasoning.",
        "Rewrite your text for a different audience (formal report vs public talk) and compare tone shifts.",
    ]

    mini_test = [
        ("Translate into German: 'The argument gains depth through precise framing.'", "Das Argument gewinnt durch präzise Rahmung an Tiefe."),
        ("Which is more C1-appropriate: vague wording or differentiated wording?", "Differentiated wording"),
        ("Provide one key word from today's vocabulary.", vocab[0][0]),
        ("At C1, what should be balanced: fluency alone or fluency + accuracy + style?", "Fluency + accuracy + style"),
        ("Name one improvement step before submission.", "Refine cohesion, register, and word choice"),
    ]

    outcome = (
        f"After Day {day_num}, you can handle '{subtopic}' with stronger C1 control and refined expression. "
        f"Your speaking is more persuasive and your writing shows advanced structure and style."
    )

    return (day_num, title, topic, eng_exp, guj_exp, vocab, sentences, tips, tasks, mini_test, outcome)


def build_days():
    days = []
    for module_index, module in enumerate(MODULES):
        start, end = module["range"]
        for day in range(start, end + 1):
            i = day - 271
            subtopic = SUBTOPICS[i]
            # Offset each module by 3 to stagger starting vocab windows beyond the 20-word bank rhythm.
            idx = (day - start) + module_index * 3
            days.append(build_day_entry(day, module, subtopic, idx))
    return days


DAYS = build_days()


def build_book():
    doc = Document()

    for section in doc.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    add_para(doc, "🇩🇪", size_pt=64, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=32, space_after=8)
    add_para(doc, "DEUTSCH LERNEN", size_pt=34, bold=True, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "C1 COURSE", size_pt=22, bold=True, color=PURPLE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "Days 271 to 330", size_pt=18, bold=True, color=ORANGE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    add_para(doc, "ગુજરાતી ભાષકો માટે C1 જર્મન અભ્યાસ", size_pt=14, bold=True, color=DARK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "For Gujarati Speakers | English + Gujarati Explanations", size_pt=11, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "60 Days · Rhetoric · Research · Strategy · C1 Mastery", size_pt=11, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)

    add_para(doc, "📖 How to Use This C1 Book", size_pt=14, bold=True, color=BLUE, space_after=4)
    tips = [
        "📅 Study one day at a time and keep your streak.",
        "✍️ Write daily; C1 improvement depends on precision editing.",
        "🗣️ Speak aloud for at least 15 minutes daily.",
        "🔁 Review errors and register shifts every week.",
        "🎯 Focus on structure, argument depth, and stylistic control.",
        "🧪 Use mock days as timed C1 exam practice.",
    ]
    for t in tips:
        add_bullet(doc, t)

    for entry in DAYS:
        add_day(doc, *entry)

    out = "German_Learning_Book_C1_Gujarati_Days_271_330.docx"
    doc.save(out)
    print(f"✅ Saved: {out}")
    print(f"   Days: {len(DAYS)}")


if __name__ == "__main__":
    build_book()
