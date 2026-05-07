# -*- coding: utf-8 -*-
"""Generate German_Learning_Book_B2_Gujarati_Days_201_270.docx (Days 201-270)"""

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
        "range": (201, 210),
        "name": "Advanced Grammar and Precision",
        "focus": "complex clause control, nominal style, and nuanced connector usage",
        "vocab": [
            ("die Nominalisierung", "nominalization", "સંજ્ઞીકરણ", "noh-mi-na-li-ZEE-rung"),
            ("die Präzision", "precision", "ચોકસાઈ", "pre-tsi-ZY-ohn"),
            ("die Nuance", "nuance", "સૂક્ષ્મ ફરક", "nu-AN-seh"),
            ("dementsprechend", "accordingly", "તે મુજબ", "deh-ment-SPRE-khend"),
            ("hingegen", "in contrast", "બીજી તરફ", "HIN-gay-gen"),
            ("somit", "thus", "આથી", "ZOH-mit"),
            ("voraussetzen", "to assume/require", "પૂર્વધારણા કરવી", "fo-ROWS-zet-sen"),
            ("das Verhältnis", "relation", "સંબંધ", "fer-HELT-nis"),
            ("der Bezug", "reference", "સંદર્ભ", "beh-TSOOK"),
            ("die Struktur", "structure", "રચના", "shtruk-TOOR"),
            ("belegen", "to support/prove", "પ્રમાણિત કરવું", "beh-LAY-gen"),
            ("ableiten", "to derive", "ઉપજાવવું", "AB-ly-ten"),
            ("einordnen", "to classify", "વર્ગીકૃત કરવું", "EYN-ord-nen"),
            ("gewichten", "to weigh", "તોલવું", "geh-VIKH-ten"),
            ("differenzieren", "to differentiate", "ફરક બતાવવો", "dif-fe-ren-TSY-ren"),
            ("das Kriterium", "criterion", "માપદંડ", "kri-TAY-ri-um"),
            ("die Folgerung", "conclusion", "નિષ્કર્ષ", "FOL-ge-rung"),
            ("mehrdeutig", "ambiguous", "અસ્પષ્ટ/બહુઅર્થવાળો", "MAYR-doy-tig"),
            ("eindeutig", "clear/unambiguous", "સ્પષ્ટ", "EYN-doy-tig"),
            ("im Hinblick auf", "with regard to", "ના સંદર્ભમાં", "im HIN-blik owf"),
        ],
    },
    {
        "range": (211, 220),
        "name": "Academic German and Argumentation",
        "focus": "summaries, formal argument chains, and evidence-based writing",
        "vocab": [
            ("die These", "thesis", "થિસિસ", "TAY-zeh"),
            ("das Argument", "argument", "દલીલ", "ar-goo-MENT"),
            ("die Begründung", "justification", "કારણસભર સમર્થન", "beh-GRÜN-dung"),
            ("widerlegen", "to refute", "ખંડન કરવું", "VEE-der-lay-gen"),
            ("der Beleg", "evidence", "પુરાવો", "beh-LEHK"),
            ("zusammenfassen", "to summarize", "સારાંશ આપવો", "tsu-ZA-men-fa-sen"),
            ("die Quelle", "source", "સ્રોત", "KVEH-leh"),
            ("die Auswertung", "evaluation", "વિશ્લેષણ", "OWS-ver-toong"),
            ("die Grafik", "chart/graphic", "ગ્રાફિક", "GRA-fik"),
            ("der Verlauf", "trend/development", "પ્રવાહ", "fer-LOWF"),
            ("zunehmen", "to increase", "વધવું", "TSOO-nay-men"),
            ("abnehmen", "to decrease", "ઘટવું", "AP-nay-men"),
            ("die Gegenposition", "counter-position", "વિરોધી દૃષ્ટિકોણ", "GAY-gen-po-si-TSY-ohn"),
            ("schlussfolgern", "to infer", "નિષ્કર્ષ કાઢવો", "SHLOOS-fol-gern"),
            ("die Kernaussage", "core statement", "મુખ્ય નિવેદન", "KERN-ow-sa-ge"),
            ("veranschaulichen", "to illustrate", "દ્રશ્યરૂપે સમજાવવું", "fer-AN-show-li-khen"),
            ("prägnant", "concise", "સંક્ષિપ્ત પરંતુ અસરકારક", "PRAYG-nant"),
            ("kohärent", "coherent", "સુસંગત", "ko-hay-RENT"),
            ("nachvollziehbar", "comprehensible", "સહજ રીતે સમજાય એવું", "nakh-FOL-tsee-bar"),
            ("objektiv", "objective", "નિષ્પક્ષ", "ob-yek-TEEF"),
        ],
    },
    {
        "range": (221, 230),
        "name": "Professional Communication",
        "focus": "leadership language, negotiation, and workplace conflict management",
        "vocab": [
            ("die Zielvereinbarung", "target agreement", "લક્ષ્ય કરાર", "TSEEL-fer-eyn-ba-rung"),
            ("die Rückmeldung", "feedback", "પ્રતિસાદ", "RÜK-mel-doong"),
            ("abstimmen", "to coordinate", "સંયોજન કરવું", "AP-shtim-men"),
            ("priorisieren", "to prioritize", "પ્રાથમિકતા આપવી", "pri-o-ri-ZEE-ren"),
            ("eskalieren", "to escalate", "તીવ્ર બનવું", "es-ka-LEE-ren"),
            ("entschärfen", "to defuse", "તણાવ ઘટાડવો", "ent-SHER-fen"),
            ("die Zuständigkeit", "responsibility", "જવાબદારી", "TSOO-shten-dig-kite"),
            ("verbindlich", "binding/committed", "બંધનકારક", "fer-BIND-likh"),
            ("fristgerecht", "on time", "સમયમર્યાદા અંદર", "FRIST-ge-rekht"),
            ("der Engpass", "bottleneck", "અડચણ બિંદુ", "ENG-pas"),
            ("die Ressource", "resource", "સંસાધન", "re-SOR-seh"),
            ("die Umsetzung", "implementation", "અમલીકરણ", "OOM-zet-soong"),
            ("nachverfolgen", "to follow up", "પાછળથી તપાસવું", "nakh-fer-FOL-gen"),
            ("die Abstimmung", "alignment", "સુમેળ", "AP-shtim-moong"),
            ("kompromissbereit", "willing to compromise", "સમાધાન માટે તૈયાર", "kom-pro-MIS-be-ryt"),
            ("zielorientiert", "goal-oriented", "લક્ષ્યકેન્દ્રિત", "TSEEL-o-ri-en-TEERT"),
            ("die Verhandlung", "negotiation", "વાટાઘાટ", "fer-HAND-loong"),
            ("nachhaltig", "sustainable", "ટકાઉ", "NAKH-hal-tig"),
            ("belastbar", "resilient", "દબાણ સહનશીલ", "beh-LAST-bar"),
            ("die Kernaufgabe", "core task", "મુખ્ય કાર્ય", "KERN-owf-ga-be"),
        ],
    },
    {
        "range": (231, 240),
        "name": "Society, Media, and Policy",
        "focus": "public discourse, media literacy, and structured opinion expression",
        "vocab": [
            ("die Medienkompetenz", "media literacy", "મીડિયા સમજ શક્તિ", "MAY-di-en-kom-pe-TENTS"),
            ("die Desinformation", "disinformation", "ભ્રામક માહિતી", "des-in-for-ma-TSY-ohn"),
            ("die Debatte", "debate", "ચર્ચા", "deh-BAT-teh"),
            ("die Maßnahme", "measure", "કદમ", "MAHS-na-meh"),
            ("die Beteiligung", "participation", "ભાગીદારી", "beh-TY-li-gung"),
            ("der Konsens", "consensus", "સામાન્ય સહમતિ", "kon-ZENS"),
            ("die Kontroverse", "controversy", "વિવાદાસ્પદ મુદ્દો", "kon-tro-VER-zeh"),
            ("regulieren", "to regulate", "નિયમન કરવું", "re-gu-LEE-ren"),
            ("der Rahmen", "framework", "ઢાંચો", "RAH-men"),
            ("die Verantwortung", "responsibility", "જવાબદારી", "fer-ANT-vort-oong"),
            ("die Transparenz", "transparency", "પારદર્શિતા", "trans-pa-RENTS"),
            ("die Perspektive", "perspective", "દૃષ્ટિકોણ", "per-spek-TEE-ve"),
            ("gleichberechtigt", "equal rights", "સમાન અધિકારવાળો", "GLYKH-be-rekh-tigt"),
            ("die Integration", "integration", "સમાવેશ", "in-te-gra-TSY-ohn"),
            ("der Zusammenhalt", "social cohesion", "સામાજિક એકતા", "tsu-ZA-men-halt"),
            ("abwägen", "to balance/weigh", "તોલવું", "AP-vay-gen"),
            ("die Datensicherheit", "data security", "ડેટા સુરક્ષા", "DA-ten-zikher-hite"),
            ("die Privatsphäre", "privacy", "ગોપનીયતા", "pri-VAHT-sfay-reh"),
            ("die Teilhabe", "inclusion/participation", "સહભાગિતા", "TYL-ha-be"),
            ("langfristig", "long-term", "દીર્ઘકાલીન", "LANG-fris-tig"),
        ],
    },
    {
        "range": (241, 250),
        "name": "B2 Exam Skills and Strategy",
        "focus": "task fulfilment, time management, and high-scoring exam language",
        "vocab": [
            ("die Aufgabenstellung", "task prompt", "પ્રશ્નની સૂચના", "OWF-ga-ben-shtel-loong"),
            ("der Zeitplan", "time plan", "સમયયોજના", "TSYTE-plan"),
            ("die Gliederung", "outline", "રચનાત્મક આયોજન", "GLI-de-rung"),
            ("der Hauptteil", "main part", "મુખ્ય ભાગ", "HOWPT-tyl"),
            ("die Einleitung", "introduction", "પ્રારંભ", "EYN-ly-toong"),
            ("der Schluss", "conclusion", "અંતિમ ભાગ", "SHLOOS"),
            ("präzisieren", "to specify", "સ્પષ્ટ કરવું", "pray-tsi-ZEE-ren"),
            ("die Kohärenz", "coherence", "સુસંગતતા", "ko-hay-RENTS"),
            ("die Relevanz", "relevance", "સંબંધિતતા", "re-leh-VANTS"),
            ("der Redemittel", "speaking phrase", "બોલવાની રચના", "RAY-de-mit-tel"),
            ("die Korrektur", "correction", "સુધારો", "kor-rek-TOOR"),
            ("überarbeiten", "to revise", "પુનઃસંપાદિત કરવું", "ü-ber-AR-by-ten"),
            ("bewerten", "to assess", "મૂલ્યાંકન કરવું", "beh-VER-ten"),
            ("der Schwerpunkt", "focus", "કેન્દ્રબિંદુ", "SHVER-punkt"),
            ("treffsicher", "precise/accurate", "એકદમ યોગ્ય", "TREF-zi-kher"),
            ("ausgewogen", "balanced", "સંતુલિત", "OWS-ge-vo-gen"),
            ("der Fehlertyp", "error type", "ભૂલનો પ્રકાર", "FAY-ler-tüp"),
            ("die Musterlösung", "model solution", "નમૂનાત્મક ઉકેલ", "MOOS-ter-lö-zoong"),
            ("zielgenau", "targeted", "નિશ્ચિત લક્ષ્યવાળો", "TSEEL-ge-now"),
            ("die Selbstkontrolle", "self-check", "સ્વ-ચકાસણી", "ZELBST-kon-tro-leh"),
        ],
    },
    {
        "range": (251, 260),
        "name": "Fluency, Register, and Style",
        "focus": "natural high-level expression, idiomatic language, and register control",
        "vocab": [
            ("der Sprachstil", "language style", "ભાષાશૈલી", "SHPRAKH-shtil"),
            ("umgangssprachlich", "colloquial", "બોલચાલની ભાષા", "OOM-gangs-shprah-khlich"),
            ("formell", "formal", "ઔપચારિક", "for-MEL"),
            ("nuanciert", "nuanced", "સૂક્ષ્મ અર્થવાળો", "nu-an-SEERT"),
            ("die Redewendung", "idiom", "રૂઢિપ્રયોગ", "RAY-de-ven-doong"),
            ("treffend", "apt/accurate", "યોગ્ય અને અસરકારક", "TREF-ent"),
            ("die Betonung", "stress/emphasis", "જોર/ઉચ્ચાર ભાર", "beh-TOH-noong"),
            ("verdeutlichen", "to clarify", "સ્પષ્ટ બનાવવું", "fer-DOYT-li-khen"),
            ("abschwächen", "to soften", "નરમ બનાવવું", "AP-shvay-khen"),
            ("verstärken", "to intensify", "મજબૂત બનાવવું", "fer-SHTER-ken"),
            ("die Eleganz", "elegance", "અભિવ્યક્તિની સુંદરતા", "eh-leh-GANTS"),
            ("der Feinschliff", "final polish", "અંતિમ ઘસારો", "FYN-shlif"),
            ("stichhaltig", "well-founded", "મજબૂત આધારવાળો", "SHTIK-hal-tig"),
            ("die Prägnanz", "conciseness", "સંક્ષિપ્ત અસર", "PRAYG-nants"),
            ("abwechslungsreich", "varied", "વિવિધતાસભર", "AP-veks-loongs-rykh"),
            ("zielgruppengerecht", "appropriate for audience", "પ્રેક્ષક મુજબ યોગ્ય", "TSEEL-gru-pen-ge-rekht"),
            ("die Ironie", "irony", "વ્યંગ્ય", "ee-ro-NEE"),
            ("überzeugen", "to convince", "મનાવવું", "ü-ber-TSOY-gen"),
            ("pointiert", "sharp/pointed", "તીક્ષ્ણ રીતે રજૂ કરેલું", "poin-TEEERT"),
            ("souverän", "confident/composed", "આત્મવિશ્વાસપૂર્ણ", "soo-ve-RAIN"),
        ],
    },
    {
        "range": (261, 270),
        "name": "Final Consolidation and B2 Launch",
        "focus": "integrated performance, mock exams, and long-term post-B2 continuity",
        "vocab": [
            ("die Gesamtauswertung", "overall evaluation", "કુલ મૂલ્યાંકન", "geh-ZAMT-ows-ver-toong"),
            ("der Leistungsstand", "performance level", "પ્રદર્શન સ્તર", "LY-stoongs-shtant"),
            ("der Meilenstein", "milestone", "મહત્ત્વનું મંચ", "MY-len-shtyn"),
            ("festigen", "to consolidate", "મજબૂત કરવું", "FES-ti-gen"),
            ("abrufen", "to retrieve", "યાદમાંથી કાઢવું", "AP-roo-fen"),
            ("der Transfer", "transfer", "રૂપાંતરાત્મક ઉપયોગ", "trans-FER"),
            ("die Anwendung", "application", "પ્રયોગ", "AN-ven-doong"),
            ("durchhalten", "to persevere", "અડગ રહેવું", "DOORKH-hal-ten"),
            ("der Feinschliff", "final polish", "અંતિમ સુધારો", "FYN-shlif"),
            ("die Routine", "routine", "રૂટિન", "roo-TEE-neh"),
            ("die Zielstufe", "target level", "લક્ષ્ય સ્તર", "TSEEL-shtu-feh"),
            ("erreichen", "to achieve", "હાંસલ કરવું", "er-RY-khen"),
            ("die Anschlussstrategie", "next-step strategy", "આગળની વ્યૂહરચના", "AN-shloos-shtra-te-GEE"),
            ("weiterführen", "to continue", "આગળ ચાલુ રાખવું", "VY-ter-fü-ren"),
            ("eigenständig", "independent", "સ્વતંત્ર રીતે", "EY-gen-shten-dig"),
            ("reflektieren", "to reflect", "પાછું જોઈ વિચારીવું", "re-flek-TEE-ren"),
            ("die Lernkurve", "learning curve", "શીખવાની વક્રતા", "LERN-kur-ve"),
            ("nachhaltig sichern", "to secure sustainably", "ટકાઉ રીતે સુનિશ્ચિત કરવું", "NAKH-hal-tig ZI-khern"),
            ("die Abschlussphase", "final phase", "અંતિમ તબક્કો", "AP-shloos-fa-zeh"),
            ("sattelfest", "fully confident", "પકડીને નિપુણ", "ZAT-tel-fest"),
        ],
    },
]

SUBTOPICS = [
    "Nominalization vs verbal style", "Advanced connector chains", "Konjunktiv I in reports", "Participle constructions", "Passive alternatives", "Complex relative clauses", "Prepositional verb patterns", "Word formation at B2", "Modal particles for nuance", "Precision rewrite workshop",
    "Structured summary writing", "Argument and counterargument", "Data commentary language", "Academic presentation openings", "Citing and referencing ideas", "Paraphrasing without loss", "Critical response writing", "Discussion moderation phrases", "Formal conclusion techniques", "Academic mini-project day",
    "Meeting facilitation language", "Negotiation framing", "Conflict de-escalation", "Professional email escalation", "Project risk communication", "Task delegation language", "Constructive feedback loops", "Client communication quality", "Decision memo writing", "Workplace simulation day",
    "Media reliability checks", "Climate and policy debate", "Migration and integration language", "Education reform argumentation", "Healthcare system comparisons", "Digital privacy discussion", "Public participation vocabulary", "Intercultural misunderstanding repair", "Volunteer and civic engagement", "Society debate simulation",
    "B2 reading strategy", "B2 listening note system", "B2 writing structure", "B2 speaking monologue", "B2 speaking interaction", "Mediation and reformulation", "Time management under exam pressure", "Frequent error correction", "Grammar consolidation sprint", "Full B2 mock set A",
    "Idioms in context", "Collocations for fluency", "Formal vs informal register", "Storytelling with tension", "Persuasive speech language", "Spontaneous speaking drills", "Accent and speed listening", "High-precision writing", "Advanced review and polish", "Full B2 mock set B",
    "Weak-point diagnosis", "Integrated skill circuit", "Long-form speaking day", "Long-form writing day", "Listening under noise", "Reading under time limit", "Final correction bootcamp", "Complete B2 mock exam", "Reflection and next plan", "B2 graduation and C1 bridge",
]


def pick_vocab(bank, start, size=15):
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
        f"Day {day_num} moves your B2 German forward with a focused lesson on '{subtopic}'. "
        f"Today you train {module['focus']}. The goal is to produce longer, precise, and well-structured German "
        f"that works in exams, work, and academic discussions."
    )
    guj_exp = (
        f"Day {day_num} માં તમે B2 સ્તરે '{subtopic}' પર focused અભ્યાસ કરો છો. આજે તમારું ધ્યાન "
        f"{module['focus']} પર છે. હેતુ એ છે કે તમે exam, work અને academic discussion માટે લાંબી, "
        f"ચોક્કસ અને સુંવાળી German બનાવી શકો."
    )

    sentences = [
        (
            f"Im heutigen Training steht {k1} im Mittelpunkt.",
            f"In today's training, {k1} is central.",
            f"આજના અભ્યાસમાં {k1} કેન્દ્રમાં છે."
        ),
        (
            f"Ich kann {k2} jetzt bewusster und präziser verwenden.",
            f"I can now use {k2} more consciously and precisely.",
            f"હવે હું {k2} ને વધુ જાગૃત રીતે અને ચોકસાઈથી વાપરી શકું છું."
        ),
        (
            f"Unsere Gruppe hat mit {k3} eine klare Lösung formuliert.",
            f"Our group formulated a clear solution using {k3}.",
            f"અમારા જૂથે {k3} વડે સ્પષ્ટ ઉકેલ રજૂ કર્યો."
        ),
        (
            "Je genauer die Struktur ist, desto überzeugender wirkt die Aussage.",
            "The clearer the structure, the more convincing the statement becomes.",
            "રચના જેટલી સ્પષ્ટ હોય, નિવેદન તેટલું જ વિશ્વસનીય લાગે છે."
        ),
        (
            "Ich überprüfe meine Formulierungen, bevor ich den Text abgebe.",
            "I check my formulations before submitting the text.",
            "હું લખાણ સબમિટ કરતાં પહેલા મારી અભિવ્યક્તિઓ ચકાસું છું."
        ),
    ]

    tips = [
        ("📌", "B2 Focus", f"Today prioritize accuracy + coherence. Use {k1} and {k2} actively in speaking and writing.", "E8F0FE"),
        ("💡", "Strategy", "Write first draft fast, then spend a second round only on connectors, register, and grammar.", "FFF8E1"),
        ("🎯", "Progress", "B2 growth comes from deliberate correction. Keep an error notebook and revisit it weekly.", "E6F4EA"),
    ]

    tasks = [
        f"Write a 160–180 word paragraph on '{subtopic}' using at least 6 target words.",
        "Record a 2-minute spoken response and self-check word order and connector quality.",
        "Rewrite your text once in a more formal register and compare both versions.",
    ]

    mini_test = [
        (f"Translate into German: 'This point is central in today's topic.'", f"Dieser Punkt steht im heutigen Thema im Mittelpunkt."),
        (f"Which phrase fits formal writing better: colloquial shortcut or structured connector?", "Structured connector"),
        (f"Give one key word from today's vocabulary.", vocab[0][0]),
        ("At B2, what should improve together: only speed or speed + precision?", "Speed + precision"),
        ("What is one self-check action before submission?", "Review grammar, connectors, and clarity"),
    ]

    outcome = (
        f"After Day {day_num}, you can handle '{subtopic}' with more control and clearer B2 expression. "
        f"Your speaking is more organized and your writing is more professional and exam-ready."
    )

    return (day_num, title, topic, eng_exp, guj_exp, vocab, sentences, tips, tasks, mini_test, outcome)


def build_days():
    days = []
    for module_index, module in enumerate(MODULES):
        start, end = module["range"]
        for day in range(start, end + 1):
            i = day - 201
            subtopic = SUBTOPICS[i]
            idx = (day - start) + module_index * 2
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
    add_para(doc, "B2 COURSE", size_pt=22, bold=True, color=PURPLE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "Days 201 to 270", size_pt=18, bold=True, color=ORANGE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    add_para(doc, "ગુજરાતી ભાષકો માટે B2 જર્મન અભ્યાસ", size_pt=14, bold=True, color=DARK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "For Gujarati Speakers | English + Gujarati Explanations", size_pt=11, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "70 Days · Advanced Grammar · Professional German · B2 Exam Mastery", size_pt=11, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)

    add_para(doc, "📖 How to Use This B2 Book", size_pt=14, bold=True, color=BLUE, space_after=4)
    tips = [
        "📅 Study one day at a time and keep your streak.",
        "✍️ Write every day; B2 improves fastest through output.",
        "🗣️ Speak aloud for at least 10 minutes daily.",
        "🔁 Review error patterns every week.",
        "🎯 Focus on precision, coherence, and register.",
        "🧪 Use mock days as timed exam practice.",
    ]
    for t in tips:
        add_bullet(doc, t)

    for entry in DAYS:
        add_day(doc, *entry)

    out = "German_Learning_Book_B2_Gujarati_Days_201_270.docx"
    doc.save(out)
    print(f"✅ Saved: {out}")
    print(f"   Days: {len(DAYS)}")


if __name__ == "__main__":
    build_book()
