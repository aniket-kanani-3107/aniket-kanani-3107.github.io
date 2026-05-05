# -*- coding: utf-8 -*-
"""Generate German_Learning_Book_A2_Gujarati_Days_121_200.docx (Days 121-200)"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─────────────────────────── helpers ────────────────────────────────────────
BLUE   = RGBColor(0x1A, 0x73, 0xE8)
YELLOW = RGBColor(0xF9, 0xAB, 0x00)
GRAY   = RGBColor(0x55, 0x55, 0x55)
GREEN  = RGBColor(0x0F, 0x96, 0x60)
RED    = RGBColor(0xD9, 0x34, 0x25)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
DARK   = RGBColor(0x20, 0x20, 0x20)


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


# ═══════════════════════════ BOOK DATA ═══════════════════════════════════════
DAYS = []

# ── Day 121 ─────────────────────────────────────────────────────────────────
DAYS.append((121,
"WELCOME TO THE B1 BRIDGE 🌟",
"A2→B1 Transition & Goals",
"Welcome to Day 121! You have completed A2 and are now starting the B1 bridge. B1 means you can handle real-life conversations, understand longer texts, and express opinions with reasons. This phase builds confidence, accuracy, and fluency step by step.",
"Day 121 માં તમારું સ્વાગત! તમે A2 પૂર્ણ કર્યું અને હવે B1 bridge શરૂ થાય છે. B1 એટલે કે તમે real-life conversations કરી શકો, લાંબા texts સમજી શકો, અને કારણો સાથે opinions આપી શકો. આ phase confidence, accuracy અને fluency વધારશે.",
[
("der Übergang","transition","પ્રવેશ/પુલ","ÜR-geh-gang"),
("das Ziel","goal","લક્ષ્ય","TSEEL"),
("die Fähigkeit","skill/ability","કુશળતા","FÄH-ig-kite"),
("das Niveau","level","સ્તર","nee-VOH"),
("fließend","fluent","સરસ રીતે બોલવું","FLEE-sent"),
("selbstständig","independent","સ્વતંત્ર","ZELPST-shtän-dig"),
("fortsetzen","to continue","આગળ વધવું","FORT-set-sen"),
("die Herausforderung","challenge","પડકાર","he-ROWS-for-deh-rung"),
("der Fortschritt","progress","પ્રગતિ","FORT-shrit"),
("die Wiederholung","revision","પુનરાવર્તન","VEE-der-hoh-lung"),
("der Plan","plan","યોજનાં","PLAHN"),
("konsequent","consistent","સતત","kon-ze-KVENT"),
("diszipliniert","disciplined","શિસ્તબદ્ધ","dis-zi-pli-NEERT"),
("motiviert","motivated","પ્રેરિત","moh-ti-VEERT"),
("die Routine","routine","રૂટિન","roo-TEE-neh"),
],
[
("Ich starte jetzt den B1-Bridge-Kurs.","I am starting the B1 bridge course now.","હું હવે B1 bridge course શરૂ કરું છું."),
("Mein Ziel ist es, fließender zu sprechen.","My goal is to speak more fluently.","મારું લક્ષ્ય વધુ fluent રીતે બોલવું છે."),
("Mit Disziplin kommt echter Fortschritt.","With discipline comes real progress.","શિસ્ત સાથે સાચી પ્રગતિ આવે છે."),
("Ich mache jeden Tag eine kurze Wiederholung.","I do a short review every day.","હું દરરોજ થોડું પુનરાવર્તન કરું છું."),
("B1 bedeutet mehr Selbstständigkeit im Alltag.","B1 means more independence in daily life.","B1 એટલે દૈનિક જીવનમાં વધુ સ્વતંત્રતા."),
],
[
("🏆","B1 Mindset","B1 is about longer texts, clearer arguments, and more natural speech. Don’t rush — build depth.","E8F0FE"),
("💡","Daily Routine","15–20 minutes of vocabulary + 20 minutes of speaking practice daily gives fast results.","FFF8E1"),
("🎯","Focus","Accuracy first, speed second. Say it correctly, then say it faster.","E6F4EA"),
],
["Write your 3 personal B1 goals in German.","Create a weekly study plan (Mon–Sun) with time blocks.","Record a 2‑minute self-introduction in German."],
[("What does 'Übergang' mean?","Transition"),("Translate: 'Mein Ziel ist fließendes Deutsch.'","My goal is fluent German."),("'konsequent' means?","Consistent"),("Translate: 'Ich mache jeden Tag Wiederholung.'","I review every day."),("B1 focuses on what?","Longer texts, reasons, and natural speech")],
"You have officially started the B1 bridge. Your foundation is strong and your next 80 days will make your German confident and natural."
))

# ── Day 122 ─────────────────────────────────────────────────────────────────
DAYS.append((122,
"KONJUNKTIV II — WÜRDE + INFINITIV 🌀",
"Polite & Hypothetical Sentences",
"Konjunktiv II is used for polite requests, hypotheticals, and unreal situations. The easiest form is: würde + infinitive. Example: Ich würde gern reisen. (I would like to travel.)",
"Konjunktiv II polite requests અને unreal situations માટે વપરાય છે. સૌથી સરળ રચના: würde + infinitive. ઉદા: Ich würde gern reisen. (હું મુસાફરી કરવા ઈચ્છું છું.)",
[
("würde","would (aux.)","હોય તો","VÜR-deh"),
("gern","gladly / like to","ખુશીથી","GERN"),
("vielleicht","maybe","કદાચ","FEE-laykht"),
("die Möglichkeit","possibility","શક્યતા","MÖG-likh-kite"),
("der Wunsch","wish","ઇચ્છા","VOONSH"),
("höflich","polite","વિનમ્ર","HÖF-likh"),
("bitten","to ask/request","વિનંતી કરવી","BIT-en"),
("reisen","to travel","મુસાફરી કરવી","RY-zen"),
("lernen","to learn","શીખવું","LEHR-nen"),
("kaufen","to buy","ખરીદવું","KOW-fen"),
("machen","to do","કરવું","MAKH-en"),
("sagen","to say","કહેવું","ZAH-gen"),
("brauchen","to need","જરૂર હોવી","BROW-khen"),
("helfen","to help","મદદ કરવી","HEL-fen"),
("besuchen","to visit","મળવા જવું","beh-ZOO-khen"),
],
[
("Ich würde gern Deutsch sprechen.","I would like to speak German.","હું German બોલવા ઈચ્છું છું."),
("Würden Sie mir bitte helfen?","Would you please help me?","શું તમે મને મદદ કરી શકો?"),
("Ich würde morgen in die Stadt gehen.","I would go to the city tomorrow.","હું કાલ શહેરમાં જઈશ/જઈશ ઈચ્છા છે."),
("Vielleicht würde ich ein neues Handy kaufen.","Maybe I would buy a new phone.","કદાચ હું નવો ફોન ખરીદું."),
("Was würdest du tun?","What would you do?","તુ શું કરત?"),
],
[
("📌","Formula","würde + infinitive at the end. Example: Ich würde heute lernen.","E8F0FE"),
("💡","Politeness","würde makes requests softer and very polite: Würden Sie…?",
"FFF8E1"),
("⚠️","Word Order","würde is the conjugated verb in position 2; infinitive stays at the end.","FFE0E0"),
],
["Write 5 sentences using 'würde' + infinitive.","Convert 5 present tense sentences into Konjunktiv II with würde.","Role-play: ask for help politely using 'Würden Sie…?'"],
[("What is the formula for Konjunktiv II (easy form)?","würde + infinitive"),("Translate: 'Ich würde gern bleiben.'","I would like to stay."),("'höflich' means?","Polite"),("Where does the infinitive go?","At the end"),("Translate: 'Würdest du mir helfen?'","Would you help me?")],
"You can now use the easiest Konjunktiv II form for polite and hypothetical sentences. This is a major B1 communication upgrade."
))

# ── Day 123 ─────────────────────────────────────────────────────────────────
DAYS.append((123,
"KONJUNKTIV II WITH MODALS 💬",
"Could / Should / Might",
"Konjunktiv II with modal verbs makes your German more polite and nuanced. Examples: Ich könnte kommen (I could come), Ich sollte lernen (I should learn), Ich dürfte fragen (I might be allowed to ask).",
"Konjunktiv II modal verbs સાથે વધુ polite અને nuanced બને છે. ઉદા: Ich könnte kommen (હું આવી શકું), Ich sollte lernen (મારે શીખવું જોઈએ), Ich dürfte fragen (હું કદાચ પૂછું).",
[
("könnte","could","શકી શકું","KÖN-teh"),
("sollte","should","જોઈએ","ZOL-teh"),
("dürfte","might be allowed","મંજूरी હોઈ શકે","DÜR-fteh"),
("müsste","would have to","કરી જ પડે","MÜS-teh"),
("wollte","would want","ઈચ્છું છું","VOL-teh"),
("dürfen","to be allowed","મંજूरी હોવી","DÜR-fen"),
("sollen","should (duty)","જોઈએ","ZOL-en"),
("müssen","must","કરી જ પડે","MÜS-en"),
("können","can","શકીવું","KÖN-en"),
("wollen","to want","ઇચ્છવું","VOL-en"),
("ratsam","advisable","ઉપયોગી","RAHT-zam"),
("die Empfehlung","recommendation","ભલામણ","em-PFAY-lung"),
("vielleicht","maybe","કદાચ","FEE-laykht"),
("besser","better","વધારે સારું","BES-er"),
("dringend","urgent","તાત્કાલિક","DRING-ent"),
],
[
("Ich könnte heute früher kommen.","I could come earlier today.","હું આજે વહેલા આવી શકું."),
("Du solltest mehr schlafen.","You should sleep more.","તને વધુ ઊંઘવી જોઈએ."),
("Ich dürfte Sie kurz etwas fragen?","May I briefly ask you something?","શું હું તમને ટૂંકું કંઈ પૂછું?"),
("Wir müssten einen neuen Termin finden.","We would have to find a new appointment.","અમે નવો appointment શોધવો પડશે."),
("Das wäre besser.","That would be better.","આ વધુ સારું હોત."),
],
[
("📌","Polite Modals","könnte / dürfte / sollte sound softer than kann / darf / soll.","E8F0FE"),
("💡","Advice","Use 'sollte' to give friendly advice: 'Du solltest…'","FFF8E1"),
("⚠️","Do not overuse","Mix with 'würde' sentences for variety.","FFE0E0"),
],
["Write 5 sentences with könnte/sollte/dürfte.","Give advice to a friend using 'sollte'.", "Create 3 polite requests using 'dürfte/könnte'."],
[("Konjunktiv II of 'können'?","könnte"),("Translate: 'Du solltest mehr lernen.'","You should study more."),("'dürfte' means?","might be allowed"),("Translate: 'Wir müssten gehen.'","We would have to go."),("Which sounds more polite: darf or dürfte?","dürfte")],
"You can now use Konjunktiv II with modal verbs — a key B1 skill for polite and natural German."
))

# ── Day 124 ─────────────────────────────────────────────────────────────────
DAYS.append((124,
"POLITE REQUESTS & OFFERS 🤝",
"Service German with Konjunktiv II",
"German service situations (shops, hotels, offices) require polite language. Today you practise requests and offers using würde/könnte/dürfte.",
"German service situations માં polite language જરૂરી છે. આજે તમે würde/könnte/dürfte સાથે requests અને offers પ્રેક્ટિસ કરો.",
[
("dürfte ich…?","may I…?","શું હું…?","DÜR-fteh ikh"),
("könnten Sie…?","could you…?","શું તમે…?","KÖN-ten zee"),
("würden Sie…?","would you…?","શું તમે…?","VÜR-den zee"),
("einverstanden","agree","સહમત","EYN-fer-shtan-den"),
("gerne","with pleasure","ખુશીથી","GER-neh"),
("leider","unfortunately","દુર્ભાગ્યે","LY-der"),
("die Bitte","the request","વિનંતી","BIT-teh"),
("die Auskunft","information","માહિતી","OWS-koonft"),
("die Alternative","alternative","વૈકલ્પિક","al-ter-na-TEE-ve"),
("die Lösung","solution","ઉકેલ","LÖ-zoong"),
("vereinbaren","to arrange","ગોઠવવું","fehr-EYN-bah-ren"),
("verschieben","to postpone","મુલતવી રાખવું","fehr-SHEE-ben"),
("kostenlos","free of charge","મફત","KOS-ten-lohs"),
("sofort","immediately","તરત","zoh-FORT"),
("möglich","possible","શક્ય","MÖG-likh"),
],
[
("Könnten Sie mir bitte helfen?","Could you please help me?","શું તમે કૃપા કરીને મને મદદ કરી શકો?"),
("Dürfte ich eine Frage stellen?","May I ask a question?","શું હું એક પ્રશ્ન પૂછું?"),
("Würden Sie das bitte wiederholen?","Would you please repeat that?","શું તમે કૃપા કરીને ફરી કહેશો?"),
("Leider ist das heute nicht möglich.","Unfortunately that is not possible today.","દુર્ભાગ્યે આજ શક્ય નથી."),
("Gerne, ich helfe Ihnen sofort.","With pleasure, I will help you immediately.","ખુશીથી, હું તરત મદદ કરું છું."),
],
[
("💡","Polite Formula","Könnten Sie…? is the safest polite request in German.","FFF8E1"),
("📌","Tone","Use 'bitte' to soften requests. It matters in German culture.","E8F0FE"),
("🇩🇪","Service Tip","Germans appreciate clear, polite language. Direct + polite = best.","E6F4EA"),
],
["Write 5 polite requests for a hotel or office.","Turn 5 direct commands into polite questions.","Role-play: ask for information at the reception desk."],
[("Translate: 'Könnten Sie das erklären?'","Could you explain that?"),("'einverstanden' means?","Agree"),("How do you say 'May I…?'","Dürfte ich…?"),("Translate: 'Würden Sie bitte warten?'","Would you please wait?"),("'möglich' means?","Possible")],
"You can now handle polite requests and offers in German — essential for professional and service situations."
))

# ── Day 125 ─────────────────────────────────────────────────────────────────
DAYS.append((125,
"KONJUNKTIV II — IRREGULAR FORMS 🔥",
"wäre, hätte, könnte, müsste",
"Some verbs have special Konjunktiv II forms that you must memorize: sein→wäre, haben→hätte, können→könnte, müssen→müsste, wollen→wollte. These are used constantly in spoken German.",
"કેટલા verbs ના ખાસ Konjunktiv II forms યાદ કરવાના છે: sein→wäre, haben→hätte, können→könnte, müssen→müsste, wollen→wollte. આ forms spoken German માં ખૂબ વપરાય છે.",
[
("wäre","would be","હોત","VÄH-reh"),
("hätte","would have","હોતું","HET-eh"),
("könnte","could","શકી હોત","KÖN-teh"),
("müsste","would have to","પડતું","MÜS-teh"),
("wollte","would want","ઈચ્છું હોત","VOL-teh"),
("Wenn ich Zeit hätte…","If I had time…","જો મારી પાસે સમય હોત…","VEN ikh TSYTE HET-eh"),
("Es wäre besser…","It would be better…","એ વધુ સારું હોત…","es VÄH-reh BES-er"),
("Ich hätte gern…","I would like…","મને ગમતું…","ikh HET-eh gern"),
("Wir könnten…","We could…","અમે કરી શકીએ…","veer KÖN-ten"),
("Du müsstest…","You would have to…","તને કરવું પડતું…","doo MÜS-test"),
("eigentlich","actually","ખરેખર","EYE-gen-likh"),
("lieber","rather / prefer","વધારે પસંદ","LEE-ber"),
("falls","in case","જો","FALS"),
("zufrieden","satisfied","સંતુષ્ટ","tsoo-FREE-den"),
("wünschenswert","desirable","ઇચ્છનીય","VÜN-shens-vert"),
],
[
("Wenn ich mehr Zeit hätte, würde ich mehr lesen.","If I had more time, I would read more.","જો મારી પાસે વધુ સમય હોત, તો હું વધુ વાંચત."),
("Es wäre besser, früher zu kommen.","It would be better to come earlier.","વહેલા આવવું વધુ સારું હોત."),
("Ich hätte gern einen Termin.","I would like an appointment.","મને appointment જોઈએ."),
("Wir könnten morgen starten.","We could start tomorrow.","અમે કાલથી શરૂ કરી શકીએ."),
("Du müsstest mehr üben.","You would have to practise more.","તને વધુ પ્રેક્ટિસ કરવી પડશે."),
],
[
("📌","Core Forms","wäre / hätte / könnte / müsste are the top 4 forms — master them first.","E8F0FE"),
("💡","Polite Use","'Ich hätte gern…' is the most polite way to order in cafés.","FFF8E1"),
("⚠️","Unreal condition","wenn + hätte/wäre + würde in main clause for unreal situations.","FFE0E0"),
],
["Write 5 unreal 'wenn' sentences using wäre/hätte.","Practice ordering with 'Ich hätte gern…' (5 items).","Convert 5 normal sentences to Konjunktiv II irregular forms."],
[("Konjunktiv II of 'sein'?","wäre"),("Translate: 'Ich hätte gern einen Kaffee.'","I would like a coffee."),("'müsste' means?","would have to"),("Translate: 'Wenn ich reich wäre…'","If I were rich…"),("Which form is used for 'haben'?","hätte")],
"Irregular Konjunktiv II forms mastered. You now sound much more natural and polite in German."
))

# ── Day 126 ─────────────────────────────────────────────────────────────────
DAYS.append((126,
"UNREAL CONDITIONS (WENN) 🌧️",
"If I were… If I had…",
"Unreal conditions use Konjunktiv II to describe hypothetical situations. Structure: Wenn + Konjunktiv II, dann + Konjunktiv II (often with würde).",
"Hypothetical situations માટે Konjunktiv II વપરાય છે. બંધારણ: Wenn + Konjunktiv II, dann + Konjunktiv II (ઘણું વખત würde સાથે).",
[
("wenn","if/when","જો","VEN"),
("dann","then","પછી","DAN"),
("falls","in case","જો","FALS"),
("wäre","would be","હોત","VÄH-reh"),
("hätte","would have","હોતું","HET-eh"),
("würde","would","હોતું","VÜR-deh"),
("reich","rich","સમૃદ્ધ","RAYKH"),
("arm","poor","ગરીબ","ARM"),
("glücklich","happy","ખુશ","GLÜK-likh"),
("frei","free","મુક્ત","FRY"),
("die Gelegenheit","opportunity","અવસર","geh-LAY-gen-hite"),
("der Traum","dream","સ્વપ્ન","TROWM"),
("verwirklichen","to realize","સાકાર કરવું","fehr-VEER-kli-khen"),
("sparen","to save (money)","બચત કરવી","SHPAR-en"),
("spenden","to donate","દાન આપવું","SHPEN-den"),
],
[
("Wenn ich reich wäre, würde ich viel reisen.","If I were rich, I would travel a lot.","જો હું અમીર હોત, તો ઘણું મુસાફરી કરત."),
("Wenn ich mehr Zeit hätte, würde ich mehr lernen.","If I had more time, I would learn more.","જો મારા પાસે વધુ સમય હોત, તો હું વધુ શીખત."),
("Wenn es nicht regnen würde, würden wir rausgehen.","If it didn't rain, we would go out.","જો વરસાદ ન પડતો, તો અમે બહાર જઈએ."),
("Falls ich krank wäre, bliebe ich zu Hause.","If I were sick, I'd stay at home.","જો હું બીમાર હોત, તો ઘેર રહેતો."),
("Mein Traum wäre ein Haus am Meer.","My dream would be a house by the sea.","મારું સપનું સમુદ્ર પાસેનું ઘર હોત."),
],
[
("📌","Structure","Wenn + Konjunktiv II, dann + Konjunktiv II. 'Dann' is optional.","E8F0FE"),
("💡","Comma Rule","German always uses a comma after the wenn-clause.","FFF8E1"),
("⚠️","Tense","Use Konjunktiv II for unreal present. For past unreal, use hätte/wäre + Partizip II.","FFE0E0"),
],
["Write 5 hypothetical sentences about your life.","Create 3 'Falls...' sentences for emergency situations.","Say 3 dreams using 'Wenn ich...'"],
[("Translate: 'Wenn ich mehr Zeit hätte, würde ich mehr lesen.'","If I had more time, I would read more."),("'der Traum' means?","Dream"),("What punctuation follows a wenn-clause?","A comma"),("Translate: 'Falls es regnet, bleiben wir zu Hause.'","If it rains, we stay at home."),("Which mood is used for unreal conditions?","Konjunktiv II")],
"You can now express hypothetical situations confidently. This is a key B1 grammar skill."
))

# ── Day 127 ─────────────────────────────────────────────────────────────────
DAYS.append((127,
"PASSIVE VOICE — PRESENT 🔄",
"Vorgangspassiv (Präsens)",
"The passive voice focuses on the action, not the doer. In German: werden + Partizip II. Example: Der Brief wird geschrieben. (The letter is being written.)",
"Passive voice action પર ધ્યાન આપે છે, કરનાર પર નહીં. German માં: werden + Partizip II. ઉદા: Der Brief wird geschrieben. (પત્ર લખાઈ રહ્યો છે.)",
[
("das Passiv","the passive voice","Passive","das pa-SIV"),
("werden","to become / to be (passive)","થવું","VEHR-den"),
("geschrieben","written","લખાયેલ","geh-SHREE-ben"),
("gebaut","built","બનાવવામાં આવેલ","geh-BOWT"),
("gemacht","done/made","કરાયેલ","geh-MAKHT"),
("geöffnet","opened","ખુલાયેલ","geh-ÖF-net"),
("geschlossen","closed","બંધ કરાયેલ","geh-SHLOS-en"),
("produziert","produced","ઉત્પાદિત","pro-doo-TSIERT"),
("verkauft","sold","વેચાયેલ","fehr-KOWFT"),
("die Ware","goods","માલ","VAH-reh"),
("der Brief","letter","પત્ર","BREEF"),
("das Haus","house","ઘર","HOWS"),
("die Rechnung","bill","બિલ","REKH-nung"),
("die Lieferung","delivery","ડિલિવરી","LEE-fehr-oong"),
("täglich","daily","દરરોજ","TÄG-likh"),
],
[
("Der Brief wird heute geschrieben.","The letter is being written today.","પત્ર આજે લખાઈ રહ્યો છે."),
("Das Haus wird neu gebaut.","The house is being built newly.","ઘર નવું બનાવવામાં આવે છે."),
("Die Ware wird morgen geliefert.","The goods are delivered tomorrow.","માલ કાલે ડિલિવર થશે."),
("Die Rechnung wird per E-Mail geschickt.","The bill is sent by email.","બિલ ઈમેઈલ દ્વારા મોકલાય છે."),
("Das Fenster wird geöffnet.","The window is being opened.","ખિડકી ખોલવામાં આવી રહી છે."),
],
[
("📌","Formula","werden (conjugated) + Partizip II at the end.","E8F0FE"),
("💡","Agent","If you mention the doer, use 'von + Dative': Das Haus wird von der Firma gebaut.","FFF8E1"),
("⚠️","Use","Passive is common in formal/official texts and announcements.","FFE0E0"),
],
["Convert 5 active sentences into passive.","Write 5 passive sentences about daily processes (emails, deliveries).","Underline the verb position in passive sentences."],
[("Passive formula?","werden + Partizip II"),("Translate: 'Der Brief wird geschrieben.'","The letter is being written."),("How to add the doer?","von + Dative"),("Passive of 'bauen'?","wird gebaut"),("Translate: 'Die Ware wird geliefert.'","The goods are delivered.")],
"You can now use passive voice in the present. This is essential for formal German and official communication."
))

# ── Day 128 ─────────────────────────────────────────────────────────────────
DAYS.append((128,
"PASSIVE VOICE — PAST 🕰️",
"Präteritum & Perfekt Passive",
"Past passive is formed with 'wurde' (Präteritum) or 'ist...worden' (Perfekt). Example: Der Brief wurde geschrieben. / Der Brief ist geschrieben worden.",
"Past passive માટે 'wurde' (Präteritum) અથવા 'ist...worden' (Perfekt) વપરાય છે. ઉદા: Der Brief wurde geschrieben. / Der Brief ist geschrieben worden.",
[
("wurde","was (passive past)","થયું હતું","VOOR-deh"),
("worden","been (passive)","થયેલ","VOR-den"),
("ist ... worden","has been","થઈ ગયું છે","ist...VOR-den"),
("gebaut","built","બનાયેલ","geh-BOWT"),
("geprüft","checked","તપાસાયેલ","geh-PRÜFT"),
("genehmigt","approved","મંજુર","geh-neh-MIKHT"),
("veröffentlicht","published","પ્રકાશિત","fehr-ÖF-fent-likht"),
("gesendet","sent","મોકલાયેલ","geh-ZEN-det"),
("erstellt","created","બનાવેલ","ehr-SHTELT"),
("am Montag","on Monday","સોમવારે","am MON-tahg"),
("letzte Woche","last week","ગયા અઠવાડિયે","LET-steh VO-kheh"),
("im Jahr 2020","in year 2020","વર્ષ 2020માં","im YAHR"),
("die Regel","rule","નિયમ","RAY-gel"),
("das Dokument","document","દસ્તાવેજ","do-koo-MENT"),
("die Entscheidung","decision","નિણર્ય","ent-SHAY-dung"),
],
[
("Der Vertrag wurde gestern unterschrieben.","The contract was signed yesterday.","કોન્ટ્રેક્ટ કાલે સહી થયો હતો."),
("Das Dokument ist geprüft worden.","The document has been checked.","દસ્તાવેજ તપાસાઈ ગયો છે."),
("Die Entscheidung wurde letzte Woche getroffen.","The decision was made last week.","નિણર્ય ગયા અઠવાડિયે થયો હતો."),
("Der Bericht ist veröffentlicht worden.","The report has been published.","રિપોર્ટ પ્રકાશિત થયો છે."),
("Die Rechnung wurde gesendet.","The bill was sent.","બિલ મોકલાયું હતું."),
],
[
("📌","Präteritum Passive","wurde + Partizip II (spoken in stories)." ,"E8F0FE"),
("💡","Perfekt Passive","ist + Partizip II + worden (formal reports)." ,"FFF8E1"),
("⚠️","Worden vs geworden","Passive uses 'worden' (not geworden)." ,"FFE0E0"),
],
["Write 5 passive sentences in the past.","Convert 3 active past sentences into passive.","Find 3 passive sentences in a news article."],
[("Passive past with 'wurde' formula?","wurde + Partizip II"),("Translate: 'Die Entscheidung wurde getroffen.'","The decision was made."),("Passive perfect formula?","ist + Partizip II + worden"),("What is the passive form of 'veröffentlichen'?","wurde veröffentlicht / ist veröffentlicht worden"),("When to use 'worden'?","In passive perfect")],
"You can now use passive in the past. This is a strong B1 grammar tool for formal and written German."
))

# ── Day 129 ─────────────────────────────────────────────────────────────────
DAYS.append((129,
"PASSIVE WITH MODAL VERBS 🛠️",
"Must be done / can be done",
"Passive with modal verbs combines werden + Partizip II + modal infinitive. Example: Die Aufgabe muss gemacht werden. (The task must be done.)",
"Passive with modal verbs: werden + Partizip II + modal infinitive. ઉદા: Die Aufgabe muss gemacht werden. (કાર્ય કરવું જ પડે.)",
[
("muss gemacht werden","must be done","કરવું જ પડે","moos ge-MAKHT VEHR-den"),
("kann erledigt werden","can be completed","પૂર્ણ થઈ શકે","kan ehr-LED-ikt VEHR-den"),
("soll vorbereitet werden","should be prepared","તૈયાર થવું જોઈએ","zol for-bah-RY-tet VEHR-den"),
("darf benutzt werden","may be used","વપરાઈ શકે","darf beh-NOOTST VEHR-den"),
("müsste repariert werden","would have to be repaired","મરમત કરવી પડશે","MÜS-teh reh-pah-REERT VEHR-den"),
("die Aufgabe","task","કાર્ય","OWF-gah-beh"),
("die Maschine","machine","મશીન","ma-SHEE-neh"),
("die Regel","rule","નિયમ","RAY-gel"),
("die Genehmigung","approval","મંજুরি","geh-NEH-mi-gung"),
("die Sicherheit","safety","સુરક્ષા","ZIKH-er-hite"),
("benutzen","to use","વાપરવું","beh-NOOT-sen"),
("reparieren","to repair","મરمت કરવી","reh-pah-REER-en"),
("erledigen","to complete","પૂરું કરવું","ehr-LED-ih-gen"),
("vorbereiten","to prepare","તૈયાર કરવું","for-bah-RY-ten"),
("erlauben","to allow","મંજુર કરવું","ehr-LOW-ben"),
],
[
("Die Aufgabe muss heute gemacht werden.","The task must be done today.","આ કાર્ય આજે કરવું જ પડશે."),
("Die Maschine kann morgen repariert werden.","The machine can be repaired tomorrow.","મશીન કાલે મરમત થઈ શકે છે."),
("Die Regeln sollen eingehalten werden.","The rules should be followed.","નિયમોનું પાલન થવું જોઈએ."),
("Der Raum darf nicht benutzt werden.","The room may not be used.","રૂમ વાપરવો મંજૂર નથી."),
("Alles muss vorbereitet werden.","Everything must be prepared.","બધું તૈયાર થવું જોઈએ."),
],
[
("📌","Structure","Modal in position 2, Partizip II + werden at the end.","E8F0FE"),
("💡","Common Use","Used in instructions, rules, and official notices.","FFF8E1"),
("⚠️","Word Order","Two verbs at the end: gemacht werden, repariert werden.","FFE0E0"),
],
["Write 5 passive+modal sentences about rules in a company.","Convert 3 active modal sentences into passive.","Make a list of 5 safety rules using 'muss/soll/darf'."],
[("Translate: 'Die Aufgabe muss gemacht werden.'","The task must be done."),("Where is the modal verb?","Position 2"),("What stays at the end?","Partizip II + werden"),("Translate: 'Der Raum darf nicht benutzt werden.'","The room may not be used."),("Passive+modal of 'reparieren'?","muss repariert werden")],
"You can now combine passive with modal verbs — a key skill for rules, instructions, and formal German."
))

# ── Day 130 ─────────────────────────────────────────────────────────────────
DAYS.append((130,
"INFINITIVE CLAUSES — UM…ZU 🎯",
"Purpose: In order to…",
"Infinitive clauses express purpose. Structure: um + zu + infinitive. Example: Ich lerne Deutsch, um in Deutschland zu arbeiten. (I learn German in order to work in Germany.)",
"Infinitive clauses purpose દર્શાવે છે. બંધારણ: um + zu + infinitive. ઉદા: Ich lerne Deutsch, um in Deutschland zu arbeiten. (હું Germany માં કામ કરવા માટે German શીખું છું.)",
[
("um … zu","in order to","માટે","oom...tsoo"),
("das Ziel","goal","લક્ષ્ય","TSEEL"),
("die Absicht","intention","ઇરાદો","AP-zikht"),
("erreichen","to reach/achieve","મેળવવું","ehr-RY-khen"),
("verbessern","to improve","સુધારવું","fehr-BES-ern"),
("sparen","to save","બચત કરવી","SHPAR-en"),
("lernen","to learn","શીખવું","LEHR-nen"),
("arbeiten","to work","કામ કરવું","AR-by-ten"),
("verstehen","to understand","સમજવું","fehr-SHTAY-en"),
("erklären","to explain","સમજાવવું","ehr-KLÄ-ren"),
("bestellen","to order","ઓર્ડર કરવું","beh-SHTEL-en"),
("besuchen","to visit","મળવા જવું","beh-ZOO-khen"),
("bewerben","to apply","અરજી કરવી","beh-VER-ben"),
("die Bewerbung","application","અરજી","be-VER-bung"),
("die Prüfung","exam","પરીક્ષા","PRÜ-fung"),
],
[
("Ich lerne Deutsch, um in Deutschland zu arbeiten.","I learn German in order to work in Germany.","હું Germanyમાં કામ કરવા માટે German શીખું છું."),
("Sie spart Geld, um ein Auto zu kaufen.","She saves money to buy a car.","તે કાર ખરીદવા માટે પૈસા બચાવે છે."),
("Wir üben täglich, um die Prüfung zu bestehen.","We practise daily in order to pass the exam.","અમે પરીક્ષા પાસ કરવા માટે દરરોજ પ્રેક્ટિસ કરીએ છીએ."),
("Ich rufe dich an, um eine Frage zu stellen.","I call you to ask a question.","હું પ્રશ્ન પૂછવા માટે તને ફોન કરું છું."),
("Er fährt nach Berlin, um seine Familie zu besuchen.","He goes to Berlin to visit his family.","તે પરિવારને મળવા Berlin જાય છે."),
],
[
("📌","Structure","Main clause + comma + um + zu + infinitive.","E8F0FE"),
("💡","Same Subject","Use um…zu when both clauses have the SAME subject.","FFF8E1"),
("⚠️","Different Subject","If subjects differ, use 'damit' + subordinate clause instead.","FFE0E0"),
],
["Write 5 sentences using um…zu.","Convert 3 sentences with 'damit' into um…zu (same subject).","Underline the infinitive in each sentence."],
[("What does um…zu express?","Purpose (in order to)"),("Translate: 'Ich lerne, um besser zu sprechen.'","I learn in order to speak better."),("When do you use um…zu?","When both clauses have the same subject"),("Translate: 'Sie spart, um zu reisen.'","She saves to travel."),("What to use if subjects differ?","damit + clause")],
"You can now use um…zu clauses to express purpose clearly — a core B1 grammar skill."
))

print("Days 121-130 appended")

# ── Day 131 ─────────────────────────────────────────────────────────────────
DAYS.append((131,
"INFINITIVE CLAUSES — OHNE/ANSTATT 🧩",
"Without doing / instead of doing",
"Two more infinitive clauses: ohne…zu (without doing) and anstatt…zu (instead of doing). Example: Er geht, ohne zu bezahlen. (He leaves without paying.)",
"બે વધુ infinitive clauses: ohne…zu (કર્યા વગર) અને anstatt…zu (આની બદલે). ઉદા: Er geht, ohne zu bezahlen. (તે ચુકવણી કર્યા વગર જાય છે.)",
[
("ohne … zu","without …ing","… કર્યા વગર","OH-neh...tsoo"),
("anstatt … zu","instead of …ing","…ની બદલે","AN-shtat...tsoo"),
("bezahlen","to pay","ચુકવવું","beh-TSAH-len"),
("warten","to wait","રાહ જોવી","VAR-ten"),
("aufhören","to stop","બંધ કરવું","OWF-hö-ren"),
("starten","to start","શરૂ કરવું","SHTAR-ten"),
("diskutieren","to discuss","ચર્ચા કરવી","dis-koo-TEE-ren"),
("streiten","to argue","વિવાદ કરવો","SHTRY-ten"),
("helfen","to help","મદદ કરવી","HEL-fen"),
("klären","to clarify","સ્પષ્ટ કરવું","KLÄ-ren"),
("versuchen","to try","પ્રયાસ કરવો","fehr-ZOO-khen"),
("vergessen","to forget","ભૂલી જવું","fehr-GES-en"),
("vermeiden","to avoid","ટાળવું","fehr-MY-den"),
("ohne Grund","without reason","કારણ વગર","OH-neh GROONT"),
("stattdessen","instead","એના બદલે","SHTAT-des-en"),
],
[
("Er geht, ohne zu bezahlen.","He leaves without paying.","તે ચુકવણી કર્યા વગર જાય છે."),
("Sie trinkt Kaffee, anstatt Tee zu trinken.","She drinks coffee instead of tea.","તે ચા બદલે કોફી પીવે છે."),
("Ich lerne, ohne Musik zu hören.","I study without listening to music.","હું સંગીત સાંભળ્યા વગર ભણું છું."),
("Anstatt zu streiten, sollten wir reden.","Instead of arguing, we should talk.","વિવાદની બદલે વાત કરવી જોઈએ."),
("Er fährt los, ohne zu warten.","He leaves without waiting.","તે રાહ જોયા વગર નીકળી જાય છે."),
],
[
("📌","Structure","Main clause + comma + ohne/anstatt + zu + infinitive.","E8F0FE"),
("💡","Same Subject Rule","Use these only when the subject is the same in both clauses.","FFF8E1"),
("⚠️","Word Order","The infinitive always stays at the end after 'zu'.","FFE0E0"),
],
["Write 5 sentences with ohne…zu.","Write 5 sentences with anstatt…zu.","Convert 3 normal sentences into ohne/anstatt clauses."],
[("Translate: 'Er geht, ohne zu bezahlen.'","He leaves without paying."),("'anstatt' means?","Instead of"),("When can you use ohne…zu?","When both clauses have the same subject"),("Translate: 'Anstatt zu warten, rufe ich an.'","Instead of waiting, I call."),("Where does the infinitive go?","At the end")],
"You can now express 'without doing' and 'instead of doing' in German — an important B1 structure."
))

# ── Day 132 ─────────────────────────────────────────────────────────────────
DAYS.append((132,
"RELATIVE CLAUSES — GENITIVE 🔗",
"dessen / deren",
"Relative clauses in Genitive show possession: der Mann, dessen Auto… (the man whose car…). For feminine/plural use deren: die Frau, deren Tasche… (the woman whose bag…).",
"Genitive relative clauses possession બતાવે: der Mann, dessen Auto… (એ માણસ જેના કાર…). feminine/plural માટે deren: die Frau, deren Tasche… (એ સ્ત્રી જેના બેગ…).",
[
("dessen","whose (m/n)","જેનું","DES-en"),
("deren","whose (f/pl)","જેની/જેના","DEE-ren"),
("der Mann, dessen…","the man whose…","એ માણસ જેના…","dehr MAN DES-en"),
("die Frau, deren…","the woman whose…","એ સ્ત્રી જેની…","dee FROW DEE-ren"),
("das Kind, dessen…","the child whose…","એ બાળક જેના…","das KINT DES-en"),
("die Leute, deren…","the people whose…","એ લોકો જેના…","dee LOY-teh DEE-ren"),
("das Auto","car","કાર","OW-toh"),
("die Tasche","bag","બેગ","TA-sheh"),
("die Firma","company","કંપની","FIR-mah"),
("der Chef","boss","માલિક","SHEF"),
("die Idee","idea","વિચાર","ee-DAY"),
("der Name","name","નામ","NAH-meh"),
("die Tochter","daughter","દીકરી","TOKH-ter"),
("der Sohn","son","દીકરો","ZOHN"),
("berühmt","famous","પ્રસિદ્ધ","beh-RÜMT"),
],
[
("Der Mann, dessen Auto dort steht, ist mein Nachbar.","The man whose car is standing there is my neighbour.","જે માણસની કાર ત્યાં ઊભી છે, તે મારો પાડોશી છે."),
("Die Frau, deren Tasche verloren ist, sucht Hilfe.","The woman whose bag is lost is looking for help.","જે સ્ત્રીનું બેગ ખોવાયું છે, તે મદદ શોધે છે."),
("Das Kind, dessen Vater Arzt ist, ist sehr fleißig.","The child whose father is a doctor is very diligent.","જે બાળકના પિતા ડૉક્ટર છે તે ખૂબ મહેનતી છે."),
("Die Firma, deren Produkte beliebt sind, wächst schnell.","The company whose products are popular grows fast.","જે કંપનીના પ્રોડક્ટ્સ લોકપ્રિય છે તે ઝડપથી વધી રહી છે."),
("Ich kenne einen Mann, dessen Name Rahul ist.","I know a man whose name is Rahul.","હું એક માણસને ઓળખું છું যার નામ Rahul છે."),
],
[
("📌","Form","dessen = masculine/neuter; deren = feminine/plural.","E8F0FE"),
("💡","Comma Rule","Relative clauses always need commas in German.","FFF8E1"),
("⚠️","Case vs Gender","Genitive shows possession; gender decides dessen/deren.","FFE0E0"),
],
["Write 5 sentences using dessen/deren.","Describe 3 people with 'whose' clauses.","Underline the possessed noun in each sentence."],
[("Translate: 'der Mann, dessen Auto…'","the man whose car…"),("'deren' is used for?","feminine and plural"),("Translate: 'die Frau, deren Tasche…'","the woman whose bag…"),("Where do you put commas?","Before and after the relative clause"),("'dessen' is used for?","masculine/neuter")],
"You can now form Genitive relative clauses with dessen/deren — a strong B1 grammar milestone."
))

# ── Day 133 ─────────────────────────────────────────────────────────────────
DAYS.append((133,
"N-DECLENSION NOUNS 🧱",
"der Student, der Kunde, der Mensch",
"Some masculine nouns add -n or -en in all cases except Nominative singular. These are called N-declension nouns: der Student → den Studenten, dem Studenten, des Studenten.",
"કેટલાક masculine nouns Nominative સિવાય બધામાં -n/-en લે છે. આ N-declension nouns છે: der Student → den Studenten, dem Studenten, des Studenten.",
[
("der Student","student","વિદ્યાર્થી","shtoo-DENT"),
("den Studenten","accusative form","વિદ્યાર્થીને","den shtoo-DEN-ten"),
("dem Studenten","dative form","વિદ્યાર્થીને","daym shtoo-DEN-ten"),
("des Studenten","genitive form","વિદ્યાર્થીનો","des shtoo-DEN-ten"),
("der Kunde","customer","ગ્રાહક","KOON-deh"),
("der Mensch","human","માનવ","MENSH"),
("der Junge","boy","છોકરો","YOON-geh"),
("der Nachbar","neighbour","પાડોશી","NAKH-bar"),
("der Kollege","colleague","કોલિગ","ko-LAY-geh"),
("der Name","name","નામ","NAH-meh"),
("der Herr","Mr./gentleman","શ્રી","HEHR"),
("der Löwe","lion","સિંહ","LÖ-veh"),
("der Experte","expert","વિશેષજ્ઞ","eks-PAIR-teh"),
("der Polizist","police officer","પોલીસ","po-li-TSIST"),
("der Präsident","president","રાષ્ટ્રપતિ","prä-zi-DENT"),
],
[
("Ich sehe den Studenten.","I see the student.","હું વિદ્યાર્થીને જોઉં છું."),
("Ich helfe dem Kunden.","I help the customer.","હું ગ્રાહકને મદદ કરું છું."),
("Der Name des Jungen ist Amir.","The boy's name is Amir.","છોકરાનું નામ Amir છે."),
("Ich spreche mit dem Kollegen.","I speak with the colleague.","હું કોલિગ સાથે વાત કરું છું."),
("Der Mensch braucht Respekt.","Humans need respect.","માનવને સન્માન જોઈએ."),
],
[
("📌","Rule","Masculine N-nouns take -n/-en in Acc/Dat/Gen singular.","E8F0FE"),
("💡","Common Groups","People and professions are often N-nouns (Student, Kunde, Kollege, Experte).","FFF8E1"),
("⚠️","Exception","der Name is special: des Namens (adds -ns).","FFE0E0"),
],
["Make a table for der Student in all cases.","Write 5 sentences with N-declension nouns.","Identify 5 N-declension nouns from a text."],
[("Accusative of 'der Student'?","den Studenten"),("Dative of 'der Kunde'?","dem Kunden"),("Genitive of 'der Name'?","des Namens"),("Are N-declension nouns mostly masculine?","Yes"),("Translate: 'Ich helfe dem Studenten.'","I help the student.")],
"You can now handle N-declension nouns correctly — a frequent source of B1 errors."
))

# ── Day 134 ─────────────────────────────────────────────────────────────────
DAYS.append((134,
"ADJECTIVE ENDINGS — STRONG DECLENSION 🧠",
"No article (strong endings)",
"When there is no article, the adjective carries the case and gender information. This is strong declension: guter Wein, kalte Milch, frisches Brot.",
"જ્યારે article નથી, ત્યારે adjective કેસ અને gender બતાવે છે. આ strong declension છે: guter Wein, kalte Milch, frisches Brot.",
[
("guter Wein","good wine","સારો wine","GOO-ter VINE"),
("kalte Milch","cold milk","ઠંડી દૂધ","KAL-teh MILKH"),
("frisches Brot","fresh bread","તાજું બ્રેડ","FRISH-es BROHT"),
("alter Freund","old friend","જૂનો મિત્ર","AL-ter FROYNT"),
("neue Arbeit","new work","નવી કામ","NOY-eh AR-byt"),
("schwerer Test","difficult test","કઠિન ટેસ્ટ","SHVAY-rer TEST"),
("langes Wochenende","long weekend","લાંબું weekend","LANG-es"),
("kleines Problem","small problem","નાનું સમસ્યા","KLY-nes pro-BLAYM"),
("großer Erfolg","big success","મોટી સફળતા","GROH-ser"),
("starker Kaffee","strong coffee","ઘટ્ટ કોફી","SHTAR-ker"),
("schöne Aussicht","beautiful view","સુંદર દૃશ્ય","SHÖ-neh OWSS-zi-kht"),
("sicherer Job","secure job","સુરક્ષિત નોકરી","ZIKH-er-er"),
("billige Lösung","cheap solution","સસતું ઉકેલ","BIL-ig-eh"),
("neuer Plan","new plan","નવો પ્લાન","NOY-er"),
("richtig","correct","સાચું","RIKH-tikh"),
],
[
("Guter Kaffee hilft am Morgen.","Good coffee helps in the morning.","સારી કોફી સવારમાં મદદ કરે છે."),
("Frisches Brot ist in Deutschland wichtig.","Fresh bread is important in Germany.","Germany માં તાજું બ્રેડ મહત્વનું છે."),
("Alter Wein schmeckt besser.","Old wine tastes better.","જૂનો વાઇન વધુ સારું લાગે છે."),
("Kalte Milch steht im Kühlschrank.","Cold milk is in the fridge.","ઠંડું દૂધ ફ્રિજમાં છે."),
("Neuer Plan, neuer Anfang.","New plan, new beginning.","નવો પ્લાન, નવી શરૂઆત."),
],
[
("📌","Rule","No article → adjective shows case/gender: -er (m), -e (f), -es (n), -e (pl).", "E8F0FE"),
("��","Remember","Strong endings look like definite articles without 'd'.", "FFF8E1"),
("⚠️","Common Error","Do not use weak endings without articles.", "FFE0E0"),
],
["Write 5 sentences without articles using strong endings.","Turn 5 sentences with 'der/die/das' into no-article versions.","Create a mini menu: 'guter Wein, frisches Brot…'"],
[("Strong ending for masculine nominative?","-er"),("Translate: 'gutes Wetter'","good weather"),("Strong ending for neuter nominative?","-es"),("Translate: 'schöne Aussicht'","beautiful view"),("When do you use strong declension?","When there is no article")],
"Strong adjective endings mastered — you can now describe things precisely without articles."
))

# ── Day 135 ─────────────────────────────────────────────────────────────────
DAYS.append((135,
"PREPOSITIONAL VERBS I 📍",
"Verbs with fixed prepositions",
"Some verbs require a fixed preposition and case. Example: warten auf + Akkusativ, denken an + Akkusativ, sprechen über + Akkusativ. You must learn them as a set.",
"કેટલાક verbs સાથે fixed preposition અને case આવે છે: warten auf + Akkusativ, denken an + Akkusativ, sprechen über + Akkusativ. આને set તરીકે યાદ રાખો.",
[
("warten auf","to wait for","રાહ જોવી","VAR-ten owf"),
("denken an","to think about","વિચાર કરવો","DEN-ken an"),
("sprechen über","to talk about","વિશે બોલવું","SHPREKH-en Ü-ber"),
("sich interessieren für","to be interested in","રસ રાખવો","in-teh-res-EE-ren für"),
("teilnehmen an","to participate in","ભાગ લેવો","TILE-nay-men an"),
("sich freuen auf","to look forward to","રાહ જોવી","FROY-en owf"),
("sich ärgern über","to be annoyed about","ખફા થવું","ÄR-gern Ü-ber"),
("bitten um","to ask for","માંગવું","BIT-en oom"),
("sorgen für","to take care of","ધ્યાન રાખવું","ZOR-gen für"),
("abhängen von","to depend on","નિર્ભર હોવું","AP-heng-en fon"),
("glauben an","to believe in","વિશ્વાસ રાખવો","GLOW-ben an"),
("vorbereiten auf","to prepare for","તૈયારી કરવી","for-bah-RY-ten owf"),
("sich bewerben um","to apply for","અરજી કરવી","beh-VER-ben oom"),
("träumen von","to dream of","સપના જોવા","TROY-men fon"),
("erkennen an","to recognize by","ઓળખવું","ehr-KEN-en an"),
],
[
("Ich warte auf den Bus.","I am waiting for the bus.","હું બસની રાહ જોઈ રહ્યો છું."),
("Wir sprechen über das Problem.","We are talking about the problem.","અમે સમસ્યા વિશે વાત કરી રહ્યા છીએ."),
("Sie interessiert sich für Kunst.","She is interested in art.","તેને કલા માં રસ છે."),
("Er nimmt an der Prüfung teil.","He participates in the exam.","તે પરીક્ષામાં ભાગ લે છે."),
("Ich freue mich auf das Wochenende.","I look forward to the weekend.","હું weekendની રાહ જોઈ રહ્યો છું."),
],
[
("📌","Learning Tip","Always learn verb + preposition + case together.","E8F0FE"),
("💡","Question Words","Akkusativ: Worauf? Woran? Worüber? | Dativ: Womit? Wovon?", "FFF8E1"),
("⚠️","Common Error","Don't change the preposition — it is fixed with the verb.","FFE0E0"),
],
["Write 10 sentences using prepositional verbs.","Create flashcards: verb + preposition + case.","Ask 5 questions using 'Worauf/Woran/Worüber'."],
[("Translate: 'Ich warte auf dich.'","I am waiting for you."),("'teilnehmen an' takes which case?","Dative"),("'sprechen über' takes which case?","Accusative"),("Translate: 'Sie freut sich auf den Urlaub.'","She looks forward to the holiday."),("'abhängen von' means?","to depend on")],
"You can now use common prepositional verbs — essential for natural B1 German."
))

# ── Day 136 ─────────────────────────────────────────────────────────────────
DAYS.append((136,
"PREPOSITIONAL VERBS II 🧭",
"Verb + preposition in real life",
"Today you extend prepositional verbs to everyday communication: complaints, requests, opinions, and feelings. These verbs appear constantly in German emails and conversations.",
"આજે તમે prepositional verbs ને real-life communication માં વાપરશો: complaints, requests, opinions, feelings. આ verbs emails અને conversations માં બહુ આવે છે.",
[
("sich beschweren über","to complain about","ફરિયાદ કરવી","beh-SHVEHR-en Ü-ber"),
("sich entschuldigen für","to apologize for","માફી માગવી","ent-SHOOL-di-gen für"),
("sich kümmern um","to take care of","ધ્યાન રાખવું","KÜM-mern oom"),
("sich verlassen auf","to rely on","વિશ્વાસ રાખવો","fehr-LAH-sen owf"),
("sich unterhalten über","to discuss","ચર્ચા કરવી","oon-ter-HAL-ten Ü-ber"),
("sich bedanken bei","to thank (someone)","આભાર માનવું","beh-DANK-en bye"),
("sich wundern über","to be surprised about","આશ્ચર્ય થવું","VOON-dern Ü-ber"),
("Angst haben vor","to be afraid of","ડર લાગવો","ANGST HAH-ben for"),
("abhängig von","dependent on","નિર્ભર","AP-hen-gikh fon"),
("bereit für","ready for","તૈયાર","beh-RY-t für"),
("stolz auf","proud of","ગર્વિત","shtolts owf"),
("zufrieden mit","satisfied with","સંતુષ્ટ","tsoo-FREE-den mit"),
("interessiert an","interested in","રસ ધરાવે","in-teh-res-EErt an"),
("denken über","think about","વિચાર કરવો","DEN-ken Ü-ber"),
("sprechen mit","to speak with","સાથે બોલવું","SHPREKH-en mit"),
],
[
("Ich beschwere mich über den Lärm.","I complain about the noise.","હું શોર વિશે ફરિયાદ કરું છું."),
("Ich entschuldige mich für die Verspätung.","I apologize for the delay.","હું મોડા થવા બદલ માફી માંગુ છું."),
("Wir kümmern uns um die Gäste.","We take care of the guests.","અમે મહેમાનોનું ધ્યાન રાખીએ છીએ."),
("Ich verlasse mich auf dich.","I rely on you.","હું તારા પર ભરોસો રાખું છું."),
("Ich bin stolz auf meinen Fortschritt.","I am proud of my progress.","હું મારી પ્રગતિ પર ગર્વ છે."),
],
[
("📌","Case Check","Most 'über' and 'für' use Akkusativ; 'mit' uses Dativ.","E8F0FE"),
("💡","Useful in emails","'Ich entschuldige mich für…' and 'Ich beschwere mich über…' are common formal phrases.","FFF8E1"),
("⚠️","Pronouns","Use the correct pronoun case after prepositions: auf dich, über ihn, mit ihr.","FFE0E0"),
],
["Write 5 complaint sentences using 'sich beschweren über'.","Write 3 apology sentences using 'sich entschuldigen für'.","Create 5 sentences describing feelings with prepositional verbs."],
[("Translate: 'Ich beschwere mich über den Service.'","I complain about the service."),("'sich verlassen auf' means?","to rely on"),("Which case follows 'mit'?","Dative"),("Translate: 'Ich bin stolz auf dich.'","I am proud of you."),("'Angst haben vor' means?","to be afraid of")],
"You can now use advanced prepositional verbs for real-life communication — a true B1 skill."
))

# ── Day 137 ─────────────────────────────────────────────────────────────────
DAYS.append((137,
"WORD ORDER — TEKAMOLO 📐",
"Time – Cause – Manner – Place",
"German word order often follows TeKaMoLo: Temporal (when), Kausal (why), Modal (how), Lokal (where). This makes sentences sound natural and organized.",
"German word order ઘણી વખત TeKaMoLo અનુસરે: Temporal (ક્યારે), Kausal (શા માટે), Modal (કઈ રીતે), Lokal (ક્યાં). આથી વાક્ય natural બને છે.",
[
("Temporal","time","સમય","tem-po-RAHL"),
("Kausal","cause","કારણ","kau-ZAHL"),
("Modal","manner","રીત","mo-DAL"),
("Lokal","place","સ્થાન","lo-KAHL"),
("gestern","yesterday","કાળે","GES-tern"),
("morgen","tomorrow","કાલે","MOR-gen"),
("wegen der Arbeit","because of work","કામના કારણે","VAY-gen dehr AR-byt"),
("mit dem Bus","by bus","બસથી","mit daym BOOS"),
("zu Hause","at home","ઘરે","tsoo HOW-zeh"),
("im Büro","in the office","ઓફિસમાં","im BÜ-roh"),
("schnell","quickly","ઝડપથી","SHNEL"),
("langsam","slowly","ધીમે","LANG-zam"),
("gerne","gladly","ખુશીથી","GER-neh"),
("draußen","outside","બહાર","DROW-sen"),
("drinnen","inside","અંદર","DRIN-en"),
],
[
("Ich arbeite heute wegen der Deadline schnell im Büro.","I work today quickly in the office because of the deadline.","હું આજે deadlineના કારણે ઝડપથી ઓફિસમાં કામ કરું છું."),
("Wir fahren morgen mit dem Zug nach Berlin.","We are traveling tomorrow by train to Berlin.","અમે કાલે ટ્રેનથી Berlin જઈએ છીએ."),
("Er bleibt heute wegen Krankheit zu Hause.","He stays at home today because of illness.","તે બીમારીના કારણે આજે ઘેર રહે છે."),
("Sie lernt abends gerne in der Bibliothek.","She likes to study in the library in the evenings.","તે સાંજે લાઇબ્રેરીમાં ખુશીથી ભણે છે."),
("Wir essen heute langsam im Garten.","We eat slowly in the garden today.","અમે આજે બગીચામાં ધીમે ધીમે ખાઈએ છીએ."),
],
[
("📌","Order","Time → Cause → Manner → Place. (TeKaMoLo)","E8F0FE"),
("💡","Flexibility","Word order can change for emphasis, but TeKaMoLo sounds natural.","FFF8E1"),
("⚠️","Too many details","If the sentence feels heavy, split it into two sentences.","FFE0E0"),
],
["Create 5 sentences using TeKaMoLo order.","Move one element to the front for emphasis in 3 sentences.","Underline Time/Cause/Manner/Place in each sentence."],
[("What does TeKaMoLo stand for?","Time – Cause – Manner – Place"),("Translate: 'Wir fahren morgen mit dem Bus in die Stadt.'","We go to the city tomorrow by bus."),("Where do you place 'weil' clauses?","After the main clause, verb to end"),("Which comes first: Time or Place?","Time"),("Translate: 'Er arbeitet heute im Büro.'","He works in the office today.")],
"You can now build natural German sentences with TeKaMoLo — a core B1 fluency skill."
))

# ── Day 138 ─────────────────────────────────────────────────────────────────
DAYS.append((138,
"TEMPORAL CLAUSES ⏳",
"als, wenn, während, bevor, nachdem",
"Temporal conjunctions describe time relationships. Use 'als' for single past events, 'wenn' for repeated/conditional events, 'während' for simultaneous actions, 'bevor' and 'nachdem' for sequence.",
"Temporal conjunctions સમય સંબંધ બતાવે છે. 'als' single past માટે, 'wenn' repeated/conditional માટે, 'während' simultaneous માટે, 'bevor'/'nachdem' sequence માટે વપરાય છે.",
[
("als","when (single past)","ત્યારે","ALS"),
("wenn","when/if (repeated)","જ્યારે/જો","VEN"),
("während","while","દરમિયાન","VÄH-rent"),
("bevor","before","પહેલાં","beh-FOR"),
("nachdem","after","પછી","nakh-DAYM"),
("immer wenn","whenever","જયારે પણ","IM-er ven"),
("jedes Mal","every time","દર વખતે","YAY-des MAHL"),
("gleich","right away","તુરંત","GLYKH"),
("später","later","પછી","SHPÄ-ter"),
("früher","earlier","પહેલાં","FRÜ-her"),
("als Kind","as a child","બાળપણમાં","als KINT"),
("plötzlich","suddenly","અચાનક","PLÖTS-likh"),
("regelmäßig","regularly","નિયમિત","RAY-gel-mä-sig"),
("die Gewohnheit","habit","આદત","geh-VOHN-hite"),
("der Moment","moment","ક્ષણ","mo-MENT"),
],
[
("Als ich klein war, spielte ich draußen.","When I was малень, I played outside.","જ્યારે હું નાનો હતો, ત્યારે બહાર રમતો હતો."),
("Wenn ich Zeit habe, lese ich.","When I have time, I read.","જ્યારે સમય હોય ત્યારે હું વાંચું છું."),
("Während ich koche, höre ich Musik.","While I cook, I listen to music.","હું રસોઈ કરું ત્યારે સંગીત સાંભળું છું."),
("Bevor wir gehen, trinken wir Kaffee.","Before we go, we drink coffee.","જવા પહેલાં અમે કોફી પીીએ છીએ."),
("Nachdem er angekommen war, rief er mich an.","After he arrived, he called me.","તે પહોંચ્યા પછી મને ફોન કર્યો."),
],
[
("📌","Verb-last","All temporal clauses are subordinate — verb goes to the end.","E8F0FE"),
("💡","als vs wenn","als = one-time past event; wenn = repeated or present/future.","FFF8E1"),
("⚠️","nachdem","After 'nachdem' use past perfect if needed (war angekommen).", "FFE0E0"),
],
["Write 5 sentences using als/wenn/während.","Create a short story using before/after clauses.","Underline the verb in each subordinate clause."],
[("'als' is used for?","Single past event"),("'wenn' is used for?","Repeated events or conditions"),("Translate: 'Während ich lerne, trinke ich Tee.'","While I study, I drink tea."),("Where does the verb go?","At the end"),("Translate: 'Bevor wir gehen, essen wir.'","Before we go, we eat.")],
"You can now express time relationships with temporal clauses — a strong B1 structure for storytelling."
))

# ── Day 139 ─────────────────────────────────────────────────────────────────
DAYS.append((139,
"INDIRECT QUESTIONS ❓",
"Ob / W-Fragen in indirect form",
"Indirect questions are used in polite speech and reporting. Example: Ich weiß nicht, ob er kommt. / Ich frage mich, warum er spät ist. Verb goes to the end.",
"Indirect questions polite speech માં વપરાય છે. ઉદા: Ich weiß nicht, ob er kommt. / Ich frage mich, warum er spät ist. Verb end માં જાય છે.",
[
("ob","whether/if","કે શું","OP"),
("warum","why","શા માટે","VA-rooom"),
("wann","when","ક્યારે","VAN"),
("wie","how","કઈ રીતે","VEE"),
("wo","where","ક્યાં","VOH"),
("wer","who","કોણ","VEHR"),
("was","what","શું","VAS"),
("Ich weiß nicht…","I don't know…","મને ખબર નથી…","ikh VICE nikht"),
("Ich frage mich…","I wonder…","હું વિચારું છું…","ikh FRAH-geh mikh"),
("Können Sie mir sagen…","Can you tell me…","શું તમે મને કહી શકો…","KÖN-en zee"),
("ich möchte wissen…","I would like to know…","હું જાણવા ઈચ્છું છું…","ikh MÖKHT-eh"),
("ob er kommt","whether he comes","કે તે આવે છે કે નહીં","OP ehr KOMT"),
("warum es so ist","why it is so","શા માટે έτσι છે","VA-rooom es zo ist"),
("wann der Zug fährt","when the train leaves","ટ્રેન ક્યારે જાય છે","VAN dehr TSOOG FÄHRT"),
("wie viel es kostet","how much it costs","કેટલી કિંમત છે","VEE feel es KOS-tet"),
],
[
("Ich weiß nicht, ob er heute kommt.","I don't know if he comes today.","મને ખબર નથી કે તે આજે આવે છે કે નહીં."),
("Können Sie mir sagen, wann der Bus fährt?","Can you tell me when the bus leaves?","શું તમે મને કહી શકો બસ ક્યારે જાય છે?"),
("Ich frage mich, warum sie traurig ist.","I wonder why she is sad.","હું વિચારું છું કે તે કેમ દુઃખી છે."),
("Ich möchte wissen, wie viel es kostet.","I would like to know how much it costs.","હું જાણવા ઈચ્છું છું કે કિંમત કેટલી છે."),
("Er erklärt, warum er zu spät war.","He explains why he was late.","તે સમજાવે છે કે તે મોડું કેમ થયું."),
],
[
("📌","Verb-last","Indirect questions are subordinate clauses — verb goes to the end.","E8F0FE"),
("💡","Polite Speech","Indirect questions sound softer and more polite.","FFF8E1"),
("⚠️","Use comma","Always use a comma before the indirect question.","FFE0E0"),
],
["Turn 5 direct questions into indirect questions.","Write 3 polite questions using 'Können Sie mir sagen…?'", "Create 5 sentences using 'ob' and W-words."],
[("Translate: 'Ich weiß nicht, ob er kommt.'","I don't know if he comes."),("Where does the verb go?","At the end"),("'Ich frage mich' means?","I wonder"),("Translate: 'Können Sie mir sagen, wie spät es ist?'","Can you tell me what time it is?"),("What word introduces yes/no indirect questions?","ob")],
"You can now form indirect questions — an essential B1 skill for polite and formal German."
))

# ── Day 140 ─────────────────────────────────────────────────────────────────
DAYS.append((140,
"MODULE REVIEW: DAYS 121–140 ✅",
"B1 Grammar Checkpoint 1",
"Great work! You covered Konjunktiv II, infinitive clauses, passive voice, genitive relatives, N-declension, strong adjectives, prepositional verbs, TeKaMoLo, temporal clauses and indirect questions. Today is your consolidation day.",
"શાબાશ! તમે Konjunktiv II, infinitive clauses, passive voice, genitive relatives, N-declension, strong adjectives, prepositional verbs, TeKaMoLo, temporal clauses અને indirect questions cover કર્યા. આજ review day છે.",
[
("die Zusammenfassung","summary","સારાંશ","tsoo-ZA-men-fas-ung"),
("wiederholen","to review","પુનરાવર્તન","vee-der-HOH-len"),
("überprüfen","to check","તપાસવું","Ü-ber-prü-fen"),
("sicher","certain","વિશ્વાસપૂર્વક","ZIKH-er"),
("der Fehler","mistake","ભૂલ","FAY-ler"),
("korrigieren","to correct","સુધારવું","ko-ri-GHEE-ren"),
("die Struktur","structure","રચના","shtrook-TOOR"),
("die Regel","rule","નિયમ","RAY-gel"),
("die Übung","exercise","અભ્યાસ","Ü-bung"),
("die Aufgabe","task","કાર્ય","OWF-gah-beh"),
("der Fortschritt","progress","પ્રગતિ","FORT-shrit"),
("bereit","ready","તૈયાર","beh-RITE"),
("selbstbewusst","confident","આત્મવિશ્વાસી","zelpst-be-VOOST"),
("wichtig","important","મહત્વનું","VIKH-tikh"),
("stabil","stable","સ્થિર","shtah-BEEL"),
],
[
("Ich wiederhole heute alle Regeln.","I review all rules today.","હું આજે બધા નિયમો ફરીથી કરું છું."),
("Ich fühle mich jetzt sicherer.","I feel more confident now.","હવે હું વધુ આત્મવિશ્વાસી લાગું છું."),
("Fehler sind normal — ich korrigiere sie.","Mistakes are normal — I correct them.","ભૂલો સામાન્ય છે — હું તેને સુધારું છું."),
("Die Struktur ist jetzt klar.","The structure is clear now.","હવે રચના સ્પષ્ટ છે."),
("Ich bin bereit für das nächste Modul.","I am ready for the next module.","હું હવે આગળના module માટે તૈયાર છું."),
],
[
("🏆","Checkpoint","If you can explain each grammar point in one sentence, you are ready to move on.","E8F0FE"),
("💡","Study Tip","Do a mini test: write 10 sentences using 10 different structures.","FFF8E1"),
("🎯","Next","Days 141–160 focus on professional life, work, and society vocabulary.","E6F4EA"),
],
["Write 10 sentences mixing grammar from Days 121–139.","Make a mistake list and correct each one.","Record a 2‑minute summary of the grammar topics in German."],
[("Translate: 'Ich bin bereit für das nächste Modul.'","I am ready for the next module."),("What is TeKaMoLo?","Time–Cause–Manner–Place"),("'dessen/deren' are used for?","Genitive relative clauses"),("Passive formula?","werden + Partizip II"),("Konjunktiv II formula?","würde + infinitive")],
"Checkpoint 1 complete. Your B1 grammar foundation is strong and ready for the next level of vocabulary and real-life topics."
))

print("Days 131-140 appended")

# ── Day 141 ─────────────────────────────────────────────────────────────────
DAYS.append((141,
"JOB APPLICATIONS 📝",
"Bewerbung basics",
"Today you learn the core vocabulary for job applications in Germany: application documents, cover letter, CV, and how to talk about your skills. This is essential for professional integration.",
"આજે તમે Germany માં job application માટેના મુખ્ય શબ્દો શીખશો: application documents, cover letter, CV, અને skills વિશે વાત કરવી.",
[
("die Bewerbung","application","અરજી","beh-VER-bung"),
("der Lebenslauf","CV / resume","સીવી","LAY-bens-lowf"),
("das Anschreiben","cover letter","કવર લેટર","AN-shry-ben"),
("die Stelle","job position","પદ","SHTEL-eh"),
("die Anzeige","job ad","જોબ જાહેરાત","AN-tsy-geh"),
("die Qualifikation","qualification","યોગ્યતા","kva-li-fi-kah-TSYON"),
("die Erfahrung","experience","અનુભવ","ehr-FAH-roong"),
("die Fähigkeit","skill","કુશળતા","FÄH-ig-kite"),
("die Ausbildung","training","ટ્રેનિંગ","OWS-bil-doong"),
("der Abschluss","degree","ડિગ્રી","AP-shloos"),
("die Motivation","motivation","પ્રેરણા","moh-ti-va-TSYON"),
("die Referenz","reference","રેફરન્સ","reh-feh-RENTS"),
("die Frist","deadline","સમયસીમા","FRIST"),
("einreichen","to submit","સબમિટ કરવું","EYN-ry-khen"),
("die Unterschrift","signature","સહી","OON-ter-shrift"),
],
[
("Ich reiche heute meine Bewerbung ein.","I submit my application today.","હું આજે મારી અરજી સબમિટ કરું છું."),
("Mein Lebenslauf ist aktuell.","My CV is up to date.","મારું CV અપડેટ છે."),
("Ich habe viel Erfahrung in der IT.","I have a lot of experience in IT.","મારે IT માં ઘણો અનુભવ છે."),
("Das Anschreiben erklärt meine Motivation.","The cover letter explains my motivation.","કવર લેટર મારી પ્રેરણા સમજાવે છે."),
("Die Frist endet am Freitag.","The deadline ends on Friday.","ડેડલાઇન શુક્રવારે પૂરી થાય છે."),
],
[
("📌","Core Docs","Bewerbung = Anschreiben + Lebenslauf + Zeugnisse.","E8F0FE"),
("💡","Tip","Keep it concise and professional. Germans value clarity.","FFF8E1"),
("🇩🇪","Standard","Applications are usually sent as a PDF by email or portal.","E6F4EA"),
],
["List your 5 strongest skills in German.","Write 3 lines for your motivation letter.","Create a checklist of application documents."],
[("'der Lebenslauf' means?","CV / resume"),("Translate: 'Ich reiche die Bewerbung ein.'","I submit the application."),("'die Frist' means?","Deadline"),("What is 'Anschreiben'?","Cover letter"),("Translate: 'Ich habe viel Erfahrung.'","I have a lot of experience.")],
"You now understand the key vocabulary for German job applications. This is step one toward working in Germany."
))

# ── Day 142 ─────────────────────────────────────────────────────────────────
DAYS.append((142,
"CV & PROFILE LANGUAGE 👔",
"Professional self-presentation",
"Today you learn phrases to describe your profile, strengths, and work style in a CV or LinkedIn profile in German.",
"આજે તમે CV અથવા LinkedIn માં પોતાનું પ્રોફાઇલ અને strengths બતાવવા માટેના phrase શીખશો.",
[
("beruflich","professional","વ્યાવસાયિક","beh-ROOF-likh"),
("zuverlässig","reliable","વિશ્વસનીય","tsoo-VER-likh"),
("teamfähig","team-oriented","ટીમ માટે યોગ્ય","TEAM-fäh-ig"),
("selbstständig","independent","સ્વતંત્ર","zelpst-SHTÄN-dig"),
("kommunikativ","communicative","સંવાદક્ષમ","koh-moo-nee-ka-TEEF"),
("analytisch","analytical","વિશ્લેષણાત્મક","ah-nah-LÜ-tish"),
("zielorientiert","goal-oriented","લક્ષ્યકેન્દ્રિત","TSEEL-or-ee-en-TEERT"),
("belastbar","resilient","દબાણ સહનશીલ","beh-LAST-bar"),
("organisiert","organized","વ્યવસ્થિત","or-ga-ni-ZEERT"),
("engagiert","committed","સમર્પિત","on-gah-ZEERT"),
("verantwortlich","responsible","જવાબદાર","fehr-ANT-vort-likh"),
("Pünktlichkeit","punctuality","સમયપાલન","PÜNKT-likh-kite"),
("Führung","leadership","નેતૃત્વ","FÜ-rung"),
("Teamarbeit","teamwork","ટીમ વર્ક","TEAM-ar-byt"),
("Problemlösung","problem solving","સમસ્યા ઉકેલ","pro-BLAYM-lö-zoong"),
],
[
("Ich bin zuverlässig und teamfähig.","I am reliable and a team player.","હું વિશ્વસનીય અને ટીમ-ઓરિએન્ટેડ છું."),
("Meine Stärken sind Analyse und Organisation.","My strengths are analysis and organization.","મારી strengths analysis અને organization છે."),
("Ich arbeite selbstständig und zielorientiert.","I work independently and goal-oriented.","હું સ્વતંત્ર અને લક્ષ્યકેન્દ્રિત રીતે કામ કરું છું."),
("Pünktlichkeit ist mir sehr wichtig.","Punctuality is very important to me.","સમયપાલન મારા માટે ખૂબ મહત્વનું છે."),
("Ich habe Erfahrung in Teamarbeit.","I have experience in teamwork.","મારે ટીમ વર્કનો અનુભવ છે."),
],
[
("💡","CV Language","Use short, clear adjectives. Germans prefer facts over exaggeration.","FFF8E1"),
("📌","Structure","Start with key strengths, then add experience and education.","E8F0FE"),
("🇩🇪","Tone","Professional tone = formal Sie in letters, but neutral in CV.","E6F4EA"),
],
["Write 6 adjectives describing yourself professionally.","Create a 4-line German profile summary.","Translate 5 strengths from English to German."],
[("'zuverlässig' means?","Reliable"),("Translate: 'Ich arbeite selbstständig.'","I work independently."),("'teamfähig' means?","Team-oriented"),("'belastbar' means?","Resilient / can handle stress"),("Translate: 'Pünktlichkeit ist mir wichtig.'","Punctuality is important to me.")],
"You can now describe your professional profile in German — an essential B1 workplace skill."
))

# ── Day 143 ─────────────────────────────────────────────────────────────────
DAYS.append((143,
"JOB INTERVIEW QUESTIONS 🎤",
"Vorstellungsgespräch",
"Interviews are a key step in getting a job. Today you learn common interview questions and how to answer them confidently in German.",
"Interview job મેળવવાનો મુખ્ય પગથિયો છે. આજે તમે common interview questions અને જવાબ આપવા માટેનું German શીખશો.",
[
("das Vorstellungsgespräch","job interview","ઇન્ટરવ્યુ","for-SHTEL-lungs-geh-SHPRÄKH"),
("Warum möchten Sie hier arbeiten?","Why do you want to work here?","તમે અહીં કેમ કામ કરવું ઇચ્છો?","VA-rooom"),
("Was sind Ihre Stärken?","What are your strengths?","તમારી strengths શું છે?","VAS zint EE-reh"),
("Was sind Ihre Schwächen?","What are your weaknesses?","તમારી weaknesses શું છે?","SHVÄKH-en"),
("Erzählen Sie etwas über себя.","Tell me about yourself.","તમારા વિશે થોડું કહો.","eR-TSEHL-en"),
("Ihre Erfahrung","your experience","તમારો અનુભવ","EE-reh ehr-FAH-roong"),
("der Karriereplan","career plan","કેરિયર પ્લાન","kah-REE-reh-plahn"),
("das Gehalt","salary","પગાર","geh-HALT"),
("verhandeln","to negotiate","નેગોશિએટ કરવું","fehr-HAN-del-n"),
("die Probezeit","probation period","પરીક્ષણ સમય","PRO-beh-tsyte"),
("die Verantwortung","responsibility","જવાબદારી","fehr-ANT-vor-toong"),
("die Erwartung","expectation","અપેક્ષા","ehr-VAR-tung"),
("die Aufgabe","task","કાર્ય","OWF-gah-beh"),
("die Frage","question","પ્રશ્ન","FRAH-geh"),
("die Antwort","answer","જવાબ","ANT-vort"),
],
[
("Warum möchten Sie bei uns arbeiten?","Why do you want to work with us?","તમે અમારા સાથે કેમ કામ કરવું ઇચ્છો?"),
("Meine Stärken sind Zuverlässigkeit und Teamarbeit.","My strengths are reliability and teamwork.","મારી strengths વિશ્વસનીયતા અને ટીમ વર્ક છે."),
("Eine Schwäche ist, dass ich manchmal zu viel Perfektion will.","A weakness is that I sometimes want too much perfection.","એક weakness એ છે કે હું ક્યારેક અતિ perfeсtion માંગું છું."),
("Ich habe drei Jahre Erfahrung in der IT.","I have three years of experience in IT.","મારે IT માં ત્રણ વર્ષનો અનુભવ છે."),
("Welche Aufgaben erwarten mich?","What tasks can I expect?","મને કયા કાર્યોની અપેક્ષા રાખવી?"),
],
[
("💡","Interview Tip","Be honest, structured, and give examples.","FFF8E1"),
("📌","Weakness","Always mention how you improve your weakness.","E8F0FE"),
("🇩🇪","Culture","Punctuality and preparation are critical in German interviews.","E6F4EA"),
],
["Write answers to 5 interview questions.","Practice: 2-minute self-introduction in German.","List 3 questions you would ask the interviewer."],
[("'Vorstellungsgespräch' means?","Job interview"),("Translate: 'Was sind Ihre Stärken?'","What are your strengths?"),("'Probezeit' means?","Probation period"),("Translate: 'Ich habe drei Jahre Erfahrung.'","I have three years of experience."),("What should you do with a weakness?","Explain how you improve it")],
"You can now handle basic job interview questions in German — a key professional milestone."
))

# ── Day 144 ─────────────────────────────────────────────────────────────────
DAYS.append((144,
"WORKPLACE COMMUNICATION 📞",
"Emails, calls, and requests",
"Today you learn office communication phrases: asking for information, giving updates, and requesting help.",
"આજે તમે ઓફિસ communication phrases શીખશો: માહિતી માંગવી, updates આપવી, અને મદદ માંગવી.",
[
("die E-Mail","email","ઈમેઈલ","EE-mail"),
("die Anfrage","request/inquiry","વિનંતી","AN-frah-geh"),
("die Rückmeldung","feedback","પ્રતિસાદ","RÜK-mel-dung"),
("das Update","update","અપડેટ","UP-dayt"),
("der Termin","appointment","મીટિંગ","tehr-MEEN"),
("die Information","information","માહિતી","in-for-ma-TSYON"),
("die Datei","file","ફાઈલ","DAI-tee"),
("die Anlage","attachment","એટેચમેન્ટ","AN-lah-geh"),
("die Frage","question","પ્રશ્ન","FRAH-geh"),
("die Antwort","answer","જવાબ","ANT-vort"),
("bitte prüfen","please check","કૃપા કરી તપાસો","PRÜ-fen"),
("weiterleiten","to forward","ફોરવર્ડ કરવું","VY-ter-ly-ten"),
("bestätigen","to confirm","પુષ્ટિ કરવી","beh-SHTÄ-ti-gen"),
("senden","to send","મોકલવું","ZEN-den"),
("erhalten","to receive","પ્રાપ્ત કરવું","ehr-HAL-ten"),
],
[
("Könnten Sie mir die Datei senden?","Could you send me the file?","શું તમે મને ફાઈલ મોકલી શકો?"),
("Ich habe das Update erhalten.","I have received the update.","મને અપડેટ મળ્યો છે."),
("Bitte prüfen Sie die Anlage.","Please check the attachment.","કૃપા કરી એટેચમેન્ટ તપાસો."),
("Ich leite die Anfrage weiter.","I am forwarding the request.","હું વિનંતી ફોરવર્ડ કરું છું."),
("Vielen Dank für Ihre Rückmeldung.","Thank you for your feedback.","તમારા પ્રતિસાદ માટે આભાર."),
],
[
("💡","Office Tone","Use formal Sie and polite requests in workplace emails.","FFF8E1"),
("📌","Subject Lines","Use clear subjects: 'Betreff: Anfrage zu…'.","E8F0FE"),
("🇩🇪","Speed","Quick responses are appreciated in German offices.","E6F4EA"),
],
["Write a short German email requesting information.","Write a reply confirming receipt of a file.","Practice 5 polite workplace requests."],
[("'die Anlage' means?","Attachment"),("Translate: 'Bitte prüfen Sie die Datei.'","Please check the file."),("'weiterleiten' means?","to forward"),("Translate: 'Ich habe Ihre E-Mail erhalten.'","I received your email."),("Formal thanks in emails?","Vielen Dank für…")],
"You can now handle basic workplace communication in German."
))

# ── Day 145 ─────────────────────────────────────────────────────────────────
DAYS.append((145,
"MEETINGS & PRESENTATIONS 📊",
"Office meetings vocabulary",
"Meetings are part of professional life. Today you learn vocabulary for discussions, agendas, and presenting ideas.",
"મીટિંગ્સ professional life નો ભાગ છે. આજે તમે discussions, agenda અને ideas present કરવા માટે vocab શીખશો.",
[
("das Meeting","meeting","મીટિંગ","MEET-ing"),
("die Agenda","agenda","એજન્ડા","ah-ZHEN-dah"),
("das Protokoll","minutes / protocol","મીટિંગ મિનિટ્સ","pro-to-KOL"),
("der Punkt","agenda item","મુદ્દો","POONKT"),
("die Diskussion","discussion","ચર્ચા","dis-koos-SYON"),
("der Vorschlag","suggestion","સૂચન","FOR-shlahg"),
("die Entscheidung","decision","નિણર્ય","ent-SHAY-dung"),
("die Präsentation","presentation","પ્રેઝન્ટેશન","pre-zen-ta-TSYON"),
("vorstellen","to present","પ્રસ્તુત કરવું","FOR-shtel-en"),
("zusammenfassen","to summarize","સારાંશ કરવું","tsoo-ZAM-en-fas-en"),
("zustimmen","to agree","સહમત થવું","TSOO-shtim-en"),
("ablehnen","to reject","નકારવું","AP-lay-nen"),
("die Deadline","deadline","ડેડલાઇન","DED-lyne"),
("der nächste Schritt","next step","આગળનો પગલું","NÄKHS-teh shrit"),
("die Priorität","priority","પ્રાથમિકતા","pri-o-ri-TÄT"),
],
[
("Die Agenda hat fünf Punkte.","The agenda has five items.","એજન્ડામાં પાંચ મુદ્દા છે."),
("Ich stelle heute die Präsentation vor.","I present the presentation today.","હું આજે પ્રેઝન્ટેશન રજૂ કરું છું."),
("Wir müssen eine Entscheidung treffen.","We have to make a decision.","અમને નિર્ણય લેવો પડશે."),
("Ich stimme dem Vorschlag zu.","I agree with the suggestion.","હું સૂચન સાથે સહમત છું."),
("Der nächste Schritt ist die Umsetzung.","The next step is implementation.","આગળનું પગલું અમલમાં લાવવાનું છે."),
],
[
("💡","Meeting Tip","Use clear, short sentences and confirm decisions.","FFF8E1"),
("📌","Protocol","Protokoll = written meeting summary, very common in Germany.","E8F0FE"),
("🇩🇪","Culture","German meetings are structured and time-focused.","E6F4EA"),
],
["Write a short meeting summary in German.","Create a sample agenda for a 30‑minute meeting.","Practice 3 phrases for agreeing/disagreeing."],
[("'die Agenda' means?","Agenda"),("Translate: 'Wir treffen eine Entscheidung.'","We make a decision."),("'zustimmen' means?","to agree"),("Translate: 'Ich lehne den Vorschlag ab.'","I reject the suggestion."),("'der nächste Schritt' means?","Next step")],
"You can now participate in meetings in German with clear, professional vocabulary."
))

# ── Day 146 ─────────────────────────────────────────────────────────────────
DAYS.append((146,
"PROJECT MANAGEMENT 🧩",
"Planning and deadlines",
"You will often discuss tasks, milestones, and deadlines at work. Today you learn project vocabulary to coordinate and plan in German.",
"કામમાં tasks, milestones અને deadlines વિશે વાત કરવી પડે છે. આજે project vocabulary શીખો.",
[
("das Projekt","project","પ્રોજેક્ટ","pro-YEKT"),
("die Aufgabe","task","કાર્ય","OWF-gah-beh"),
("die Deadline","deadline","ડેડલાઇન","DED-lyne"),
("der Meilenstein","milestone","માઇલસ્ટોન","MY-len-shtyne"),
("der Zeitplan","timeline","સમયપત્રક","TSYTE-plahn"),
("das Ziel","goal","લક્ષ્ય","TSEEL"),
("die Ressource","resource","સ્રોત","reh-ZOOR-seh"),
("das Budget","budget","બજેટ","BOO-zhet"),
("die Priorität","priority","પ્રાથમિકતા","pri-o-ri-TÄT"),
("verschieben","to postpone","મુલતવી કરવું","fehr-SHEE-ben"),
("erreichen","to achieve","મેળવવું","ehr-RY-khen"),
("planen","to plan","યોજના બનાવવી","PLAH-nen"),
("umsetzen","to implement","અમલમાં લાવવું","OOM-zet-sen"),
("verzögert","delayed","મોડું","fehr-TZÖ-gert"),
("pünktlich","on time","સમય પર","PÜNKT-likh"),
],
[
("Das Projekt hat eine Deadline am Freitag.","The project has a deadline on Friday.","પ્રોજેક્ટની ડેડલાઇન શુક્રવારે છે."),
("Wir müssen den Zeitplan anpassen.","We must adjust the timeline.","અમને સમયપત્રક બદલવું પડશે."),
("Der Meilenstein wurde erreicht.","The milestone was reached.","માઇલસ્ટોન પ્રાપ્ત થયો."),
("Das Budget ist begrenzt.","The budget is limited.","બજેટ સીમિત છે."),
("Die Aufgabe wurde um eine Woche verschoben.","The task was postponed by one week.","કાર્ય એક અઠવાડિયા માટે મુલતવી થયું."),
],
[
("💡","Planning","Use clear dates and responsibilities in German projects.","FFF8E1"),
("📌","Status","'verzögert' (delayed) and 'pünktlich' are key status words.","E8F0FE"),
("🇩🇪","Precision","German project communication values precision and transparency.","E6F4EA"),
],
["Create a project plan with 3 milestones in German.","Write 5 sentences about deadlines and delays.","Practice saying dates in German."],
[("'der Meilenstein' means?","Milestone"),("Translate: 'Die Aufgabe ist verzögert.'","The task is delayed."),("'der Zeitplan' means?","Timeline"),("Translate: 'Wir planen das Budget.'","We plan the budget."),("'umsetzen' means?","to implement")],
"You can now talk about projects, deadlines, and planning in German — a key B1 workplace ability."
))

# ── Day 147 ─────────────────────────────────────────────────────────────────
DAYS.append((147,
"CUSTOMER SERVICE 💬",
"Handling clients politely",
"Customer service requires polite, clear language. Today you learn how to greet customers, handle problems, and offer solutions in German.",
"Customer service માં polite અને clear language જરૂરી છે. આજે તમે ગ્રાહકો સાથે વાત, સમસ્યા હેન્ડલ અને solutions આપવાનું German શીખશો.",
[
("der Kunde / die Kundin","customer","ગ્રાહક","KOON-deh"),
("die Anfrage","inquiry","વિનંતી","AN-frah-geh"),
("die Beschwerde","complaint","ફરિયાદ","beh-SHVER-deh"),
("die Lösung","solution","ઉકેલ","LÖ-zoong"),
("die Rückgabe","return","રીટર્ન","RÜK-gah-beh"),
("die Garantie","warranty","ગેરંટી","ga-ran-TEE"),
("der Umtausch","exchange","એક્સચેન્જ","OOM-towsh"),
("zufrieden","satisfied","સંતુષ્ટ","tsoo-FREE-den"),
("unzufrieden","unsatisfied","અસંતુષ્ટ","OON-tsoo-FREE-den"),
("entschuldigen","to apologize","માફી માગવી","ent-SHOOL-di-gen"),
("prüfen","to check","ચકાસવું","PRÜ-fen"),
("bearbeiten","to process","પ્રોસેસ કરવું","beh-AR-by-ten"),
("schnell","quickly","ઝડપથી","SHNEL"),
("gerecht","fair","ન્યાયસંગત","geh-REKHT"),
("die Zufriedenheit","satisfaction","સંતોષ","tsoo-FREE-den-hite"),
],
[
("Guten Tag, wie kann ich Ihnen helfen?","Good day, how can I help you?","નમસ્તે, હું તમને કેવી રીતે મદદ કરી શકું?"),
("Es tut mir leid für die Unannehmlichkeiten.","I am sorry for the inconvenience.","અસુવિધા માટે મને ખેદ છે."),
("Wir prüfen Ihre Anfrage sofort.","We will check your request immediately.","અમે તમારી વિનંતી તરત તપાસીએ છીએ."),
("Der Umtausch ist möglich.","An exchange is possible.","એક્સચેન્જ શક્ય છે."),
("Ihre Zufriedenheit ist uns wichtig.","Your satisfaction is important to us.","તમારો સંતોષ અમારા માટે મહત્વનો છે."),
],
[
("💡","Tone","Polite tone and quick response build trust.","FFF8E1"),
("📌","Key Phrase","'Es tut mir leid' is essential in service German.","E8F0FE"),
("🇩🇪","Culture","German customers expect clear rules and fairness.","E6F4EA"),
],
["Write 5 customer service responses in German.","Create a short script for a return/exchange.","Practice apology phrases for service problems."],
[("Translate: 'Es tut mir leid.'","I am sorry."),("'die Beschwerde' means?","Complaint"),("Translate: 'Wir prüfen das.'","We check it."),("'der Umtausch' means?","Exchange"),("'Ihre Zufriedenheit' means?","Your satisfaction")],
"You can now handle basic customer service situations in German with polite and professional language."
))

# ── Day 148 ─────────────────────────────────────────────────────────────────
DAYS.append((148,
"SALARY & CONTRACTS 💶",
"Gehalt, Vertrag, Arbeitszeit",
"Understanding salary and contract terms is essential for working in Germany. Today you learn contract vocabulary, benefits, and working hours.",
"Germany માં કામ કરવા માટે salary અને contract terms સમજવા જરૂરી છે. આજે તમે contract vocabulary, benefits અને working hours શીખશો.",
[
("das Gehalt","salary","પગાર","geh-HALT"),
("der Vertrag","contract","કરાર","fehr-TRAHG"),
("die Arbeitszeit","working hours","કામનો સમય","AR-byts-tsyte"),
("die Vollzeit","full-time","ફુલ-ટાઇમ","FOL-tsyte"),
("die Teilzeit","part-time","પાર્ટ-ટાઇમ","TILE-tsyte"),
("die Überstunden","overtime","ઓવરટાઇમ","Ü-ber-SHTOON-den"),
("der Urlaub","paid leave","રજા","OOR-lowp"),
("die Probezeit","probation period","પરીક્ષણ સમય","PRO-beh-tsyte"),
("die Kündigung","termination","નોટિસ/છૂટા","KÜN-di-gung"),
("die Frist","notice period","નોટિસ સમય","FRIST"),
("der Bonus","bonus","બોનસ","BOH-noos"),
("die Versicherung","insurance","વિમા","fehr-ZIKH-e-rung"),
("der Tarifvertrag","collective agreement","ટેરિફ કરાર","ta-REEF-fehr-TRAHG"),
("brutto","gross","ગ્રોસ","BROOT-toh"),
("netto","net","નેટ","NET-toh"),
],
[
("Das Gehalt beträgt 3.000 Euro brutto.","The salary is 3,000 euros gross.","પગાર 3000 યુરો બ્રુટો છે."),
("Die Probezeit dauert sechs Monate.","The probation period lasts six months.","પરીક્ષણ સમય છ મહિના છે."),
("Ich arbeite in Vollzeit.","I work full-time.","હું ફુલ-ટાઇમ કામ કરું છું."),
("Der Vertrag enthält 30 Tage Urlaub.","The contract includes 30 days of leave.","કરારમાં 30 દિવસની રજા છે."),
("Die Kündigungsfrist beträgt vier Wochen.","The notice period is four weeks.","નોટિસ સમય ચાર અઠવાડિયા છે."),
],
[
("💡","Brutto vs Netto","Always ask: Is the salary gross (brutto) or net (netto)?",
"FFF8E1"),
("📌","Contracts","Read the contract carefully: hours, leave, and notice period.","E8F0FE"),
("🇩🇪","Legal","Employment contracts are binding; keep a signed copy.","E6F4EA"),
],
["List 5 questions you would ask about a job contract.","Write a sentence explaining your preferred working time.","Translate your current salary in German (if applicable)."],
[("'brutto' means?","Gross"),("Translate: 'Ich arbeite Teilzeit.'","I work part-time."),("'die Kündigung' means?","Termination"),("Translate: 'Der Vertrag enthält Urlaub.'","The contract includes leave."),("'die Überstunden' means?","Overtime")],
"You now understand basic salary and contract vocabulary in German — essential for working in Germany."
))

# ── Day 149 ─────────────────────────────────────────────────────────────────
DAYS.append((149,
"OFFICE TOOLS & IT 🖥️",
"Workplace technology vocabulary",
"Modern offices rely on technology. Today you learn IT and office tool vocabulary to communicate efficiently at work.",
"આજે તમે ઓફિસ IT અને tools સંબંધિત શબ્દો શીખશો જેથી કામમાં સરળતાથી સંચાર કરી શકો.",
[
("der Computer","computer","કમ્પ્યુટર","kom-PYOO-ter"),
("der Bildschirm","screen","સ્ક્રીન","BIL-shirm"),
("die Tastatur","keyboard","કીબોર્ડ","tas-ta-TOOR"),
("die Maus","mouse","માઉસ","MOWS"),
("der Drucker","printer","પ્રિન્ટર","DRÜK-er"),
("die Software","software","સોફ્ટવેર","SOFT-ver"),
("die Hardware","hardware","હાર્ડવેર","HARD-ver"),
("das Passwort","password","પાસવર્ડ","PAS-vort"),
("die Datei","file","ફાઈલ","DAI-tee"),
("speichern","to save","સેવ કરવું","SHPI-khern"),
("öffnen","to open","ખોલવું","ÖF-nen"),
("drucken","to print","પ્રિન્ટ કરવું","DRÜK-en"),
("aktualisieren","to update","અપડેટ કરવું","ak-too-ah-lee-ZEE-ren"),
("die Verbindung","connection","કનેક્શન","fehr-BIN-dung"),
("die Störung","fault/error","તકલીફ","SHTÖ-rung"),
],
[
("Der Drucker funktioniert nicht.","The printer is not working.","પ્રિન્ટર કામ નથી કરતો."),
("Ich speichere die Datei.","I save the file.","હું ફાઈલ સેવ કરું છું."),
("Bitte aktualisieren Sie die Software.","Please update the software.","કૃપા કરીને સોફ્ટવેર અપડેટ કરો."),
("Ich habe mein Passwort vergessen.","I forgot my password.","હું પાસવર્ડ ભૂલી ગયો છું."),
("Es gibt eine Störung im Netzwerk.","There is an error in the network.","નેટવર્કમાં તકલીફ છે."),
],
[
("💡","IT Tip","Learn common error phrases — they appear in tickets and emails.","FFF8E1"),
("📌","Office German","Short commands are common: öffnen, speichern, drucken.","E8F0FE"),
("🇩🇪","Culture","In German offices, clear written instructions are valued.","E6F4EA"),
],
["Write 5 IT problem sentences in German.","Create a mini guide: how to print a document.","Translate 5 tech words from your daily work."],
[("'der Drucker' means?","Printer"),("Translate: 'Ich speichere die Datei.'","I save the file."),("'die Störung' means?","Fault/error"),("Translate: 'Bitte öffnen Sie das Dokument.'","Please open the document."),("'die Verbindung' means?","Connection")],
"You can now speak about office technology and basic IT tasks in German."
))

# ── Day 150 ─────────────────────────────────────────────────────────────────
DAYS.append((150,
"MODULE REVIEW: DAYS 141–150 ✅",
"Work & Professional Communication Review",
"Today you review job applications, interview language, workplace communication, meetings, projects, customer service, contracts, and office IT vocabulary.",
"આજે તમે job applications, interview language, workplace communication, meetings, projects, customer service, contracts અને office IT vocabulary review કરો.",
[
("der Überblick","overview","ઝાંખી","Ü-ber-blik"),
("die Vorbereitung","preparation","તૈયારી","for-bah-RY-tung"),
("die Übung","exercise","અભ્યાસ","Ü-bung"),
("die Wiederholung","review","પુનરાવર્તન","VEE-der-hoh-lung"),
("der Fortschritt","progress","પ્રગતિ","FORT-shrit"),
("die Kompetenz","competence","ક્ષમતા","kom-peh-TENTS"),
("beruflich","professional","વ્યાવસાયિક","beh-ROOF-likh"),
("klar","clear","સ્પષ્ટ","KLAHR"),
("strukturiert","structured","રચનાત્મક","shtrook-too-REERT"),
("sicher","confident","આત્મવિશ્વાસ","ZIKH-er"),
("die Erfahrung","experience","અનુભવ","ehr-FAH-roong"),
("die Anfrage","request","વિનંતી","AN-frah-geh"),
("die Entscheidung","decision","નિણર્ય","ent-SHAY-dung"),
("die Verantwortung","responsibility","જવાબદારી","fehr-ANT-vort-oong"),
("das Ziel","goal","લક્ષ્ય","TSEEL"),
],
[
("Ich fasse die wichtigsten Punkte zusammen.","I summarize the most important points.","હું મહત્વના મુદ્દાઓનો સારાંશ કરું છું."),
("Meine berufliche Sprache ist jetzt klarer.","My professional language is now clearer.","મારી વ્યાવસાયિક ભાષા હવે વધુ સ્પષ્ટ છે."),
("Ich fühle mich sicherer im Vorstellungsgespräch.","I feel more confident in the interview.","હું ઇન્ટરવ્યુમાં વધુ આત્મવિશ્વાસ અનુભવું છું."),
("Ich kann Anfragen professionell beantworten.","I can answer requests professionally.","હું વિનંતીઓને વ્યાવસાયિક રીતે જવાબ આપી શકું છું."),
("Mein nächstes Ziel: flüssig im Job sprechen.","My next goal: speak fluently at work.","મારું આગલું લક્ષ્ય: કામમાં fluent બોલવું."),
],
[
("🏆","Checkpoint","If you can write a short formal email and answer 5 interview questions, you are ready.","E8F0FE"),
("💡","Practice","Record yourself answering interview questions in German.","FFF8E1"),
("🎯","Next Module","Days 151–170 focus on society, media, and advanced daily life.","E6F4EA"),
],
["Write a professional email about a delay.","Answer 5 interview questions in German.","Create a 1-page German CV draft."],
[("Translate: 'Meine berufliche Sprache ist klar.'","My professional language is clear."),("What is 'Vorstellungsgespräch'?","Job interview"),("'die Anfrage' means?","Request/inquiry"),("Translate: 'Ich fasse zusammen.'","I summarize."),("What is the next module about?","Society, media, advanced daily life")],
"Module review complete. Your professional German has leveled up significantly."
))

print("Days 141-150 appended")

# ── Day 151 ─────────────────────────────────────────────────────────────────
DAYS.append((151,
"NEWS & MEDIA 🗞️",
"Nachrichten und Medien",
"Understanding news and media helps you integrate into German society. Today you learn vocabulary for newspapers, TV news, and basic media reporting.",
"સમાચાર અને મીડિયા સમજવાથી German societyમાં જોડાવું સરળ બને છે. આજે newspaper, TV news અને reporting માટે vocab શીખશો.",
[
("die Nachrichten","news","સમાચાર","NAKH-rikh-ten"),
("die Zeitung","newspaper","અખબાર","TSOY-toong"),
("der Artikel","article","લેખ","AR-ti-kel"),
("die Schlagzeile","headline","હેડલાઇન","SHLAHG-tsy-leh"),
("der Bericht","report","રિપોર્ટ","beh-RIKHT"),
("die Meldung","news report","સમાચાર રિપોર્ટ","MEL-dung"),
("das Fernsehen","television","ટીવી","FEHRN-zay-en"),
("das Radio","radio","રેડિયો","RAH-dee-oh"),
("die Quelle","source","સ્રોત","KVEL-leh"),
("die Sendung","broadcast","પ્રસરણ","ZEN-dung"),
("aktuell","current","હાલનું","ak-TOO-el"),
("berichten","to report","રિપોર્ટ કરવું","beh-RIKH-ten"),
("veröffentlichen","to publish","પ્રકાશિત કરવું","fehr-ÖF-fent-likhen"),
("die Meinung","opinion","મત","MY-nung"),
("objektiv","objective","નિષ્પક્ષ","ob-YEK-tiv"),
],
[
("Ich lese jeden Morgen die Nachrichten.","I read the news every morning.","હું દર સવાર સમાચાર વાંચું છું."),
("Die Schlagzeile ist heute sehr wichtig.","The headline is very important today.","આજની હેડલાઇન બહુ મહત્વની છે."),
("Der Bericht wurde gestern veröffentlicht.","The report was published yesterday.","રિપોર્ટ ગઈકાલે પ્રકાશિત થયો હતો."),
("Im Fernsehen gibt es eine Sondersendung.","There is a special broadcast on TV.","ટીવી પર ખાસ કાર્યક્રમ છે."),
("Wir prüfen die Quelle der Meldung.","We check the source of the report.","અમે સમાચારનો સ્રોત તપાસીએ છીએ."),
],
[
("💡","Media Tip","Use multiple sources for balanced information.","FFF8E1"),
("📌","Objectivity","'objektiv' is a key word in German media discussions.","E8F0FE"),
("🇩🇪","Habit","Many Germans watch Tagesschau at 8 pm for daily news.","E6F4EA"),
],
["Write a 5-line summary of a news story in German.","List 5 German media sources you know.","Describe the difference between opinion and report in German."],
[("'die Schlagzeile' means?","Headline"),("Translate: 'Ich lese die Zeitung.'","I read the newspaper."),("'die Quelle' means?","Source"),("Translate: 'Der Bericht wurde veröffentlicht.'","The report was published."),("'objektiv' means?","Objective")],
"You can now follow basic German news and media discussions."
))

# ── Day 152 ─────────────────────────────────────────────────────────────────
DAYS.append((152,
"CIVIC LIFE & POLITICS 🏛️",
"Democracy vocabulary",
"Today you learn basic civic and political vocabulary to understand discussions about society, elections, and government.",
"આજે તમે સમાજ, ચૂંટણી અને સરકાર વિશે વાત સમજવા માટે civic/political vocab શીખશો.",
[
("die Regierung","government","સરકાર","reh-GEE-roong"),
("die Wahl","election","ચૂંટણી","VAHL"),
("der Bürger","citizen","નાગરિક","BÜR-ger"),
("die Demokratie","democracy","લોકશાહી","deh-moh-kra-TEE"),
("das Gesetz","law","કાયદો","geh-ZETS"),
("die Partei","party","પાર્ટી","par-TAI"),
("der Politiker","politician","રાજકારણી","po-li-TEE-ker"),
("der Staat","state","રાજ્ય","SHTAHT"),
("die Meinung","opinion","મત","MY-nung"),
("die Debatte","debate","ચર્ચા","deh-BAH-teh"),
("die Freiheit","freedom","સ્વતંત્રતા","FRY-hite"),
("die Verantwortung","responsibility","જવાબદારી","fehr-ANT-vort-oong"),
("das Recht","right","હક","REKHT"),
("die Stimme","vote","મત","SHTIM-meh"),
("wählen","to vote","મતદાન કરવું","VÄH-len"),
],
[
("Die Bürger wählen eine neue Regierung.","Citizens elect a new government.","નાગરિકો નવી સરકાર ચૂંટે છે."),
("Demokratie braucht Verantwortung.","Democracy needs responsibility.","લોકશાહી માટે જવાબદારી જરૂરી છે."),
("Das Gesetz gilt für alle.","The law applies to everyone.","કાયદો દરેક પર લાગુ પડે છે."),
("Wir diskutieren über politische Themen.","We discuss political topics.","અમે રાજકીય મુદ્દાઓ પર ચર્ચા કરીએ છીએ."),
("Meine Stimme ist wichtig.","My vote is important.","મારો મત મહત્વનો છે."),
],
[
("💡","Civic Tip","Learn neutral, respectful language for political discussions.","FFF8E1"),
("📌","Key Verb","wählen = to vote, wählen lassen = to get elected.","E8F0FE"),
("🇩🇪","Germany","Germany is a parliamentary democracy with federal states (Bundesländer).","E6F4EA"),
],
["Write 5 sentences about democracy in German.","List 3 rights you consider important.","Practice explaining your opinion on an election result."],
[("'die Regierung' means?","Government"),("Translate: 'Ich gehe wählen.'","I go to vote."),("'die Freiheit' means?","Freedom"),("Translate: 'Das Gesetz gilt für alle.'","The law applies to everyone."),("'die Debatte' means?","Debate")],
"You now have the basic vocabulary to understand civic and political discussions in German."
))

# ── Day 153 ─────────────────────────────────────────────────────────────────
DAYS.append((153,
"ENVIRONMENT ADVANCED 🌱",
"Climate, energy, sustainability",
"Today you expand your environmental vocabulary to talk about climate policy, renewable energy, and sustainability in Germany.",
"આજે તમે climate policy, renewable energy અને sustainability વિશે વાત કરવા માટે vocab વિસ્તારો.",
[
("der Klimaschutz","climate protection","જલવાયુ સંરક્ષણ","KLEE-ma-shoots"),
("die Nachhaltigkeit","sustainability","ટકાઉપણું","NAKH-hal-tikh-kite"),
("erneuerbar","renewable","નવસર્જનશીલ","ehr-NOY-er-bar"),
("die Windenergie","wind energy","પવન ઊર્જા","VINT-eh-ner-ghee"),
("die Solarenergie","solar energy","સૌર ઊર્જા","zo-LAR-eh-ner-ghee"),
("die Emission","emission","ઉત્સર્જન","eh-MIS-yon"),
("der CO₂-Ausstoß","CO2 emissions","CO₂ ઉત્સર્જન","koo-tsvay OWSS-shtohs"),
("die Förderung","subsidy/support","પ્રોત્સાહન","FÖR-deh-rung"),
("das Recycling","recycling","રીસાયક્લિંગ","ree-SAI-kling"),
("die Müllvermeidung","waste reduction","કચરો ઓછો કરવો","MÜL-fehr-MY-doong"),
("die Umweltpolitik","environmental policy","પર્યાવરણ નીતિ","OOM-velt-po-li-TEEK"),
("die Ressource","resource","સ્રોત","reh-ZOOR-seh"),
("sparsam","economical","મિતવ્યયી","SHPAR-zam"),
("bewusst","conscious","જાગૃત","beh-VOOST"),
("die Zukunft","future","ભવિષ્ય","TSOO-koonft"),
],
[
("Klimaschutz ist eine große Aufgabe.","Climate protection is a big task.","જલવાયુ સંરક્ષણ મોટું કામ છે."),
("Deutschland investiert in erneuerbare Energien.","Germany invests in renewable energy.","Germany નવસર્જનશીલ ઊર્જામાં રોકાણ કરે છે."),
("Wir müssen CO₂-Emissionen reduzieren.","We must reduce CO₂ emissions.","અમે CO₂ ઉત્સર્જન ઘટાડવું જોઈએ."),
("Recycling spart Ressourcen.","Recycling saves resources.","રીસાયક્લિંગ સ્રોત બચાવે છે."),
("Nachhaltigkeit ist wichtig für die Zukunft.","Sustainability is important for the future.","ટકાઉપણું ભવિષ્ય માટે મહત્વનું છે."),
],
[
("💡","Key Phrase","'Wir müssen Emissionen reduzieren' is common in German debates.","FFF8E1"),
("📌","Vocabulary","erneuerbar = renewable, nachhaltig = sustainable.","E8F0FE"),
("🇩🇪","Context","Germany’s Energiewende is a national energy transition strategy.","E6F4EA"),
],
["Write 5 sentences about sustainability in German.","List 5 eco-friendly actions you do.","Explain what Energiewende means in your own words."],
[("'erneuerbar' means?","Renewable"),("Translate: 'Wir müssen Emissionen reduzieren.'","We must reduce emissions."),("'die Nachhaltigkeit' means?","Sustainability"),("Translate: 'Recycling spart Ressourcen.'","Recycling saves resources."),("'die Umweltpolitik' means?","Environmental policy")],
"You can now discuss climate and sustainability topics — a major part of modern German society."
))

# ── Day 154 ─────────────────────────────────────────────────────────────────
DAYS.append((154,
"SOCIAL MEDIA & DIGITAL LIFE 📱",
"Online communication",
"Today you learn vocabulary for social media, online communication, and digital habits — useful for modern life in Germany.",
"આજે તમે social media, online communication અને digital habits માટે vocab શીખશો.",
[
("die sozialen Medien","social media","સોશિયલ મીડિયા","zo-TSYAL-en MAY-dee-en"),
("der Account","account","એકાઉન્ટ","ak-KOWNT"),
("das Profil","profile","પ્રોફાઇલ","pro-FEEL"),
("posten","to post","પોસ્ટ કરવું","POHS-ten"),
("teilen","to share","શેર કરવું","TY-len"),
("liken","to like","લાઈક કરવું","LY-ken"),
("kommentieren","to comment","કમેન્ટ કરવું","kom-men-TEE-ren"),
("folgen","to follow","ફોલો કરવું","FOL-gen"),
("die Nachricht","message","સંદેશ","NAKH-rikh"),
("die Privatsphäre","privacy","ગોપનીયતા","pree-VAHT-sfä-reh"),
("die Sicherheit","security","સુરક્ષા","ZIKH-er-hite"),
("das Passwort","password","પાસવર્ડ","PAS-vort"),
("hochladen","to upload","અપલોડ કરવું","HOKH-lah-den"),
("herunterladen","to download","ડાઉનલોડ કરવું","heh-ROON-ter-lah-den"),
("die Zeitverschwendung","time waste","સમય બગાડ","TSYTE-fer-SHVEN-dung"),
],
[
("Ich teile ein Foto auf Instagram.","I share a photo on Instagram.","હું Instagram પર ફોટો શેર કરું છું."),
("Bitte ändere dein Passwort.","Please change your password.","કૃપા કરીને તમારો પાસવર્ડ બદલો."),
("Soziale Medien können Zeitverschwendung sein.","Social media can be a waste of time.","સોશિયલ મીડિયા સમય બગાડ હોઈ શકે છે."),
("Ich achte auf meine Privatsphäre.","I pay attention to my privacy.","હું મારી ગોપનીયતા વિશે સાવચેત છું."),
("Er lädt ein Video hoch.","He uploads a video.","તે વિડિઓ અપલોડ કરે છે."),
],
[
("💡","Digital Tip","Use strong passwords and two-factor authentication.","FFF8E1"),
("📌","Language","German uses many English tech words but conjugates them.","E8F0FE"),
("🇩🇪","Balance","Digital wellbeing is a common topic in Germany.","E6F4EA"),
],
["Write 5 sentences about your social media habits.","Create 3 safety rules for online life in German.","Write a short message to a friend via WhatsApp in German."],
[("'die Privatsphäre' means?","Privacy"),("Translate: 'Ich poste ein Foto.'","I post a photo."),("'hochladen' means?","To upload"),("Translate: 'Bitte teile das nicht.'","Please don't share that."),("'die Zeitverschwendung' means?","Waste of time")],
"You can now discuss digital life and social media in German."
))

# ── Day 155 ─────────────────────────────────────────────────────────────────
DAYS.append((155,
"EDUCATION & UNIVERSITY 🎓",
"Studying at a higher level",
"Today you expand academic vocabulary: lectures, seminars, research, and academic tasks — useful for university life in Germany.",
"આજે તમે academic vocabulary શીખશો: lectures, seminars, research અને academic tasks.",
[
("die Vorlesung","lecture","લેક્ટર","FOR-lay-zoong"),
("das Seminar","seminar","સેમિનાર","zeh-mee-NAR"),
("die Forschung","research","શોધ","FOR-shoong"),
("die These","thesis","થિસિસ","TAY-zeh"),
("die Prüfung","exam","પરીક્ષા","PRÜ-fung"),
("die Note","grade","ગ્રેડ","NOH-teh"),
("die Bibliothek","library","લાઇબ્રેરી","bib-lee-oh-TEHK"),
("der Dozent","lecturer","ડોઝન્ટ","do-TSEHNT"),
("die Abgabe","submission","સબમિશન","AP-gah-beh"),
("die Frist","deadline","ડેડલાઇન","FRIST"),
("das Modul","module","મોડ્યુલ","MOH-dool"),
("die Hausarbeit","term paper","ટર્મ પેપર","HOWS-ar-byt"),
("das Praktikum","internship","ઇન્ટર્નશિપ","prak-TEE-koom"),
("die Teilnahme","participation","ભાગ લેવું","TILE-nah-meh"),
("die Anwesenheit","attendance","હાજરી","AN-vay-zen-hite"),
],
[
("Ich besuche eine Vorlesung in Informatik.","I attend a lecture in computer science.","હું કમ્પ્યુટર સાયન્સમાં લેક્ટર સાંભળું છું."),
("Die Hausarbeit muss bis Montag abgegeben werden.","The term paper must be submitted by Monday.","ટર્મ પેપર સોમવાર સુધીમાં આપવું છે."),
("Die Note war sehr gut.","The grade was very good.","ગ્રેડ ખૂબ સારો હતો."),
("Ich mache ein Praktikum im Sommer.","I do an internship in the summer.","હું ઉનાળામાં ઇન્ટર્નશિપ કરું છું."),
("Die Anwesenheit im Seminar ist wichtig.","Attendance in the seminar is important.","સેમિનારમાં હાજરી મહત્વની છે."),
],
[
("💡","Academic Tip","Keep track of deadlines (Fristen) and submission rules.","FFF8E1"),
("📌","University","German universities often have many seminars and group projects.","E8F0FE"),
("🇩🇪","Culture","Participation and punctuality are expected in classes.","E6F4EA"),
],
["Write a short plan for your study week in German.","List 5 academic tasks you do regularly.","Describe your favourite subject in German."],
[("'die Vorlesung' means?","Lecture"),("Translate: 'Ich muss die Hausarbeit abgeben.'","I must submit the term paper."),("'das Praktikum' means?","Internship"),("Translate: 'Die Teilnahme ist wichtig.'","Participation is important."),("'die Anwesenheit' means?","Attendance")],
"You can now talk about university life in German — a key skill for students in Germany."
))

# ── Day 156 ─────────────────────────────────────────────────────────────────
DAYS.append((156,
"HOUSING & RENTING 🏠",
"Contracts, rent, and repairs",
"Housing is a critical topic in Germany. Today you learn vocabulary for renting, contracts, and apartment issues.",
"Germany માં housing મહત્વપૂર્ણ છે. આજે renting, contracts અને apartment issues માટે vocab શીખશો.",
[
("die Miete","rent","ભાડું","MEE-teh"),
("der Vermieter","landlord","મકાનમાલિક","fehr-MEE-ter"),
("der Mieter","tenant","ભાડૂત","MEE-ter"),
("die Nebenkosten","additional costs","અતિરિક્ત ખર્ચ","NAY-ben-kos-ten"),
("die Kaution","deposit","ડિપોઝિટ","KOW-tsyohn"),
("der Mietvertrag","rental contract","ભાડા કરાર","MEET-fehr-trahg"),
("die Wohnung","apartment","ફ્લેટ","VOH-nung"),
("die Reparatur","repair","મરામત","reh-pah-RA-TOOR"),
("der Schaden","damage","નુકસાન","SHAH-den"),
("der Hausmeister","caretaker","હાઉસમાસ્ટર","HOWS-my-ster"),
("kündigen","to terminate","નોટિસ આપવી","KÜN-di-gen"),
("die Kündigungsfrist","notice period","નોટિસ સમય","KÜN-di-gungs-frist"),
("die Besichtigung","viewing","વિઝિટ","beh-ZIKH-ti-gung"),
("möbliert","furnished","ફર્નિશ્ડ","MÖ-bliert"),
("unmöbliert","unfurnished","અનફર્નિશ્ડ","oon-MÖ-bliert"),
],
[
("Die Miete beträgt 800 Euro kalt.","The rent is 800 euros cold (without utilities).","ભાડું 800 યુરો (કોલ્ડ) છે."),
("Der Vermieter repariert die Heizung.","The landlord repairs the heating.","મકાનમાલિક હીટિંગ ઠીક કરે છે."),
("Die Nebenkosten sind im Vertrag erklärt.","The additional costs are explained in the contract.","અતિરિક્ત ખર્ચ કરારમાં લખેલા છે."),
("Ich habe einen Schaden in der Küche.","I have damage in the kitchen.","મારા રસોડામાં નુકસાન છે."),
("Die Kaution wird am Anfang bezahlt.","The deposit is paid at the beginning.","ડિપોઝિટ શરૂઆતમાં ચૂકવાય છે."),
],
[
("💡","Cold vs Warm","Kaltmiete = rent without utilities; Warmmiete = including utilities.","FFF8E1"),
("📌","Repairs","Report repairs to the landlord or Hausmeister in writing.","E8F0FE"),
("🇩🇪","Market","Housing is competitive in big cities — apply quickly.","E6F4EA"),
],
["Write a short email reporting a repair problem.","List 5 questions for apartment viewing.","Explain the difference between Kaltmiete and Warmmiete."],
[("'die Kaution' means?","Deposit"),("Translate: 'Ich kündige den Mietvertrag.'","I terminate the rental contract."),("'die Nebenkosten' means?","Additional costs"),("Translate: 'Der Schaden ist groß.'","The damage is big."),("'möbliert' means?","Furnished")],
"You can now discuss renting and housing issues in German — essential for life in Germany."
))

# ── Day 157 ─────────────────────────────────────────────────────────────────
DAYS.append((157,
"CITY ADMINISTRATION 🏛️",
"Anmeldung, Bürgeramt, documents",
"Germany has many administrative steps (Anmeldung, visas, permits). Today you learn vocabulary for city offices and official documents.",
"Germany માં Anmeldung અને permits જેવી ઘણી પ્રક્રિયાઓ છે. આજે city offices અને documents માટે vocab શીખશો.",
[
("das Bürgeramt","citizen office","બ્યુર્ગરઅમ્ટ","BÜR-ger-amt"),
("die Anmeldung","registration","રજીસ્ટ્રેશન","AN-mel-doong"),
("die Meldebescheinigung","registration certificate","રજીસ્ટ્રેશન પત્ર","MEL-deh-be-shy-ni-gung"),
("der Termin","appointment","નિયુક્તિ","tehr-MEEN"),
("der Ausweis","ID card","ઓળખ પત્ર","OWS-vise"),
("der Reisepass","passport","પાસપોર્ટ","RY-zeh-pass"),
("das Formular","form","ફોર્મ","for-MOO-lar"),
("die Gebühr","fee","ફી","geh-BÜR"),
("die Adresse","address","સરનામું","ah-DRES-eh"),
("die Aufenthaltserlaubnis","residence permit","રેસિડન્સ પરમિટ","OWF-ent-halts-er-LOWB-nis"),
("die Bescheinigung","certificate","પ્રમાણપત્ર","beh-SHY-ni-gung"),
("abholen","to collect","લઈ જવું","AP-ho-len"),
("einreichen","to submit","જમા કરવું","EYN-ry-khen"),
("unterschreiben","to sign","સહી કરવી","OON-ter-shry-ben"),
("bearbeiten","to process","પ્રોસેસ કરવું","beh-AR-by-ten"),
],
[
("Ich habe einen Termin beim Bürgeramt.","I have an appointment at the citizen office.","મારી Bürgeramt માં appointment છે."),
("Ich muss mich heute anmelden.","I must register today.","હું આજે Anmeldung કરવો પડશે."),
("Bitte unterschreiben Sie hier.","Please sign here.","કૃપા કરીને અહીં સહી કરો."),
("Die Gebühr beträgt 15 Euro.","The fee is 15 euros.","ફી 15 યુરો છે."),
("Die Aufenthaltserlaubnis ist gültig.","The residence permit is valid.","રેસિડન્સ પરમિટ માન્ય છે."),
],
[
("💡","Preparation","Bring passport, rental contract, and forms to the Bürgeramt.","FFF8E1"),
("📌","Appointments","Many offices require online appointments — book early.","E8F0FE"),
("🇩🇪","Formality","Always use formal Sie at offices.","E6F4EA"),
],
["Write a short checklist for Anmeldung.","Create 5 sentences for the Bürgeramt conversation.","Practice asking for help with a form."],
[("'das Bürgeramt' means?","Citizen office"),("Translate: 'Ich muss mich anmelden.'","I must register."),("'die Gebühr' means?","Fee"),("Translate: 'Bitte unterschreiben Sie.'","Please sign."),("'die Aufenthaltserlaubnis' means?","Residence permit")],
"You can now navigate German city administration vocabulary with confidence."
))

# ── Day 158 ─────────────────────────────────────────────────────────────────
DAYS.append((158,
"LONG‑DISTANCE TRAVEL 🚆",
"Intercity trains, tickets, delays",
"Today you learn advanced travel vocabulary for long-distance trains, reservations, cancellations, and refunds.",
"આજે તમે લાંબી મુસાફરી માટે trains, reservations, cancellations અને refunds અંગે vocab શીખશો.",
[
("die Reservierung","reservation","બુકિંગ","reh-zehr-VEE-roong"),
("die Sitzplatzreservierung","seat reservation","સીટ બુકિંગ","SITS-plats-reh-zehr-VEE-roong"),
("der Fahrplan","timetable","સમયપત્રક","FAR-plahn"),
("die Verspätung","delay","મોડું","fehr-SHPÄ-tung"),
("der Anschluss","connection","કનેક્શન","AN-shloos"),
("der Ausfall","cancellation","રદ થવું","OWS-fal"),
("die Erstattung","refund","રિફંડ","ehr-SHTAT-tung"),
("die Rückfahrt","return trip","રિટર્ન ટ્રિપ","RÜK-fart"),
("die Hinfahrt","outbound trip","આગળની મુસાફરી","HIN-fart"),
("umsteigen","to change (trains)","ટ્રેન બદલી","OOM-shty-gen"),
("der Bahnsteig","platform","પ્લેટફોર્મ","BAHN-shtyge"),
("pünktlich","on time","સમય પર","PÜNKT-likh"),
("die Durchsage","announcement","જાહેરાત","DOORKH-zah-geh"),
("das Ticket","ticket","ટિકિટ","TIK-et"),
("gültig","valid","માન્ય","GÜL-tikh"),
],
[
("Mein Zug hat 20 Minuten Verspätung.","My train is 20 minutes late.","મારી ટ્રેન 20 મિનિટ મોડું છે."),
("Ich brauche eine Sitzplatzreservierung.","I need a seat reservation.","મને સીટ રિઝર્વેશન જોઈએ."),
("Der Anschlusszug fällt aus.","The connecting train is cancelled.","કનેક્શન ટ્રેન રદ થઈ ગઈ છે."),
("Kann ich eine Erstattung bekommen?","Can I get a refund?","શું મને રિફંડ મળશે?"),
("Der Bahnsteig steht auf der Anzeige.","The platform is shown on the display.","પ્લેટફોર્મ ડિસ્પ્લે પર દેખાય છે."),
],
[
("💡","DB Tip","Keep the ticket and delay info for refunds.","FFF8E1"),
("📌","Phrases","'Zug fällt aus' = train canceled. 'Verspätung' = delay.","E8F0FE"),
("🇩🇪","Travel","ICE trains often require seat reservations for comfort.","E6F4EA"),
],
["Write a short complaint about a delayed train.","Practice 5 travel questions at a station.","Explain the difference between Hin- und Rückfahrt."],
[("'die Verspätung' means?","Delay"),("Translate: 'Der Zug fällt aus.'","The train is cancelled."),("'die Erstattung' means?","Refund"),("Translate: 'Ich brauche eine Reservierung.'","I need a reservation."),("'gültig' means?","Valid")],
"You can now manage long-distance travel issues in German with confidence."
))

# ── Day 159 ─────────────────────────────────────────────────────────────────
DAYS.append((159,
"FINANCE & INSURANCE 💳",
"Banking and personal finance",
"Today you expand financial vocabulary: account statements, transfers, and insurance terms.",
"આજે તમે finance vocabulary વિસ્તારો: account statement, transfers, insurance terms.",
[
("das Konto","account","એકાઉન્ટ","KON-toh"),
("der Kontostand","balance","બેલેન્સ","KON-toh-shtant"),
("der Kontoauszug","statement","સ્ટેટમેન્ટ","KON-toh-ows-tsoog"),
("überweisen","to transfer","ટ્રાન્સફર કરવું","Ü-ber-vye-zen"),
("die Überweisung","transfer","ટ્રાન્સફર","Ü-ber-vye-zung"),
("die Abbuchung","debit","ડેબિટ","AP-book-ung"),
("die Gutschrift","credit","ક્રેડિટ","GOOT-shrift"),
("die Versicherung","insurance","વીમા","fehr-ZIKH-er-ung"),
("die Haftpflicht","liability insurance","લાયબિલિટી ઇન્સ્યોરન્સ","HAFT-flikht"),
("die Krankenversicherung","health insurance","હેલ્થ ઇન્સ્યોરન્સ","KRANK-en-fehr-ZIKH-er-ung"),
("die Prämie","premium","પ્રિમિયમ","PRÄ-mee-eh"),
("der Vertrag","contract","કરાર","fehr-TRAHG"),
("kündigen","to cancel","કૅન્સલ કરવું","KÜN-di-gen"),
("die Laufzeit","duration","સમયગાળો","LOWF-tsyte"),
("die Gebühr","fee","ફી","geh-BÜR"),
],
[
("Ich überprüfe meinen Kontoauszug.","I check my account statement.","હું મારો account statement તપાસું છું."),
("Die Überweisung dauert zwei Tage.","The transfer takes two days.","ટ્રાન્સફર માટે બે દિવસ લાગે છે."),
("Die Krankenversicherung ist verpflichtend.","Health insurance is mandatory.","હેલ્થ ઇન્સ્યોરન્સ ફરજિયાત છે."),
("Die Prämie ist dieses Jahr gestiegen.","The premium increased this year.","આ વર્ષે પ્રિમિયમ વધ્યું છે."),
("Ich möchte den Vertrag kündigen.","I want to cancel the contract.","હું કરાર રદ કરવું ઈચ્છું છું."),
],
[
("💡","Finance Tip","Check your Kontoauszug regularly for fees and errors.","FFF8E1"),
("📌","Insurance","In Germany, health insurance is mandatory; liability insurance is strongly recommended.","E8F0FE"),
("🇩🇪","Banking","Most transfers use IBAN and take 1–2 days.","E6F4EA"),
],
["Write 5 sentences about your banking activities.","List 3 types of insurance in German.","Practice asking a bank to explain a fee."],
[("'Kontoauszug' means?","Account statement"),("Translate: 'Ich überweise 100 Euro.'","I transfer 100 euros."),("'die Prämie' means?","Premium"),("Translate: 'Die Krankenversicherung ist verpflichtend.'","Health insurance is mandatory."),("'kündigen' means?","To cancel")],
"You can now discuss financial and insurance topics in German."
))

# ── Day 160 ─────────────────────────────────────────────────────────────────
DAYS.append((160,
"MODULE REVIEW: DAYS 151–160 ✅",
"Society & Media Review",
"You covered news, politics, environment, digital life, education, housing, administration, travel, and finance. Today is review and consolidation.",
"તમે news, politics, environment, digital life, education, housing, administration, travel અને finance cover કર્યા. આજે review day છે.",
[
("die Wiederholung","review","પુનરાવર્તન","VEE-der-hoh-lung"),
("zusammenfassen","summarize","સારાંશ કરવું","tsoo-ZAM-en-fas-en"),
("die Struktur","structure","રચના","shtrook-TOOR"),
("die Information","information","માહિતી","in-for-ma-TSYON"),
("die Gesellschaft","society","સમાજ","geh-ZELL-shaft"),
("die Medien","media","મીડિયા","MAY-dee-en"),
("die Verwaltung","administration","પ્રશાસન","fehr-VAL-tung"),
("die Umwelt","environment","પર્યાવરણ","OOM-velt"),
("die Finanzierung","financing","વિત્તીય","fi-nan-TSI-rung"),
("die Mobilität","mobility","મોબિલિટી","mo-bi-li-TÄT"),
("die Bildung","education","શિક્ષણ","BIL-dung"),
("die Wohnung","apartment","ફ્લેટ","VOH-nung"),
("die Nachricht","message/news","સમાચાર","NAKH-rikh"),
("die Debatte","debate","ચર્ચા","deh-BAH-teh"),
("die Entscheidung","decision","નિણર્ય","ent-SHAY-dung"),
],
[
("Ich fasse die Themen dieser Woche zusammen.","I summarize the topics of this week.","હું આ અઠવાડિયાના વિષયોનું સારાંશ કરું છું."),
("Ich kann Nachrichten besser verstehen.","I can understand news better.","હું હવે સમાચાર વધુ સારી રીતે સમજું છું."),
("Finanzierung und Verträge sind klarer geworden.","Financing and contracts have become clearer.","વિત્ત અને કરારો હવે વધુ સ્પષ્ટ છે."),
("Ich kann über Umwelt und Politik sprechen.","I can talk about environment and politics.","હું પર્યાવરણ અને રાજনীতি વિશે બોલી શકું છું."),
("Ich bin bereit für das nächste Modul.","I am ready for the next module.","હું આગળના module માટે તૈયાર છું."),
],
[
("🏆","Checkpoint","If you can explain a news item in German, you are ready to move on.","E8F0FE"),
("💡","Practice","Watch a short German news video and write a summary.","FFF8E1"),
("🎯","Next","Days 161–180 focus on health, wellbeing, and social life.","E6F4EA"),
],
["Write a 6‑sentence summary of a news story.","List 10 words you learned in this module.","Record a 1‑minute speech about sustainability."],
[("Translate: 'Ich bin bereit für das nächste Modul.'","I am ready for the next module."),("'die Verwaltung' means?","Administration"),("'die Medien' means?","Media"),("Translate: 'Ich kann Nachrichten verstehen.'","I can understand news."),("What is next module about?","Health, wellbeing, social life")],
"Module review complete. Your vocabulary for society and media is solid."
))

print("Days 151-160 appended")

# ── Day 161 ─────────────────────────────────────────────────────────────────
DAYS.append((161,
"HEALTHCARE SYSTEM 🩺",
"Doctor visits, insurance, appointments",
"Today you learn advanced healthcare vocabulary: specialist doctors, referrals, and health insurance procedures in Germany.",
"આજે તમે advanced healthcare vocabulary શીખશો: specialist doctors, referrals અને health insurance પ્રક્રિયા.",
[
("der Facharzt","specialist doctor","સ્પેશિયાલિસ્ટ","FAKH-arzt"),
("die Überweisung","referral","રિફરલ","Ü-ber-vye-zung"),
("die Krankenkasse","health insurance fund","હેલ્થ ઇન્સ્યોરન્સ","KRANK-en-kas-seh"),
("die Sprechstunde","consultation hours","કન્સલ્ટેશન સમય","SHPREKH-shtoon-deh"),
("die Praxis","clinic","ક્લિનિક","PRAK-sis"),
("die Behandlung","treatment","ઉપચાર","beh-HAND-lung"),
("die Diagnose","diagnosis","નિદાન","dee-ahg-NOH-zeh"),
("der Befund","medical finding","રિપોર્ટ","beh-FOONT"),
("die Untersuchung","examination","ચેકઅપ","oon-ter-ZOO-kung"),
("die Wartezeit","waiting time","વેઇટિંગ સમય","VAR-teh-tsyte"),
("der Termin","appointment","નિયુક્તિ","tehr-MEEN"),
("die Praxisgebühr","clinic fee","ફી","PRAK-sis-geh-BÜR"),
("die Versichertenkarte","insurance card","ઇન્સ્યોરન્સ કાર્ડ","fehr-ZIKH-er-ten-kar-teh"),
("die Notaufnahme","emergency room","ઇમર્જન્સી","NOHT-owf-nah-meh"),
("die Bescheinigung","certificate","પ્રમાણપત્ર","beh-SHY-ni-gung"),
],
[
("Ich brauche eine Überweisung zum Facharzt.","I need a referral to a specialist.","મને સ્પેશિયાલિસ્ટ માટે રિફરલ જોઈએ."),
("Die Wartezeit ist leider lang.","The waiting time is unfortunately long.","વેઇટિંગ સમય દુર્ભાગ્યે લાંબો છે."),
("Bitte zeigen Sie Ihre Versichertenkarte.","Please show your insurance card.","કૃપા કરીને ઇન્સ્યોરન્સ કાર્ડ બતાવો."),
("Die Diagnose wird heute besprochen.","The diagnosis will be discussed today.","આજે નિદાન અંગે ચર્ચા થશે."),
("Ich gehe in die Notaufnahme.","I go to the emergency room.","હું ઇમર્જન્સી રૂમમાં જઈ રહ્યો છું."),
],
[
("💡","Tip","In Germany you often need a referral to see a specialist.","FFF8E1"),
("📌","Insurance Card","Always carry your Versichertenkarte.","E8F0FE"),
("🇩🇪","Waiting Times","For specialists, waiting times can be several weeks.","E6F4EA"),
],
["Write a dialogue for booking a specialist appointment.","List 5 health-related documents you might need.","Explain how to ask for a referral in German."],
[("'die Überweisung' means?","Referral"),("Translate: 'Ich brauche einen Facharzt.'","I need a specialist."),("'die Versichertenkarte' means?","Insurance card"),("Translate: 'Die Wartezeit ist lang.'","The waiting time is long."),("'die Notaufnahme' means?","Emergency room")],
"You can now navigate advanced healthcare procedures in German."
))

# ── Day 162 ─────────────────────────────────────────────────────────────────
DAYS.append((162,
"SYMPTOMS & DIAGNOSIS 🧪",
"Describing conditions in detail",
"Today you learn how to describe symptoms precisely and understand common diagnoses in German.",
"આજે તમે symptoms ને ચોક્કસ રીતે વર્ણવવા અને સામાન્ય diagnoses સમજવા શીખશો.",
[
("der Schmerz","pain","દર્દ","SHMERTS"),
("stechend","stabbing","ચુભતો","SHTEH-khent"),
("dumpf","dull","મંદ","DOOMPf"),
("der Husten","cough","ખાંસી","HOOS-ten"),
("die Entzündung","inflammation","સોજો","ent-TZÜN-dung"),
("die Infektion","infection","ચેપ","in-fek-TSYON"),
("die Allergie","allergy","એલર્જી","al-LER-ghee"),
("schwindelig","dizzy","ચક્કર","SHVIN-de-likh"),
("übel","nauseous","મન ઊલટી","Ü-bel"),
("Fieber","fever","તાવ","FEE-ber"),
("die Temperatur","temperature","તાપમાન","tem-peh-rah-TOOR"),
("der Blutdruck","blood pressure","બ્લડ પ્રેશર","BLOOT-drook"),
("die Tablette","tablet/pill","ગોળી","ta-BLET-teh"),
("die Nebenwirkung","side effect","સાઇડ ઇફેક્ટ","NAY-ben-veer-kung"),
("die Besserung","improvement","સુધારો","BES-er-ung"),
],
[
("Ich habe stechende Schmerzen im Rücken.","I have stabbing pain in my back.","મને પીઠમાં ચુભતો દુખાવો છે."),
("Mir ist schwindelig.","I feel dizzy.","મને ચક્કર આવે છે."),
("Ich habe Fieber und Husten.","I have fever and cough.","મને તાવ અને ખાંસી છે."),
("Die Diagnose ist eine Entzündung.","The diagnosis is an inflammation.","નિદાન સોજો છે."),
("Gibt es Nebenwirkungen?","Are there side effects?","શું સાઇડ ઇફેક્ટ્સ છે?"),
],
[
("💡","Precision","Describe pain: stechend, dumpf, stark, leicht.","FFF8E1"),
("📌","Key Phrase","Mir ist… (I feel…) is common for symptoms.","E8F0FE"),
("🇩🇪","Doctor Talk","Doctors often ask: Seit wann? Wie stark? Wo genau?", "E6F4EA"),
],
["Write 5 symptom sentences using 'Mir ist…'.", "Describe pain in three ways.", "Practice asking about side effects in German."],
[("'schwindelig' means?","Dizzy"),("Translate: 'Ich habe Fieber.'","I have a fever."),("'die Nebenwirkung' means?","Side effect"),("Translate: 'Die Diagnose ist eine Infektion.'","The diagnosis is an infection."),("What does 'stechend' describe?","Stabbing pain")],
"You can now describe symptoms and diagnoses precisely in German."
))

# ── Day 163 ─────────────────────────────────────────────────────────────────
DAYS.append((163,
"PHARMACY & MEDICATION 💊",
"Medicines, dosage, instructions",
"Today you learn pharmacy vocabulary: dosage, prescriptions, and how to ask for medicines in German.",
"આજે તમે pharmacy vocabulary શીખશો: dosage, prescriptions અને medicines વિશે પૂછવું.",
[
("die Apotheke","pharmacy","ફાર્મસી","ah-po-TEH-keh"),
("das Rezept","prescription","પ્રિસ્ક્રિપ્શન","reh-TSEPT"),
("die Dosierung","dosage","ડોઝ","do-SEE-roong"),
("die Tablette","pill","ગોળી","ta-BLET-teh"),
("der Sirup","syrup","સીરપ","ZEE-roop"),
("die Salbe","ointment","મલમ","SAL-beh"),
("das Schmerzmittel","painkiller","પેઇનકિલર","SHMERTS-mit-el"),
("das Antibiotikum","antibiotic","એન્ટીબાયોટિક","an-tee-by-OH-ti-koom"),
("ohne Rezept","without prescription","પ્રિસ્ક્રિપ્શન વગર","OH-neh reh-TSEPT"),
("einnehmen","to take (medicine)","દવા લેવા","EYN-nay-men"),
("dreimal täglich","three times daily","દિવસમાં 3 વાર","DRY-mal TÄG-likh"),
("vor dem Essen","before meals","ખોરાક પહેલાં","for dem ES-en"),
("nach dem Essen","after meals","ખોરાક પછી","nakh dem ES-en"),
("die Packung","package","પેક","PAK-ung"),
("die Haltbarkeit","expiry","સમાપ્તિ","HALT-bar-kite"),
],
[
("Ich brauche ein Schmerzmittel ohne Rezept.","I need a painkiller without prescription.","મને પ્રિસ્ક્રિપ્શન વગર પેઇનકિલર જોઈએ."),
("Wie ist die Dosierung?","What is the dosage?","ડોઝ કેટલો છે?"),
("Nehmen Sie die Tablette dreimal täglich.","Take the tablet three times daily.","ગોળી દિવસમાં ત્રણ વાર લો."),
("Bitte vor dem Essen einnehmen.","Please take before meals.","કૃપા કરીને ભોજન પહેલાં લો."),
("Ist die Packung noch haltbar?","Is the package still valid?","પેક હજુ માન્ય છે?"),
],
[
("💡","Key Verb","einnehmen = to take medicine.","FFF8E1"),
("📌","Prescription","Rezept = paper prescription; e-Rezept is digital.","E8F0FE"),
("🇩🇪","Pharmacy","Pharmacists often give advice without a doctor.","E6F4EA"),
],
["Write a pharmacy dialogue asking for a medicine.","Practice dosage instructions in German.","List 5 medicines you know in German."],
[("'die Dosierung' means?","Dosage"),("Translate: 'vor dem Essen einnehmen.'","take before meals"),("'ohne Rezept' means?","without prescription"),("Translate: 'Ich brauche ein Antibiotikum.'","I need an antibiotic."),("'die Apotheke' means?","Pharmacy")],
"You can now handle pharmacy conversations in German."
))

# ── Day 164 ─────────────────────────────────────────────────────────────────
DAYS.append((164,
"FITNESS & NUTRITION 🥗",
"Healthy lifestyle vocabulary",
"Today you learn how to talk about fitness, diet, and healthy routines in German.",
"આજે તમે fitness, diet અને healthy routine વિશે German માં વાત કરશો.",
[
("die Gesundheit","health","આરોગ્ય","geh-ZOON-hite"),
("fit","fit","ફિટ","FIT"),
("trainieren","to train","વ્યાયામ કરવો","tray-NEE-ren"),
("die Ernährung","nutrition","પોષણ","er-NÄHR-oong"),
("ausgewogen","balanced","સંતુલિત","OWS-geh-VOH-gen"),
("die Kalorien","calories","કૅલરીઝ","ka-LOH-ree-en"),
("das Protein","protein","પ્રોટીન","pro-TEEN"),
("die Kohlenhydrate","carbs","કાર્બ્સ","koh-len-HY-drah-teh"),
("die Vitamine","vitamins","વિટામિન્સ","vee-TAH-mee-neh"),
("das Gemüse","vegetables","શાકભાજી","geh-MÜ-zeh"),
("das Obst","fruit","ફળ","OPST"),
("weniger","less","ઓછું","VAY-ni-ger"),
("mehr","more","વધુ","MEHR"),
("regelmäßig","regularly","નિયમિત","RAY-gel-mä-sig"),
("das Ziel","goal","લક્ષ્ય","TSEEL"),
],
[
("Ich trainiere dreimal pro Woche.","I train three times per week.","હું અઠવાડિયામાં ત્રણ વાર ટ્રેનિંગ કરું છું."),
("Eine ausgewogene Ernährung ist wichtig.","A balanced diet is important.","સંતુલિત આહાર મહત્વનો છે."),
("Ich esse mehr Gemüse und weniger Zucker.","I eat more vegetables and less sugar.","હું વધુ શાકભાજી અને ઓછું ખાંડ ખાઉં છું."),
("Regelmäßiger Sport verbessert die Gesundheit.","Regular sport improves health.","નિયમિત વ્યાયામ આરોગ્ય સુધારે છે."),
("Mein Ziel ist mehr Fitness.","My goal is more fitness.","મારું લક્ષ્ય વધુ ફિટનેસ છે."),
],
[
("💡","Habit","Small daily habits bring big results.","FFF8E1"),
("📌","Balance","Use ausgewogen for a balanced diet.","E8F0FE"),
("🇩🇪","Culture","Sports clubs and gyms are very common in Germany.","E6F4EA"),
],
["Write your weekly fitness plan in German.","List 5 healthy foods in German.","Write 3 sentences about your nutrition goals."],
[("'die Ernährung' means?","Nutrition"),("Translate: 'Ich trainiere regelmäßig.'","I train regularly."),("'ausgewogen' means?","Balanced"),("Translate: 'Ich esse weniger Zucker.'","I eat less sugar."),("'die Gesundheit' means?","Health")],
"You can now talk about fitness and nutrition in German."
))

# ── Day 165 ─────────────────────────────────────────────────────────────────
DAYS.append((165,
"MENTAL HEALTH & WELLBEING 🧠",
"Stress, balance, relaxation",
"Mental health is an important topic. Today you learn vocabulary for stress, feelings, and relaxation in German.",
"Mental health મહત્વપૂર્ણ છે. આજે તમે stress, feelings અને relaxation વિશે vocab શીખશો.",
[
("der Stress","stress","તણાવ","SHTRES"),
("gestresst","stressed","તણાવમાં","geh-SHTREST"),
("entspannt","relaxed","આરામદાયક","ent-SHPANT"),
("die Ruhe","calm","શાંતિ","ROO-eh"),
("die Balance","balance","સંતુલન","ba-LAN-seh"),
("die Erholung","recovery","આરામ","ehr-HOH-lung"),
("die Pause","break","વિરામ","POW-zeh"),
("sich entspannen","to relax","આરામ કરવો","zikh ent-SHPAN-en"),
("überfordert","overwhelmed","અતિભારિત","Ü-ber-for-dert"),
("sich Sorgen machen","to worry","ચિંતા કરવી","ZOR-gen"),
("zufrieden","satisfied","સંતોષિત","tsoo-FREE-den"),
("glücklich","happy","ખુશ","GLÜK-likh"),
("die Unterstützung","support","આધાર","oon-ter-SHTÜTS-ung"),
("das Gespräch","conversation","વાતચીત","geh-SHPRÄKH"),
("die Beratung","counseling","પરામર્શ","beh-RAH-tung"),
],
[
("Ich fühle mich gestresst.","I feel stressed.","હું તણાવમાં છું."),
("Ich brauche eine Pause.","I need a break.","મને વિરામ જોઈએ."),
("Ein Gespräch hilft mir.","A conversation helps me.","વાતચીત મને મદદ કરે છે."),
("Ich versuche, mich zu entspannen.","I try to relax.","હું આરામ કરવાનો પ્રયાસ કરું છું."),
("Balance zwischen Arbeit und Freizeit ist wichtig.","Balance between work and free time is important.","કામ અને ફ્રી ટાઇમ વચ્ચે સંતુલન મહત્વનું છે."),
],
[
("💡","Self-care","Regular breaks and sleep improve wellbeing.","FFF8E1"),
("📌","Language","Use 'Ich fühle mich…' to describe emotions.","E8F0FE"),
("🇩🇪","Culture","Mental health awareness is growing in Germany.","E6F4EA"),
],
["Write 5 sentences about your stress and relaxation.","List 5 activities that help you relax.","Practice asking for support in German."],
[("'gestresst' means?","Stressed"),("Translate: 'Ich brauche eine Pause.'","I need a break."),("'die Erholung' means?","Recovery/relaxation"),("Translate: 'Ich fühle mich entspannt.'","I feel relaxed."),("'die Unterstützung' means?","Support")],
"You can now talk about mental health and wellbeing in German."
))

# ── Day 166 ─────────────────────────────────────────────────────────────────
DAYS.append((166,
"EMERGENCIES & FIRST AID 🚑",
"Safety and urgent situations",
"Today you learn emergency vocabulary, how to call for help, and basic first aid phrases in German.",
"આજે તમે emergency vocabulary, મદદ બોલાવવી, અને basic first aid phrases શીખશો.",
[
("der Notruf","emergency call","ઇમર્જન્સી કોલ","NOHT-roof"),
("die 112","emergency number","112","hundert-zwoelf"),
("der Unfall","accident","અકસ્માત","OON-fal"),
("die Verletzung","injury","ઇજા","fehr-LET-zung"),
("bluten","to bleed","લોહી વહેવું","BLOO-ten"),
("bewusstlos","unconscious","બેભાન","beh-VOOST-lohs"),
("die Rettung","rescue","રેસ્ક્યુ","RET-toong"),
("der Krankenwagen","ambulance","એમ્બ્યુલન્સ","KRANK-en-vah-gen"),
("dringend","urgent","તાત્કાલિક","DRING-ent"),
("sofort","immediately","તરત","zoh-FORT"),
("helfen","to help","મદદ કરવી","HEL-fen"),
("beruhigen","to calm down","શાંત કરવું","beh-ROO-ig-en"),
("die Erste Hilfe","first aid","ફર્સ્ટ એઇડ","ER-steh HIL-feh"),
("stabilisieren","to stabilize","સ્થિર કરવું","shta-bi-li-ZEER-en"),
("die Gefahr","danger","જોખમ","geh-FAHR"),
],
[
("Rufen Sie den Notruf 112!","Call emergency number 112!","ઇમર્જન્સી 112 પર કોલ કરો!"),
("Es gab einen Unfall.","There was an accident.","અકસ્માત થયો છે."),
("Er ist bewusstlos.","He is unconscious.","તે બેભાન છે."),
("Bitte schicken Sie einen Krankenwagen.","Please send an ambulance.","કૃપા કરીને એમ્બ્યુલન્સ મોકલો."),
("Ich leiste Erste Hilfe.","I give first aid.","હું ફર્સ્ટ એઇડ આપું છું."),
],
[
("💡","Emergency Tip","Stay calm and give clear information: location, what happened, number of injured.","FFF8E1"),
("📌","Number","112 works across the EU for emergencies.","E8F0FE"),
("🇩🇪","Key Phrase","'Ich brauche Hilfe!' = I need help.","E6F4EA"),
],
["Write a short emergency call script in German.","List 5 important emergency words in German.","Practice giving your address clearly."],
[("Emergency number in Germany?","112"),("Translate: 'Er ist bewusstlos.'","He is unconscious."),("'die Verletzung' means?","Injury"),("Translate: 'Bitte schicken Sie einen Krankenwagen.'","Please send an ambulance."),("'die Erste Hilfe' means?","First aid")],
"You can now handle emergency situations with basic German vocabulary and phrases."
))

# ── Day 167 ─────────────────────────────────────────────────────────────────
DAYS.append((167,
"FAMILY & RELATIONSHIPS 👨‍👩‍👧",
"Talking about close relationships",
"Today you learn vocabulary to talk about relationships, family roles, and life events.",
"આજે તમે સંબંધો, પરિવાર અને જીવનની ઘટનાઓ વિશે vocabulary શીખશો.",
[
("die Beziehung","relationship","સંબંધ","beh-TSIE-hoong"),
("verheiratet","married","વિવાહિત","fehr-HY-rah-tet"),
("geschieden","divorced","છૂટાછેડા","geh-SHIE-den"),
("der Partner","partner","પાર્ટનર","PARt-ner"),
("die Partnerin","partner (f)","પાર્ટનર (સ્ત્રી)","PARt-neh-rin"),
("die Familie","family","પરિવાર","fa-MEE-lee-eh"),
("die Verwandten","relatives","સગા","fehr-VANT-en"),
("der Enkel","grandson","પૌત્ર","ENG-kel"),
("die Enkelin","granddaughter","પૌત્રી","ENG-keh-lin"),
("die Hochzeit","wedding","લગ્ન","HOCH-tsyte"),
("die Trennung","separation","અલગાવ","TREN-noong"),
("die Verantwortung","responsibility","જવાબદારી","fehr-ANT-vort-oong"),
("sich verlieben","to fall in love","પ્રેમમાં પડવું","zikh fehr-LEE-ben"),
("sich verloben","to get engaged","સગાઈ કરવી","zikh fehr-LOH-ben"),
("das Zusammenleben","living together","સાથે રહેવું","tsoo-ZAM-en-lay-ben"),
],
[
("Sie sind seit fünf Jahren verheiratet.","They have been married for five years.","તેઓ પાંચ વર્ષથી વિવાહિત છે."),
("Ich habe eine enge Beziehung zu meiner Familie.","I have a close relationship with my family.","મારે મારા પરિવાર સાથે નજીકનો સંબંધ છે."),
("Wir sind seit kurzem verlobt.","We have recently gotten engaged.","અમે તાજેતરમાં સગાઈ કરી છે."),
("Die Hochzeit findet im Sommer statt.","The wedding takes place in summer.","લગ્ન ઉનાળામાં થશે."),
("Nach der Trennung wohnen sie getrennt.","After the separation they live separately.","અલગ થયા પછી તેઓ અલગ રહે છે."),
],
[
("💡","Time Expressions","Use 'seit' + time for relationship duration.","FFF8E1"),
("📌","Neutral Language","Use respectful terms when discussing relationships.","E8F0FE"),
("🇩🇪","Culture","Civil marriage is required in Germany before any церemony.","E6F4EA"),
],
["Write 5 sentences about your family in German.","Describe a family event (wedding, birthday) in German.","Practice using 'seit' with relationship durations."],
[("'verheiratet' means?","Married"),("Translate: 'Wir sind verlobt.'","We are engaged."),("'die Hochzeit' means?","Wedding"),("Translate: 'Ich habe enge Beziehung.'","I have a close relationship."),("'die Verwandten' means?","Relatives")],
"You can now speak about relationships and family in German more precisely."
))

# ── Day 168 ─────────────────────────────────────────────────────────────────
DAYS.append((168,
"CHILDCARE & SCHOOL 🧒",
"Parents and education vocabulary",
"Today you learn vocabulary for childcare, school communication, and parent-teacher interactions.",
"આજે તમે childcare, school communication અને parent-teacher interactions માટે vocab શીખશો.",
[
("das Kind","child","બાળક","KINT"),
("die Betreuung","care","કાળજી","beh-TROY-oong"),
("die Kita","daycare","કિટા","KEE-tah"),
("die Schule","school","શાળા","SHOO-leh"),
("die Lehrerin","teacher","શિક્ષિકા","LAY-rer-in"),
("das Zeugnis","report card","રિપોર્ટ કાર્ડ","TSOYGH-nis"),
("die Hausaufgaben","homework","હોમવર્ક","HOWS-owf-gah-ben"),
("die Eltern","parents","માતા-પિતા","EL-tern"),
("der Elternabend","parents' evening","પેરેન્ટ મીટિંગ","EL-tern-ah-bent"),
("die Entwicklung","development","વિકાસ","ent-VIK-lung"),
("die Förderung","support/assistance","મદદ","FÖR-deh-rung"),
("die Anmeldung","registration","રજીસ્ટ્રેશન","AN-mel-doong"),
("die Gebühr","fee","ફી","geh-BÜR"),
("die Pause","break","વિરામ","POW-zeh"),
("der Stundenplan","timetable","ટાઇમટેબલ","SHTOON-den-plahn"),
],
[
("Mein Kind geht in die Kita.","My child goes to daycare.","મારું બાળક કિટા જાય છે."),
("Die Hausaufgaben sind heute viel.","There is a lot of homework today.","આજે ઘણું હોમવર્ક છે."),
("Wir haben nächste Woche Elternabend.","We have parent evening next week.","આગલા અઠવાડિયે પેરેન્ટ મીટિંગ છે."),
("Das Zeugnis war sehr gut.","The report card was very good.","રિપોર્ટ કાર્ડ ખૂબ સારો હતો."),
("Der Stundenplan ist geändert.","The timetable has changed.","ટાઇમટેબલ બદલાયો છે."),
],
[
("💡","School Contact","Teachers appreciate polite, clear messages.","FFF8E1"),
("📌","Key Word","Elternabend = parents’ evening, very common in Germany.","E8F0FE"),
("🇩🇪","Daycare","Kita places can be limited — apply early.","E6F4EA"),
],
["Write a short note to a teacher in German.","List 5 school-related words you use often.","Describe your child’s school day in German."],
[("'die Kita' means?","Daycare"),("Translate: 'Die Hausaufgaben sind schwer.'","The homework is difficult."),("'der Stundenplan' means?","Timetable"),("Translate: 'Wir haben Elternabend.'","We have parents' evening."),("'das Zeugnis' means?","Report card")],
"You can now discuss childcare and school topics in German."
))

# ── Day 169 ─────────────────────────────────────────────────────────────────
DAYS.append((169,
"SOCIAL EVENTS & INVITATIONS 🎉",
"Planning and accepting invites",
"Today you learn phrases for invitations, accepting, declining politely, and planning events in German.",
"આજે તમે invitations આપવી, સ્વીકારવી, નમ્ર રીતે નકારવી અને events પ્લાન કરવાના phrases શીખશો.",
[
("die Einladung","invitation","નિમંત્રણ","EYN-lah-doong"),
("einladen","to invite","આમંત્રિત કરવું","EYN-lah-den"),
("zusagen","to accept","સ્વીકારવું","TSOO-zah-gen"),
("absagen","to cancel","નકારવું","AP-zah-gen"),
("leider","unfortunately","દુર્ભાગ્યે","LY-der"),
("gern","gladly","ખુશીથી","GERN"),
("die Feier","celebration","પાર્ટી","FY-er"),
("die Planung","planning","યોજનાબદ્ધ","PLAH-nung"),
("der Zeitpunkt","time/date","સમય","TSYTE-punkt"),
("der Ort","place","જગ્યા","ORT"),
("mitbringen","to bring along","સાથે લાવવું","MIT-bring-en"),
("Geschenk","gift","ભેટ","geh-SHENK"),
("die Zusage","acceptance","સ્વીકાર","TSOO-zah-geh"),
("die Absage","decline","નકાર","AP-zah-geh"),
("sich treffen","to meet","મળવું","zikh TREF-fen"),
],
[
("Danke für die Einladung! Ich komme gern.","Thanks for the invitation! I'd love to come.","આમંત્રણ માટે આભાર! હું ખુશીથી આવીશ."),
("Leider kann ich nicht kommen.","Unfortunately I can't come.","દુર્ભાગ્યે હું આવી શકતો નથી."),
("Wann und wo treffen wir uns?","When and where do we meet?","અમે ક્યારે અને ક્યાં મળીએ?"),
("Soll ich etwas mitbringen?","Should I bring something?","શું હું કંઈ સાથે લાવું?"),
("Die Feier beginnt um 18 Uhr.","The celebration starts at 6 pm.","પાર્ટી 6 વાગે શરૂ થાય છે."),
],
[
("💡","Polite Decline","Always give a short reason when declining an invitation.","FFF8E1"),
("📌","Time","Use clear time/date expressions to avoid confusion.","E8F0FE"),
("🇩🇪","Etiquette","Bringing a small gift is polite in Germany.","E6F4EA"),
],
["Write an invitation message for a birthday.","Write a polite decline with a reason.","Create a plan for a small event in German."],
[("'einladen' means?","to invite"),("Translate: 'Leider kann ich nicht kommen.'","Unfortunately I can't come."),("'mitbringen' means?","to bring along"),("Translate: 'Wann treffen wir uns?'","When do we meet?"),("'die Feier' means?","Celebration / party")],
"You can now handle invitations and social event planning in German."
))

# ── Day 170 ─────────────────────────────────────────────────────────────────
DAYS.append((170,
"MODULE REVIEW: DAYS 161–170 ✅",
"Health & Social Life Review",
"You reviewed healthcare, symptoms, pharmacy, fitness, mental health, emergencies, family, childcare, and social events. Today is consolidation.",
"તમે healthcare, symptoms, pharmacy, fitness, mental health, emergencies, family, childcare અને social events cover કર્યા. આજે review day છે.",
[
("die Zusammenfassung","summary","સારાંશ","tsoo-ZAM-en-fas-ung"),
("die Gesundheit","health","આરોગ્ય","geh-ZOON-hite"),
("die Sicherheit","safety","સુરક્ષા","ZIKH-er-hite"),
("die Unterstützung","support","આધાર","oon-ter-SHTÜTS-ung"),
("die Familie","family","પરિવાર","fa-MEE-lee-eh"),
("die Einladung","invitation","નિમંત્રણ","EYN-lah-doong"),
("die Praxis","clinic","ક્લિનિક","PRAK-sis"),
("die Apotheke","pharmacy","ફાર્મસી","ah-po-TEH-keh"),
("die Betreuung","care","કાળજી","beh-TROY-oong"),
("die Entspannung","relaxation","આરામ","ent-SHPAN-nung"),
("die Notfall","emergency","ઇમર્જન્સી","NOHT-fal"),
("die Verantwortung","responsibility","જવાબદારી","fehr-ANT-vort-oong"),
("die Gewohnheit","habit","આદત","geh-VOHN-hite"),
("die Beziehung","relationship","સંબંધ","beh-TSIE-hoong"),
("die Planung","planning","યોજનાબદ્ધ","PLAH-nung"),
],
[
("Ich fühle mich sicherer in medizinischen Situationen.","I feel more confident in medical situations.","હું હવે medical situations માં વધુ આત્મવિશ્વાસી છું."),
("Ich kann Symptome besser beschreiben.","I can describe symptoms better.","હું હવે symptoms વધુ સારી રીતે વર્ણવી શકું છું."),
("Ich kann Einladungen höflich annehmen oder ablehnen.","I can accept or decline invitations politely.","હું invitations નમ્ર રીતે સ્વીકારી અથવા નકારી શકું છું."),
("Ich kann über Gesundheit und Familie sprechen.","I can talk about health and family.","હું આરોગ્ય અને પરિવાર વિશે વાત કરી શકું છું."),
("Ich bin bereit für das nächste Modul.","I am ready for the next module.","હું આગળના module માટે તૈયાર છું."),
],
[
("🏆","Checkpoint","If you can explain a symptom and write a polite invitation, you are ready.","E8F0FE"),
("💡","Practice","Write a short health diary for one week.","FFF8E1"),
("🎯","Next","Days 171–190 focus on communication, presentations, and advanced language.","E6F4EA"),
],
["Write 5 sentences about your health routine.","Create a 1‑minute speech about family or relationships.","Review 20 vocabulary words from this module."],
[("Translate: 'Ich kann Symptome beschreiben.'","I can describe symptoms."),("'die Einladung' means?","Invitation"),("'die Apotheke' means?","Pharmacy"),("Translate: 'Ich bin bereit.'","I am ready."),("What is next module about?","Communication, presentations, advanced language")],
"Module review complete. Your health and social German is strong and ready to expand further."
))

print("Days 161-170 appended")

# ── Day 171 ─────────────────────────────────────────────────────────────────
DAYS.append((171,
"PRESENTATIONS & PUBLIC SPEAKING 🎤",
"Structure and delivery",
"Today you learn vocabulary to structure a presentation and speak in public confidently in German.",
"આજે તમે German માં presentation structure અને public speaking માટેનું vocab શીખશો.",
[
("die Präsentation","presentation","પ્રેઝન્ટેશન","pre-zen-ta-TSYON"),
("die Einleitung","introduction","પરિચય","EYN-ly-tung"),
("der Hauptteil","main part","મુખ્ય ભાગ","HOWPT-tyl"),
("der Schluss","conclusion","નિષ્કર્ષ","SHLOOS"),
("das Thema","topic","વિષય","TAY-ma"),
("das Ziel","goal","લક્ષ્ય","TSEEL"),
("die Folie","slide","સ્લાઇડ","FOH-lee-eh"),
("das Beispiel","example","ઉદાહરણ","bye-SHPLEEL"),
("die Grafik","chart","ગ્રાફ","GRAH-fik"),
("erklären","to explain","સમજાવવું","ehr-KLÄ-ren"),
("betonen","to emphasize","જોર આપવો","beh-TOH-nen"),
("zusammenfassen","to summarize","સારાંશ કરવું","tsoo-ZAM-en-fas-en"),
("die Frage","question","પ્રશ્ન","FRAH-geh"),
("das Publikum","audience","શ્રોતાઓ","POO-bli-koom"),
("aufmerksam","attentive","ધ્યાનપૂર્વક","OWF-merk-zam"),
],
[
("Heute halte ich eine Präsentation.","Today I give a presentation.","આજે હું પ્રેઝન્ટેશન આપું છું."),
("In der Einleitung stelle ich das Thema vor.","In the introduction I present the topic.","પરિચયમાં હું વિષય રજૂ કરું છું."),
("Im Hauptteil erkläre ich die wichtigsten Punkte.","In the main part I explain the most important points.","મુખ્ય ભાગમાં હું મુખ્ય મુદ્દાઓ સમજાવું છું."),
("Am Schluss fasse ich alles zusammen.","At the end I summarize everything.","અંતે હું બધું સારાંશ કરું છું."),
("Gibt es Fragen aus dem Publikum?","Are there questions from the audience?","શ્રોતાઓ તરફથી પ્રશ્નો છે?"),
],
[
("💡","Structure","Einleitung → Hauptteil → Schluss is the standard structure.","FFF8E1"),
("📌","Clarity","Use examples and graphics to make points clear.","E8F0FE"),
("🇩🇪","Tip","Speak slowly and make eye contact.","E6F4EA"),
],
["Create a 5‑slide outline for a presentation in German.","Write 3 opening sentences for a presentation.","Practice a 1‑minute summary of a topic."],
[("'die Einleitung' means?","Introduction"),("Translate: 'Am Schluss fasse ich zusammen.'","At the end I summarize."),("'das Publikum' means?","Audience"),("Translate: 'Ich erkläre die Grafik.'","I explain the chart."),("'betonen' means?","to emphasize")],
"You can now structure and deliver a presentation in German with confidence."
))

# ── Day 172 ─────────────────────────────────────────────────────────────────
DAYS.append((172,
"ARGUMENTS & DEBATE ⚖️",
"Pro & Contra language",
"Today you learn how to build arguments, express pros and cons, and debate politely in German.",
"આજે તમે arguments બનાવવાની, pros/cons જણાવવાની અને નમ્ર રીતે debate કરવાની ભાષા શીખશો.",
[
("das Argument","argument","દલીલ","ar-goo-MENT"),
("die Begründung","reasoning","કારણ","beh-GRÜN-dung"),
("der Vorteil","advantage","લાભ","FOR-tyl"),
("der Nachteil","disadvantage","નુકસાન","NAKH-tyl"),
("einerseits … andererseits","on the one hand…on the other","એક તરફ...બીજી તરફ","EY-ner-zyts"),
("ich bin dafür","I am in favor","હું તરફ છું","ikh bin da-FÜR"),
("ich bin dagegen","I am against","હું વિરુદ્ધ છું","ikh bin dah-GAY-gen"),
("meiner Meinung nach","in my opinion","મારી મત મુજબ","MY-ner MY-nung nakh"),
("zum Beispiel","for example","ઉદાહરણરૂપ","tsoom bye-SHPLEEL"),
("außerdem","moreover","ઉપરાંત","OWS-er-daym"),
("trotzdem","nevertheless","છતા પણ","TROTS-daym"),
("überzeugen","to convince","મનાવવું","Ü-ber-TSOY-gen"),
("zustimmen","to agree","સહમત થવું","TSOO-shtim-en"),
("widersprechen","to contradict","વિરોધ કરવો","VEE-der-shpreh-khen"),
("der Kompromiss","compromise","સમજૂતી","KOM-proh-mis"),
],
[
("Ein Vorteil ist die Zeitersparnis.","An advantage is time saving.","એક લાભ સમય બચત છે."),
("Ein Nachteil ist der hohe Preis.","A disadvantage is the high price.","એક નુકસાન ઊંચી કિંમત છે."),
("Meiner Meinung nach ist das sinnvoll.","In my opinion, this is sensible.","મારી મત મુજબ આ યોગ્ય છે."),
("Einerseits ist es günstig, andererseits ist die Qualität schlecht.","On one hand it's cheap, on the other the quality is bad.","એક તરફ સસ્તું છે, બીજી તરફ ગુણવત્તા ખરાબ છે."),
("Wir brauchen einen Kompromiss.","We need a compromise.","અમને સમજૂતી જોઈએ."),
],
[
("💡","Debate","Use 'einerseits/andererseits' to show balance.","FFF8E1"),
("📌","Politeness","Disagree politely: 'Ich sehe das anders.'", "E8F0FE"),
("🇩🇪","Style","German debates prefer clear reasons and logic.","E6F4EA"),
],
["Write 5 pro/contra sentences about a topic.","Create a short debate about working from home.","Use 'außerdem' and 'trotzdem' in 3 sentences."],
[("'der Vorteil' means?","Advantage"),("Translate: 'Ich bin dagegen.'","I am against it."),("'der Kompromiss' means?","Compromise"),("Translate: 'Meiner Meinung nach…'","In my opinion…"),("'widersprechen' means?","To contradict")],
"You can now express arguments and debate politely in German."
))

# ── Day 173 ─────────────────────────────────────────────────────────────────
DAYS.append((173,
"NEGOTIATION LANGUAGE 🤝",
"Discussing terms and agreements",
"Negotiation skills are useful at work and in daily life. Today you learn phrases for negotiating politely in German.",
"Negotiation skills કામ અને daily life માટે જરૂરી છે. આજે તમે polite negotiation phrases શીખશો.",
[
("verhandeln","to negotiate","સંવાદ કરી નક્કી કરવું","fehr-HAN-deln"),
("das Angebot","offer","ઓફર","AN-geh-boht"),
("der Preis","price","કિંમત","PRYSS"),
("der Rabatt","discount","છૂટ","ra-BAT"),
("die Bedingung","condition","શરત","beh-DIN-goong"),
("die Vereinbarung","agreement","સમજૂતી","fehr-EYN-bah-rung"),
("zustimmen","to agree","સહમત થવું","TSOO-shtim-en"),
("ablehnen","to reject","નકારવું","AP-lay-nen"),
("verfügbar","available","ઉપલબ્ધ","fehr-FÜG-bar"),
("flexibel","flexible","લવચીક","fleK-SI-bel"),
("verbindlich","binding","બાંધકામક","fehr-BIN-dlikh"),
("unverbindlich","non-binding","બિન-બાંધકામક","oon-fehr-BIN-dlikh"),
("der Vorschlag","suggestion","સૂચન","FOR-shlahg"),
("das Ziel","goal","લક્ષ્ય","TSEEL"),
("einigen","to agree (reach agreement)","મળતા થવું","EYE-ni-gen"),
],
[
("Wir müssen den Preis verhandeln.","We need to negotiate the price.","અમને કિંમત પર negotiation કરવું પડશે."),
("Können Sie einen Rabatt anbieten?","Can you offer a discount?","શું તમે છૂટ આપી શકો?"),
("Das Angebot ist unverbindlich.","The offer is non-binding.","ઓફર બિન-બાંધકામક છે."),
("Wir einigen uns auf 50 Euro.","We agree on 50 euros.","અમે 50 યુરો પર સહમત થઈએ છીએ."),
("Die Bedingungen sind flexibel.","The conditions are flexible.","શરતો લવચીક છે."),
],
[
("💡","Negotiation Tip","Be polite and clear; ask for a counteroffer.","FFF8E1"),
("📌","Binding","Verbindlich = binding; unverbindlich = no obligation.","E8F0FE"),
("🇩🇪","Culture","Germans like clear terms and written agreements.","E6F4EA"),
],
["Write 5 negotiation sentences about price or time.","Create a short dialogue negotiating a service.","Use 'verbindlich/unverbindlich' in 3 sentences."],
[("'verhandeln' means?","To negotiate"),("Translate: 'Können Sie einen Rabatt anbieten?'","Can you offer a discount?"),("'die Bedingung' means?","Condition"),("Translate: 'Wir einigen uns.'","We agree."),("'verbindlich' means?","Binding")],
"You can now negotiate politely in German."
))

# ── Day 174 ─────────────────────────────────────────────────────────────────
DAYS.append((174,
"IDIOMS & EXPRESSIONS 💬",
"Common phrases in daily German",
"Idioms make your German sound natural. Today you learn common expressions and how to use them.",
"Idiomsથી તમારી German વધુ natural લાગે છે. આજે common expressions શીખશો.",
[
("Keine Ahnung!","No idea!","કોઈ ખ્યાલ નથી!","KY-neh OW-noong"),
("Das ist mir egal.","I don't care.","મને ફરક નથી."),
("Ich drücke dir die Daumen.","I keep my fingers crossed for you.","હું તારા માટે શુભેચ્છા આપું છું."),
("Alles klar!","All clear / okay!","બરાબર છે!","AL-es KLAHR"),
("Es geht mir gut.","I am fine.","હું સારું છું."),
("Das macht Sinn.","That makes sense.","તે સમજમાં આવે છે."),
("Keine Sorge.","No worries.","ચિંતા નહિ."),
("Ich bin dran.","It's my turn.","મારો વારો છે."),
("Da hast du recht.","You are right about that.","તારી વાત સાચી છે."),
("Auf jeden Fall.","In any case / definitely.","નિશ્ચિત રીતે."),
("Das ist nicht mein Ding.","That's not my thing.","એ મારી વસ્તુ નથી."),
("Gib mir Bescheid.","Let me know.","મને જણાવજે."),
("Ich melde mich.","I'll get back to you.","હું સંપર્ક કરીશ."),
("Das ist es wert.","It's worth it.","તે લાયક છે."),
("Keine Zeit.","No time.","સમય નથી."),
],
[
("Keine Sorge, alles wird gut.","No worries, everything will be fine.","ચિંતા નહિ, બધું ઠીક થશે."),
("Auf jeden Fall komme ich morgen.","I will definitely come tomorrow.","હું કાલે ચોક્કસ આવીશ."),
("Gib mir Bescheid, wenn du fertig bist.","Let me know when you're done.","જ્યારે તું તૈયાર થે ત્યારે મને જણાવજે."),
("Das ist nicht mein Ding.","That's not my thing.","એ મારી વસ્તુ નથી."),
("Da hast du recht.","You're right about that.","તારી વાત સાચી છે."),
],
[
("💡","Natural Speech","Idioms make you sound more like a native speaker.","FFF8E1"),
("📌","Usage","Learn when to use informal vs formal expressions.","E8F0FE"),
("🇩🇪","Tip","Practice idioms in real conversations.","E6F4EA"),
],
["Use 5 idioms in your own sentences.","Write a short chat message using 3 idioms.","Translate 5 everyday expressions from your language into German."],
[("'Keine Ahnung' means?","No idea"),("Translate: 'Das macht Sinn.'","That makes sense."),("'Ich melde mich' means?","I'll get back to you"),("Translate: 'Auf jeden Fall.'","Definitely"),("'Gib mir Bescheid' means?","Let me know")],
"You now know common German idioms to sound more natural in conversation."
))

# ── Day 175 ─────────────────────────────────────────────────────────────────
DAYS.append((175,
"FORMAL COMPLAINTS ✉️",
"Writing a complaint letter",
"Today you learn formal complaint language: describing a problem, requesting a solution, and staying polite.",
"આજે તમે formal complaint language શીખશો: સમસ્યા વર્ણવવી, ઉકેલ માંગવો અને polite રહેવું.",
[
("die Beschwerde","complaint","ફરિયાદ","beh-SHVER-deh"),
("der Mangel","defect","ખામી","MAN-gel"),
("die Reklamation","complaint/claim","રિક્લેમેશન","reh-kla-ma-TSYON"),
("die Rückerstattung","refund","રીફંડ","RÜK-er-SHTAT-tung"),
("der Umtausch","exchange","એક્સચેન્જ","OOM-towsh"),
("entsprechen","to correspond","મિલવું","ent-SHPREKH-en"),
("unacceptable","unacceptable","અસ્વીકાર્ય","oon-ak-tsep-TAH-bel"),
("der Fehler","error","ભૂલ","FAY-ler"),
("die Frist","deadline","સમયસીમા","FRIST"),
("die Lösung","solution","ઉકેલ","LÖ-zoong"),
("freundlich","friendly","મૈત્રીપૂર્ણ","FROYNT-likh"),
("höflich","polite","વિનમ્ર","HÖF-likh"),
("bitte um","request for","વિનંતી","BIT-teh oom"),
("unzufrieden","unsatisfied","અસંતુષ્ટ","OON-tsoo-FREE-den"),
("die Bestätigung","confirmation","પુષ્ટિ","beh-SHTÄ-ti-gung"),
],
[
("Ich möchte eine Beschwerde einreichen.","I would like to file a complaint.","હું ફરિયાદ દાખલ કરવી ઈચ્છું છું."),
("Das Produkt hat einen Mangel.","The product has a defect.","પ્રોડક્ટમાં ખામી છે."),
("Ich bitte um eine Rückerstattung.","I request a refund.","હું રિફંડ માંગું છું."),
("Bitte bestätigen Sie den Erhalt.","Please confirm receipt.","કૃપા કરીને પ્રાપ્તી પુષ્ટિ કરો."),
("Ich bin mit dem Service unzufrieden.","I am unsatisfied with the service.","હું સેવા થી અસંતુષ્ટ છું."),
],
[
("💡","Tone","Stay polite and factual in complaints.","FFF8E1"),
("📌","Structure","Problem → evidence → request → deadline.","E8F0FE"),
("🇩🇪","Legal","Written complaints are common and taken seriously in Germany.","E6F4EA"),
],
["Write a short complaint letter about a defective product.","List 3 solutions you could request.","Practice 3 polite complaint sentences."],
[("'die Beschwerde' means?","Complaint"),("Translate: 'Ich bitte um eine Rückerstattung.'","I request a refund."),("'der Mangel' means?","Defect"),("Translate: 'Ich bin unzufrieden.'","I am unsatisfied."),("'die Reklamation' means?","Claim/complaint")],
"You can now write a formal complaint in German in a professional and polite way."
))

# ── Day 176 ─────────────────────────────────────────────────────────────────
DAYS.append((176,
"INSTRUCTIONS & PROCEDURES 🧾",
"Step-by-step language",
"Today you learn how to give and understand instructions in German using sequencing words and commands.",
"આજે તમે step‑by‑step instructions German માં આપવાનું શીખશો.",
[
("zuerst","first","પ્રથમ","tsoo-ERST"),
("dann","then","પછી","DAN"),
("danach","after that","ત્યારપછી","dah-NAKH"),
("schließlich","finally","અંતે","SHLES-likh"),
("öffnen","to open","ખોલવું","ÖF-nen"),
("drücken","to press","દબાવવું","DRÜK-en"),
("auswählen","to select","પસંદ કરવું","OWS-vä-len"),
("bestätigen","to confirm","પુષ્ટિ કરવી","beh-SHTÄ-ti-gen"),
("installieren","to install","ઇન્સ્ટોલ કરવું","in-sta-LEE-ren"),
("speichern","to save","સેવ કરવું","SHPI-khern"),
("abschließen","to finish","પુર્ણ કરવું","AP-shlee-sen"),
("die Anleitung","instruction manual","માર્ગદર્શિકા","AN-ly-tung"),
("der Schritt","step","પગલું","SHRIT"),
("vorsichtig","carefully","સાવધાનીથી","FOR-zikh-tikh"),
("wichtig","important","મહત્વનું","VIKH-tikh"),
],
[
("Zuerst öffnen Sie die App.","First open the app.","પ્રથમ એપ ખોલો."),
("Dann wählen Sie die Sprache aus.","Then select the language.","પછી ભાષા પસંદ કરો."),
("Danach bestätigen Sie mit OK.","After that confirm with OK.","ત્યારપછી OK દબાવો."),
("Schließlich speichern Sie die Änderungen.","Finally save the changes.","અંતે ફેરફારો સેવ કરો."),
("Bitte lesen Sie die Anleitung.","Please read the manual.","કૃપા કરીને માર્ગદર્શિકા વાંચો."),
],
[
("💡","Clarity","Use simple verbs and sequence words.","FFF8E1"),
("📌","Formal","In instructions, use imperative (Sie-form) for politeness.","E8F0FE"),
("🇩🇪","Context","German instructions are precise and step-by-step.","E6F4EA"),
],
["Write a 5‑step instruction for an app.","Give instructions for making tea in German.","Create a short manual with 4 steps."],
[("'zuerst' means?","First"),("Translate: 'Dann drücken Sie OK.'","Then press OK."),("'die Anleitung' means?","Instruction manual"),("Translate: 'Schließlich speichern.'","Finally save."),("'vorsichtig' means?","Carefully")],
"You can now give clear instructions in German."
))

# ── Day 177 ─────────────────────────────────────────────────────────────────
DAYS.append((177,
"STORYTELLING ADVANCED 📚",
"Connecting events and emotions",
"Today you learn storytelling connectors and emotions to make stories richer in German.",
"આજે તમે storytelling connectors અને emotions શીખશો જેથી વાર્તાઓ વધુ જીવંત બને.",
[
("plötzlich","suddenly","અચાનક","PLÖTS-likh"),
("zum Glück","luckily","સૌભાગ્યે","tsoom GLÜK"),
("leider","unfortunately","દુર્ભાગ્યે","LY-der"),
("deshalb","therefore","એથી","DES-halb"),
("danach","after that","ત્યાર પછી","dah-NAKH"),
("schließlich","finally","અંતે","SHLES-likh"),
("die Überraschung","surprise","આશ્ચર્ય","Ü-ber-RAH-shung"),
("die Angst","fear","ડર","ANGST"),
("die Freude","joy","ખુશી","FROY-deh"),
("enttäuscht","disappointed","નિરાશ","ent-TÄOOSHT"),
("aufgeregt","excited","ઉત્સાહિત","OWF-geh-REHKHT"),
("mutig","brave","હિંમતવાળું","MOO-tikh"),
("das Ende","ending","અંત","EN-deh"),
("die Wendung","turning point","મોડ","VEN-doong"),
("erzählen","to tell","કહેવું","ehr-TSÄH-len"),
],
[
("Plötzlich hat es stark geregnet.","Suddenly it rained heavily.","અચાનક ભારે વરસાદ થયો."),
("Zum Glück hatten wir einen Schirm.","Luckily we had an umbrella.","સૌભાગ્યે અમારી પાસે છત્રી હતી."),
("Leider war der Zug verspätet.","Unfortunately the train was late.","દુર્ભાગ્યે ટ્રેન મોડે હતી."),
("Deshalb sind wir zu Fuß gegangen.","Therefore we went on foot.","એથી અમે પગપાળા ગયા."),
("Am Ende waren alle glücklich.","In the end everyone was happy.","અંતે બધા ખુશ હતા."),
],
[
("💡","Story Flow","Use connectors to show time and reason.","FFF8E1"),
("📌","Emotion Words","Mix emotions to make stories vivid.","E8F0FE"),
("��🇪","Tip","Germans appreciate clear sequence and detail.","E6F4EA"),
],
["Write a short story (8 sentences) using 5 connectors.","Add 3 emotion words to a story.","Tell a memory from your childhood in German."],
[("'zum Glück' means?","Luckily"),("Translate: 'Leider war der Zug verspätet.'","Unfortunately the train was late."),("'die Angst' means?","Fear"),("Translate: 'Plötzlich hat es geregnet.'","Suddenly it rained."),("'die Wendung' means?","Turning point")],
"You can now tell richer stories in German with emotions and connectors."
))

# ── Day 178 ─────────────────────────────────────────────────────────────────
DAYS.append((178,
"FORMAL PRESENTATION Q&A 🙋",
"Handling questions",
"Today you learn phrases for handling questions after a presentation: asking for clarification, answering politely, and responding to criticism.",
"આજે તમે presentation પછી questions handle કરવા માટે phrases શીખશો.",
[
("die Frage","question","પ્રશ્ન","FRAH-geh"),
("die Antwort","answer","જવાબ","ANT-vort"),
("klarstellen","to clarify","સ્પષ્ટ કરવું","KLAHR-shtel-en"),
("nachfragen","to ask again","ફરી પૂછવું","NAKH-frah-gen"),
("darauf eingehen","to address","તે પર વાત કરવી","DAR-owf EYN-geh-en"),
("zustimmen","to agree","સહમત થવું","TSOO-shtim-en"),
("kritisieren","to criticize","ટીકા કરવી","kri-ti-TSIE-ren"),
("die Kritik","criticism","ટીકા","kri-TEEK"),
("die Anmerkung","remark","ટિપ્પણી","AN-merk-ung"),
("genau","exactly","બરાબર","geh-NOW"),
("unklar","unclear","અસ્પષ્ટ","oon-KLAHR"),
("erläutern","to explain","સમજાવવું","ehr-LOY-tern"),
("noch einmal","once again","ફરીથી","nokh EYN-mal"),
("vielen Dank","many thanks","ઘણો આભાર","FEE-len DANK"),
("gerne","gladly","ખુશીથી","GER-neh"),
],
[
("Vielen Dank für Ihre Frage.","Thank you for your question.","તમારા પ્રશ્ન માટે આભાર."),
("Könnten Sie das bitte noch einmal erklären?","Could you explain that again?","કૃપા કરીને ફરી સમજાવી શકો?"),
("Gerne gehe ich darauf ein.","I will gladly address that.","હું ખુશીથી તે પર વાત કરીશ."),
("Das ist eine berechtigte Kritik.","That is a justified criticism.","આ યોગ્ય ટીકા છે."),
("Lassen Sie mich das kurz klarstellen.","Let me clarify that briefly.","મને ટૂંકું સ્પષ્ટ કરવા દો."),
],
[
("💡","Polite Q&A","Start answers with thanks and a short summary.","FFF8E1"),
("📌","Handling Criticism","Acknowledge criticism before responding.","E8F0FE"),
("🇩🇪","Tip","Stay calm and structured in Q&A sessions.","E6F4EA"),
],
["Write 5 Q&A phrases for presentations.","Practice answering a critical question politely.","Create 3 clarification questions in German."],
[("'darauf eingehen' means?","to address"),("Translate: 'Vielen Dank für Ihre Frage.'","Thank you for your question."),("'die Kritik' means?","Criticism"),("Translate: 'Könnten Sie das bitte wiederholen?'","Could you repeat that, please?"),("'unklar' means?","Unclear")],
"You can now handle presentation questions professionally in German."
))

# ── Day 179 ─────────────────────────────────────────────────────────────────
DAYS.append((179,
"CULTURAL ETIQUETTE 🇩🇪",
"Formal vs informal behavior",
"Today you learn etiquette vocabulary for formal and informal situations in Germany: greetings, titles, and polite behavior.",
"આજે તમે Germany માં formal/informal behavior માટે etiquette vocabulary શીખશો.",
[
("die Höflichkeit","politeness","વિનમ્રતા","HÖF-likh-kite"),
("die Begrüßung","greeting","અભિવાદન","beh-GRÜ-sung"),
("die Anrede","form of address","સંબંધ",
"AN-reh-deh"),
("Herr / Frau","Mr / Ms","શ્રી/શ્રીમતી","HEHR / FROW"),
("Siezen","to use formal 'Sie'","Formal બોલવું","ZEE-tsen"),
("duzen","to use 'du'","તુ બોલવું","DOO-tsen"),
("der Titel","title","ઉપાધિ","TEE-tel"),
("die Visitenkarte","business card","વિઝિટિંગ કાર્ડ","vi-ZI-ten-kar-teh"),
("pünktlich","punctual","સમયપાલન","PÜNKT-likh"),
("das Geschenk","gift","ભેટ","geh-SHENK"),
("einladen","to invite","આમંત્રિત કરવું","EYN-lah-den"),
("bedanken","to thank","આભાર માનવું","beh-DANK-en"),
("sich entschuldigen","to apologize","માફી માંગવી","ent-SHOOL-di-gen"),
("angemessen","appropriate","યોગ્ય","AN-geh-mes-en"),
("respektvoll","respectful","સન્માનપૂર્વક","reh-SPEKT-fol"),
],
[
("In Deutschland ist Pünktlichkeit sehr wichtig.","In Germany punctuality is very important.","Germany માં સમયપાલન બહુ મહત્વનું છે."),
("Wir siezen unseren Chef.","We use formal 'Sie' with our boss.","અમે અમારા બોસને formal 'Sie' કહીએ છીએ."),
("Darf ich dich duzen?","May I use 'du' with you?","શું હું તને 'du' કહી શકું?"),
("Vielen Dank für die Einladung.","Thank you for the invitation.","આમંત્રણ માટે આભાર."),
("Respektvolles Verhalten ist angemessen.","Respectful behavior is appropriate.","સન્માનપૂર્વક વર્તન યોગ્ય છે."),
],
[
("💡","Sie vs du","Use Sie in formal situations until invited to use du.","FFF8E1"),
("📌","Punctuality","Arrive 5 minutes early for meetings.","E8F0FE"),
("🇩🇪","Gifts","Small gifts are common when visiting someone’s home.","E6F4EA"),
],
["Write 5 polite greetings for formal situations.","Create a short dialogue using Sie and du.","List 5 etiquette rules for Germany."],
[("'Siezen' means?","Use formal Sie"),("Translate: 'Darf ich dich duzen?'","May I use du with you?"),("'die Höflichkeit' means?","Politeness"),("Translate: 'Pünktlichkeit ist wichtig.'","Punctuality is important."),("'respektvoll' means?","Respectful")],
"You can now handle formal and informal etiquette in German culture."
))

# ── Day 180 ─────────────────────────────────────────────────────────────────
DAYS.append((180,
"MODULE REVIEW: DAYS 171–180 ✅",
"Communication & Presentation Review",
"You covered presentations, debate, negotiation, idioms, formal complaints, instructions, storytelling, Q&A, and etiquette. Today is review and consolidation.",
"તમે presentations, debate, negotiation, idioms, formal complaints, instructions, storytelling, Q&A અને etiquette cover કર્યા. આજે review day છે.",
[
("die Kommunikation","communication","સંચાર","ko-moo-ni-ka-TSYON"),
("die Präsentation","presentation","પ્રેઝન્ટેશન","pre-zen-ta-TSYON"),
("die Diskussion","discussion","ચર્ચા","dis-koos-SYON"),
("die Verhandlung","negotiation","વાટાઘાટ","fehr-HAND-lung"),
("die Beschwerde","complaint","ફરિયાદ","beh-SHVER-deh"),
("die Anleitung","instruction","માર્ગદર્શિકા","AN-ly-tung"),
("die Geschichte","story","વાર્તા","geh-SHIKH-teh"),
("die Frage","question","પ્રશ્ન","FRAH-geh"),
("die Kritik","criticism","ટીકા","kri-TEEK"),
("die Höflichkeit","politeness","વિનમ્રતા","HÖF-likh-kite"),
("das Argument","argument","દલીલ","ar-goo-MENT"),
("die Lösung","solution","ઉકેલ","LÖ-zoong"),
("die Struktur","structure","રચના","shtrook-TOOR"),
("klar","clear","સ્પષ્ટ","KLAHR"),
("sicher","confident","આત્મવિશ્વાસ","ZIKH-er"),
],
[
("Ich kann jetzt besser präsentieren.","I can present better now.","હું હવે વધુ સારી રીતે પ્રેઝેન્ટ કરી શકું છું."),
("Ich kann höflich verhandeln.","I can negotiate politely.","હું નમ્ર રીતે વાટાઘાટ કરી શકું છું."),
("Idiome machen meine Sprache natürlicher.","Idioms make my language more natural.","Idioms મારી ભાષા વધુ natural બનાવે છે."),
("Ich kann mit Kritik umgehen.","I can handle criticism.","હું ટીકા સંભાળી શકું છું."),
("Ich bin bereit für das nächste Modul.","I am ready for the next module.","હું આગળના module માટે તૈયાર છું."),
],
[
("🏆","Checkpoint","Give a 2‑minute presentation and answer 2 questions.","E8F0FE"),
("💡","Practice","Write a formal complaint and present a short argument.","FFF8E1"),
("🎯","Next","Days 181–200 focus on exam preparation and final mastery.","E6F4EA"),
],
["Write a short presentation (intro + 3 points + conclusion).", "Prepare 5 debate sentences with pro/contra.","Review 20 vocabulary words from this module."],
[("Translate: 'Ich kann besser präsentieren.'","I can present better."),("'die Verhandlung' means?","Negotiation"),("'die Kritik' means?","Criticism"),("Translate: 'Ich bin bereit.'","I am ready."),("What is next module about?","Exam preparation and final mastery")],
"Module review complete. Your communication German is clear and confident."
))

print("Days 171-180 appended")

# ── Day 181 ─────────────────────────────────────────────────────────────────
DAYS.append((181,
"READING STRATEGIES 📖",
"B1 exam reading skills",
"Today you learn how to read longer German texts: skimming, scanning, and identifying key information.",
"આજે તમે longer German texts વાંચવાની strategies શીખશો: skimming, scanning અને key information શોધવી.",
[
("das Lesen","reading","વાંચન","LAY-zen"),
("der Text","text","લખાણ","TEKST"),
("die Überschrift","heading","શીર્ષક","Ü-ber-shrift"),
("die Aussage","statement","કથન","OWS-zah-geh"),
("die Information","information","માહિતી","in-for-ma-TSYON"),
("der Kontext","context","સંદર્ભ","KON-tekst"),
("das Detail","detail","વિગત","deh-TAIL"),
("markieren","to highlight","હાઇલાઇટ કરવું","mar-KEE-ren"),
("überfliegen","to skim","ઝડપથી વાંચવું","Ü-ber-FLEE-gen"),
("nachschlagen","to look up","ડિક્શનરીમાં જોવું","NAKH-shlah-gen"),
("zusammenfassen","to summarize","સારાંશ કરવું","tsoo-ZAM-en-fas-en"),
("der Abschnitt","paragraph","પેરાગ્રાફ","AP-shnit"),
("das Thema","topic","વિષય","TAY-ma"),
("wichtig","important","મહત્વનું","VIKH-tikh"),
("unwichtig","unimportant","અમહત્વનું","OON-vikh-tikh"),
],
[
("Ich überfliege zuerst die Überschrift.","I first skim the heading.","હું પહેલા શીર્ષક ઝડપથી વાંચું છું."),
("Wichtige Informationen markiere ich.","I highlight important information.","હું મહત્વની માહિતી માર્ક કરું છું."),
("Der Kontext hilft beim Verstehen.","Context helps with understanding.","સંદર્ભ સમજવામાં મદદ કરે છે."),
("Ich fasse den Text kurz zusammen.","I summarize the text briefly.","હું લખાણનો ટૂંકો સારાંશ કરું છું."),
("Ich schlage unbekannte Wörter nach.","I look up unknown words.","હું અજાણ્યા શબ્દો ડિક્શનરીમાં જોઉં છું."),
],
[
("💡","Reading Tip","Read the heading and first sentence of each paragraph first.","FFF8E1"),
("📌","Scan","Look for dates, names, and numbers for quick answers.","E8F0FE"),
("🇩🇪","Exam","B1 reading tasks often focus on main ideas, not details.","E6F4EA"),
],
["Summarize a short German article in 5 sentences.","Underline key words in a paragraph.","Practice skimming: read a text in 2 minutes and explain its topic."],
[("'überfliegen' means?","To skim"),("Translate: 'Ich markiere die Information.'","I highlight the information."),("'der Abschnitt' means?","Paragraph"),("Translate: 'Der Kontext hilft.'","Context helps."),("'nachschlagen' means?","To look up")],
"You can now read German texts more efficiently using exam strategies."
))

# ── Day 182 ─────────────────────────────────────────────────────────────────
DAYS.append((182,
"LISTENING STRATEGIES 🎧",
"B1 exam listening skills",
"Today you learn how to improve German listening: focus on keywords, predict context, and catch numbers, dates, and names.",
"આજે તમે listening strategies શીખશો: keywords પર ધ્યાન, context અનુમાન, અને numbers/dates/names પકડવા.",
[
("das Hören","listening","સાંભળવું","HÖ-ren"),
("die Aufnahme","recording","રેકોર્ડિંગ","OWF-nah-meh"),
("das Gespräch","conversation","વાતચીત","geh-SHPRÄKH"),
("der Sprecher","speaker","વક્તા","SHPREKH-er"),
("die Information","information","માહિતી","in-for-ma-TSYON"),
("das Schlüsselwort","keyword","કીવર્ડ","SHLÜS-el-vort"),
("das Thema","topic","વિષય","TAY-ma"),
("die Zahl","number","અંક","TSAL"),
("das Datum","date","તારીખ","DAH-toom"),
("der Name","name","નામ","NAH-meh"),
("notieren","to note down","નોટ કરવું","no-TEE-ren"),
("verpassen","to miss","ચૂકવું","fehr-PAS-en"),
("die Geschwindigkeit","speed","ગતિ","geh-SHWIN-dig-kite"),
("langsam","slow","ધીમે","LANG-zam"),
("deutlich","clearly","સ્પષ્ટ","DOYT-likh"),
],
[
("Ich höre auf Schlüsselwörter.","I listen for keywords.","હું કીવર્ડ પર ધ્યાન આપું છું."),
("Ich notiere Zahlen und Namen.","I note down numbers and names.","હું અંકો અને નામો નોટ કરું છું."),
("Das Thema war Reisen.","The topic was travel.","વિષય મુસાફરી હતો."),
("Bitte sprechen Sie deutlich.","Please speak clearly.","કૃપા કરીને સ્પષ્ટ બોલો."),
("Ich habe die Information nicht verpasst.","I did not miss the information.","મેં માહિતી ચૂકી નથી."),
],
[
("💡","Listening Tip","Predict the topic from the title before you listen.","FFF8E1"),
("📌","Focus","Don’t panic if you miss a word — focus on the main idea.","E8F0FE"),
("🇩🇪","Exam","Listen twice if possible; use the second round to confirm details.","E6F4EA"),
],
["Listen to a short German podcast and write 5 keywords.","Practice writing down dates and times from audio.","Summarize a 1‑minute audio clip in 3 sentences."],
[("'das Schlüsselwort' means?","Keyword"),("Translate: 'Ich notiere Zahlen.'","I note numbers."),("'deutlich' means?","Clearly"),("Translate: 'Das Thema war…'","The topic was…"),("'verpassen' means?","To miss")],
"You can now listen more strategically and confidently in German."
))

# ── Day 183 ─────────────────────────────────────────────────────────────────
DAYS.append((183,
"WRITING STRATEGIES ✍️",
"B1 письма and emails",
"Today you focus on writing clear, structured German texts: formal and informal emails, with proper openings and closings.",
"આજે તમે clear અને structured German writing શીખશો: formal અને informal emails, યોગ્ય openings/closings સાથે.",
[
("der Text","text","લખાણ","TEKST"),
("die Struktur","structure","રચના","shtrook-TOOR"),
("der Absatz","paragraph","પેરાગ્રાફ","AP-zats"),
("die Einleitung","introduction","પરિચય","EYN-ly-tung"),
("der Schluss","closing","અંત","SHLOOS"),
("formell","formal","ઔપચારિક","for-MEL"),
("informell","informal","અનૌપચારિક","in-for-MEL"),
("Sehr geehrte…","Dear (formal)","આદરણીય","zayr geh-EHR-teh"),
("Liebe Grüße","best regards","શુભેચ્છા","LEE-beh GRÜ-seh"),
("Mit freundlichen Grüßen","sincerely","આદરસહિત","mit FROYNT-likh-en GRÜ-sen"),
("der Grund","reason","કારણ","GROONT"),
("die Bitte","request","વિનંતી","BIT-teh"),
("die Antwort","reply","જવાબ","ANT-vort"),
("der Betreff","subject","વિષય","beh-TREFF"),
("danke","thanks","આભાર","DANK-eh"),
],
[
("Der Text hat eine klare Struktur.","The text has a clear structure.","લખાણની રચના સ્પષ્ટ છે."),
("Ich schreibe eine formelle E-Mail.","I write a formal email.","હું ઔપચારિક ઇમેઈલ લખું છું."),
("Der Betreff lautet: Anfrage.","The subject is: Inquiry.","વિષય છે: વિનંતી."),
("Vielen Dank für Ihre Antwort.","Thank you for your reply.","તમારા જવાબ માટે આભાર."),
("Liebe Grüße aus Berlin.","Best regards from Berlin.","Berlin થી શુભેચ્છા."),
],
[
("��","Writing Tip","Short paragraphs make texts easier to read.","FFF8E1"),
("📌","Subject Line","Always add a clear Betreff.","E8F0FE"),
("🇩🇪","Exam","B1 writing tasks require clear structure and politeness.","E6F4EA"),
],
["Write a formal email requesting information.","Write an informal email inviting a friend.","Create a template with opening, body, closing."],
[("'formell' means?","Formal"),("Translate: 'Mit freundlichen Grüßen.'","Sincerely"),("'der Betreff' means?","Subject"),("Translate: 'Ich schreibe eine E-Mail.'","I write an email."),("'die Bitte' means?","Request")],
"You can now write structured German emails suitable for the B1 exam."
))

# ── Day 184 ─────────────────────────────────────────────────────────────────
DAYS.append((184,
"SPEAKING EXAM PRACTICE 🗣️",
"B1 oral task patterns",
"Today you learn common B1 speaking tasks: describing pictures, giving opinions, and asking questions.",
"આજે તમે B1 speaking task patterns શીખશો: picture description, opinions, અને questions પૂછવી.",
[
("das Bild","picture","ચિત્ર","BILT"),
("beschreiben","to describe","વર્ણવવું","beh-SHRY-ben"),
("die Meinung","opinion","મત","MY-nung"),
("der Eindruck","impression","છાપ","EYN-drook"),
("die Person","person","વ્યક્તિ","pehr-ZOHN"),
("die Situation","situation","પરિસ્થિતિ","zoo-ah-TSYON"),
("im Hintergrund","in the background","પૃષ્ઠભૂમિમાં","im HIN-ter-groont"),
("im Vordergrund","in the foreground","આગળ","im FOR-der-groont"),
("wahrscheinlich","probably","શક્યતઃ","VAHR-shhyne-likh"),
("vielleicht","maybe","કદાચ","FEE-laykht"),
("fragen","to ask","પુછવું","FRAH-gen"),
("antworten","to answer","જવાબ આપવો","ANT-vort-en"),
("sich vorstellen","to imagine","કલ્પના કરવી","zikh FOR-shtel-en"),
("zustimmen","to agree","સહમત થવું","TSOO-shtim-en"),
("ein Beispiel geben","to give an example","ઉદાહરણ આપવું","bye-SHPLEEL"),
],
[
("Auf dem Bild sehe ich zwei Personen.","In the picture I see two people.","ચિત્રમાં હું બે લોકો જોઈ રહ્યો છું."),
("Im Hintergrund ist ein Park.","In the background there is a park.","પૃષ્ઠભૂમિમાં એક પાર્ક છે."),
("Meiner Meinung nach ist die Situation freundlich.","In my opinion the situation is friendly.","મારી મત મુજબ પરિસ્થિતિ મિત્રપૂર્ણ છે."),
("Wahrscheinlich warten sie auf den Bus.","They are probably waiting for the bus.","શક્ય છે કે તેઓ બસની રાહ જોઈ રહ્યાં છે."),
("Kann ich Ihnen eine Frage stellen?","May I ask you a question?","શું હું તમને પ્રશ્ન પુછું?"),
],
[
("💡","Speaking Tip","Use simple sentences and speak clearly.","FFF8E1"),
("📌","Structure","Describe: who, where, what, feelings.","E8F0FE"),
("🇩🇪","Exam","B1 speaking requires interaction and questions.","E6F4EA"),
],
["Describe a photo in 6 sentences.","Practice asking 5 questions about a topic.","Give a short opinion on a daily-life topic."],
[("'im Vordergrund' means?","In the foreground"),("Translate: 'Auf dem Bild sehe ich…'","In the picture I see…"),("'der Eindruck' means?","Impression"),("Translate: 'Wahrscheinlich…'","Probably…"),("'beschreiben' means?","To describe")],
"You can now handle common B1 speaking tasks with confidence."
))

# ── Day 185 ─────────────────────────────────────────────────────────────────
DAYS.append((185,
"GRAMMAR REVIEW I 🔁",
"Konjunktiv II, passive, clauses",
"Today you review key grammar: Konjunktiv II, passive voice, infinitive clauses, and subordinate clauses.",
"આજે તમે મુખ્ય grammar review કરશો: Konjunktiv II, passive voice, infinitive clauses, subordinate clauses.",
[
("Konjunktiv II","subjunctive","કન્ઝુંક્ટિવ II","kon-YOONK-tiv"),
("Passiv","passive","પેસિવ","pa-SIV"),
("um…zu","in order to","માટે","oom-tsoo"),
("ohne…zu","without doing","બિન","OH-neh-tsoo"),
("anstatt…zu","instead of","બદલે","AN-shtat-tsoo"),
("weil","because","કારણ કે","VILE"),
("obwohl","although","છતાં","op-VOHL"),
("während","while","દરમ્યાન","VÄH-rent"),
("werden","to be (passive)","થવું","VEHR-den"),
("Partizip II","past participle","ભૂતરૂપ","par-tee-TSEEP"),
("Verb am Ende","verb at the end","verb અંતે","FERB am EN-deh"),
("Satzklammer","sentence bracket","વાક્ય બંધારણ","ZATS-klam-er"),
("TeKaMoLo","word order","ક્રમ","teh-ka-mo-lo"),
("Relativsatz","relative clause","રિલેટિવ ક્લોઝ","reh-lah-TEEF-zats"),
("Genitiv","genitive case","genitiv","geh-neh-TEEF"),
],
[
("Wenn ich Zeit hätte, würde ich mehr lesen.","If I had time, I would read more.","જો સમય હોત, તો હું વધુ વાંચત."),
("Der Brief wird geschrieben.","The letter is being written.","પત્ર લખાઈ રહ્યો છે."),
("Ich lerne Deutsch, um in Deutschland zu arbeiten.","I learn German to work in Germany.","હું Germanyમાં કામ કરવા માટે German શીખું છું."),
("Ich gehe, ohne zu bezahlen.","I leave without paying.","હું ચૂકવણી કર્યા વગર જઈ રહ્યો છું."),
("Obwohl es regnet, gehe ich spazieren.","Although it rains, I go for a walk.","છતાં વરસાદ છે, હું ચાલવા જાઉં છું."),
],
[
("💡","Review","Mix grammar topics in your sentences to build fluency.","FFF8E1"),
("📌","Verb Position","Remember: subordinate clause → verb at the end.","E8F0FE"),
("🇩🇪","Practice","Write 10 sentences combining these structures.","E6F4EA"),
],
["Write 10 sentences using at least 3 grammar structures each.","Make a list of your top 5 grammar weaknesses.","Practice 5 passive sentences with modals."],
[("Konjunktiv II formula?","würde + infinitive"),("Passive formula?","werden + Partizip II"),("Where does the verb go in subordinate clause?","At the end"),("Translate: 'Ich lerne, um zu bestehen.'","I learn in order to pass."),("'obwohl' means?","Although")],
"Grammar review complete — your core structures are stable."
))

# ── Day 186 ─────────────────────────────────────────────────────────────────
DAYS.append((186,
"VOCABULARY REVIEW I 📚",
"High-frequency B1 words",
"Today you review high-frequency B1 vocabulary across topics: work, travel, health, media, and emotions.",
"આજે તમે high-frequency B1 vocabulary review કરશો: work, travel, health, media, emotions.",
[
("die Herausforderung","challenge","પડકાર","he-ROWS-for-deh-rung"),
("die Erfahrung","experience","અનુભવ","ehr-FAH-roong"),
("die Entscheidung","decision","નિણર્ય","ent-SHAY-dung"),
("die Möglichkeit","possibility","શક્યતા","MÖG-likh-kite"),
("die Entwicklung","development","વિકાસ","ent-VIK-lung"),
("die Nachricht","message/news","સમાચાર","NAKH-rikh"),
("die Meinung","opinion","મત","MY-nung"),
("die Sicherheit","safety","સુરક્ષા","ZIKH-er-hite"),
("die Gesundheit","health","આરોગ્ય","geh-ZOON-hite"),
("die Verantwortung","responsibility","જવાબદારી","fehr-ANT-vort-oong"),
("die Planung","planning","યોજનાબદ્ધ","PLAH-nung"),
("die Beziehung","relationship","સંબંધ","beh-TSIE-hoong"),
("die Qualität","quality","ગુણવત્તા","kva-li-TÄT"),
("die Gesellschaft","society","સમાજ","geh-ZELL-shaft"),
("die Lösung","solution","ઉકેલ","LÖ-zoong"),
],
[
("Die Entscheidung war schwierig.","The decision was difficult.","નિણર્ય મુશ્કેલ હતો."),
("Gesundheit ist das Wichtigste.","Health is the most important.","આરોગ્ય સૌથી મહત્વનું છે."),
("Ich habe eine Lösung gefunden.","I found a solution.","હું ઉકેલ શોધ્યો છે."),
("Die Qualität ist sehr gut.","The quality is very good.","ગુણવત્તા ખૂબ સારી છે."),
("Meine Meinung ist anders.","My opinion is different.","મારો મત અલગ છે."),
],
[
("💡","Memory","Group words by topic and revise weekly.","FFF8E1"),
("📌","Frequency","These words appear in most B1 texts.","E8F0FE"),
("🇩🇪","Tip","Use each word in a sentence to remember it.","E6F4EA"),
],
["Write 10 sentences using the words above.","Create flashcards for 15 high-frequency words.","Make a mind map of B1 vocabulary topics."],
[("'die Möglichkeit' means?","Possibility"),("Translate: 'Die Qualität ist wichtig.'","Quality is important."),("'die Gesellschaft' means?","Society"),("Translate: 'Ich habe eine Lösung.'","I have a solution."),("'die Herausforderung' means?","Challenge")],
"Vocabulary review complete. High-frequency words are now stronger in your memory."
))

# ── Day 187 ─────────────────────────────────────────────────────────────────
DAYS.append((187,
"MOCK TEST: READING 🧪",
"B1 reading practice",
"Today you do a mock reading test. Focus on main ideas, headings, and context clues.",
"આજે તમે reading mock test કરો. મુખ્ય વિચારો, headings અને context clues પર ધ્યાન આપો.",
[
("die Aufgabe","task","કાર્ય","OWF-gah-beh"),
("das Verständnis","understanding","સમજ","fehr-SHTEND-nis"),
("richtig/falsch","true/false","સાચું/ખોટું","Rikh-tikh/FALSH"),
("die Aussage","statement","કથન","OWS-zah-geh"),
("die Auswahl","choice","પસંદગી","OWS-vahl"),
("der Hinweis","hint","સૂચન","HIN-vize"),
("der Kontext","context","સંદર્ભ","KON-tekst"),
("zusammenfassen","summarize","સારાંશ કરવું","tsoo-ZAM-en-fas-en"),
("überfliegen","skim","ઝડપથી વાંચવું","Ü-ber-FLEE-gen"),
("markieren","highlight","હાઇલાઇટ","mar-KEE-ren"),
("wichtig","important","મહત્વનું","VIKH-tikh"),
("unwichtig","unimportant","અમહત્વનું","OON-vikh-tikh"),
("die Zeit","time","સમય","TSYTE"),
("die Lösung","solution","ઉકેલ","LÖ-zoong"),
("die Prüfung","exam","પરીક્ષા","PRÜ-fung"),
],
[
("Ich lese zuerst die Überschrift.","I first read the heading.","હું પહેલા હેડિંગ વાંચું છું."),
("Ich markiere wichtige Informationen.","I highlight important information.","હું મહત્વની માહિતી માર્ક કરું છું."),
("Der Kontext hilft beim Verstehen.","Context helps understanding.","સંદર્ભ સમજવામાં મદદ કરે છે."),
("Ich überprüfe meine Antworten.","I check my answers.","હું મારા જવાબ તપાસું છું."),
("Die Zeit ist begrenzt.","Time is limited.","સમય સીમિત છે."),
],
[
("💡","Mock Tip","Use 2–3 minutes to skim before detailed reading.","FFF8E1"),
("📌","Strategy","Focus on keywords in each question.","E8F0FE"),
("🇩🇪","Exam","B1 reading is about gist + details.","E6F4EA"),
],
["Do a 15‑minute reading exercise from a B1 workbook.","Write a 5‑sentence summary of a text.","Check your answers and note mistakes."],
[("'richtig/falsch' means?","True/false"),("Translate: 'Ich markiere die Information.'","I highlight the information."),("'der Kontext' means?","Context"),("Translate: 'Die Zeit ist begrenzt.'","Time is limited."),("'überfliegen' means?","To skim")],
"Mock reading practice complete. Your reading strategies are exam-ready."
))

# ── Day 188 ─────────────────────────────────────────────────────────────────
DAYS.append((188,
"MOCK TEST: LISTENING 🎧",
"B1 listening practice",
"Today you do a mock listening test. Focus on key facts, numbers, and context.",
"આજે તમે listening mock test કરો. keywords, numbers અને context પર ધ્યાન આપો.",
[
("das Hörverstehen","listening comprehension","સાંભળીને સમજવું","HÖR-fehr-SHTEN"),
("die Aufnahme","recording","રેકોર્ડિંગ","OWF-nah-meh"),
("das Gespräch","conversation","વાતચીત","geh-SHPRÄKH"),
("die Aufgabe","task","કાર્ય","OWF-gah-beh"),
("die Auswahl","choice","પસંદગી","OWS-vahl"),
("die Zahl","number","અંક","TSAL"),
("das Datum","date","તારીખ","DAH-toom"),
("die Uhrzeit","time","સમય","OOR-tsyte"),
("das Detail","detail","વિગત","deh-TAIL"),
("das Ergebnis","result","પરિણામ","ehr-GAYP-nis"),
("notieren","to note down","નોટ કરવું","no-TEE-ren"),
("verpassen","to miss","ચૂકવું","fehr-PAS-en"),
("wiederholen","to repeat","પુનરાવર્તન","vee-der-HOH-len"),
("genau","exactly","બરાબર","geh-NOW"),
("klar","clear","સ્પષ્ટ","KLAHR"),
],
[
("Ich höre die Aufnahme zweimal.","I listen to the recording twice.","હું રેકોર્ડિંગ બે વાર સાંભળું છું."),
("Ich notiere Zahlen und Daten.","I note numbers and dates.","હું અંકો અને તારીખો નોંધું છું."),
("Die Details sind wichtig.","The details are important.","વિગતો મહત્વની છે."),
("Ich überprüfe meine Antworten.","I check my answers.","હું મારા જવાબ તપાસું છું."),
("Das Ergebnis ist gut.","The result is good.","પરિણામ સારું છે."),
],
[
("💡","Mock Tip","Listen for signal words: zuerst, danach, schließlich.","FFF8E1"),
("📌","Note","Write short keywords, not full sentences.","E8F0FE"),
("🇩🇪","Exam","B1 listening often includes announcements and dialogues.","E6F4EA"),
],
["Do a 10‑minute listening exercise and write 5 facts.","Practice writing down phone numbers and dates.","Summarize the audio in 3 sentences."],
[("'das Hörverstehen' means?","Listening comprehension"),("Translate: 'Ich notiere Zahlen.'","I note numbers."),("'die Uhrzeit' means?","Time"),("Translate: 'Ich höre zweimal.'","I listen twice."),("'verpassen' means?","To miss")],
"Mock listening practice complete. Your listening skills are exam-ready."
))

# ── Day 189 ─────────────────────────────────────────────────────────────────
DAYS.append((189,
"MOCK TEST: WRITING ✍️",
"B1 writing practice",
"Today you do a mock writing task: write a formal email or informal message with clear structure.",
"આજે તમે writing mock task કરો: formal email અથવા informal message સ્પષ્ટ structure સાથે લખો.",
[
("die Aufgabe","task","કાર્ય","OWF-gah-beh"),
("die Struktur","structure","રચના","shtrook-TOOR"),
("der Absatz","paragraph","પેરાગ્રાફ","AP-zats"),
("die Einleitung","introduction","પરિચય","EYN-ly-tung"),
("der Schluss","closing","અંત","SHLOOS"),
("formell","formal","ઔપચારિક","for-MEL"),
("informell","informal","અનૌપચારિક","in-for-MEL"),
("der Betreff","subject","વિષય","beh-TREFF"),
("die Bitte","request","વિનંતી","BIT-teh"),
("die Antwort","reply","જવાબ","ANT-vort"),
("die Frist","deadline","સમયસીમા","FRIST"),
("danke","thanks","આભાર","DANK-eh"),
("höflich","polite","વિનમ્ર","HÖF-likh"),
("klar","clear","સ્પષ્ટ","KLAHR"),
("senden","to send","મોકલવું","ZEN-den"),
],
[
("Ich schreibe einen formellen Brief.","I write a formal letter.","હું ઔપચારિક પત્ર લખું છું."),
("Der Betreff ist klar.","The subject is clear.","વિષય સ્પષ્ટ છે."),
("Die Bitte steht im zweiten Absatz.","The request is in the second paragraph.","વિનંતી બીજા પેરાગ્રાફમાં છે."),
("Vielen Dank für Ihre Antwort.","Thank you for your reply.","તમારા જવાબ માટે આભાર."),
("Ich sende die E-Mail heute.","I send the email today.","હું આજે ઇમેઈલ મોકલૂ છું."),
],
[
("💡","Writing Tip","Always include a clear request and reason.","FFF8E1"),
("📌","Formal Closing","Use 'Mit freundlichen Grüßen' in formal emails.","E8F0FE"),
("🇩🇪","Exam","B1 writing requires clarity and polite tone.","E6F4EA"),
],
["Write a formal email requesting information about a course.","Write an informal message to cancel a meeting.","Check your text for structure and mistakes."],
[("'der Betreff' means?","Subject"),("Translate: 'Ich schreibe einen formellen Brief.'","I write a formal letter."),("'die Bitte' means?","Request"),("Translate: 'Ich sende die E-Mail.'","I send the email."),("'formell' means?","Formal")],
"Mock writing practice complete. Your written German is exam-ready."
))

# ── Day 190 ─────────────────────────────────────────────────────────────────
DAYS.append((190,
"MODULE REVIEW: DAYS 181–190 ✅",
"Exam preparation review",
"You practiced reading, listening, writing, speaking, and grammar review. Today is consolidation and confidence building.",
"તમે reading, listening, writing, speaking અને grammar review કર્યું. આજે consolidation day છે.",
[
("die Prüfung","exam","પરીક્ષા","PRÜ-fung"),
("die Vorbereitung","preparation","તૈયારી","for-bah-RY-tung"),
("die Strategie","strategy","રણનીતિ","shtra-TEH-gee"),
("die Aufgabe","task","કાર્ય","OWF-gah-beh"),
("die Lösung","solution","ઉકેલ","LÖ-zoong"),
("das Ergebnis","result","પરિણામ","ehr-GAYP-nis"),
("sicher","confident","આત્મવિશ્વાસ","ZIKH-er"),
("klar","clear","સ્પષ્ટ","KLAHR"),
("schnell","fast","ઝડપી","SHNEL"),
("genau","accurate","ચોક્કસ","geh-NOW"),
("die Wiederholung","review","પુનરાવર્તન","VEE-der-hoh-lung"),
("das Ziel","goal","લક્ષ્ય","TSEEL"),
("die Zeit","time","સમય","TSYTE"),
("die Übung","exercise","અભ્યાસ","Ü-bung"),
("die Sicherheit","confidence","આત્મવિશ્વાસ","ZIKH-er-hite"),
],
[
("Ich fühle mich sicherer für die Prüfung.","I feel more confident for the exam.","હું પરીક્ષા માટે વધુ આત્મવિશ્વાસી છું."),
("Die Strategie ist klar.","The strategy is clear.","રણનીતિ સ્પષ્ટ છે."),
("Ich übe jeden Tag ein bisschen.","I practise a little every day.","હું રોજ થોડું અભ્યાસ કરું છું."),
("Das Ergebnis wird besser.","The result is getting better.","પરિણામ વધુ સારું બની રહ્યું છે."),
("Ich bin bereit für das Finale.","I am ready for the finale.","હું અંતિમ માટે તૈયાર છું."),
],
[
("🏆","Checkpoint","If you can do a full mock test, you are ready for final days.","E8F0FE"),
("💡","Routine","Short daily practice beats long irregular sessions.","FFF8E1"),
("🎯","Next","Days 191–200 are the final stretch and graduation.","E6F4EA"),
],
["Do a full B1 mock test (reading/listening/writing).", "Write a 2‑minute speaking summary about your progress.","Review your top 20 weak words."],
[("'die Prüfung' means?","Exam"),("Translate: 'Ich fühle mich sicherer.'","I feel more confident."),("'die Strategie' means?","Strategy"),("Translate: 'Ich übe jeden Tag.'","I practise every day."),("What is next?","Final stretch and graduation")],
"Module review complete. You are now ready for the final 10 days."
))

print("Days 181-190 appended")

# ── Day 191 ─────────────────────────────────────────────────────────────────
DAYS.append((191,
"GRAMMAR REVIEW II 🔁",
"Full B1 grammar sweep",
"Today you review all major B1 structures: Konjunktiv II, passive, clauses, adjective endings, and prepositional verbs.",
"આજે તમે તમામ B1 grammar structures review કરશો: Konjunktiv II, passive, clauses, adjective endings અને prepositional verbs.",
[
("Konjunktiv II","subjunctive","કન્ઝુંક્ટિવ II","kon-YOONK-tiv"),
("Passiv","passive","પેસિવ","pa-SIV"),
("Relativsatz","relative clause","રિલેટિવ ક્લોઝ","reh-lah-TEEF-zats"),
("Nebensatz","subordinate clause","અધિન વાક્ય","NAY-ben-zats"),
("Adjektivendungen","adjective endings","adjective endings","ad-YEK-tiv-END-ung-en"),
("Präpositionen","prepositions","પ્રિપોઝિશન્સ","prä-po-zi-TSYON-en"),
("TeKaMoLo","word order","ક્રમ","teh-ka-mo-lo"),
("Genitiv","genitive case","genitiv","geh-neh-TEEF"),
("N-Deklination","N-declension","N-ડિક્લેન્શન","EN-dek-li-NA-tsyohn"),
("Infinitivsatz","infinitive clause","ઇન્ફિનિટીવ ક્લોઝ","in-fin-ee-TEEF-zats"),
("weil","because","કારણ કે","VILE"),
("obwohl","although","છતાં","op-VOHL"),
("während","while","દરમ્યાન","VÄH-rent"),
("dessen/deren","whose","જેનું/જેની","DES-en/DEE-ren"),
("sich freuen auf","look forward","રાહ જોવી","FROY-en owf"),
],
[
("Wenn ich Zeit hätte, würde ich mehr lesen.","If I had time, I would read more.","જો સમય હોત, તો હું વધુ વાંચત."),
("Der Vertrag wird unterschrieben.","The contract is being signed.","કરાર સહી થઈ રહ્યો છે."),
("Der Mann, dessen Auto dort steht, ist mein Nachbar.","The man whose car is there is my neighbor.","જે માણસની કાર ત્યાં છે તે મારો પડોશી છે."),
("Ich lerne, um die Prüfung zu bestehen.","I learn in order to pass the exam.","હું પરીક્ષા પાસ કરવા માટે શીખું છું."),
("Obwohl es regnet, gehen wir spazieren.","Although it rains, we go for a walk.","છતાં વરસાદ છે, અમે ચાલવા જઈએ છીએ."),
],
[
("💡","Review","Mix different grammar points in one paragraph.","FFF8E1"),
("📌","Accuracy","Check verb position in every subordinate clause.","E8F0FE"),
("🇩🇪","Practice","Write a 10‑sentence story using at least 5 structures.","E6F4EA"),
],
["Write 10 mixed-grammar sentences.","Underline verb positions in your sentences.","Make a checklist of grammar topics you mastered."],
[("Konjunktiv II formula?","würde + infinitive"),("Passive formula?","werden + Partizip II"),("Where does the verb go in a subordinate clause?","At the end"),("'dessen/deren' are used for?","Genitive relative clauses"),("TeKaMoLo stands for?","Time–Cause–Manner–Place")],
"Your grammar foundation is strong and ready for the final stage."
))

# ── Day 192 ─────────────────────────────────────────────────────────────────
DAYS.append((192,
"VOCABULARY REVIEW II 📦",
"Final high-frequency words",
"Today you review high-frequency vocabulary across all B1 topics to strengthen your active word bank.",
"આજે તમે બધા B1 topics માટે high-frequency vocab review કરશો.",
[
("die Möglichkeit","possibility","શક્યતા","MÖG-likh-kite"),
("die Erfahrung","experience","અનુભવ","ehr-FAH-roong"),
("die Herausforderung","challenge","પડકાર","he-ROWS-for-deh-rung"),
("die Entscheidung","decision","નિણર્ય","ent-SHAY-dung"),
("die Verantwortung","responsibility","જવાબદારી","fehr-ANT-vort-oong"),
("die Qualität","quality","ગુણવત્તા","kva-li-TÄT"),
("die Entwicklung","development","વિકાસ","ent-VIK-lung"),
("die Gesellschaft","society","સમાજ","geh-ZELL-shaft"),
("die Lösung","solution","ઉકેલ","LÖ-zoong"),
("die Beziehung","relationship","સંબંધ","beh-TSIE-hoong"),
("die Sicherheit","safety","સુરક્ષા","ZIKH-er-hite"),
("die Gesundheit","health","આરોગ્ય","geh-ZOON-hite"),
("die Planung","planning","યોજનાબદ્ધ","PLAH-nung"),
("die Kommunikation","communication","સંચાર","ko-moo-ni-ka-TSYON"),
("das Ziel","goal","લક્ષ્ય","TSEEL"),
],
[
("Die Entscheidung war nicht leicht.","The decision was not easy.","નિણર્ય સરળ નહોતો."),
("Gesundheit ist mein wichtigstes Ziel.","Health is my most important goal.","આરોગ્ય મારું સૌથી મહત્વનું લક્ષ્ય છે."),
("Gute Kommunikation ist wichtig.","Good communication is important.","સારો સંચાર મહત્વનો છે."),
("Wir suchen eine Lösung.","We are looking for a solution.","અમે ઉકેલ શોધીએ છીએ."),
("Die Gesellschaft verändert sich.","Society is changing.","સમાજ બદલાઈ રહ્યો છે."),
],
[
("💡","Tip","Repeat these words weekly in sentences.","FFF8E1"),
("📌","Active Use","Speak them aloud to make them active vocabulary.","E8F0FE"),
("🇩🇪","Exam","These words appear in most B1 tasks.","E6F4EA"),
],
["Write 10 sentences using this vocabulary.","Create flashcards for 15 high-frequency words.","Record a 1‑minute speech using at least 8 of these words."],
[("'die Möglichkeit' means?","Possibility"),("Translate: 'Wir suchen eine Lösung.'","We are looking for a solution."),("'die Gesellschaft' means?","Society"),("Translate: 'Gute Kommunikation ist wichtig.'","Good communication is important."),("'das Ziel' means?","Goal")],
"Your high-frequency vocabulary is solid and ready for final tests."
))

# ── Day 193 ─────────────────────────────────────────────────────────────────
DAYS.append((193,
"MOCK SPEAKING TEST 🗣️",
"Simulated B1 speaking",
"Today you simulate a B1 speaking test: introduction, picture description, and discussion.",
"આજે તમે B1 speaking test simulate કરશો: introduction, picture description અને discussion.",
[
("die Prüfung","exam","પરીક્ષા","PRÜ-fung"),
("die Aufgabe","task","કાર્ય","OWF-gah-beh"),
("die Beschreibung","description","વર્ણન","beh-SHRY-bung"),
("die Meinung","opinion","મત","MY-nung"),
("der Eindruck","impression","છાપ","EYN-drook"),
("die Frage","question","પ્રશ્ન","FRAH-geh"),
("antworten","to answer","જવાબ આપવો","ANT-vort-en"),
("diskutieren","to discuss","ચર્ચા કરવી","dis-koo-TEE-ren"),
("zustimmen","to agree","સહમત થવું","TSOO-shtim-en"),
("ablehnen","to disagree","નકારવું","AP-lay-nen"),
("vorschlagen","to suggest","સૂચવવું","FOR-shlah-gen"),
("vergleichen","to compare","તુલના કરવી","fehr-GLY-khen"),
("sich vorstellen","to imagine","કલ્પના કરવી","zikh FOR-shtel-en"),
("nennen","to mention","ઉલ્લેખ કરવો","NEN-en"),
("abschließen","to finish","સમાપ્ત કરવું","AP-shlee-sen"),
],
[
("Ich stelle mich kurz vor.","I introduce myself briefly.","હું પોતાનો ટૂંકો પરિચય આપું છું."),
("Auf dem Bild sehe ich eine Familie.","In the picture I see a family.","ચિત્રમાં હું એક પરિવાર જોઈ રહ્યો છું."),
("Meiner Meinung nach ist das sehr wichtig.","In my opinion this is very important.","મારી મત મુજબ આ ખૂબ મહત્વનું છે."),
("Ich stimme zu, aber ich habe einen anderen Punkt.","I agree, but I have another point.","હું સહમત છું, પરંતુ મારી પાસે બીજો મુદ્દો છે."),
("Zum Schluss fasse ich kurz zusammen.","Finally I summarize briefly.","અંતે હું ટૂંકું સારાંશ કરું છું."),
],
[
("💡","Speaking Tip","Use simple, clear sentences and connect ideas.","FFF8E1"),
("📌","Timing","Keep each part within the time limit.","E8F0FE"),
("🇩🇪","Exam","Be friendly and interactive; ask a question to your partner.","E6F4EA"),
],
["Practice a 2‑minute introduction.","Describe a picture in 6 sentences.","Do a 3‑minute discussion about travel."],
[("'die Beschreibung' means?","Description"),("Translate: 'Ich stelle mich vor.'","I introduce myself."),("'diskutieren' means?","To discuss"),("Translate: 'Zum Schluss fasse ich zusammen.'","Finally I summarize."),("'vorschlagen' means?","To suggest")],
"Mock speaking practice complete. Your speaking confidence is high."
))

# ── Day 194 ─────────────────────────────────────────────────────────────────
DAYS.append((194,
"MOCK WRITING: FORMAL ✉️",
"Complaint or request letter",
"Today you practice a formal writing task: a complaint or request with clear structure and polite tone.",
"આજે તમે formal writing task practice કરશો: complaint અથવા request letter, સ્પષ્ટ structure સાથે.",
[
("formell","formal","ઔપચારિક","for-MEL"),
("die Beschwerde","complaint","ફરિયાદ","beh-SHVER-deh"),
("die Anfrage","request","વિનંતી","AN-frah-geh"),
("der Grund","reason","કારણ","GROONT"),
("die Bitte","request","વિનંતી","BIT-teh"),
("die Lösung","solution","ઉકેલ","LÖ-zoong"),
("der Betreff","subject","વિષય","beh-TREFF"),
("die Frist","deadline","સમયસીમા","FRIST"),
("die Bestätigung","confirmation","પુષ્ટિ","beh-SHTÄ-ti-gung"),
("höflich","polite","વિનમ્ર","HÖF-likh"),
("entschuldigen","to apologize","માફી માગવી","ent-SHOOL-di-gen"),
("erhalten","to receive","પ્રાપ્ત કરવું","ehr-HAL-ten"),
("senden","to send","મોકલવું","ZEN-den"),
("danke","thanks","આભાર","DANK-eh"),
("Antwort","reply","જવાબ","ANT-vort"),
],
[
("Betreff: Reklamation",
"Subject: Complaint",
"વિષય: ફરિયાદ"),
("Ich bitte um eine schnelle Lösung.","I request a quick solution.","હું ઝડપી ઉકેલ માંગું છું."),
("Vielen Dank für Ihre Antwort.","Thank you for your reply.","તમારા જવાબ માટે આભાર."),
("Ich habe das Produkt am Montag erhalten.","I received the product on Monday.","મેં પ્રોડક્ટ સોમવારે પ્રાપ્ત કરી છે."),
("Bitte bestätigen Sie den Erhalt.","Please confirm receipt.","કૃપા કરીને પ્રાપ્તી પુષ્ટિ કરો."),
],
[
("💡","Structure","Problem → request → deadline → polite closing.","FFF8E1"),
("📌","Tone","Be factual and respectful.","E8F0FE"),
("🇩🇪","Exam","B1 writing expects clear structure and correct formality.","E6F4EA"),
],
["Write a formal complaint about a defective product.","Write a request letter for information.","Check your letter for structure and polite closing."],
[("'der Betreff' means?","Subject"),("Translate: 'Ich bitte um eine Lösung.'","I request a solution."),("'höflich' means?","Polite"),("Translate: 'Bitte bestätigen Sie.'","Please confirm."),("'die Frist' means?","Deadline")],
"Formal writing practice complete. Your structure and tone are exam-ready."
))

# ── Day 195 ─────────────────────────────────────────────────────────────────
DAYS.append((195,
"MOCK WRITING: INFORMAL 📨",
"Messages to friends",
"Today you practice informal writing: inviting, canceling, or giving updates to friends.",
"આજે તમે informal writing practice કરશો: મિત્રો ને invite કરવું, cancel કરવું અથવા updates આપવું.",
[
("informell","informal","અનૌપચારિક","in-for-MEL"),
("einladen","to invite","આમંત્રિત કરવું","EYN-lah-den"),
("absagen","to cancel","નકારવું","AP-zah-gen"),
("verschieben","to postpone","મુલતવી કરવું","fehr-SHEE-ben"),
("sich treffen","to meet","મળવું","zikh TREF-fen"),
("die Zeit","time","સમય","TSYTE"),
("der Ort","place","જગ્યા","ORT"),
("leider","unfortunately","દુર્ભાગ્યે","LY-der"),
("gern","gladly","ખુશીથી","GERN"),
("die Nachricht","message","સંદેશ","NAKH-rikh"),
("schicken","to send","મોકલવું","SHIK-en"),
("antworten","to answer","જવાબ આપવો","ANT-vort-en"),
("danke","thanks","આભાર","DANK-eh"),
("Grüße","greetings","શુભેચ્છા","GRÜ-seh"),
("bis bald","see you soon","جل्दी મળીએ","bis BALT"),
],
[
("Hi! Hast du morgen Zeit?","Hi! Do you have time tomorrow?","હાય! કાલે સમય છે?"),
("Leider muss ich absagen.","Unfortunately I have to cancel.","દુર્ભાગ્યે મને cancel કરવું પડશે."),
("Können wir das Treffen verschieben?","Can we postpone the meeting?","શું આપણે મળવાનું મુલતવી રાખીએ?"),
("Ich komme gern.","I'd love to come.","હું ખુશીથી આવીશ."),
("Bis bald! Viele Grüße.","See you soon! Best regards.","جل્દી મળીએ! શુભેચ્છા."),
],
[
("💡","Informal Tone","Use 'du', short sentences, and friendly greetings.","FFF8E1"),
("📌","Emoji","Emojis are common in informal messages, but keep it balanced.","E8F0FE"),
("🇩🇪","Common","WhatsApp is the most common informal channel.","E6F4EA"),
],
["Write an invitation to a friend.","Write a short message canceling plans politely.","Send a thank-you message in German."],
[("'absagen' means?","To cancel"),("Translate: 'Können wir verschieben?'","Can we postpone?"),("'bis bald' means?","See you soon"),("Translate: 'Ich komme gern.'","I’d love to come."),("'Grüße' means?","Greetings")],
"Informal writing practice complete. Your friendly German is natural and clear."
))

# ── Day 196 ─────────────────────────────────────────────────────────────────
DAYS.append((196,
"FINAL LISTENING DRILL 🎧",
"Speed and comprehension",
"Today you do a final listening drill focusing on speed, clarity, and recognizing key phrases.",
"આજે તમે final listening drill કરશો: speed, clarity અને key phrases ઓળખવા પર ધ્યાન.",
[
("das Tempo","speed","ઝડપ","TEM-po"),
("schnell","fast","ઝડપી","SHNEL"),
("langsam","slow","ધીમે","LANG-zam"),
("deutlich","clear","સ્પષ્ટ","DOYT-likh"),
("schwer verständlich","hard to understand","સમજવામાં કઠિન","SHVER fer-SHTEND-likh"),
("der Akzent","accent","ઉચ્ચાર","ak-TSENT"),
("die Wiederholung","repetition","પુનરાવર્તન","VEE-der-hoh-lung"),
("die Zusammenfassung","summary","સારાંશ","tsoo-ZAM-en-fas-ung"),
("das Stichwort","keyword","કીવર્ડ","SHTIKH-vort"),
("das Detail","detail","વિગત","deh-TAIL"),
("die Frage","question","પ્રશ્ન","FRAH-geh"),
("die Antwort","answer","જવાબ","ANT-vort"),
("verstehen","to understand","સમજવું","fehr-SHTAY-en"),
("notieren","to note","નોટ કરવું","no-TEE-ren"),
("verbessern","to improve","સુધારવું","fehr-BES-ern"),
],
[
("Das Tempo ist heute schneller.","The speed is faster today.","આજે ઝડપ વધુ છે."),
("Bitte sprechen Sie deutlich.","Please speak clearly.","કૃપા કરીને સ્પષ્ટ બોલો."),
("Ich notiere die Stichwörter.","I note the keywords.","હું કીવર્ડ નોટ કરું છું."),
("Ich fasse den Text kurz zusammen.","I summarize the text briefly.","હું લખાણનો ટૂંકો સારાંશ કરું છું."),
("Ich verstehe immer mehr.","I understand more and more.","હું વધુ અને વધુ સમજું છું."),
],
[
("💡","Drill","Listen once for gist, once for details.","FFF8E1"),
("📌","Accent","Different accents are normal — focus on context.","E8F0FE"),
("🇩🇪","Tip","Shadowing helps improve listening and speaking together.","E6F4EA"),
],
["Do a 5‑minute shadowing exercise.","Write a 3‑sentence summary after listening.","Repeat a 1‑minute audio until you understand 90%."],
[("'das Tempo' means?","Speed"),("Translate: 'Ich notiere die Stichwörter.'","I note the keywords."),("'deutlich' means?","Clearly"),("Translate: 'Ich verstehe mehr.'","I understand more."),("'der Akzent' means?","Accent")],
"Final listening drill complete. Your comprehension is strong and fast."
))

# ── Day 197 ─────────────────────────────────────────────────────────────────
DAYS.append((197,
"FINAL READING DRILL 📄",
"Speed + accuracy",
"Today you do a final reading drill: faster reading with accurate answers.",
"આજે તમે final reading drill કરશો: ઝડપી વાંચન અને સાચા જવાબ.",
[
("das Lesen","reading","વાંચન","LAY-zen"),
("die Geschwindigkeit","speed","ઝડપ","geh-SHWIN-dig-kite"),
("die Genauigkeit","accuracy","ચોકસાઈ","geh-NOW-ig-kite"),
("der Absatz","paragraph","પેરાગ્રાફ","AP-zats"),
("die Frage","question","પ્રશ્ન","FRAH-geh"),
("die Antwort","answer","જવાબ","ANT-vort"),
("die Information","information","માહિતી","in-for-ma-TSYON"),
("das Detail","detail","વિગત","deh-TAIL"),
("das Thema","topic","વિષય","TAY-ma"),
("zusammenfassen","summarize","સારાંશ કરવું","tsoo-ZAM-en-fas-en"),
("überfliegen","skim","ઝડપથી વાંચવું","Ü-ber-FLEE-gen"),
("markieren","highlight","હાઇલાઇટ","mar-KEE-ren"),
("richtig/falsch","true/false","સાચું/ખોટું","Rikh-tikh/FALSH"),
("der Kontext","context","સંદર્ભ","KON-tekst"),
("die Lösung","solution","ઉકેલ","LÖ-zoong"),
],
[
("Ich lese den Text schneller als früher.","I read the text faster than before.","હું પહેલા કરતાં ઝડપથી વાંચું છું."),
("Die Antworten sind genauer.","The answers are more accurate.","જવાબ વધુ ચોક્કસ છે."),
("Der Kontext hilft bei schwierigen Fragen.","Context helps with difficult questions.","સંદર્ભ મુશ્કેલ પ્રશ્નોમાં મદદ કરે છે."),
("Ich markiere wichtige Details.","I highlight important details.","હું મહત્વની વિગતો માર્ક કરું છું."),
("Ich überprüfe alle Lösungen.","I check all solutions.","હું બધા ઉકેલો તપાસું છું."),
],
[
("💡","Drill","Set a timer to practice speed.","FFF8E1"),
("📌","Accuracy","Accuracy matters more than speed in exams.","E8F0FE"),
("🇩🇪","Tip","Read the questions first, then the text.","E6F4EA"),
],
["Do a 10‑minute timed reading task.","Write 5 correct answers from a short text.","Summarize the text in 3 sentences."],
[("'die Genauigkeit' means?","Accuracy"),("Translate: 'Ich lese schneller.'","I read faster."),("'überfliegen' means?","To skim"),("Translate: 'Ich überprüfe die Lösungen.'","I check the solutions."),("'richtig/falsch' means?","True/false")],
"Final reading drill complete. Your speed and accuracy are strong."
))

# ── Day 198 ─────────────────────────────────────────────────────────────────
DAYS.append((198,
"SELF-ASSESSMENT & GOALS 🎯",
"Reflect and plan",
"Today you evaluate your progress and set clear goals for the next stage (B1 exam or B2 bridge).",
"આજે તમે તમારું મૂલ્યાંકન કરો અને આગળના stage માટે goals નક્કી કરો.",
[
("die Selbsteinschätzung","self-assessment","સ્વમૂલ્યાંકન","ZELPST-ayn-SHET-sung"),
("das Ziel","goal","લક્ષ્ય","TSEEL"),
("der Fortschritt","progress","પ્રગતિ","FORT-shrit"),
("die Stärke","strength","મજબૂતી","SHTER-keh"),
("die Schwäche","weakness","કમજોરી","SHVÄKH-eh"),
("die Motivation","motivation","પ્રેરણા","moh-ti-va-TSYON"),
("der Plan","plan","યોજનાબદ્ધ","PLAHN"),
("die Strategie","strategy","રણનીતિ","shtra-TEH-gee"),
("das Ergebnis","result","પરિણામ","ehr-GAYP-nis"),
("regelmäßig","regularly","નિયમિત","RAY-gel-mä-sig"),
("konsequent","consistent","સતત","kon-ze-KVENT"),
("die Verbesserung","improvement","સુધારો","fehr-BES-er-ung"),
("die Übung","practice","અભ્યાસ","Ü-bung"),
("das Feedback","feedback","પ્રતિસાદ","FEEB-bek"),
("die Vorbereitung","preparation","તૈયારી","for-bah-RY-tung"),
],
[
("Meine Stärke ist das Sprechen.","My strength is speaking.","મારી મજબૂતી બોલવું છે."),
("Meine Schwäche ist das Schreiben.","My weakness is writing.","મારી કમજોરી લખવું છે."),
("Ich setze mir klare Ziele.","I set myself clear goals.","હું સ્પષ્ટ લક્ષ્યો નક્કી કરું છું."),
("Ich übe regelmäßig und konsequent.","I practise regularly and consistently.","હું નિયમિત અને સતત અભ્યાસ કરું છું."),
("Feedback hilft mir, mich zu verbessern.","Feedback helps me improve.","પ્રતિસાદ મને સુધારવામાં મદદ કરે છે."),
],
[
("💡","Reflection","Honest self-assessment helps you improve faster.","FFF8E1"),
("📌","Plan","Set short-term and long-term goals.","E8F0FE"),
("🇩🇪","Next Step","Decide if you want to take the B1 exam soon.","E6F4EA"),
],
["Write your 3 strengths and 3 weaknesses in German.","Create a 30‑day study plan.","Write a paragraph about your German journey."],
[("'die Selbsteinschätzung' means?","Self-assessment"),("Translate: 'Ich setze mir Ziele.'","I set goals."),("'konsequent' means?","Consistent"),("Translate: 'Feedback hilft mir.'","Feedback helps me."),("'die Verbesserung' means?","Improvement")],
"You have a clear picture of your progress and goals."
))

# ── Day 199 ─────────────────────────────────────────────────────────────────
DAYS.append((199,
"FINAL MOCK TEST DAY 📝",
"Full B1 practice",
"Today you take a full mock test: reading, listening, writing, and speaking. Focus on time management and confidence.",
"આજે તમે full mock test કરો: reading, listening, writing, speaking. સમય વ્યવસ્થાપન અને આત્મવિશ્વાસ પર ધ્યાન આપો.",
[
("die Prüfung","exam","પરીક્ષા","PRÜ-fung"),
("die Vorbereitung","preparation","તૈયારી","for-bah-RY-tung"),
("die Zeit","time","સમય","TSYTE"),
("die Aufgabe","task","કાર્ય","OWF-gah-beh"),
("die Lösung","solution","ઉકેલ","LÖ-zoong"),
("die Strategie","strategy","રણનીતિ","shtra-TEH-gee"),
("die Konzentration","concentration","ધ્યાન","kon-tsen-tra-TSYON"),
("die Pause","break","વિરામ","POW-zeh"),
("die Bewertung","evaluation","મૂલ્યાંકન","beh-VER-tung"),
("die Verbesserung","improvement","સુધારો","fehr-BES-er-ung"),
("die Aufgabe erledigen","to complete a task","કાર્ય પૂર્ણ કરવું","er-LED-igen"),
("genau","accurate","ચોક્કસ","geh-NOW"),
("pünktlich","on time","સમય પર","PÜNKT-likh"),
("ruhig","calm","શાંત","ROO-ikh"),
("bereit","ready","તૈયાર","beh-RITE"),
],
[
("Ich mache heute einen vollständigen Test.","I do a full test today.","હું આજે સંપૂર્ણ ટેસ્ટ કરું છું."),
("Zeitmanagement ist wichtig.","Time management is important.","સમય વ્યવસ્થાપન મહત્વનું છે."),
("Ich bleibe ruhig und konzentriert.","I stay calm and focused.","હું શાંત અને એકાગ્ર રહું છું."),
("Nach dem Test analysiere ich die Fehler.","After the test I analyze the mistakes.","ટેસ્ટ પછી હું ભૂલો વિશ્લેષણ કરું છું."),
("Ich bin bereit für die Prüfung.","I am ready for the exam.","હું પરીક્ષા માટે તૈયાર છું."),
],
[
("💡","Mock Tip","Simulate real exam conditions: time limit, no phone.","FFF8E1"),
("📌","Focus","Do not rush; accuracy first.","E8F0FE"),
("🇩🇪","Mindset","Confidence grows with practice.","E6F4EA"),
],
["Complete a full B1 mock test today.","Review all mistakes and rewrite corrections.","Write a 1‑minute summary of your performance."],
[("'die Konzentration' means?","Concentration"),("Translate: 'Zeitmanagement ist wichtig.'","Time management is important."),("'bereit' means?","Ready"),("Translate: 'Ich bleibe ruhig.'","I stay calm."),("'die Bewertung' means?","Evaluation")],
"Final mock test completed. You are ready for the real exam."
))

# ── Day 200 ─────────────────────────────────────────────────────────────────
DAYS.append((200,
"GRADUATION DAY 🎓",
"B1 Bridge Completion & What's Next",
"Congratulations! You have completed 80 days of the B1 bridge (Days 121–200). You now have solid B1 skills in grammar, vocabulary, and communication. Today we celebrate your achievement and plan your next steps toward B2.",
"અભિનંદન! તમે B1 bridge ના 80 દિવસ પૂર્ણ કર્યા. હવે તમારી B1 skills મજબૂત છે. આજે ઉજવણી અને આગળના B2 માટે યોજના બનાવીએ છીએ.",
[
("der Abschluss","completion","પૂર્ણતા","AP-shloos"),
("die Urkunde","certificate","પ્રમાણપત્ર","OOR-koon-deh"),
("die Kompetenz","competence","ક્ષમતા","kom-peh-TENTS"),
("fortgeschritten","advanced","ઉન્નત","FORT-geh-shrit-en"),
("die Vorbereitung","preparation","તૈયારી","for-bah-RY-tung"),
("das Sprachniveau","language level","ભાષા સ્તર","SHPRAKH-nee-VOH"),
("die Zertifizierung","certification","પ્રમાણપત્ર","tser-tee-fee-TSYON"),
("das Goethe-Institut","Goethe Institute","Goethe Institute","GÖ-teh"),
("die Prüfung","exam","પરીક્ષા","PRÜ-fung"),
("weiterlernen","to keep learning","આગળ શીખવું","VY-ter-LEHR-nen"),
("die Zukunft","future","ભવિષ્ય","TSOO-koonft"),
("das Ziel","goal","લક્ષ્ય","TSEEL"),
("der nächste Schritt","next step","આગળનું પગલું","NÄKHS-teh shrit"),
("stolz","proud","ગર્વિત","SHTOLTS"),
("weitermachen","to continue","ચાલુ રાખવું","VY-ter-makh-en"),
],
[
("Ich bin stolz auf meinen Fortschritt.","I am proud of my progress.","હું મારી પ્રગતિ પર ગર્વિત છું."),
("Mein Sprachniveau ist jetzt deutlich besser.","My language level is clearly better now.","મારો ભાષા સ્તર હવે ઘણો સુધર્યો છે."),
("Ich plane, die B1-Prüfung zu machen.","I plan to take the B1 exam.","હું B1 પરીક્ષા આપવાનું આયોજન કરું છું."),
("Der nächste Schritt ist B2.","The next step is B2.","આગળનું પગલું B2 છે."),
("Ich mache weiter — jeden Tag.","I will continue — every day.","હું આગળ ચાલુ રાખીશ — દરરોજ."),
],
[
("🏆","Achievement","You completed the B1 bridge with 80 days of focused practice.","E8F0FE"),
("🎯","What's Next","B2 topics: advanced reading, complex grammar, formal writing, debate, professional German.","E6F4EA"),
("🌟","Next Steps","1) Take the B1 exam. 2) Start B2. 3) Speak with native speakers weekly.","FFF8E1"),
],
["Write a reflection about your 200‑day journey.","Set 3 goals for the next 6 months.","Celebrate today — you earned it!"],
[("'der Abschluss' means?","Completion"),("Translate: 'Ich bin stolz.'","I am proud."),("'das Goethe-Institut' means?","Goethe Institute"),("Translate: 'Der nächste Schritt ist B2.'","The next step is B2."),("'weitermachen' means?","To continue")],
"🎓 Congratulations! You completed Day 200. Your German is strong, confident, and ready for the next level."
))

print("Days 191-200 appended")

# ═══════════════════════ MAIN FUNCTION ════════════════════════════════════

def build_book():
    doc = Document()

    # Page size A4
    for section in doc.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    # Title page
    add_para(doc, "🇩🇪", size_pt=60, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=40, space_after=8)
    add_para(doc, "DEUTSCH LERNEN", size_pt=32, bold=True, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "A2 → B1 Bridge Course", size_pt=18, bold=True, color=YELLOW, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "Days 121 to 200", size_pt=16, bold=True, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    add_para(doc, "ગુજરાતી ભાષકો માટે B1 Bridge જર્મન અભ્યાસ", size_pt=13, bold=True, color=DARK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "For Gujarati Speakers | English + Gujarati Explanations", size_pt=11, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "80 Days · B1 Grammar · Real-Life German · Exam Practice", size_pt=11, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)
    doc.add_paragraph()

    # How to use
    add_para(doc, "📖 How to Use This Book", size_pt=14, bold=True, color=BLUE, space_after=4)
    tips = [
        "📅 Study ONE day at a time — consistency beats speed.",
        "✍️  Write vocabulary in a notebook every day.",
        "🗣️  Speak every sentence out loud — B1 needs confidence.",
        "🔁  Review previous days weekly (spaced repetition).",
        "🎯  Complete the practice tasks and mini tests honestly.",
        "🇩🇪  Think in German as much as possible.",
    ]
    for t in tips:
        add_bullet(doc, t)
    doc.add_paragraph()

    # B1 bridge intro
    doc.add_page_break()
    add_para(doc, "🟧  B1 BRIDGE  —  THE TRANSITION", size_pt=18, bold=True, color=GREEN, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=20, space_after=4)
    add_para(doc, "Days 121 – 200", size_pt=14, bold=True, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    add_para(doc, "This bridge course takes you from A2 to confident B1 communication. After 80 days you will:", size_pt=11, space_after=4)
    achievements = [
        "✅ Use B1 grammar: Konjunktiv II, passive voice, complex clauses",
        "✅ Communicate professionally in emails, meetings, and interviews",
        "✅ Discuss news, society, health, and culture with confidence",
        "✅ Handle exams with clear strategies for reading, listening, and writing",
        "✅ Present, debate, and negotiate politely in German",
    ]
    for a in achievements:
        add_bullet(doc, a)
    add_para(doc, "🎯 Motivation: B1 is where German becomes truly useful in real life. Keep going!", size_pt=11, bold=True, color=BLUE, space_before=6, space_after=4)
    doc.add_paragraph()

    # Days 121-200
    for entry in DAYS:
        add_day(doc, *entry)

    # Save
    out = "German_Learning_Book_A2_Gujarati_Days_121_200.docx"
    doc.save(out)
    print(f"✅ Saved: {out}")
    print(f"   Days: {len(DAYS)}")

build_book()
