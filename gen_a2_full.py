# -*- coding: utf-8 -*-
"""Generate German_Learning_Book_A2_Gujarati.docx  (Days 61-120)"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

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
    if align:   p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after  = Pt(space_after)
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
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
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
    """rows = list of (german, english, gujarati, sound)"""
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    headers = ["🇩🇪 German","🇬🇧 English","ગુજરાતી","🔊 Sound like…"]
    bg_colors = ["1A73E8","1A73E8","1A73E8","1A73E8"]
    for i,(h,bg) in enumerate(zip(headers,bg_colors)):
        set_cell_bg(hdr[i], bg)
        cell_para(hdr[i], h, bold=True, size_pt=10, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
    for row_data in rows:
        row = table.add_row().cells
        colors = [DARK, DARK, DARK, GRAY]
        for i,(val,clr) in enumerate(zip(row_data, colors)):
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
    """
    vocab: list of (de, en, gu, sound)
    sentences: list of (de, en, gu)
    tips: list of (emoji, label, text, bg)
    tasks: list of str
    mini_test: list of (question, answer)
    """
    doc.add_page_break()
    # Day header
    add_para(doc, f"📘  DAY {day_num}", size_pt=21, bold=True, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=2)
    add_para(doc, title, size_pt=14, bold=True, color=YELLOW, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_para(doc, f"📚 Topic: {topic}", size_pt=10.5, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    doc.add_paragraph()
    # English explanation
    add_para(doc, "📖  English Explanation", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    add_para(doc, eng_exp, size_pt=10.5, space_after=6)
    # Gujarati explanation
    add_para(doc, "🇮🇳  ગુજરાતી સ્પષ્ટીકરણ", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    add_para(doc, guj_exp, size_pt=10.5, space_after=6)
    # Vocabulary
    add_para(doc, "📋  Vocabulary List", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    add_vocab_table(doc, vocab)
    # Sentences
    add_para(doc, "💬  Sentence Examples — Real Life", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    for de,en,gu in sentences:
        add_sentence_table(doc, de, en, gu)
    # Tips
    for emoji, label, text, bg in tips:
        add_tip_box(doc, emoji, label, text, bg)
    # Practice tasks
    add_para(doc, "✍️  Practice Tasks", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    for task in tasks:
        add_bullet(doc, task)
    doc.add_paragraph()
    # Mini test
    add_para(doc, "🧪  Mini Test", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    for i,(q,a) in enumerate(mini_test,1):
        add_para(doc, f"Q{i}. {q}", size_pt=10.5, space_after=1)
        add_para(doc, f"   ✅ {a}", size_pt=10.5, color=GREEN, space_after=3)
    doc.add_paragraph()
    # Final outcome
    add_para(doc, "🏆  Final Outcome", size_pt=10.5, bold=True, color=BLUE, space_after=2)
    add_para(doc, outcome, size_pt=10.5, space_after=8)

def add_level_test(doc, test_num, title, days_covered, questions):
    """questions: list of (question, answer)"""
    doc.add_page_break()
    add_para(doc, f"📝  LEVEL TEST {test_num}", size_pt=21, bold=True, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=2)
    add_para(doc, title, size_pt=14, bold=True, color=YELLOW, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_para(doc, f"📚 Covering: {days_covered}", size_pt=10.5, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    doc.add_paragraph()
    for i,(q,a) in enumerate(questions,1):
        add_para(doc, f"Q{i}. {q}", size_pt=10.5, bold=True, space_after=1)
        add_para(doc, f"   ✅ {a}", size_pt=10.5, color=GREEN, space_after=5)
    doc.add_paragraph()

# ═══════════════════════════ BOOK DATA ══════════════════════════════════════

# Each entry: (day_num, title, topic, eng_exp, guj_exp, vocab, sentences, tips, tasks, mini_test, outcome)
DAYS = []

# ── Day 61 ─────────────────────────────────────────────────────────────────
DAYS.append((61,
"WELCOME TO A2! 🚀",
"A1 Review + A2 Introduction",
"Congratulations! You have completed A1. You now know greetings, numbers, family, food, colours, and basic grammar. A2 takes you further — you will learn to talk about past events, give directions, shop, travel, express opinions, and much more. A2 German makes you a real communicator!",
"અભિનંદન! તમે A1 પૂર્ણ કર્યું. હવે A2 શરૂ થાય છે. A2 માં તમે ભૂતકાળ, દિશાઓ, ખરીદી, મુસાફરી, અભિપ્રાય — આ બધું બોલી શકશો. A2 એ તમને ખરો German communicator બનાવે છે!",
[
("sprechen","to speak","બોલવું","SHPREH-khen"),
("verstehen","to understand","સમજવું","fehr-SHTAY-en"),
("schreiben","to write","લખવું","SHRY-ben"),
("lesen","to read","વાંચવું","LAY-zen"),
("hören","to listen","સાંભળવું","HÖ-ren"),
("wiederholen","to revise","ફરી કરવું","VEE-der-ho-len"),
("lernen","to learn","શીખવું","LEHR-nen"),
("üben","to practise","પ્રેક્ટિસ કરવી","ÜH-ben"),
("verbessern","to improve","સુધારવું","fehr-BES-ern"),
("erinnern","to remember","યાદ કરવું","ehr-IN-ern"),
("Fortschritt","progress","પ્રગતિ","FORT-shrit"),
("Ziel","goal","ધ્યેય","TSEEL"),
("täglich","daily","દરરોજ","TÄG-likh"),
("weiter","further / continue","આગળ","VY-ter"),
("Erfolg","success","સફળતા","ehr-FOLK"),
],
[
("Ich lerne täglich Deutsch.","I learn German every day.","હું દરરોજ જર્મન શીખું છું."),
("Ich habe A1 erfolgreich abgeschlossen.","I have successfully completed A1.","મેં A1 સફળતાપૂર્વક પૂર્ણ કર્યું."),
("Mein Ziel ist es, fließend Deutsch zu sprechen.","My goal is to speak German fluently.","મારો ધ્યેય છે કે હું ફ્લુઅન્ટ જર્મન બોલું."),
("Ich übe jeden Tag eine Stunde.","I practise for one hour every day.","હું દરરોજ એક કલાક પ્રેક્ટિસ કરું છું."),
("Fortschritt kommt mit Geduld.","Progress comes with patience.","ધૈર્ય સાથે પ્રગતિ આવે છે."),
],
[
("🏆","A1 Achievement","You know 400+ words, present tense verbs, articles, and basic sentences. That is REAL progress!","E8F0FE"),
("🎯","A2 Goals","By Day 120 you will: use Dativ case, talk in the past, give directions, shop, travel, and express opinions.","E6F4EA"),
("💡","Tip","Write a short German diary entry every day — even 3 sentences. It builds grammar + vocabulary together.","FFF8E1"),
],
["Write 5 things you already know in German (A1 review).","Write your 3 A2 goals (what you want to say in German).","Say aloud: 'Ich lerne täglich Deutsch. Mein Ziel ist Erfolg!'"],
[("What does 'Fortschritt' mean?","Progress"),("Translate: 'Ich übe jeden Tag.'","I practise every day."),("Which level comes after A1?","A2"),("What does 'Ziel' mean?","Goal"),("Translate: 'weiter lernen'","to continue learning")],
"You have officially started A2! Your foundation is solid. The next 60 days will transform your German from basic to confident communication."
))

# ── Day 62 ─────────────────────────────────────────────────────────────────
DAYS.append((62,
"THE DATIVE CASE 🔵",
"Dativ — The Third German Case",
"German has 4 cases: Nominative (subject), Accusative (direct object), Dative (indirect object), and Genitive (possession). Today we master the Dative case. The Dative shows to whom or for whom something is done. Articles change in Dative: der→dem, die→der, das→dem, die(plural)→den.",
"જર્મનમાં 4 કેસ છે. Dativ એ indirect object (જેના માટે / જેને) દર્શાવે છે. Dativ માં articles બદલાય: der→dem, die→der, das→dem, plural→den. ઉદા: Ich gebe dem Mann das Buch. (હું માણસને પુસ્તક આપું છું.)",
[
("der Dativ","dative case","ત્રીજો કારક","dah-TEEF"),
("dem","the (m/n Dativ)","(m/n) ને","DAYM"),
("der","the (f Dativ)","(f) ને","DAIR"),
("den","the (pl Dativ)","(plural) ને","DAYN"),
("einem","a (m/n Dativ)","(m/n) ને","EYE-nem"),
("einer","a (f Dativ)","(f) ને","EYE-ner"),
("geben","to give","આપવું","GAY-ben"),
("helfen","to help (+ Dat)","મદદ કરવી","HEL-fen"),
("danken","to thank (+ Dat)","ધન્યવાદ આપવો","DANK-en"),
("gehören","to belong to (+ Dat)","સંબંધિત હોવું","geh-HÖ-ren"),
("zeigen","to show","બતાવવું","TSAY-gen"),
("schreiben","to write to","ને લખવું","SHRY-ben"),
("mit","with (+ Dativ)","સાથે","MIT"),
("aus","from / out of (+ Dativ)","માંથી","OUSE"),
("bei","at / near (+ Dativ)","પાસે","BY"),
],
[
("Ich gebe dem Kind ein Buch.","I give the child a book.","હું બાળકને એક પુસ્તક આપું છું."),
("Sie hilft der Frau.","She helps the woman.","તે સ્ત્રીને મદદ કરે છે."),
("Das gehört dem Mann.","That belongs to the man.","તે માણસનું છે."),
("Er dankt dem Lehrer.","He thanks the teacher.","તે શિક્ષકને ધન્યવાد આપે છે."),
("Ich wohne bei meiner Familie.","I live with my family.","હું મારા પરિવાર સાથે રહું છું."),
],
[
("📌","Case Summary","Nominative=subject | Accusative=direct obj | Dative=indirect obj | Think: 'I give IT (acc) to HIM (dat)'","E8F0FE"),
("⚠️","Common Mistake","Do NOT use 'dem' for feminine! Feminine Dative = 'der'. Example: Ich helfe der Frau (NOT dem Frau).","FFE0E0"),
("💡","Memory Trick","Dative = 'to whom / for whom'. Ask: 'Wem?' (To whom?) to find the Dative.","FFF8E1"),
],
["Make a table: Nominative / Accusative / Dative for der/die/das/die(pl).","Write 5 sentences using Dative (helfen, geben, danken).","Say aloud: 'Ich helfe dem Mann. Ich gebe der Frau ein Geschenk.'"],
[("What case is used for the indirect object?","Dative (Dativ)"),("'Der Mann' in Dative becomes…?","dem Mann"),("'Die Frau' in Dative becomes…?","der Frau"),("Which question word finds Dative?","Wem? (To whom?)"),("Translate: 'Ich gebe dem Kind das Spielzeug.'","I give the child the toy.")],
"You now understand the Dative case. This unlocks hundreds of German verbs and prepositions that require Dative. Great milestone!"
))

# ── Day 63 ─────────────────────────────────────────────────────────────────
DAYS.append((63,
"ACCUSATIVE vs DATIVE REVIEW 🔄",
"Case Contrast Practice",
"Today we drill the difference between Accusative (direct object — Wen/Was?) and Dative (indirect object — Wem?). This is one of the most important distinctions in German grammar. Master this and the rest of grammar becomes much easier.",
"આજે Accusative (direct object — Wen/Was?) અને Dative (indirect object — Wem?) નો ફરક drill કરીએ. આ German grammar નો સૌથી મહત્ત્વપૂર્ણ ભેદ છે. આ સ્પષ્ટ થઈ જાય તો grammar ઘણી સહેલી બની જાય.",
[
("der Akkusativ","accusative case","બીજો કારક","ah-ku-zah-TEEF"),
("den","the (m Acc)","(m) ને","DAYN"),
("die","the (f Acc)","(f) ને","DEE"),
("das","the (n Acc)","(n) ને","DAS"),
("einen","a (m Acc)","(m) ને","EYE-nen"),
("kaufen","to buy (+ Acc)","ખરીદવું","KOW-fen"),
("sehen","to see (+ Acc)","જોવું","ZAY-en"),
("brauchen","to need (+ Acc)","જોઈએ","BROW-khen"),
("finden","to find (+ Acc)","શોધવું","FIN-den"),
("besuchen","to visit (+ Acc)","મળવા જવું","beh-ZOO-khen"),
("schenken","to give as gift","ભેટ આપવી","SHENK-en"),
("bringen","to bring","લાવવું","BRING-en"),
("empfehlen","to recommend","ભલામણ કરવી","em-PFAY-len"),
("schicken","to send","મોકલવું","SHIK-en"),
("erklären","to explain","સમજાવવું","ehr-KLEH-ren"),
],
[
("Ich kaufe den Mantel. (Acc)","I buy the coat. (Acc)","હું ઓવરકોટ ખરીદું છું. (Acc)"),
("Ich gebe dem Kind den Mantel. (Dat+Acc)","I give the child the coat.","હું બાળકને ઓવરکوट આપું છું."),
("Sie sieht den Mann. (Acc)","She sees the man. (Acc)","તે માણસને જુએ છે."),
("Er schenkt der Frau Blumen. (Dat+Acc)","He gives the woman flowers.","તે સ્ત્રીને ફૂલ ભેટ આપે છે."),
("Ich brauche einen Arzt. (Acc)","I need a doctor.","મને ડૉક્ટર જોઈએ."),
],
[
("🔑","Key Rule","Ask 'Wen/Was?' → Accusative. Ask 'Wem?' → Dative. Two objects in one sentence: Dative comes FIRST.","E8F0FE"),
("💡","Order Tip","Ich gebe [Wem? → dem Kind] [Was? → das Buch]. Dative before Accusative (unless pronoun).","FFF8E1"),
("⚠️","Mistake to Avoid","'Ich helfe den Mann' is WRONG. 'helfen' takes Dative! Correct: 'Ich helfe dem Mann'","FFE0E0"),
],
["Write 5 Accusative sentences and 5 Dative sentences.","Transform: Change the Nominative article to Acc/Dat in 10 nouns.","Explain to yourself in Gujarati: 'Wem' vs 'Wen'."],
[("'Wen/Was?' finds which case?","Accusative"),("'Wem?' finds which case?","Dative"),("In 'Ich gebe dem Mann das Buch' — 'dem Mann' is…?","Dative (indirect object)"),("'das Buch' in the same sentence is…?","Accusative (direct object)"),("Translate: 'Ich sehe die Frau.'","I see the woman.")],
"You can now distinguish Accusative and Dative confidently. This is a huge grammar victory. Keep practising with real sentences!"
))

# ── Day 64 ─────────────────────────────────────────────────────────────────
DAYS.append((64,
"DATIVE PREPOSITIONS 🗺️",
"Prepositions Always Taking Dative",
"Some German prepositions ALWAYS take the Dative case, no matter what. Learn this list by heart: aus, bei, mit, nach, seit, von, zu, gegenüber, außer. After these words, articles always become Dative forms.",
"કેટલીક German prepositions હંમેશા Dative case લે છે. આ list ગોખી લો: aus, bei, mit, nach, seit, von, zu, gegenüber, außer. આ શબ્દો પછી articles હંમેશા Dative form માં આવે.",
[
("aus","out of / from","માંથી","OUSE"),
("bei","at / near / with","પાસે / ઘરે","BY"),
("mit","with","સાથે","MIT"),
("nach","after / to (cities)","પછી / તરફ","NAKH"),
("seit","since / for (time)","થી","ZITE"),
("von","from / of / by","થી / નો","FON"),
("zu","to (people/places)","પ્રત્યે / તરફ","TSOO"),
("gegenüber","opposite","સામે","geh-gen-Ü-ber"),
("außer","except / apart from","સિવાય","OWS-ehr"),
("ab","from (time/place)","થી","UP"),
("nach Hause","to home (direction)","ઘરે (દિશા)","nakh HOW-zeh"),
("zu Hause","at home (location)","ઘરે (સ્થળ)","tsoo HOW-zeh"),
("vom (von+dem)","from the","..ના/ની થી","FOM"),
("beim (bei+dem)","at the","..ના/ની ઘરે","BYME"),
("zum (zu+dem)","to the (m/n)","..ની તરફ","TSOOM"),
],
[
("Ich komme aus Deutschland.","I come from Germany.","હું જર્મનીથી આવું છું."),
("Ich wohne bei meiner Tante.","I live at my aunt's place.","હું મારી માસીના ઘરે રહું છું."),
("Ich fahre mit dem Bus.","I travel by bus.","હું bus થી જઉं છું."),
("Seit drei Jahren lerne ich Deutsch.","I have been learning German for three years.","ત્રણ વર્ષથી હું જર્મન શીખી રહ્યો/ી છું."),
("Ich gehe zum Arzt.","I am going to the doctor.","હું ડૉक्टर પાસે જઉं છું."),
],
[
("📌","Dative-only Prepositions","aus · bei · mit · nach · seit · von · zu · gegenüber · außer · ab. Memorise them as a chant!","E8F0FE"),
("💡","Contraction Tip","von+dem=vom | bei+dem=beim | zu+dem=zum | zu+der=zur. These contractions are used in everyday speech.","FFF8E1"),
("🔊","Pronunciation","'gegenüber' = geh-gen-Ü-ber. The ü sound = say 'ee' and round your lips like 'oo'.","E6F4EA"),
],
["Chant the 9 Dative prepositions 5 times until you can say them without looking.","Write 2 sentences for each Dative preposition (18 sentences total).","Practice contractions: vom, beim, zum, zur in sentences."],
[("Which preposition means 'since/for (time)'?","seit"),("'Ich fahre ___ Bus' — fill in?","mit dem"),("Translate: 'Ich wohne bei meinem Bruder.'","I live at my brother's place."),("von+dem contracts to…?","vom"),("zu+der contracts to…?","zur")],
"You now know all Dative-only prepositions. These appear in 70% of everyday German sentences — you'll see them everywhere!"
))

# ── Day 65 ─────────────────────────────────────────────────────────────────
DAYS.append((65,
"ACCUSATIVE PREPOSITIONS ➡️",
"Prepositions Always Taking Accusative",
"Just like Dative-only prepositions, some prepositions ALWAYS take Accusative: durch, für, gegen, ohne, um, bis, entlang, wider. After these words, articles always take Accusative forms. Remember: für (for) is one of the most common ones!",
"Dative-only prepositions ની જેમ, કેટલીક prepositions હંમેશા Accusative case લે છે: durch, für, gegen, ohne, um, bis, entlang, wider. ઉદ: 'für mich' (મારા માટે), 'ohne dich' (તારા વગર). 'für' સૌથી common preposition છે!",
[
("durch","through","માં થઈ","DOOKH"),
("für","for","માટે","FÜR"),
("gegen","against","વિરુદ્ધ","GAY-gen"),
("ohne","without","વિના / વગર","OH-neh"),
("um","around / at (time)","ની આસ-પાસ / વાગ્યે","OOM"),
("bis","until / up to","સુધી","BIS"),
("entlang","along (after noun)","ને અનુસરી","ent-LANG"),
("wider","against (formal)","વિરુદ્ધ (formal)","VEE-der"),
("für mich","for me","મારા માટે","für MIKH"),
("für dich","for you","તારા માટે","für DIKH"),
("ohne mich","without me","મારા વિના","OH-neh mikh"),
("um 8 Uhr","at 8 o'clock","8 વાગ્યે","oom AKT oor"),
("durch den Park","through the park","park માં થઈ","dookh dayn park"),
("gegen den Wind","against the wind","પવન વિરુદ્ધ","gay-gen dayn vint"),
("bis Montag","until Monday","સોમવાર સુધી","bis MON-tahg"),
],
[
("Ich kaufe das Geschenk für dich.","I buy the gift for you.","હું તારા માટે ભેટ ખરીદું છું."),
("Wir gehen durch den Park.","We walk through the park.","અમે park માં થઈ ચાલીએ છીએ."),
("Ohne dich bin ich verloren.","Without you I am lost.","તારા વગર હું ખોવાઈ ગયો/ગઈ."),
("Das Spiel beginnt um 18 Uhr.","The game starts at 6 pm.","રમત 6 વાગ્યે શરૂ થાય."),
("Ich bin gegen diese Idee.","I am against this idea.","હું આ idea વિરુદ્ધ છું."),
],
[
("📌","Accusative-only Prepositions","durch · für · gegen · ohne · um · bis · entlang · wider. Tip: 'dufgobuw' as a mnemonic!","E8F0FE"),
("💡","'für' Tip","'für' is the most-used Accusative preposition. 'Das ist für dich' = This is for you. Use it daily!","FFF8E1"),
("⚠️","'um' Warning","'um' has two meanings: 'around' (place) and 'at' (time). Context decides: 'um den See' (around the lake) vs 'um 9 Uhr' (at 9 o'clock).","FFE0E0"),
],
["Write 2 sentences for each Accusative preposition.","Compare: make 5 sentence pairs using Dative prepositions AND Accusative prepositions.","Make a card: Dative preps on blue, Accusative preps on red."],
[("Which preposition means 'without'?","ohne"),("'für' always takes which case?","Accusative"),("Translate: 'Ich warte bis Freitag.'","I wait until Friday."),("'um 7 Uhr' means…?","at 7 o'clock"),("Translate: 'durch die Stadt'","through the city")],
"Accusative prepositions are now in your toolkit. Together with Dative prepositions, you can now form complex, natural-sounding German sentences!"
))

# ── Day 66 ─────────────────────────────────────────────────────────────────
DAYS.append((66,
"TWO-WAY PREPOSITIONS 🔀",
"Wechselpräpositionen — Location vs Direction",
"Nine prepositions in German can take EITHER Dative (location — where?) OR Accusative (direction — where to?). These are: an, auf, hinter, in, neben, über, unter, vor, zwischen. Rule: Wo? (where) = Dative | Wohin? (where to) = Accusative.",
"German ની 9 prepositions Dative (ક્યાં?) OR Accusative (ક્યાં તરફ?) — બંને case લઈ શકે: an, auf, hinter, in, neben, über, unter, vor, zwischen. Rule: Wo? (ક્યાં?) = Dative | Wohin? (ક્યાં જઈ?) = Accusative.",
[
("an","at / on (vertical)","પર (ઊભું)","AN"),
("auf","on (horizontal)","પર (આડું)","OWF"),
("hinter","behind","પાછળ","HIN-ter"),
("in","in / into","માં","IN"),
("neben","next to","બાજુ","NAY-ben"),
("über","over / above","ઉપર","Ü-ber"),
("unter","under / below","નીચે","OON-ter"),
("vor","in front of","સામે","FOR"),
("zwischen","between","વચ્ચે","TSVISH-en"),
("liegen","to lie (location)","પડ્યું/ક્યાં છે","LEE-gen"),
("hängen","to hang (location)","ટa​ǵgayel હોવું","HENG-en"),
("legen","to lay/put (direction)","મૂકવું","LAY-gen"),
("stellen","to place upright (dir.)","ઊભું કરવું","SHTEL-en"),
("hängen","to hang (direction)","ટàngavvun","HENG-en"),
("stehen","to stand (location)","ઊભું/ક્યાં છે","SHTAY-en"),
],
[
("Das Buch liegt auf dem Tisch. (Wo?→Dat)","The book is lying on the table.","પુस्तक ટेবल पर छे. (Dativ)"),
("Ich lege das Buch auf den Tisch. (Wohin?→Acc)","I put the book on the table.","હું পुस्तक ટेবল पर मूकूं छूं. (Akkusativ)"),
("Das Bild hängt an der Wand. (Wo?→Dat)","The picture hangs on the wall.","ছवि दीवाल पर छे. (Dativ)"),
("Er hängt das Bild an die Wand. (Wohin?→Acc)","He hangs the picture on the wall.","ते ছवि दीवाल पर टांगे छे. (Akkusativ)"),
("Die Katze sitzt unter dem Stuhl.","The cat sits under the chair.","बिल्ली खुरशी नीचे बेठी छे."),
],
[
("🔑","The Magic Question","Wo? (Where is it?) → Dative | Wohin? (Where is it going?) → Accusative. Ask the question first, then choose the case.","E8F0FE"),
("💡","Location Verbs","liegen, stehen, hängen, sitzen, stecken = location verbs → use Dative.","FFF8E1"),
("💡","Direction Verbs","legen, stellen, hängen, setzen, stecken = direction verbs → use Accusative.","E6F4EA"),
],
["Draw a room and describe where things ARE (Wo? → Dativ).","Now write what you PUT where (Wohin? → Akkusativ).","Make a two-column table: Wo?(Dativ) | Wohin?(Akkusativ)"],
[("'Wo?' requires which case?","Dative"),("'Wohin?' requires which case?","Accusative"),("Translate: 'Das Buch liegt auf dem Tisch.'","The book is on the table."),("How many two-way prepositions are there?","9"),("Name 3 two-way prepositions.","an, auf, in (or any 3 from the list)")],
"Two-way prepositions mastered! This is the key to describing locations and movements in German. You are now thinking like a German speaker!"
))

# ── Day 67 ─────────────────────────────────────────────────────────────────
DAYS.append((67,
"MODAL VERBS I — KÖNNEN & MÜSSEN 🔧",
"können (can) and müssen (must)",
"Modal verbs express ability, permission, necessity or desire. They always pair with an infinitive verb at the end of the sentence. Today: können (can/to be able to) and müssen (must/have to). Note the irregular present tense conjugations!",
"Modal verbs ability, permission, necessity અથવा desire express કરે છે. ते infinitive verb ना साथे आवे छे — infinitive sentence ना अंते. आज: können (शकवूं) और müssen (जोईए / करवूं ज पडशे). तेमनी conjugation irregular छे, ध्यान रखो!",
[
("können","can / to be able to","શકવું","KÖN-en"),
("ich kann","I can","હું શkі","ikh KAN"),
("du kannst","you can","તું શkes","doo KANST"),
("er/sie/es kann","he/she/it can","તે શkه","ehr KAN"),
("wir können","we can","આपणे शkीए","veer KÖN-en"),
("ihr könnt","you all can","તमे शkो","eer KÖNT"),
("sie/Sie können","they/you(formal) can","ते/आप शkे","zee KÖN-en"),
("müssen","must / have to","જ  કરવું પડે","MÜS-en"),
("ich muss","I must","मारे करवूं ज पडे","ikh MOOS"),
("du musst","you must","तारे करवूं ज पडे","doo MOOST"),
("er/sie muss","he/she must","तेणे करवूं ज पडे","ehr MOOS"),
("wir müssen","we must","आपणे करवूं ज पडे","veer MÜS-en"),
("ihr müsst","you all must","तमारे करवूं ज पडे","eer MÜST"),
("Infinitiv","infinitive (sentence-end)","मूळ क्रिया (अंते)","in-fin-ee-TEEF"),
("Satzklammer","sentence bracket","वाक्य-संरचना","ZATS-klam-er"),
],
[
("Ich kann Deutsch sprechen.","I can speak German.","હું German bolī shakun chhun."),
("Kannst du mir helfen?","Can you help me?","Šhakhe tun maney madad karavī?"),
("Wir müssen jetzt gehen.","We must go now.","Apanē havē javun padshē."),
("Du musst mehr üben.","You must practise more.","Tarē vadhu abhyās karavō padshē."),
("Er kann gut kochen.","He can cook well.","Tē sārī rītē rāndhan karī shakē chhe."),
],
[
("📌","Word Order Rule","Modal verb goes to position 2. Infinitive goes to the END: 'Ich [kann] Deutsch [sprechen].'","E8F0FE"),
("💡","können tip","Use 'können' when asking for help: 'Können Sie mir helfen?' = Can you help me? Very useful in Germany!","FFF8E1"),
("⚠️","Spelling","'müssen' has ü (umlaut). ich muss / du musst — no umlaut in the conjugated forms! Common mistake.","FFE0E0"),
],
["Conjugate können and müssen for all 6 persons.","Write 5 sentences with können and 5 with müssen.","Ask 3 questions using 'Kannst du…?' and answer them."],
[("Where does the infinitive go in a modal verb sentence?","At the end of the sentence"),("'Ich ___ schwimmen.' (can) → fill in?","kann"),("Translate: 'Du musst schlafen.'","You must sleep."),("'können' in 3rd person singular is…?","kann"),("Translate: 'Wir können morgen kommen.'","We can come tomorrow.")],
"können and müssen are two of the most-used German verbs. You can now express ability and necessity. Real-life German conversations just became possible!"
))

# ── Day 68 ─────────────────────────────────────────────────────────────────
DAYS.append((68,
"MODAL VERBS II — WOLLEN, SOLLEN, DÜRFEN, MÖGEN 🎯",
"The Other Four Modal Verbs",
"German has 6 modal verbs. You learned können and müssen. Now the remaining four: wollen (want to), sollen (supposed to / should), dürfen (may / allowed to), mögen/möchten (like / would like). All follow the same word-order rule: modal at position 2, infinitive at the end.",
"German ना 6 modal verbs छे. तमे können अने müssen शीख्यां. हवे बाकी चार: wollen (इच्छवूं), sollen (जोईए / कहेवायूं छे), dürfen (मंजूरी), mögen/möchten (गमवूं / इच्छवूं). बधां same rule follow कर: modal position 2, infinitive अंते.",
[
("wollen","to want to","ઈच्छवूं","VOL-en"),
("ich will","I want to","मने करवूं छे","ikh VIL"),
("sollen","to be supposed to","कहेवायूं छे / जोईए","ZOL-en"),
("ich soll","I am supposed to","मारे करवानूं छे","ikh ZOL"),
("dürfen","may / to be allowed to","मंजूरी/अधिकार छे","DÜR-fen"),
("ich darf","I may","मने मंजूरी छे","ikh DARF"),
("mögen","to like","गमवूं","MÖ-gen"),
("ich mag","I like","मने गमे छे","ikh MAHG"),
("möchten","would like to","इच्छा छे (polite)","MÖKh-ten"),
("ich möchte","I would like","मने गमशे","ikh MÖKh-teh"),
("du magst","you like","तने गमे छे","doo MAHGST"),
("wir wollen","we want to","आपणे इच्छीए","veer VOL-en"),
("ihr dürft","you all may","तमने मंजूरी छे","eer DÜRFT"),
("er soll","he is supposed to","तेणे करवानूं छे","ehr ZOL"),
("sie möchten","they would like","तेओ इच्छे छे","zee MÖKh-ten"),
],
[
("Ich will Arzt werden.","I want to become a doctor.","मारे doctor बनवूं छे."),
("Du sollst um 8 Uhr da sein.","You should be there at 8 o'clock.","तारे 8 वागे त्यां हाजर हوवूं जोईए."),
("Darf ich hier sitzen?","May I sit here?","शूं हूं अहीं बेसी शकूं?"),
("Ich mag Musik sehr.","I like music a lot.","मने music खूब गमे छे."),
("Ich möchte einen Kaffee, bitte.","I would like a coffee, please.","मारे एक coffee जोईए, कृपा करी."),
],
[
("📌","All 6 Modal Verbs","können (can) | müssen (must) | wollen (want) | sollen (should) | dürfen (may) | mögen (like). Learn as a set!","E8F0FE"),
("💡","möchten tip","'möchten' is the polite form of 'mögen'. Use 'Ich möchte…' instead of 'Ich will…' in shops — it sounds much more polite!","FFF8E1"),
("🇩🇪","Cultural Tip","In a restaurant or shop, always use 'Ich möchte…' (I would like…). Using 'Ich will…' sounds demanding to German ears.","E6F4EA"),
],
["Conjugate all 4 new modal verbs for all 6 persons.","Write a short paragraph about your day using at least 4 different modal verbs.","Practise: order food/drink using 'Ich möchte…' sentences."],
[("Polite way to say 'I want' in a shop?","Ich möchte (I would like)"),("Translate: 'Darf ich das Fenster öffnen?'","May I open the window?"),("'Wir ___ morgen kommen.' (want to)","wollen"),("What does 'sollen' express?","Being supposed to / obligation from others"),("Translate: 'Er mag keine Spinnen.'","He doesn't like spiders.")],
"All 6 modal verbs are now yours! These are the backbone of everyday German. Combine them with vocabulary and you can express almost any basic idea."
))

# ── Day 69 ─────────────────────────────────────────────────────────────────
DAYS.append((69,
"SEPARABLE VERBS ✂️",
"Trennbare Verben — Verbs That Split!",
"German has verbs that split into two parts in a sentence! The prefix (like auf-, an-, ab-, ein-, aus-, mit-, vor-, zurück-) jumps to the END of the sentence, while the verb stem stays at position 2. Example: aufmachen (to open) → Ich mache die Tür auf. (I open the door.)",
"German ना Separable verbs sentence मां split थाय छे! Prefix (auf-, an-, ab-, ein-, aus-...) sentence ना अंते जाय छे, अने verb stem position 2 पर रहे छे. दा.त.: aufmachen (खोलवूं) → 'Ich mache die Tür auf.' - prefix 'auf' अंते गयूं!",
[
("aufmachen","to open","ઉઘاडवूं","OWF-makh-en"),
("zumachen","to close","बंध करवूं","TSOO-makh-en"),
("anrufen","to call (phone)","phone कर","AN-roo-fen"),
("einkaufen","to go shopping","खरीदी कर","EYN-kow-fen"),
("aufstehen","to get up","उठवूं","OWF-shtay-en"),
("einschlafen","to fall asleep","सूई जवूं","EYN-shlah-fen"),
("ankommen","to arrive","पहोंचवूं","AN-kom-en"),
("abfahren","to depart","रवाना थवूं","AB-fah-ren"),
("mitnehmen","to take along","साथे लेवूं","MIT-nay-men"),
("vorstellen","to introduce","परिचय कराववो","FOR-shtel-en"),
("zurückkommen","to come back","पाछा आववूं","tsoo-RÜK-kom-en"),
("fernsehen","to watch TV","TV जोवूं","FEHRN-zay-en"),
("ausziehen","to move out / undress","निकळवूं","OWS-tsee-en"),
("umsteigen","to change (transport)","गाडी बदलवी","OOM-shty-gen"),
("aufräumen","to tidy up","साफ करवूं","OWF-roy-men"),
],
[
("Ich stehe um 7 Uhr auf.","I get up at 7 o'clock.","हूं 7 वागे उठूं छूं."),
("Wir kaufen heute ein.","We go shopping today.","आपणे आज खरीदी करीए छीए."),
("Er ruft seine Mutter an.","He calls his mother.","ते तेनी मांने phone करे छे."),
("Der Zug fährt um 10 Uhr ab.","The train departs at 10 o'clock.","ट्रेन 10 वागे रवाना थाय छे."),
("Kommst du morgen zurück?","Are you coming back tomorrow?","शूं तूं काले पाछो/पाछी आवशे?"),
],
[
("📌","Split Rule","Prefix goes to the END of the sentence (or clause). Verb stem stays at position 2. In infinitive form, prefix and verb are joined.","E8F0FE"),
("💡","How to Identify","If a verb has a stressed prefix (AUF-machen, EIN-kaufen) it is separable. Dictionary shows: aufmachen, sep. [auf-]","FFF8E1"),
("⚠️","Subordinate Clauses","In subordinate clauses, the verb goes to the end and the prefix re-joins: '…weil ich früh aufstehe.' (because I get up early)","FFE0E0"),
],
["Write 10 separable verb sentences (one for each verb above).","Make a table: Infinitive | Split form sentence | Meaning.","Conjugate: aufstehen for all 6 persons with time expressions."],
[("In 'Ich mache das Fenster auf' — where is the prefix?","At the end of the sentence"),("Translate: 'Wir steigen in München um.'","We change (trains) in Munich."),("'aufräumen' means?","To tidy up"),("Translate: 'Er schläft um 22 Uhr ein.'","He falls asleep at 10 pm."),("In infinitive form, separable verbs are…?","Written as one word: aufmachen, einkaufen, etc.")],
"Separable verbs unlocked! Now you can describe daily routines, travel plans, and actions in detail. You are speaking real German now!"
))

# ── Day 70 ─────────────────────────────────────────────────────────────────
DAYS.append((70,
"AT THE SHOP 🛒",
"Im Geschäft — Shopping in German",
"Shopping is a daily activity and one of the first real-life German situations you'll encounter. Today you learn all the vocabulary and phrases needed to shop confidently in Germany — asking for items, sizes, prices, and making a purchase.",
"ขरीदी रोजिंदी प्रवृत्ति छे अने जर्मनी मां पहेली real-life situation छे. आज तमे जर्मनी मां confidently खरीदी करवा माटे vocabulary अने phrases शीखशो — वस्तु मागवी, size, किंमत, खरीदी करवी.",
[
("das Geschäft","the shop / store","दुकान","das geh-SHEFT"),
("der Laden","the shop (small)","નаनी दुकान","dehr LAH-den"),
("das Kaufhaus","the department store","मोटी दुकान","das KOW-hows"),
("der Supermarkt","the supermarket","supermarket","dehr ZOO-per-markt"),
("die Kasse","the cash register / checkout","काउंटर / कैश","dee KAS-eh"),
("der Verkäufer","the salesperson (m)","दुकानदार (पुरुष)","dehr fehr-KOY-fer"),
("die Verkäuferin","the salesperson (f)","दुकानदार (स्त्री)","dee fehr-KOY-fer-in"),
("der Preis","the price","किंमत","dehr PRYS"),
("die Größe","the size","size","dee GRÖ-seh"),
("die Farbe","the colour","रंग","dee FAR-beh"),
("günstig","cheap / affordable","ससती","GÜN-stikh"),
("teuer","expensive","महंगूं","TOY-er"),
("der Rabatt","the discount","छूट","dehr ra-BAT"),
("der Kassenbon","the receipt","receipt","dehr KAS-en-bon"),
("bezahlen","to pay","ચूकववूं","beh-TSAH-len"),
],
[
("Was kostet das?","How much does that cost?","आनी किंमत केटली छे?"),
("Haben Sie das in Größe M?","Do you have this in size M?","शूं आ M size मां छे?"),
("Ich hätte gerne einen Rabatt.","I would like a discount.","मारे छूट जोईए."),
("Kann ich mit Karte zahlen?","Can I pay by card?","शूं हूं card थी चूकवी शकूं?"),
("Wo ist die Kasse, bitte?","Where is the checkout, please?","काउंटर क्यां छे, कृपा?"),
],
[
("🇩🇪","Cultural Tip","In Germany, bags cost money in supermarkets. Always say 'Eine Tüte, bitte' (a bag please) if you need one.","E6F4EA"),
("💡","Useful Phrase","'Ich schaue mich nur um.' = I'm just looking. (Said when a salesperson asks if they can help.)","FFF8E1"),
("🔊","Pronunciation","'Kaufhaus' = KOW-hows. 'Preis' = PRYS (rhymes with 'price'). 'günstig' = GÜN-stikh.","E8F0FE"),
],
["Role-play: You want to buy a shirt. Write the full conversation.","Visit any online German shop and read 5 product descriptions.","Write 5 shopping questions using 'Haben Sie…?' and 'Was kostet…?'"],
[("Translate: 'Was kostet das?'","How much does that cost?"),("'teuer' means…?","Expensive"),("How do you say 'the checkout'?","die Kasse"),("Translate: 'Ich möchte das kaufen.'","I would like to buy this."),("'günstig' means…?","Cheap / affordable")],
"Shopping vocabulary mastered! Your first trip to a German store is now fully covered. You can ask prices, sizes, pay — and even ask for a discount!"
))

print("Days 61-70 data defined. Continuing...")

# ── Day 71 ─────────────────────────────────────────────────────────────────
DAYS.append((71,
"CLOTHES & FASHION 👗",
"Kleidung — Talking About Clothes",
"Today you learn German vocabulary for clothing and fashion. Whether you're shopping, describing outfits, or doing laundry — these words are essential for everyday life in Germany.",
"આज Kleidung (कपडा) ना शब्दो शीखशो. खरीदी, outfit ना वर्णन, या कपडा धोवा — आ शब्दो रोजिंदी जिंदगी मां जरूरी छे.",
[
("das Hemd","the shirt (button-up)","शर्ट","das HEMD"),
("die Bluse","the blouse","blouse","dee BLOO-zeh"),
("die Hose","the trousers","पैंट","dee HOH-zeh"),
("der Rock","the skirt","स्कर्ट","dehr ROK"),
("das Kleid","the dress","फ्रॉक / ड्रेस","das KLITE"),
("die Jacke","the jacket","जैकेट","dee YAK-eh"),
("der Mantel","the coat","ओवरकोट","dehr MAN-tel"),
("der Pullover","the jumper/sweater","sweater","dehr pool-OH-ver"),
("die Schuhe (pl)","the shoes","जूता","dee SHOO-eh"),
("die Socken (pl)","the socks","मोजा","dee ZOK-en"),
("die Unterwäsche","underwear","अंदरना कपडा","dee OON-ter-vesh-eh"),
("die Mütze","the hat / beanie","टोपी","dee MÜ-tseh"),
("der Schal","the scarf","स्कार्फ","dehr SHAHL"),
("die Handschuhe","gloves","दस्ताना","dee HANT-shoo-eh"),
("anziehen","to put on (clothes)","पहेरवूं","AN-tsee-en"),
],
[
("Ich ziehe die Jacke an.","I put on the jacket.","हूं जेकेट पहेरूं छूं."),
("Das Kleid ist sehr schön.","The dress is very beautiful.","ड्रेस खूब सुंदर छे."),
("Welche Größe haben Sie?","What size do you take?","तमारी size केटली छे?"),
("Ich suche eine rote Hose.","I'm looking for red trousers.","मारे लाल pants जोईए."),
("Diese Schuhe sind zu eng.","These shoes are too tight.","आ जूता खूब तंग छे."),
],
[
("💡","Anziehen is separable","'anziehen' splits: 'Ich ziehe die Jacke an.' The 'an' goes to the end. Same for 'ausziehen' (to take off).","FFF8E1"),
("🇩🇪","German Sizes","German shoe sizes are EU sizes. Clothing: XS/S/M/L/XL or numeric (36/38/40...). Women's size 38 ≈ UK 12.","E6F4EA"),
("🔊","Pronunciation","'Kleid' = KLITE (rhymes with 'light'). 'Schuhe' = SHOO-eh. 'Mütze' = MÜ-tseh.","E8F0FE"),
],
["Describe what you are wearing right now in German.","Write 5 sentences: 'Ich trage…' (I am wearing…) with different clothing.","Make a seasonal wardrobe list: Summer / Winter clothes in German."],
[("Translate: 'Das ist zu groß.'","That is too big."),("'die Hose' means?","Trousers / pants"),("'anziehen' is which type of verb?","Separable verb"),("Translate: 'Ich möchte dieses Hemd kaufen.'","I would like to buy this shirt."),("'die Schuhe' means?","Shoes")],
"You can now talk about clothing in German — shop for clothes, describe outfits, and discuss fashion. Germany has world-class fashion cities like Berlin and Hamburg!"
))

# ── Day 72 ─────────────────────────────────────────────────────────────────
DAYS.append((72,
"AT THE MARKET 🥦",
"Auf dem Markt — Buying Fresh Food",
"Germany has wonderful weekly markets (Wochenmärkte) where you can buy fresh vegetables, fruits, bread, and cheese. Today you learn market vocabulary, quantities, and how to ask for what you want.",
"Germany मां wonderful weekly markets (Wochenmärkte) छे — fresh vegetables, fruits, bread, cheese मळे. आज market vocabulary, quantities, अने मागवाना phrases शीखशो.",
[
("der Markt","the market","बाजार","dehr MARKT"),
("der Stand","the stall","स्टॉल","dehr SHTANT"),
("das Gemüse","vegetables","शाकभाजी","das geh-MÜ-zeh"),
("das Obst","fruit","फळ","das OPST"),
("das Brot","bread","रोटी / ब्रेड","das BROHT"),
("der Käse","cheese","पनीर / चीज़","dehr KÄ-zeh"),
("das Fleisch","meat","मांस","das FLYSH"),
("der Fisch","fish","मछली","dehr FISH"),
("das Kilogramm","kilogram","किलोग्राम","das kee-lo-GRAM"),
("das Gramm","gram","ग्राम","das GRAM"),
("das Stück","piece / item","टुकडो / नंग","das SHTÜK"),
("die Tüte","the bag","थेलो","dee TÜ-teh"),
("frisch","fresh","تازूं","FRISH"),
("reif","ripe","पाकूं","RYFE"),
("der Preis","the price","किंमत","dehr PRYS"),
],
[
("Ich hätte gerne ein Kilo Tomaten.","I would like one kilo of tomatoes.","मारे एक किलो टमेटा जोईए."),
("Was kostet das Brot?","How much does the bread cost?","ब्रेड ना केटला पैसा छे?"),
("Sind die Äpfel frisch?","Are the apples fresh?","शूं सेब ताजा छे?"),
("Ich nehme 500 Gramm Käse.","I'll take 500 grams of cheese.","हूं 500 ग्राम पनीर लईश."),
("Haben Sie noch Erdbeeren?","Do you still have strawberries?","शूं तमारी पासे स्ट्रॉबेरी छे?"),
],
[
("🇩🇪","Market Culture","German markets open early (7am!) and close by noon. Always greet: 'Guten Morgen!' before asking. Germans love fresh market produce.","E6F4EA"),
("💡","Quantity Phrases","'Ein Kilo…' | '500 Gramm…' | 'Zwei Stück…' | 'Eine Tüte…' | 'Einen Bund…' (a bunch). Learn these quantity words!","FFF8E1"),
("🔊","Pronunciation","'Gemüse' = geh-MÜ-zeh. 'Käse' = KÄ-zeh. 'Stück' = SHTÜK.","E8F0FE"),
],
["Role-play: You are at the market. Write a conversation buying 5 different items.","Learn the German names for 10 vegetables you eat regularly.","Write a shopping list for a week in German."],
[("How do you say 'How much does this cost?'","Was kostet das?"),("'frisch' means?","Fresh"),("Translate: 'Zwei Kilo Kartoffeln, bitte.'","Two kilos of potatoes, please."),("'das Gemüse' means?","Vegetables"),("Where do Germans buy fresh produce?","At the Wochenmarkt (weekly market)")],
"Market vocabulary is yours! You can now navigate a German market, ask for prices, quantities, and freshness. This is practical German at its finest!"
))

# ── Day 73 ─────────────────────────────────────────────────────────────────
DAYS.append((73,
"FOOD & COOKING 🍳",
"Essen und Kochen — In the Kitchen",
"Today you learn vocabulary for cooking methods, kitchen items, and describing food. Germans love home cooking — learning this vocabulary helps you read recipes, discuss meals, and feel at home in a German kitchen.",
"आज cooking methods, kitchen items, अने food describe करवाना words शीखशो. Germans घरना खाना प्रेम करे छे — आ vocabulary recipe वांचवा, खाना बद्दल वात करवा, अने German kitchen मां comfortable feel करवा मदद करे.",
[
("kochen","to cook / boil","रांधवूं","KO-khen"),
("backen","to bake","बेक करवूं","BAK-en"),
("braten","to fry / roast","तळवूं / भूनवूं","BRAH-ten"),
("schneiden","to cut","काटवूं","SHNY-den"),
("rühren","to stir","हलाववूं","RÜ-ren"),
("kochen (boil)","to boil","उकाळवूं","KO-khen"),
("das Rezept","the recipe","रेसिपी","das reh-TSEPT"),
("die Zutat","the ingredient","ઘटक","dee TSOO-taht"),
("der Topf","the pot","तपेलूं","dehr TOPF"),
("die Pfanne","the frying pan","कड़ाई","dee PFAN-eh"),
("der Backofen","the oven","ओवन","dehr BAK-oh-fen"),
("der Kühlschrank","the fridge","फ्रिज","dehr KÜL-shrank"),
("salzig","salty","खारूं","ZAL-tsikh"),
("süß","sweet","मीठूं","ZÜSS"),
("scharf","spicy / hot","तीखूं","SHARP"),
],
[
("Ich koche heute Abend Suppe.","I am cooking soup this evening.","आज सांजे हूं soup रांधूं छूं."),
("Das schmeckt sehr gut!","That tastes very good!","आ खूब स्वादिष्ट छे!"),
("Kannst du Pasta kochen?","Can you cook pasta?","शूं तूं pasta रांधी शके?"),
("Das Essen ist zu scharf.","The food is too spicy.","खाना खूब तीखूं छे."),
("Ich brauche das Rezept.","I need the recipe.","मारे recipe जोईए."),
],
[
("🇩🇪","German Food Culture","Germany is famous for Brot (bread) — over 300 types! Also Wurst (sausage), Sauerkraut, Pretzels. Meals are hearty and filling.","E6F4EA"),
("💡","Taste Words","süß (sweet) | salzig (salty) | scharf (spicy) | sauer (sour) | bitter (bitter) | lecker (delicious!). 'Das ist lecker!' is the most useful compliment.","FFF8E1"),
("🔊","Pronunciation","'kochen' = KO-khen. 'Pfanne' = PFAN-eh. 'Kühlschrank' = KÜL-shrank.","E8F0FE"),
],
["Write a recipe for your favourite dish in German (5-8 steps).","Describe today's meals in German: Frühstück / Mittagessen / Abendessen.","Learn 10 German spices and herbs (Gewürze und Kräuter)."],
[("'backen' means?","To bake"),("Translate: 'Das schmeckt lecker!'","That tastes delicious!"),("How do you say 'the fridge'?","der Kühlschrank"),("'scharf' means?","Spicy / hot"),("Translate: 'Ich schneide das Gemüse.'","I cut the vegetables.")],
"Kitchen and cooking vocabulary complete! You can now follow a German recipe, describe food, and impress any German host with your culinary German!"
))

# ── Day 74 ─────────────────────────────────────────────────────────────────
DAYS.append((74,
"AT THE RESTAURANT 🍽️",
"Im Restaurant — Dining Out in German",
"Going to a restaurant in Germany requires specific vocabulary and polite phrases. Today you learn how to make a reservation, order food and drinks, ask for the bill, and behave like a local. German restaurant etiquette is slightly different from other cultures!",
"Germany मां restaurant मां जवा specific vocabulary अने polite phrases जोईए. आज reservation करवी, खाना-पीणूं order करवूं, bill मागवूं — अने local जेमां behave करवूं शीखशो.",
[
("das Restaurant","the restaurant","रेस्टोरां","das res-toh-RAHNG"),
("der Tisch","the table","टेबल","dehr TISH"),
("die Speisekarte","the menu","menu","dee SHPY-zeh-kar-teh"),
("die Vorspeise","the starter","starter","dee FOR-shpy-zeh"),
("das Hauptgericht","the main course","main course","das HOWPT-geh-rikht"),
("die Nachspeise","the dessert","dessert","dee NAKH-shpy-zeh"),
("der Kellner","the waiter (m)","waiter","dehr KEL-ner"),
("die Kellnerin","the waitress (f)","waitress","dee KEL-ner-in"),
("die Rechnung","the bill","bill","dee REKH-nung"),
("das Trinkgeld","the tip","tip","das TRINK-gelt"),
("reservieren","to reserve","reserve करवूं","reh-zehr-VEE-ren"),
("bestellen","to order","order करवूं","beh-SHTEL-en"),
("empfehlen","to recommend","ભlaman करवी","em-PFAY-len"),
("bringen","to bring","लावvun","BRING-en"),
("zahlen","to pay","चूकववूं","TSAH-len"),
],
[
("Ich möchte einen Tisch für zwei reservieren.","I would like to reserve a table for two.","मारे बे जण माटे टेबल reserve करवूं छे."),
("Was empfehlen Sie heute?","What do you recommend today?","आज शूं ભlamण करशो?"),
("Ich nehme das Wiener Schnitzel.","I'll have the Wiener Schnitzel.","हूं Wiener Schnitzel लईश."),
("Die Rechnung, bitte.","The bill, please.","bill लावो, कृपा."),
("Hat es Ihnen geschmeckt?","Did you enjoy the food?","खाना गम्यूं ने?"),
],
[
("🇩🇪","Tipping in Germany","Tip by rounding up or adding ~10%. Say 'Stimmt so' (keep the change) or mention the total you want to pay: 'Ich zahle 25 Euro.'","E6F4EA"),
("💡","Calling Waiter","To get a waiter's attention: 'Entschuldigung!' (Excuse me!) — do NOT snap fingers or shout. That is considered rude in Germany.","FFF8E1"),
("💡","Table Manners","Wait to be seated. Do NOT start eating until everyone has food. The host often says 'Guten Appetit!' before eating.","E8F0FE"),
],
["Write a full restaurant conversation: reservation → ordering → paying.","Learn the names of 5 famous German dishes (Wiener Schnitzel, Sauerbraten, Bratwurst, Käsespätzle, Schwarzwälder Kirschtorte).","Practise: 'Ich möchte…' for ordering 3 different meals."],
[("How do you ask for the bill?","Die Rechnung, bitte."),("'die Speisekarte' means?","The menu"),("Translate: 'Was empfehlen Sie?'","What do you recommend?"),("How do you call a waiter politely?","Entschuldigung!"),("'Trinkgeld' means?","Tip (gratuity)")],
"Restaurant German is now yours! You can dine out in Germany like a confident, polite guest. Guten Appetit — enjoy your meal!"
))

# ── Day 75 ─────────────────────────────────────────────────────────────────
DAYS.append((75,
"ORDERING FOOD & DRINKS 🥤",
"Bestellen — From Menu to Table",
"Today you focus specifically on ordering phrases, drink vocabulary, dietary requirements, and politely handling issues at a restaurant or café. Very practical for everyday German life!",
"आज specifically ordering phrases, drinks vocabulary, dietary requirements, अने restaurant/café मां problems handle करवा नी polite language शीखशो. आ रोजिंदी जिंदगी मां खूब जरूरी छे!",
[
("das Getränk","the drink","पीण","das geh-TRÄNK"),
("das Wasser","water","पाणी","das VAS-er"),
("der Saft","juice","जूस","dehr ZAFT"),
("das Bier","beer","बीयर","das BEER"),
("der Wein","wine","वाईन","dehr VINE"),
("der Kaffee","coffee","कॉफी","dehr KAF-ay"),
("der Tee","tea","चाय","dehr TAY"),
("die Milch","milk","दूध","dee MILKH"),
("vegetarisch","vegetarian","शाकाहारी","veh-geh-TAH-rish"),
("vegan","vegan","vegan","veh-GAHN"),
("laktosefrei","lactose-free","लैक्टोज-मुक्त","lak-TOH-zeh-fry"),
("glutenfrei","gluten-free","ग्लूटेन-मुक्त","gloo-TEN-fry"),
("die Allergie","the allergy","एलर्जी","dee ah-lehr-GHEE"),
("noch einmal","once more / again","फरीथी","nokh EYN-mahl"),
("das stimmt","that is correct","सही छे","das SHTIMMT"),
],
[
("Ich bin Vegetarier. Haben Sie vegetarische Gerichte?","I am vegetarian. Do you have vegetarian dishes?","हूं शाकाहारी छूं. शूं vegetarian खाना छे?"),
("Ich hätte gerne ein Glas Wasser, bitte.","I would like a glass of water, please.","मारे एक गिलास पाणी जोईए, कृपा."),
("Ich habe eine Nussallergie.","I have a nut allergy.","मने nutती एलर्जी छे."),
("Entschuldigung, das ist nicht meine Bestellung.","Excuse me, this is not my order.","माफ करो, आ मारी order नथी."),
("Noch einmal dasselbe, bitte.","The same again, please.","फरीथी एज, कृपा."),
],
[
("💡","Allergy Tip","Always state allergies clearly: 'Ich bin allergisch gegen Nüsse/Gluten/Milch.' German kitchens take allergies seriously.","FFE0E0"),
("🇩🇪","Tap Water Tip","In Germany, tap water (Leitungswasser) is safe to drink but restaurants rarely serve it free. Ask: 'Stilles Wasser oder Sprudelwasser?' (Still or sparkling?)","E6F4EA"),
("🔊","Pronunciation","'Getränk' = geh-TRÄNK. 'vegetarisch' = veh-geh-TAH-rish. 'Allergie' = ah-lehr-GHEE.","E8F0FE"),
],
["Write your dietary requirements in German (even if you have none, practise the vocabulary).","Memorise: 'Ich hätte gerne…' + 5 different drinks and foods.","Write a dialogue where you have a dietary restriction and explain it."],
[("Translate: 'Ich bin allergisch gegen Gluten.'","I am allergic to gluten."),("'vegetarisch' means?","Vegetarian"),("How do you say 'once more / again'?","noch einmal"),("Translate: 'Ein Glas Wasser, bitte.'","A glass of water, please."),("'stilles Wasser' means?","Still (non-sparkling) water")],
"You can now handle any food situation in Germany — ordering, dietary needs, drink requests, and fixing mix-ups. You're a confident restaurant-goer in German!"
))

# ── Day 76 ─────────────────────────────────────────────────────────────────
DAYS.append((76,
"MONEY & PAYING 💶",
"Geld und Bezahlen — Handling Money in German",
"Money matters! Today you learn euro and cent vocabulary, how to ask prices, haggle politely, understand change, and use German banking vocabulary. Essential for daily life and travel in Germany.",
"पैसा महत्त्वपूर्ण! आज euro/cent vocabulary, किंमत पूछवी, politely किंमत negotiate करवी, बाकी पैसा समजवूं, अने German banking vocabulary शीखशो. जर्मनी मां daily life माटे जरूरी.",
[
("der Euro","the euro","€","dehr OY-roh"),
("der Cent","the cent","सेंट","dehr TSENT"),
("das Bargeld","cash","रोकडा","das BAR-gelt"),
("die Kreditkarte","credit card","क्रेडिट कार्ड","dee kreh-DEET-kar-teh"),
("die EC-Karte","debit card","डेबिट कार्ड","dee ay-TSAY-kar-teh"),
("der Geldautomat","the ATM","ATM","dehr GELT-ow-toh-maht"),
("das Wechselgeld","the change (money)","बाकी पैसा","das VEK-sel-gelt"),
("der Betrag","the amount","रकम","dehr beh-TRAHG"),
("günstig","cheap","ससतूं","GÜN-stikh"),
("kostenlos","free of charge","मफत","KOS-ten-lohs"),
("Stimmt so.","Keep the change.","बाकी रहेवा दो.","SHTIMMT zo"),
("Kann ich mit Karte zahlen?","Can I pay by card?","शूं card थी चूकवी शकूं?","kan ikh..."),
("Wie viel kostet das?","How much does it cost?","आनी किंमत केटली?","vee feel..."),
("Haben Sie es kleiner?","Do you have anything smaller?","शूं छूटा पैसा छे?","HAH-ben zee..."),
("die Quittung","the receipt","receipt","dee KVIT-ung"),
],
[
("Das kostet fünfzehn Euro fünfzig.","That costs fifteen euros fifty.","आ 15 euro 50 cent नूं छे."),
("Kann ich mit EC-Karte zahlen?","Can I pay by debit card?","शूं EC card थी चूकवी शकूं?"),
("Haben Sie Wechselgeld für 50 Euro?","Do you have change for 50 euros?","शूं 50 euro ना छूटा पैसा छे?"),
("Wo ist der nächste Geldautomat?","Where is the nearest ATM?","सौथी नजीक ATM क्यां छे?"),
("Das ist kostenlos.","That is free of charge.","आ मफत छे."),
],
[
("🇩🇪","Germany Cash Culture","Germany is still a cash-heavy country! Many shops, restaurants, and markets are 'Nur Bargeld' (cash only). Always carry cash!","E6F4EA"),
("💡","Price Reading","'3,50 €' is read as 'drei Euro fünfzig'. The comma is decimal in Germany (NOT a period). €1.000 means one thousand euros!","FFF8E1"),
("💡","Stimmt so","'Stimmt so' literally means 'That's right' and is used to tell the waiter/cashier to keep the change as a tip.","E8F0FE"),
],
["Practise reading 10 prices aloud in German: 4,50 / 12,99 / 0,75 / 100,00 etc.","Write 5 payment scenarios in German dialogue form.","Find out: how do you say your country's currency in German?"],
[("'Bargeld' means?","Cash"),("How do you say 'Keep the change'?","Stimmt so."),("Translate: 'Das kostet zwanzig Euro.'","That costs twenty euros."),("What does 'Geldautomat' mean?","ATM"),("In Germany, is cash or card more common?","Cash (Bargeld) is very common")],
"Money vocabulary sorted! You can pay, ask for change, find ATMs, and handle all financial transactions in German. Very practical for real life in Germany!"
))

# ── Day 77 ─────────────────────────────────────────────────────────────────
DAYS.append((77,
"LARGE NUMBERS & COUNTING 🔢",
"Zahlen 100–1.000.000",
"In A1 you learned numbers 1-100. Today you go much further: hundreds, thousands, millions. You'll also learn how to say years, phone numbers, addresses, and prices with large numbers — all essential for real German communication.",
"A1 मां 1-100 शीख्यां. आज खूब आगे जईए: सैकडा, हजारो, लाखो. वर्षो, phone numbers, addresses, अने मोटी किंमतो पण शीखशो — real German communication माटे जरूरी.",
[
("hundert","hundred","સो","HOON-dert"),
("zweihundert","two hundred","बसो","TSVY-hoon-dert"),
("tausend","thousand","हजार","TOW-zend"),
("zweitausend","two thousand","बे हजार","TSVY-tow-zend"),
("zehntausend","ten thousand","दस हजार","TSAYN-tow-zend"),
("hunderttausend","hundred thousand","एक लाख","HOON-dert-tow-zend"),
("eine Million","one million","दस लाख / दस लाख","EY-neh mil-YOHN"),
("hundertfünfzig","150","एक सो पचास","HOON-dert-FÜNFtsikh"),
("dreitausendvierhundert","3,400","ते हजार चार सो","DRY-tow-zend-FEER-hoon-dert"),
("das Jahr","the year","वर्ष","das YAHR"),
("neunzehnhundert","1900","ओगणीसो","NOYN-tsayn-HOON-dert"),
("zweitausendvierundzwanzig","2024","बे हजार चोवीस","..."),
("die Telefonnummer","phone number","फोन नंबर","dee teh-leh-FON-noom-er"),
("die Postleitzahl","the postal code","ZIP/postal code","dee POST-lyte-tsahl"),
("die Hausnummer","house number","घर नंबर","dee HOWS-noom-er"),
],
[
("Das Haus kostet dreihunderttausend Euro.","The house costs 300,000 euros.","घर 3 लाख euro नूं छे."),
("Meine Telefonnummer ist 0176 24681357.","My phone number is 0176 24681357.","मारो phone नंबर 0176 24681357 छे."),
("Deutschland hat über achtzig Millionen Einwohner.","Germany has over 80 million inhabitants.","Germany मां 8 करोडथी वधु लोको रहे छे."),
("Das Jahr zweitausendvierundzwanzig war besonders.","The year 2024 was special.","वर्ष 2024 खास हतूं."),
("Meine Postleitzahl ist 10115.","My postal code is 10115.","मारी ZIP code 10115 छे."),
],
[
("📌","Number Building","German numbers are built logically: hundert + tausend + million. 345 = dreihundert-fünf-und-vierzig. Practise building!","E8F0FE"),
("💡","Year Tip","Years 1100-1999: say as hundreds. 1985 = neunzehnhundertfünfundachtzig. 2000+ = zweitausend + rest. 2024 = zweitausendvierundzwanzig.","FFF8E1"),
("🔊","Phone Numbers","Read German phone numbers digit by digit or in pairs: 0-1-7-6 or 01-76. Practice with real German phone number formats.","E6F4EA"),
],
["Write your phone number, postal code, and birth year in German words.","Practise: count from 100 to 1000 in steps of 100.","Write out 5 large prices (house, car, phone, etc.) in German words."],
[("How do you say 1000 in German?","tausend"),("Translate: 'Das kostet fünfhundert Euro.'","That costs five hundred euros."),("How do you say the year 2000?","zweitausend"),("'eine Million' means?","One million"),("How do you say 250?","zweihundertfünfzig")],
"Large numbers are now in your German toolkit! You can say prices, years, phone numbers, and addresses — essential for real life and paperwork in Germany!"
))

# ── Day 78 ─────────────────────────────────────────────────────────────────
DAYS.append((78,
"TRANSPORTATION 🚆",
"Verkehrsmittel — Getting Around Germany",
"Germany has excellent public transport — trains (Deutsche Bahn), trams, U-Bahn (subway), S-Bahn (city trains), and buses. Today you learn all transport vocabulary, buying tickets, and asking about schedules.",
"Germany नी public transport excellent छे — trains, trams, U-Bahn, S-Bahn, buses. आज transport vocabulary, ticket खरीदवूं, अने schedule पूछवूं शीखशो.",
[
("das Verkehrsmittel","means of transport","वाहन","das fehr-KEHRS-mit-el"),
("der Zug","the train","ट्रेन","dehr TSOOG"),
("die U-Bahn","the subway / metro","metro","dee OO-bahn"),
("die S-Bahn","city/suburban rail","शहेरी ट्रेन","dee ES-bahn"),
("die Straßenbahn","the tram","ट्राम","dee SHTRAH-sen-bahn"),
("der Bus","the bus","बस","dehr BOOS"),
("das Flugzeug","the aeroplane","विमान","das FLOO K-tsoyg"),
("das Fahrrad","the bicycle","साइकल","das FAR-raht"),
("das Auto","the car","कार","das OW-toh"),
("der Bahnhof","the train station","रेल्वे स्टेशन","dehr BAHN-hohf"),
("die Haltestelle","the stop (bus/tram)","stop","dee HAL-teh-shtel-eh"),
("die Fahrkarte","the ticket","ticket","dee FAR-kar-teh"),
("abfahren","to depart (sep. verb)","रवाना थवूं","AB-fah-ren"),
("ankommen","to arrive (sep. verb)","पहोंचवूं","AN-kom-en"),
("umsteigen","to change/transfer (sep.)","गाडी बदलवी","OOM-shty-gen"),
],
[
("Wann fährt der nächste Zug nach Berlin?","When does the next train to Berlin depart?","Berlin माटे अगली ट्रेन क्यारे छूटे?"),
("Ich möchte eine Fahrkarte nach München kaufen.","I would like to buy a ticket to Munich.","मारे Munich माटे ticket जोईए."),
("Muss ich umsteigen?","Do I have to change trains?","शूं मारे ट्रेन बदलवी पडशे?"),
("Der Zug hat zehn Minuten Verspätung.","The train is ten minutes late.","ट्रेन दस मिनिट मोडी छे."),
("Wo ist die nächste U-Bahn-Station?","Where is the nearest metro station?","नजीक नी metro station क्यां छे?"),
],
[
("🇩🇪","Deutsche Bahn","Germany's national railway is the DB (Deutsche Bahn). ICE trains are the fastest. Book early for cheap prices on bahn.de!","E6F4EA"),
("💡","Verspätung","'Verspätung' (delay) is unfortunately common in Germany! If your train is late: 'Der Zug hat X Minuten Verspätung.'","FFF8E1"),
("💡","Ticket Types","Einzelfahrt (single), Hin- und Rückfahrt (return), Tageskarte (day pass), Monatskarte (monthly pass). Learn these!","E8F0FE"),
],
["Plan a trip from your city to 3 German cities — write the transport sentences.","Learn the German public transport vocabulary for your city.","Write 5 questions you might ask at a train station."],
[("'die Fahrkarte' means?","The ticket"),("Translate: 'Wann kommt der Bus an?'","When does the bus arrive?"),("'umsteigen' means?","To change / transfer (trains/buses)"),("What is the German national railway called?","Deutsche Bahn (DB)"),("Translate: 'Ich fahre mit der U-Bahn.'","I travel by metro.")],
"Transport vocabulary complete! You can navigate trains, buses, and metros in Germany. The whole country is now open to you — Deutschland entdecken! (Discover Germany!)"
))

# ── Day 79 ─────────────────────────────────────────────────────────────────
DAYS.append((79,
"ASKING FOR DIRECTIONS 🗺️",
"Wegbeschreibung — Finding Your Way",
"Getting lost is part of every travel adventure! Today you learn how to ask for and give directions in German — turning left/right, distances, landmarks, and politely asking strangers for help.",
"खोवाई जवूं दर travel adventure नो भाग छे! आज German मां directions पूछवी अने आपवी शीखशो — डाबे/जमणे वळवूं, अंतर, landmark, अने strangers ने politely पूछवूं.",
[
("die Richtung","direction","दिशा","dee RIKH-tung"),
("links","left","ডाबे","LINKS"),
("rechts","right","जमणे","REKHTS"),
("geradeaus","straight ahead","सीधूं आगे","geh-RAH-deh-ows"),
("abbiegen","to turn (sep.)","वळवूं","AB-bee-gen"),
("die Kreuzung","the crossroads","चौराहो","dee KROY-tsung"),
("die Ampel","the traffic light","ट्राफिक लाईट","dee AM-pel"),
("die Brücke","the bridge","पुल","dee BRÜ-keh"),
("der Platz","the square","चौक","dehr PLATS"),
("die Straße","the street","रस्तो","dee SHTRAH-seh"),
("weit","far","दूर","VITE"),
("nah / neben","near / next to","नजीक","NAH / NAY-ben"),
("um die Ecke","around the corner","खूणे वळींने","oom dee EK-eh"),
("gegenüber","opposite","सामे","geh-gen-Ü-ber"),
("Entschuldigung","excuse me / sorry","माफ करो","ent-SHOOL-di-gung"),
],
[
("Entschuldigung, wie komme ich zum Bahnhof?","Excuse me, how do I get to the train station?","माफ करो, रेल्वे स्टेशन केम जवाय?"),
("Gehen Sie geradeaus, dann links.","Go straight ahead, then turn left.","सीधूं जाओ, पछी डाबे वळो."),
("Die Post ist um die Ecke.","The post office is around the corner.","post office खूणे वळींने छे."),
("Wie weit ist das Stadtzentrum?","How far is the city centre?","शहेर नूं centre केटलूं दूर छे?"),
("Das ist ungefähr fünf Minuten zu Fuß.","That is about 5 minutes on foot.","ते लगभग 5 मिनिट चालीने छे."),
],
[
("💡","Polite Start","Always start with 'Entschuldigung' when asking a stranger. Germans appreciate politeness and will happily help!","FFF8E1"),
("🇩🇪","Navigation Tip","In Germany, Google Maps works perfectly. But knowing directions in German helps when locals give you verbal directions quickly.","E6F4EA"),
("📌","Key Pattern","'Wie komme ich zu/zum/zur + destination?' = How do I get to...? zu + dem = zum (m/n), zu + der = zur (f)","E8F0FE"),
],
["Write directions from your home to the nearest shop in German.","Practise: give directions to 5 landmarks in a German city (use a map).","Role-play: a tourist asks you for directions — write both sides of the conversation."],
[("How do you say 'straight ahead'?","geradeaus"),("Translate: 'Biegen Sie links ab.'","Turn left."),("'die Kreuzung' means?","The crossroads / intersection"),("How do you ask 'How far is it?'","Wie weit ist das?"),("'um die Ecke' means?","Around the corner")],
"Directions vocabulary mastered! You can now navigate any German city, ask strangers for help, and understand the answers. Keine Angst vor dem Verirren! (No fear of getting lost!)"
))

# ── Day 80 ─────────────────────────────────────────────────────────────────
DAYS.append((80,
"MID-BOOK REVIEW TEST 📝",
"Days 61–80 Comprehensive Review",
"Congratulations on reaching Day 80! You have covered cases (Dative/Accusative), prepositions, modal verbs, separable verbs, shopping, cooking, restaurant, money, transport, and directions. Today is a full review to consolidate everything.",
"Day 80 पहोंचवा बदल अभिनंदन! तमे cases, prepositions, modal verbs, separable verbs, shopping, cooking, restaurant, money, transport, अने directions cover कर्यां. आज full review छे.",
[
("Wiederholung","revision / review","ઉजળी / review","vee-der-HOH-lung"),
("der Rückblick","the review / look back","ठलकवूं","dehr RÜK-blik"),
("zusammenfassen","to summarise","सारांश काढवो","tsoo-ZA-men-fas-en"),
("überprüfen","to check / verify","चकासवूं","Ü-ber-prü-fen"),
("der Fehler","the mistake","ભूल","dehr FAY-ler"),
("korrigieren","to correct","सुधारवूं","ko-ri-GHEE-ren"),
("die Übung","the exercise","exercise","dee Ü-bung"),
("der Test","the test","test","dehr TEST"),
("das Ergebnis","the result","परिणाम","das ehr-GAY P-nis"),
("bestehen","to pass (a test)","pass करवूं","beh-SHTAY-en"),
("durchfallen","to fail","fail थवूं","DOOKH-fal-en"),
("die Note","the grade","ग्रेड","dee NOH-teh"),
("wiederholen","to revise","फरी शीखवूं","vee-der-HOH-len"),
("stark","strong","मजबूत","SHTARK"),
("schwach","weak","कमजोर","SHVAKH"),
],
[
("Ich wiederhole alle Vokabeln.","I revise all the vocabulary.","हूं बधां vocabulary revise करूं छूं."),
("Ich habe den Test bestanden!","I passed the test!","में test pass कर्यो!"),
("Mein schwacher Punkt ist die Grammatik.","My weak point is the grammar.","मारो कमजोर भाग grammar छे."),
("Ich übe täglich, um besser zu werden.","I practise daily to get better.","हूं better थवा दररोज practice करूं छूं."),
("Fortschritt braucht Zeit und Geduld.","Progress needs time and patience.","प्रगति माटे समय अने धीरज जोईए."),
],
[
("🏆","Achievement","By Day 80 you know: 4 German cases, 20+ prepositions, 6 modal verbs, 30+ separable verbs, and 400+ A2 vocabulary words!","E8F0FE"),
("💡","Review Strategy","Go back over Days 61-79. Re-read any day that feels unclear. Test yourself on vocabulary without looking.","FFF8E1"),
("🎯","Next Steps","Days 81-120 cover: travel, health, past tense, future tense, work, adjective endings, relative clauses, and more. Stay strong!","E6F4EA"),
],
["Review Days 61-79: flip back through each day's vocabulary.","Write a 10-sentence German paragraph using knowledge from Days 61-79.","Make your personal 'weak points' list and review those topics."],
[("Which case shows the indirect object?","Dative (Dativ)"),("Name 3 Dative-only prepositions.","aus, bei, mit, nach, seit, von, zu (any 3)"),("Translate: 'Ich möchte mit der U-Bahn fahren.'","I would like to travel by metro."),("'Wie viel kostet das?' means?","How much does this cost?"),("Translate: 'Ich muss links abbiegen.'","I have to turn left.")],
"Day 80 milestone achieved! You have covered the foundations of A2. The next 40 days will complete your A2 journey and prepare you for B1 German. Weiter so! (Keep it up!)"
))

print("Days 71-80 appended successfully!")

# ── Day 81 ─────────────────────────────────────────────────────────────────
DAYS.append((81,
"TRAVEL & HOLIDAYS ✈️",
"Reisen und Urlaub — Planning a Trip",
"Germany is a hub for European travel! Today you learn vocabulary for planning holidays, booking accommodation, and talking about travel experiences. Whether going to the Alps, the coast, or neighbouring countries — these phrases are essential.",
"Germany European travel नो hub छे! आज holiday plan करवी, accommodation book करवूं, अने travel experience बद्दल वात करवा vocabulary शीखशो.",
[
("die Reise","the trip / journey","ट्रिप / प्रवास","dee RY-zeh"),
("der Urlaub","the holiday / vacation","vacation","dehr OOR-lowp"),
("buchen","to book","book करवूं","BOO-khen"),
("das Hotel","the hotel","hotel","das hoh-TEL"),
("die Pension","the guesthouse / B&B","gest house","dee pen-ZYOHN"),
("das Zimmer","the room","ओरडो","das TSIM-er"),
("die Reservierung","the reservation","reservation","dee reh-zehr-VEE-rung"),
("die Abreise","the departure (hotel)","hotel checkout","dee AB-ry-zeh"),
("die Ankunft","the arrival","आगमन","dee AN-koonft"),
("der Reisepass","the passport","passport","dehr RY-zeh-pas"),
("das Visum","the visa","visa","das VEE-zoom"),
("der Koffer","the suitcase","suitcase","dehr KOF-er"),
("packen","to pack","पैक करवूं","PAK-en"),
("das Ausland","abroad / foreign country","विदेश","das OWS-lant"),
("die Sehenswürdigkeit","the sight / attraction","ध्यान खेंचनारूं स्थळ","dee ZAY-ens-vür-dikh-kite"),
],
[
("Ich möchte ein Zimmer für zwei Nächte buchen.","I would like to book a room for two nights.","मारे बे रात माटे room book करवूं छे."),
("Mein Reisepass ist abgelaufen.","My passport has expired.","मारो passport expire थई गयो."),
("Wo sind die wichtigsten Sehenswürdigkeiten?","Where are the most important sights?","सौथी महत्त्वना ध्यान खेंचनारां स्थळो क्यां छे?"),
("Ich packe meinen Koffer für den Urlaub.","I am packing my suitcase for the holiday.","हूं vacation माटे suitcase पैक करूं छूं."),
("Wann beginnt die Reise?","When does the journey begin?","प्रवास क्यारे शरू थाय?"),
],
[
("🇩🇪","Top German Destinations","Bavaria (Bayern), Black Forest (Schwarzwald), Rhine Valley (Rheintal), Berlin, Munich, Hamburg, Cologne (Köln), Neuschwanstein Castle!","E6F4EA"),
("💡","Hotel Phrases","Check-in: 'Ich habe eine Reservierung.' Check-out: 'Ich möchte auschecken.' Room problem: 'Es gibt ein Problem mit dem Zimmer.'","FFF8E1"),
("💡","Travel Documents","In Germany (Schengen), EU citizens need ID card only. Non-EU need passport + visa. Always have docs ready!","E8F0FE"),
],
["Plan a 3-day trip to Germany in German — where you go, where you stay, what you see.","Write 5 sentences about your dream vacation using German vocabulary.","Learn the names of 5 German cities and what they are famous for."],
[("'der Urlaub' means?","Holiday / vacation"),("Translate: 'Ich möchte ein Zimmer buchen.'","I would like to book a room."),("'der Koffer' means?","Suitcase"),("How do you say 'passport'?","der Reisepass"),("Translate: 'Wir fahren ins Ausland.'","We are going abroad.")],
"Travel vocabulary complete! You can plan, book, pack, and navigate a trip to Germany or through Europe in German. Deutschland, wir kommen! (Germany, we're coming!)"
))

# ── Day 82 ─────────────────────────────────────────────────────────────────
DAYS.append((82,
"WEATHER ☀️🌧️",
"Das Wetter — Talking About the Weather",
"The weather is a universal conversation starter! Germans love talking about Wetter (weather). Today you learn weather vocabulary, how to describe weather conditions, and how to understand German weather forecasts.",
"Weather बद्दल वात कर्वी universal conversation starter छे! Germans Wetter (weather) बद्दल खूब वात करे छे. आज weather vocabulary, conditions describe करवी, अने German weather forecast समजवूं शीखशो.",
[
("das Wetter","the weather","हवामान","das VET-er"),
("die Temperatur","the temperature","तापमान","dee tem-peh-rah-TOOR"),
("der Grad","degree (temperature)","डिग्री","dehr GRAHD"),
("sonnig","sunny","तडकाळूं","ZON-ikh"),
("bewölkt","cloudy","वादळाळूं","beh-VÖLKT"),
("regnerisch","rainy","वरसादी","REGH-neh-rish"),
("windig","windy","पवनाळूं","VIN-dikh"),
("neblig","foggy","धुम्मसियूं","NAY-blikh"),
("kalt","cold","ठंडूं","KALT"),
("warm","warm","गरम","WARM"),
("heiß","hot","ઉकाળો","HICE"),
("der Schnee","snow","बरफ","dehr SHNAY"),
("der Regen","rain","वरसाद","dehr RAY-gen"),
("das Gewitter","thunderstorm","वीजळी-वंटोळ","das geh-VIT-er"),
("die Wettervorhersage","the weather forecast","weather forecast","dee VET-er-for-hehr-zah-geh"),
],
[
("Wie ist das Wetter heute?","How is the weather today?","आज हवामान केवूं छे?"),
("Es regnet stark.","It is raining heavily.","खूब वरसाद पडे छे."),
("Morgen wird es sonnig und warm.","Tomorrow it will be sunny and warm.","काले तडको अने गरमी रहेशे."),
("Die Temperatur beträgt 25 Grad.","The temperature is 25 degrees.","तापमान 25 डिग्री छे."),
("Im Winter schneit es in Bayern.","In winter it snows in Bavaria.","शियाळामां Bavaria मां बरफ पडे छे."),
],
[
("🇩🇪","German Weather","German summers are warm (20-30°C). Winters can be cold (-5 to 5°C) with snow in the south. Autumn is rainy. April weather changes hourly!","E6F4EA"),
("💡","'Es' as Subject","Weather verbs use 'es' as subject: 'Es regnet' (it rains), 'Es schneit' (it snows), 'Es ist kalt' (it is cold). 'es' is impersonal here.","FFF8E1"),
("🔊","Forecast Language","'Morgen wird es…' (Tomorrow it will be…) | 'Heute Nachmittag…' (This afternoon…) | 'In der Nacht…' (During the night…)","E8F0FE"),
],
["Write a 5-day weather forecast for a German city in German.","Describe today's weather where you live in German.","Write 5 sentences about seasonal weather in Germany using the vocabulary."],
[("Translate: 'Es ist sehr kalt heute.'","It is very cold today."),("'sonnig' means?","Sunny"),("How do you say 'thunderstorm'?","das Gewitter"),("Translate: 'Wie viel Grad hat es?'","How many degrees is it?"),("'der Schnee' means?","Snow")],
"Weather vocabulary complete! You can now have one of the most common German conversations. Small talk about Wetter breaks the ice in any German situation!"
))

# ── Day 83 ─────────────────────────────────────────────────────────────────
DAYS.append((83,
"SEASONS & NATURE 🌸🍂",
"Jahreszeiten und Natur — The Natural World",
"Germany's four seasons are distinct and beautiful. Today you learn the seasons, months, nature vocabulary, and how to describe the natural environment. This also helps you connect weather, time, and outdoor activities.",
"Germany ना चार ऋतु अलग अने सुंदर छे. आज seasons, months, nature vocabulary, अने natural environment describe करवूं शीखशो.",
[
("der Frühling","spring","वसंत","dehr FRÜ-ling"),
("der Sommer","summer","ग्रीष्म / ઉनाळो","dehr ZOM-er"),
("der Herbst","autumn","शरद ऋतु","dehr HERPST"),
("der Winter","winter","शिशिर / शियाळो","dehr VIN-ter"),
("der Wald","the forest","जंगल","dehr VALT"),
("der Berg","the mountain","पहाड","dehr BEHRG"),
("der See","the lake","तलाव","dehr ZAY"),
("der Fluss","the river","नदी","dehr FLOOS"),
("das Meer","the sea","समुद्र","das MEHR"),
("die Wiese","the meadow","घास नो मेदान","dee VEE-zeh"),
("der Baum","the tree","ઝাড़","dehr BOWM"),
("die Blume","the flower","फूल","dee BLOO-meh"),
("das Tier","the animal","प्राणी","das TEER"),
("der Vogel","the bird","पक्षी","dehr FOH-gel"),
("die Landschaft","the landscape","प्राकृतिक दृश्य","dee LANT-shaft"),
],
[
("Im Frühling blühen die Blumen.","In spring the flowers bloom.","वसंत मां फूल खीले छे."),
("Der Schwarzwald ist wunderschön.","The Black Forest is wonderful.","Schwarzwald अति सुंदर छे."),
("Im Sommer fahre ich ans Meer.","In summer I go to the sea.","ऊनाळा मां हूं समुद्र किनारे जाउं छूं."),
("Die Bäume werden im Herbst bunt.","The trees become colourful in autumn.","शरद ऋतु मां झाड रंगीन थाय छे."),
("Bayern hat viele wunderschöne Berge.","Bavaria has many beautiful mountains.","Bavaria मां ઘणा સुंदर पहाडो छे."),
],
[
("🇩🇪","Nature in Germany","Germany has the Black Forest (Schwarzwald), Rhine River (Rhein), Bavarian Alps, and Baltic Sea coast. Nature lovers paradise!","E6F4EA"),
("💡","'im + Season'","'Im Frühling' / 'Im Sommer' / 'Im Herbst' / 'Im Winter' = In spring/summer/autumn/winter. Use 'im' for all four seasons!","FFF8E1"),
("🔊","'der See' vs 'das Meer'","'der See' = lake (inland). 'das Meer' = sea/ocean. Don't confuse them! 'Am See' = at the lake. 'Am Meer' = at the sea.","FFE0E0"),
],
["Describe your favourite season in German (at least 5 sentences).","Write about nature where you live in German.","Learn the German names of 5 animals and 5 plants found in Germany."],
[("'der Herbst' means?","Autumn"),("Translate: 'Im Winter ist es kalt.'","In winter it is cold."),("'der Wald' means?","The forest"),("How do you say 'the mountain'?","der Berg"),("Translate: 'Ich liebe den Frühling.'","I love spring.")],
"Seasons and nature vocabulary is yours! You can describe the natural world, talk about seasons, and appreciate Germany's stunning landscapes in German!"
))

# ── Day 84 ─────────────────────────────────────────────────────────────────
DAYS.append((84,
"BODY PARTS 🫀",
"Der Körper — The Human Body",
"Knowing body parts is essential for medical situations, sports, describing appearance, and general conversation. Today you learn the main body parts, health expressions, and how to say where something hurts.",
"Body parts medical situations, sports, appearance describe करवा, अने सामान्य conversation माटे जरूरी छे. आज main body parts, health expressions, अने ક્यां दर्द छे ते शीखशो.",
[
("der Körper","the body","शरीर","dehr KÖR-per"),
("der Kopf","the head","माथूं","dehr KOPF"),
("das Gesicht","the face","चहेरो","das geh-ZIKHT"),
("das Auge","the eye","आंख","das OW-geh"),
("das Ohr","the ear","कान","das OHR"),
("die Nase","the nose","नाक","dee NAH-zeh"),
("der Mund","the mouth","मों","dehr MOONT"),
("der Zahn","the tooth","दांत","dehr TSAHN"),
("der Hals","the throat / neck","ग़ळूं / गरदन","dehr HALS"),
("die Schulter","the shoulder","खभो","dee SHOOL-ter"),
("der Arm","the arm","बाजु","dehr ARM"),
("die Hand","the hand","हाथ","dee HANT"),
("der Rücken","the back","पीठ","dehr RÜ-ken"),
("das Bein","the leg","पग","das BINE"),
("der Fuß","the foot","पाव","dehr FOOS"),
],
[
("Wo tut es Ihnen weh?","Where does it hurt?","ク्यां दर्द थाय छे?"),
("Mein Kopf tut weh.","My head hurts.","मारूं माथूं दुखे छे."),
("Ich habe Halsschmerzen.","I have a sore throat.","मारूं ग़ळूं दुखे छे."),
("Mein Rücken schmerzt seit drei Tagen.","My back has been hurting for three days.","त्रण दिवसथी मारी पीठ दुखे छे."),
("Ich habe Zahnschmerzen.","I have toothache.","मारा दांत दुखे छे."),
],
[
("📌","Body Part + schmerzen","German: [body part] + schmerzen = -ache. Kopfschmerzen (headache), Bauchschmerzen (stomach ache), Zahnschmerzen (toothache). Very useful!","E8F0FE"),
("💡","'weh tun' vs 'schmerzen'","'Es tut weh' (it hurts) is casual. 'Es schmerzt' is more formal/medical. Both are correct!","FFF8E1"),
("🔊","Plural Changes","Auge→Augen, Ohr→Ohren, Hand→Hände, Fuß→Füße, Zahn→Zähne. Body part plurals are irregular — learn them individually!","FFE0E0"),
],
["Write 10 sentences: '[Body part] tut weh.' or 'Ich habe [body part]-schmerzen.'","Label a body diagram with German terms.","Practise saying the body parts while touching them (kinesthetic learning)."],
[("Translate: 'Mein Bauch tut weh.'","My stomach hurts."),("How do you say 'headache' in German?","Kopfschmerzen"),("'das Knie' (knee) — can you guess its plural?","die Knie"),("Translate: 'Ich habe Rückenschmerzen.'","I have back pain."),("How do you ask 'Where does it hurt?'","Wo tut es Ihnen weh?")],
"Body vocabulary complete! Whether at the doctor, the gym, or just describing yourself, you can now communicate about the human body in German!"
))

# ── Day 85 ─────────────────────────────────────────────────────────────────
DAYS.append((85,
"AT THE DOCTOR 🏥",
"Beim Arzt — A Medical Visit in German",
"Medical appointments require specific vocabulary and polite language. Today you learn how to describe symptoms, understand a doctor's questions, and handle a medical consultation in German. Very important for life in Germany!",
"Medical appointments specific vocabulary अने polite language मांगे छे. आज symptoms describe करवां, doctor ना प्रश्नो समजवां, अने German मां medical consultation handle करवूं शीखशो.",
[
("der Arzt","the doctor (m)","ডॉक्टर (पुरुष)","dehr ARTST"),
("die Ärztin","the doctor (f)","ডॉक्टर (स्त्री)","dee ÄRT-stin"),
("die Praxis","the doctor's surgery","clinic","dee PRAK-sis"),
("der Termin","the appointment","appointment","dehr tehr-MEEN"),
("das Symptom","the symptom","symptom","das züm-TOHM"),
("das Fieber","fever","तاव","das FEE-ber"),
("der Husten","cough","खांसी","dehr HOOS-ten"),
("der Schnupfen","runny nose","शरदी","dehr SHNOO P-fen"),
("die Erkältung","cold (illness)","शरदी / ठंडी","dee ehr-KÄL-tung"),
("die Grippe","flu / influenza","फ्लू","dee GRIP-eh"),
("das Rezept","the prescription","prescription","das reh-TSEPT"),
("das Medikament","the medication","दवाई","das meh-dee-kah-MENT"),
("die Apotheke","the pharmacy","दवाखानूं","dee ah-poh-TAY-keh"),
("allergisch gegen","allergic to","एलर्जी","ah-LEHR-gish GAY-gen"),
("die Krankenversicherung","health insurance","health insurance","dee KRANK-en-fehr-zikheh-rung"),
],
[
("Ich hätte gerne einen Termin beim Arzt.","I would like an appointment with the doctor.","मारे doctor पासे appointment जोईए."),
("Ich habe seit drei Tagen Fieber.","I have had a fever for three days.","त्रण दिवसथी मने ताव छे."),
("Was für Symptome haben Sie?","What symptoms do you have?","तमने केवा symptoms छे?"),
("Ich bin gegen Penicillin allergisch.","I am allergic to penicillin.","मने Penicillin थी एलर्जी छे."),
("Nehmen Sie dieses Medikament dreimal täglich.","Take this medication three times daily.","आ दवाई दिवस मां त्रण वखत लो."),
],
[
("🇩🇪","German Healthcare","Germany has one of the best healthcare systems in the world (gesetzliche Krankenversicherung = public health insurance). Register with a Hausarzt (family doctor) when you arrive.","E6F4EA"),
("💡","'seit' + Dative","For how long you've had a symptom, use 'seit': 'Ich habe seit zwei Tagen Schmerzen.' (I've had pain for two days.) 'seit' + Dative!","FFF8E1"),
("💡","Pharmacy","German pharmacies (Apotheke) have a green cross ✚ symbol. They can advise on minor ailments without a doctor visit.","E8F0FE"),
],
["Write a dialogue: patient visits doctor with cold symptoms.","List 8 symptoms in German and how you would describe them.","Learn the emergency number in Germany: 112 (Notruf). Practise: 'Ich brauche einen Krankenwagen!'"],
[("What is the emergency number in Germany?","112"),("'das Fieber' means?","Fever"),("Translate: 'Ich habe Husten und Schnupfen.'","I have a cough and runny nose."),("Where do you collect a prescription?","At the Apotheke (pharmacy)"),("How do you say 'health insurance'?","die Krankenversicherung")],
"Medical German is now in your toolkit! This is one of the most important practical skills for living in Germany. You can describe symptoms, understand doctors, and navigate healthcare!"
))

print("Days 81-85 appended!")

# ── Day 86 ─────────────────────────────────────────────────────────────────
DAYS.append((86,
"THE PAST TENSE — PERFEKT I 📜",
"Perfekt mit 'haben' — Talking About the Past",
"Germans mainly use the Perfekt tense for speaking about past events (not the Präteritum like English). Perfekt is formed with: haben/sein + past participle (Partizip II). Today: Perfekt with 'haben'. Rule for regular verbs: ge + stem + t. Example: machen → gemacht.",
"Germans spoken German मां mainly Perfekt tense वापरे छे. Perfekt बनावानी रीत: haben/sein + past participle (Partizip II). आज 'haben' सात Perfekt. Regular verbs: ge + stem + t. दा.त.: machen → gemacht.",
[
("das Perfekt","present perfect tense","Perfekt tense","das pehr-FEKT"),
("das Partizip II","past participle","भूतकाळ नो रूप","das par-tee-TSEEP TSVY"),
("gemacht","made / done (machen)","किया","geh-MAKHT"),
("gespielt","played (spielen)","खेल्यूं","geh-SHPEELT"),
("gekauft","bought (kaufen)","खरीद्यूं","geh-KOWFT"),
("gelernt","learnt (lernen)","शीख्यूं","geh-LEHRNT"),
("gearbeitet","worked (arbeiten)","काम किया","geh-AR-by-tet"),
("gesagt","said (sagen)","कह्यूं","geh-ZAHGT"),
("gefragt","asked (fragen)","पूछ्यूं","geh-FRAHGT"),
("gehört","heard (hören)","सांभळ्यूं","geh-HÖHRT"),
("gesucht","looked for (suchen)","ढूंढ्यूं","geh-ZOOKHT"),
("gewartet","waited (warten)","राह्यूं","geh-VAR-tet"),
("Ich habe … gemacht","I have done / I did","में किया","ikh hah-beh...geh-MAKHT"),
("Hast du … gespielt?","Did you play?","शूं तें खेल्यूं?","hast doo...geh-SHPEELT"),
("Er hat … gekauft","He bought","तेणे खरीद्यूं","ehr hat...geh-KOWFT"),
],
[
("Ich habe gestern Deutsch gelernt.","I learnt German yesterday.","में काल German शीख्यूं."),
("Hast du das Buch gelesen?","Did you read the book?","शूं तें पुस्तक वांच्यूं?"),
("Wir haben Fußball gespielt.","We played football.","आपणे football खेल्यूं."),
("Sie hat das Essen gekauft.","She bought the food.","तेणे खाणूं खरीद्यूं."),
("Er hat lange gewartet.","He waited for a long time.","तेणे ઘणो समय राहेवूं पड्यूं."),
],
[
("📌","Perfekt Formula","haben (conjugated) + [ge...t] at the end. Subject → haben → middle of sentence → Partizip II at END: 'Ich habe [Buch] gelesen.'","E8F0FE"),
("💡","Partizip II Rule","Regular weak verbs: ge- + stem + -t. arbeiten→gearbeitet (add -et if stem ends in -t/-d).","FFF8E1"),
("⚠️","Separable Verbs","Separable verb Partizip II: prefix + ge + stem + t. einkaufen → eingekauft. anrufen → angerufen.","FFE0E0"),
],
["Conjugate 10 regular verbs into Partizip II.","Write 5 'Ich habe…' sentences about what you did yesterday.","Make a past-tense diary entry (3-5 sentences) for today."],
[("Partizip II of 'machen'?","gemacht"),("Partizip II of 'spielen'?","gespielt"),("Translate: 'Ich habe Deutsch gelernt.'","I learned German."),("Where does Partizip II go in the sentence?","At the end"),("Partizip II of 'einkaufen'?","eingekauft")],
"Perfekt with 'haben' is now yours! This is how Germans talk about the past in conversation. You can now tell stories, describe your day, and share memories in German!"
))

# ── Day 87 ─────────────────────────────────────────────────────────────────
DAYS.append((87,
"THE PAST TENSE — PERFEKT II 🚶",
"Perfekt mit 'sein' — Movement & Change Verbs",
"Some verbs form the Perfekt with 'sein' instead of 'haben'. These are mainly: 1) verbs of movement (gehen, fahren, fliegen…) and 2) verbs of change of state (aufwachen, werden…). Important: 'sein' and 'bleiben' also use 'sein'!",
"केटलाक verbs Perfekt 'sein' सात बनावे छे — 'haben' नहीं! मुख्यत्वे: 1) movement verbs (जवूं, जावूं, उडवूं...) अने 2) state change verbs (जागवूं, बनवूं...). 'sein' अने 'bleiben' पण 'sein' वापरे छे!",
[
("gegangen","gone (gehen)","गया","geh-GANG-en"),
("gefahren","driven/gone (fahren)","गाडी मां गया","geh-FAH-ren"),
("geflogen","flown (fliegen)","उड्यूं","geh-FLOH-gen"),
("gelaufen","run/walked (laufen)","चाल्यूं","geh-LOW-fen"),
("gekommen","come (kommen)","आव्यूं","geh-KOM-en"),
("aufgestanden","got up (aufstehen)","उठ्यूं","owf-geh-SHTAN-den"),
("eingeschlafen","fallen asleep (einschlafen)","सूई गया","eyn-geh-SHLAH-fen"),
("geblieben","stayed (bleiben)","रह्यूं","geh-BLEE-ben"),
("geworden","become (werden)","बन्यूं","geh-VOR-den"),
("passiert","happened (passieren)","थयूं","pa-SEERT"),
("Ich bin gegangen","I went","हूं गया/गई","ikh bin geh-GANG-en"),
("Er ist gefahren","He drove / he went","तेणे ड्राइव किया","ehr ist geh-FAH-ren"),
("Wir sind gekommen","We came","आपणे आव्यां","veer zint geh-KOM-en"),
("Sie ist geblieben","She stayed","ते रही","zee ist geh-BLEE-ben"),
("Bist du gelaufen?","Did you walk/run?","शूं तूं चाल्यो/ी?","bist doo geh-LOW-fen"),
],
[
("Ich bin gestern nach Hause gegangen.","I went home yesterday.","हूं काल घेर गया/गई."),
("Wir sind mit dem Zug nach Berlin gefahren.","We went to Berlin by train.","आपणे ट्रेन मां Berlin गया."),
("Er ist um 7 Uhr aufgestanden.","He got up at 7 o'clock.","ते 7 वागे उठ्यो."),
("Das Kind ist eingeschlafen.","The child fell asleep.","बाळक सूई गयूं."),
("Was ist passiert?","What happened?","शूं थयूं?"),
],
[
("🔑","sein or haben?","Movement → sein. Change of state → sein. Everything else usually → haben. When in doubt, check a dictionary!","E8F0FE"),
("⚠️","Agreement","With 'sein', Partizip II does NOT change. But common error: do NOT add endings. 'Ich bin gegangen' NOT 'ich bin gegangene'.","FFE0E0"),
("💡","Mnemonic","Think: verbs that answer 'Where did you GO?' use sein. 'I went, I came, I flew, I ran, I stayed' — all movement/state.","FFF8E1"),
],
["List 10 verbs that use 'sein' in Perfekt with their Partizip II.","Write 5 sentences about a trip you took using 'sein' Perfekt.","Write a mini story (5 sentences) mixing haben and sein Perfekt."],
[("Partizip II of 'gehen'?","gegangen"),("'Ich ___ nach Hause gegangen.' — fill in?","bin"),("Translate: 'Sie ist nach München geflogen.'","She flew to Munich."),("Does 'bleiben' use haben or sein?","sein"),("Translate: 'Wir sind früh aufgestanden.'","We got up early.")],
"Perfekt with 'sein' complete! Now you know the full Perfekt tense — you can tell any past-tense story in German. This is a massive German grammar milestone!"
))

# ── Day 88 ─────────────────────────────────────────────────────────────────
DAYS.append((88,
"IRREGULAR PAST PARTICIPLES 🔀",
"Starke Verben — Strong Verbs in Perfekt",
"Many common German verbs have IRREGULAR past participles (Partizip II). These must be memorised individually — they do not follow the ge-...-t pattern. Examples: schreiben→geschrieben, essen→gegessen, trinken→getrunken.",
"ઘणा common German verbs ना IRREGULAR past participles छे — तेमणे individually याद करवा पडे. ge-...-t pattern नहीं. दा.त.: schreiben→geschrieben, essen→gegessen, trinken→getrunken.",
[
("geschrieben","written (schreiben)","लख्यूं","geh-SHREE-ben"),
("gegessen","eaten (essen)","खाध्यूं","geh-ES-en"),
("getrunken","drunk (trinken)","पीधूं","geh-TRUNK-en"),
("gesehen","seen (sehen)","जोयूं","geh-ZAY-en"),
("gelesen","read (lesen)","वांच्यूं","geh-LAY-zen"),
("gesprochen","spoken (sprechen)","बोल्यूं","geh-SHPRO-khen"),
("geschlafen","slept (schlafen)","सूतूं","geh-SHLAH-fen"),
("genommen","taken (nehmen)","लीधूं","geh-NOM-en"),
("gerufen","called (rufen)","बोलाव्यूं","geh-ROO-fen"),
("getroffen","met (treffen)","मळ्यूं","geh-TROF-en"),
("gedacht","thought (denken)","विचार्यूं","geh-DAKHT"),
("gewusst","known (wissen)","जाण्यूं","geh-VOOST"),
("gefunden","found (finden)","શोध्यूं","geh-FOON-den"),
("gehalten","held/stopped (halten)","पकड्यूं","geh-HAL-ten"),
("angefangen","begun (anfangen)","शरू किया","AN-geh-fang-en"),
],
[
("Ich habe heute viel gegessen.","I ate a lot today.","में आज ઘணूं खाध्यूं."),
("Hast du mein Buch gesehen?","Have you seen my book?","शूं तें मारूं पुस्तक जोयूं?"),
("Wir haben uns gestern getroffen.","We met yesterday.","आपणे काल मळ्यां."),
("Er hat Deutsch gesprochen.","He spoke German.","तेणे German बोल्यूं."),
("Ich habe schlecht geschlafen.","I slept badly.","में ઘटकी ઊंघ किया."),
],
[
("📌","Learning Tip","Learn irregular Partizip II in groups by pattern: -iben→-ieben (schreiben→geschrieben), -inken→-unken (trinken→getrunken). Pattern groups help!","E8F0FE"),
("💡","Verb List Strategy","Create flashcards: Infinitive | Partizip II | Meaning. Review 5 new verbs daily. Repetition builds memory!","FFF8E1"),
("🎯","Most Common","The 15 verbs above cover 80% of everyday past conversations. Master these first before learning more irregular forms.","E6F4EA"),
],
["Create flashcards for all 15 irregular Partizip II forms.","Write a story (8 sentences) using only irregular past participles.","Group the verbs by vowel change pattern: ei→ie, i→u, a→a, etc."],
[("Partizip II of 'essen'?","gegessen"),("Partizip II of 'trinken'?","getrunken"),("Translate: 'Ich habe das Buch gelesen.'","I read the book."),("Partizip II of 'schlafen'?","geschlafen"),("Translate: 'Wir haben uns gestern getroffen.'","We met yesterday.")],
"Irregular Partizip II forms conquered! The most common ones are now in your memory. Your German storytelling ability just jumped to a whole new level!"
))

# ── Day 89 ─────────────────────────────────────────────────────────────────
DAYS.append((89,
"SIMPLE PAST — PRÄTERITUM 📖",
"Vergangenheit: sein, haben & Modal Verbs",
"While spoken German uses Perfekt, written German (books, news, narratives) uses the Präteritum (simple past). You MUST know Präteritum for: sein (war), haben (hatte), and modal verbs (konnte, musste, wollte…). These are used even in speech!",
"Spoken German Perfekt वापरे, पण लखाण मां (पुस्तक, newspaper, कहानी) Präteritum वापराय. sein (war), haben (hatte), अने modal verbs (konnte, musste, wollte...) Präteritum speech मां पण वपराय — याद कर!",
[
("war","was/were (sein Prät.)","हतो/ती/तूं","VAHR"),
("hatte","had (haben Prät.)","पास हती","HAT-eh"),
("konnte","could (können Prät.)","शकतो/ती","KON-teh"),
("musste","had to (müssen Prät.)","पडतूं","MOOS-teh"),
("wollte","wanted to (wollen Prät.)","इच्छतो/ती","VOL-teh"),
("durfte","was allowed (dürfen Prät.)","मंजूरी हती","DURF-teh"),
("sollte","was supposed to (sollen)","करवानूं हतूं","ZOL-teh"),
("mochte","liked (mögen Prät.)","गमतूं","MOKH-teh"),
("ich war","I was","हूं हतो/ती","ikh VAHR"),
("du warst","you were","तूं हतो/ती","doo VAHRST"),
("er/sie/es war","he/she/it was","ते हतो/ती","ehr VAHR"),
("wir waren","we were","आपणे हतां","veer VAH-ren"),
("ich hatte","I had","मारी पासे हतूं","ikh HAT-eh"),
("du hattest","you had","तारी पासे हतूं","doo HAT-est"),
("er hatte","he had","तारी पासे हतूं","ehr HAT-eh"),
],
[
("Ich war gestern krank.","I was ill yesterday.","हूं काल बीमार हतो/ती."),
("Wir hatten keine Zeit.","We had no time.","आपणे पासे समय न हतो."),
("Er konnte nicht kommen.","He could not come.","ते आवी न शक्यो."),
("Sie musste früh aufstehen.","She had to get up early.","तेणे जल्दी उठवूं पड्यूं."),
("Als Kind wollte ich Pilot werden.","As a child I wanted to become a pilot.","बाळपण मां हूं Pilot बनवा इच्छतो/ती."),
],
[
("📌","Präteritum Priority","Only 4 groups need Präteritum in spoken German: sein, haben, modals, and narrative verbs. For everything else use Perfekt in conversation.","E8F0FE"),
("💡","sein Präteritum","war (sing) / waren (plural). Ich war, du warst, er war, wir waren, ihr wart, sie/Sie waren. Learn this like 'was/were'!","FFF8E1"),
("🔊","Pronunciation","'war' = VAHR (rhymes with 'far'). 'hatte' = HAT-eh. 'konnte' = KON-teh. All clean, short sounds.","E6F4EA"),
],
["Conjugate sein and haben in Präteritum for all 6 persons.","Write a short narrative (5 sentences) using war/hatte/konnte/musste.","Translate: 5 English sentences using 'was' / 'had' / 'could' / 'had to'."],
[("Präteritum of 'sein' (ich)?","ich war"),("Präteritum of 'haben' (wir)?","wir hatten"),("Translate: 'Sie konnte nicht schlafen.'","She could not sleep."),("Präteritum of 'müssen' (er)?","er musste"),("Translate: 'Wir waren sehr müde.'","We were very tired.")],
"Präteritum essentials complete! sein, haben, and modals in the past tense are now yours. You can read German stories and speak about the past fluently!"
))

# ── Day 90 ─────────────────────────────────────────────────────────────────
DAYS.append((90,
"WORK & PROFESSIONS 💼",
"Arbeit und Berufe — Talking About Work",
"Work and professions are common conversation topics in Germany. Today you learn job vocabulary, how to say what you do for work, describe your workplace, and ask others about their profession.",
"Work अने professions Germany मां common conversation topics छे. आज job vocabulary, काम बद्दल कहेवूं, workplace describe करवूं, अने बीजांना profession पूछवूं शीखशो.",
[
("der Beruf","the profession/job","व्यवसाय","dehr beh-ROOF"),
("arbeiten","to work","काम करवूं","AR-by-ten"),
("der Arbeitgeber","the employer","employer","dehr AR-byt-gay-ber"),
("der Arbeitnehmer","the employee","employee","dehr AR-byt-nay-mer"),
("das Büro","the office","ऑफिस","das Bü-ROH"),
("die Firma","the company","company","dee FIR-mah"),
("der Arzt / die Ärztin","doctor","ডाक्टर","dehr ARTST"),
("der Lehrer / die Lehrerin","teacher","शिक्षक","dehr LAY-rer"),
("der Ingenieur","engineer","इंजीनियर","dehr in-zheh-n-YÖR"),
("der Anwalt","lawyer","वकील","dehr AN-valt"),
("der Verkäufer","salesperson","दुकानदार","dehr fehr-KOY-fer"),
("der Fahrer","driver","ड्राइवर","dehr FAH-rer"),
("die Krankenschwester","nurse","नर्स","dee KRANK-en-shves-ter"),
("der Koch","chef/cook","रसोइयो","dehr KOKH"),
("der Programmierer","programmer","programmer","dehr pro-gram-EE-rer"),
],
[
("Was sind Sie von Beruf?","What is your profession?","तमारो व्यवसाय शूं छे?"),
("Ich bin Ingenieur von Beruf.","I am an engineer by profession.","हूं व्यवसाय एं इंजीनियर छूं."),
("Ich arbeite bei einer großen Firma.","I work at a large company.","हूं एक मोटी company मां काम करूं छूं."),
("Wie lange arbeiten Sie schon hier?","How long have you been working here?","तमे अहीं केटलां वर्षोथी काम करो छो?"),
("Mein Traumberuf ist Arzt.","My dream profession is doctor.","मारो dream job ডाक्टर छे."),
],
[
("🇩🇪","German Work Culture","Germany has a strong work-life balance culture. Work hours are typically 8am-5pm. Punctuality (Pünktlichkeit) is extremely important!","E6F4EA"),
("💡","Gender in Jobs","Most job titles have masculine and feminine forms: Lehrer/Lehrerin, Arzt/Ärztin, Programmierer/Programmiererin. Always use the correct gender!","FFF8E1"),
("💡","Von Beruf","'Was sind Sie von Beruf?' (What do you do for work?) is the formal question. 'Was machst du beruflich?' is informal. Both are common!","E8F0FE"),
],
["Write a CV introduction in German: your name, profession, company, years of experience.","Write 5 sentences describing your current or dream job.","Learn 10 more professions in German beyond the list above."],
[("How do you ask 'What is your job?'","Was sind Sie von Beruf?"),("'das Büro' means?","The office"),("Translate: 'Ich arbeite als Lehrerin.'","I work as a (female) teacher."),("'die Firma' means?","The company"),("What German word means 'punctuality'?","Pünktlichkeit")],
"Work and profession vocabulary is yours! You can introduce yourself professionally in German and hold a workplace conversation. Career-ready German achieved!"
))

print("Days 86-90 appended!")

# ── Day 91 ─────────────────────────────────────────────────────────────────
DAYS.append((91,
"DAILY ROUTINE ⏰",
"Der Tagesablauf — A Day in Your Life",
"Describing your daily routine is one of the most common German conversation topics. Today you learn time expressions, reflexive verbs (sich waschen, sich anziehen), and how to narrate a typical day from morning to night.",
"Daily routine describe करवी German conversation मां सौथी common topics मांथी एक छे. आज time expressions, reflexive verbs (sich waschen, sich anziehen), अने सवार थी रात सुधी नो दिवस narrate करवूं शीखशो.",
[
("aufwachen","to wake up (sep.)","जागवूं","OWF-vakh-en"),
("aufstehen","to get up (sep.)","उठवूं","OWF-shtay-en"),
("sich waschen","to wash oneself (refl.)","नाहवूं","zikh VA-shen"),
("sich anziehen","to get dressed (refl.)","कपडा पहेरवां","zikh AN-tsee-en"),
("frühstücken","to have breakfast","नाश्तो करवो","FRÜ-shtük-en"),
("zur Arbeit fahren","to go to work","कामे जवूं","tsoor AR-byt FAH-ren"),
("zu Mittag essen","to have lunch","दुपारनूं जमवूं","tsoo MIT-tahg ES-en"),
("nach Hause kommen","to come home","घेर आववूं","nakh HOW-zeh KOM-en"),
("sich ausruhen","to rest (refl.)","आराम करवो","zikh OWS-roo-en"),
("Abendessen machen","to cook dinner","रात नूं खाणूं बनाववूं","AH-bent-ES-en MAK-en"),
("fernsehen","to watch TV (sep.)","TV जोवूं","FEHRN-zay-en"),
("sich duschen","to shower (refl.)","shower लेवो","zikh DOOSH-en"),
("ins Bett gehen","to go to bed","सूवा जवूं","ins BET GAY-en"),
("einschlafen","to fall asleep (sep.)","सूई जवूं","EYN-shlah-fen"),
("meistens","usually / mostly","सामान्यत:","MICE-tens"),
],
[
("Ich wache jeden Morgen um 6 Uhr auf.","I wake up every morning at 6 o'clock.","हूं दરરोज सवारे 6 वागे जागूं छूं."),
("Nach dem Frühstück fahre ich zur Arbeit.","After breakfast I go to work.","नाश्ता पछी हूं कामे जाउं छूं."),
("Meistens esse ich zu Mittag in der Kantine.","I usually eat lunch in the canteen.","सामान्यत: हूं canteen मां दुपारनूं जमूं छूं."),
("Abends sehe ich eine Stunde fern.","In the evenings I watch TV for one hour.","सांजे हूं एक कलाक TV जोउं छूं."),
("Ich schlafe um 23 Uhr ein.","I fall asleep at 11 pm.","हूं 11 वागे સूई जाउं छूं."),
],
[
("📌","Reflexive Verbs","Reflexive verbs use 'sich' as the reflexive pronoun: ich→mich, du→dich, er/sie/es→sich, wir→uns, ihr→euch, sie/Sie→sich. Learn these!","E8F0FE"),
("💡","Time Expressions","morgens (in the morning) | mittags (at noon) | abends (in the evening) | nachts (at night) | täglich (daily) | meistens (usually)","FFF8E1"),
("🇩🇪","German Day","Germans typically start work early (7-8am), have lunch 12-1pm, and finish by 5pm. Dinner (Abendessen) is often the lighter meal!","E6F4EA"),
],
["Write your complete daily routine in German using the vocabulary.","Conjugate 'sich waschen' and 'sich anziehen' for all 6 persons.","Write your ideal daily routine in German (not what you do but what you wish you did!)."],
[("Translate: 'Ich stehe um 7 Uhr auf.'","I get up at 7 o'clock."),("Reflexive pronoun for 'wir'?","uns"),("'meistens' means?","Usually / mostly"),("Translate: 'Er geht um 22 Uhr ins Bett.'","He goes to bed at 10 pm."),("'sich duschen' means?","To shower (oneself)")],
"Daily routine vocabulary mastered! You can now describe your typical day from morning to night in German — a perfect topic for getting-to-know-you conversations in Germany!"
))

# ── Day 92 ─────────────────────────────────────────────────────────────────
DAYS.append((92,
"HOBBIES & FREE TIME 🎨",
"Freizeit und Hobbys — What Do You Love?",
"Hobbies and free-time activities are great conversation topics for making friends in Germany. Today you learn hobby vocabulary, how to express likes/dislikes, and how to invite someone to join an activity.",
"Hobbies अने free-time activities Germany मां दोस्त बनाववा माटे great conversation topics छे. आज hobby vocabulary, likes/dislikes express करवां, अने किसीने activity मां invite करवूं शीखशो.",
[
("das Hobby","the hobby","hobby","das HOB-ee"),
("die Freizeit","free time","ਮੁਕ੍ਤ समय","dee FRY-tsyte"),
("lesen","to read","वांचवूं","LAY-zen"),
("schreiben","to write","लखवूं","SHRY-ben"),
("zeichnen","to draw","दोरवूं","TSYKH-nen"),
("malen","to paint","रंग करवो","MAH-len"),
("Musik hören","to listen to music","music सांभळवी","moo-ZEEK HÖ-ren"),
("singen","to sing","ગाववूं","ZING-en"),
("tanzen","to dance","ड्रांस करवो","TANT-sen"),
("kochen","to cook","रांधवूं","KO-khen"),
("reisen","to travel","प्रवास करवो","RY-zen"),
("fotografieren","to photograph","photo लेवी","fo-toh-gra-FEE-ren"),
("im Internet surfen","to surf the internet","internet करवूं","im in-ter-NET ZOOR-fen"),
("Ich interessiere mich für…","I am interested in…","मने ... मां रस छे","ikh in-teh-ress-EE-reh mikh für"),
("Ich habe keine Lust auf…","I don't feel like…","मने ... नी इच्छा नथी","ikh HAH-beh KY-neh loost owf"),
],
[
("Mein Hobby ist Fotografieren.","My hobby is photography.","मारो hobby photography छे."),
("Ich interessiere mich sehr für Musik.","I am very interested in music.","मने music मां खूब रस छे."),
("Was machst du in deiner Freizeit?","What do you do in your free time?","तूं फ्री time मां शूं करे?"),
("Ich male gerne Bilder.","I like to paint pictures.","मने ছবि दोरवी/रंगवी गमे छे."),
("Hast du Lust, heute Abend tanzen zu gehen?","Do you feel like going dancing tonight?","शूं आज सांजे dance करवा जवानी इच्छा छे?"),
],
[
("💡","'gerne' Trick","'gerne' means 'gladly / with pleasure'. Adding it after a verb = 'I like to do [verb]'. 'Ich lese gerne' = I like reading. Very natural!","FFF8E1"),
("🇩🇪","German Hobbies","Germans love: Wandern (hiking), Radfahren (cycling), Fußball, Gartenarbeit (gardening), and reading. Vereine (clubs) are very popular!","E6F4EA"),
("📌","Interest Expression","'Ich interessiere mich für + Akkusativ' = I'm interested in... 'Ich interessiere mich für Musik/Sport/Kunst.'","E8F0FE"),
],
["Write about your hobbies in German (5-7 sentences).","Ask 3 people (real or imaginary) about their hobbies and write their answers.","Describe a typical Saturday in German including your free-time activities."],
[("Translate: 'Ich lese gerne.'","I like to read."),("'die Freizeit' means?","Free time"),("How do you say 'I'm interested in…'?","Ich interessiere mich für…"),("Translate: 'Was machst du am Wochenende?'","What do you do at the weekend?"),("'tanzen' means?","To dance")],
"Hobby and free-time vocabulary complete! You can talk about what you love doing, invite friends to activities, and make meaningful connections in German society!"
))

# ── Day 93 ─────────────────────────────────────────────────────────────────
DAYS.append((93,
"SPORT & FITNESS 🏃",
"Sport treiben — Sports in German",
"Sport is a universal language and Germany is a sports-loving nation (Fußball is a religion!). Today you learn sports vocabulary, how to talk about sporting activities, and express enthusiasm for sport.",
"Sport universal language छे अने Germany sports-loving nation छे (Fußball तो religion छे!). आज sports vocabulary, sporting activities बद्दल वात करवी, अने sport माटे enthusiasm express करवूं शीखशो.",
[
("Sport treiben","to do sport","sport करवो","SHPORT TRY-ben"),
("das Fußball","football/soccer","फुटबॉल","das FOOS-bal"),
("das Tennis","tennis","टेनिस","das TEN-is"),
("das Schwimmen","swimming","तरणूं","das SHVIM-en"),
("das Laufen","running","दोडवूं","das LOW-fen"),
("das Radfahren","cycling","साइकलिंग","das RAHT-fah-ren"),
("das Wandern","hiking","पहाड मां चालवूं","das VAN-dern"),
("das Skifahren","skiing","skiing","das SHEE-fah-ren"),
("die Sporthalle","sports hall","sports hall","dee SHPORT-hal-eh"),
("das Fitnessstudio","the gym","gym","das FIT-nes-shtu-dee-oh"),
("das Mannschaft","the team","team","das MAN-shaft"),
("der Sieg","the victory","जीत","dehr ZEEK"),
("die Niederlage","the defeat","हार","dee NEE-der-lah-geh"),
("trainieren","to train","train करवूं","tray-NEE-ren"),
("gewinnen","to win","जीतवूं","geh-VIN-en"),
],
[
("Ich treibe dreimal die Woche Sport.","I do sport three times a week.","हूं अठवाडिये त्रण वखत sport करूं छूं."),
("Mein Lieblingsverein ist Bayern München.","My favourite club is Bayern Munich.","मारी favourite team Bayern Munich छे."),
("Ich gehe jeden Morgen laufen.","I go running every morning.","हूं दररोज सवारे दोडवा जाउं छूं."),
("Das Spiel hat 2:1 geendet.","The game ended 2-1.","रमत 2:1 एं पूरी थई."),
("Schwimmen ist gut für die Gesundheit.","Swimming is good for health.","तरणूं आरोग्ય माटे सारूं छे."),
],
[
("🇩🇪","German Football","The Bundesliga is Germany's football league. FC Bayern München is the most famous club. The national team won the World Cup 4 times (1954, 1974, 1990, 2014).","E6F4EA"),
("��","Sport verb","In German: 'Ich spiele Fußball/Tennis' (I play football/tennis). But: 'Ich gehe schwimmen/laufen/wandern' (I go swimming/running/hiking). Two patterns!","FFF8E1"),
("📌","Result Format","Football results: 'zwei zu eins' (2:1). 'Das Spiel steht 3:0' = The game stands at 3-0.","E8F0FE"),
],
["Write about your favourite sport in German (5 sentences).","Describe a sporting event (real or imaginary) in German using past tense.","Research: how do you say 5 other sports in German not listed above?"],
[("Translate: 'Ich spiele gerne Tennis.'","I like to play tennis."),("'das Wandern' means?","Hiking"),("How do you say 'to win'?","gewinnen"),("Translate: 'Deutschland hat gewonnen!'","Germany won!"),("'die Mannschaft' means?","The team")],
"Sports vocabulary is yours! You can cheer, discuss results, talk about your fitness routine, and bond with Germans over sport — especially Fußball!"
))

# ── Day 94 ─────────────────────────────────────────────────────────────────
DAYS.append((94,
"THE HOME 🏠",
"Zuhause — Rooms and Furniture",
"Being able to describe your home is essential for daily conversation, finding accommodation, and everyday life in Germany. Today you learn rooms of the house, furniture, and how to describe your living space.",
"घर describe करवूं daily conversation, accommodation शोधवी, अने जर्मनी मां रोजिंदी जिंदगी माटे जरूरी छे. आज घरना rooms, furniture, अने रहेठाण describe करवूं शीखशो.",
[
("das Haus","the house","घर","das HOWS"),
("die Wohnung","the flat / apartment","apartment","dee VOH-nung"),
("das Zimmer","the room","ओरडो","das TSIM-er"),
("die Küche","the kitchen","रसोड","dee KÜ-khe"),
("das Wohnzimmer","the living room","બेठक खंड","das VOHN-tsim-er"),
("das Schlafzimmer","the bedroom","सूवाना ओरडो","das SHLAHF-tsim-er"),
("das Badezimmer","the bathroom","bathroom","das BAH-deh-tsim-er"),
("die Toilette","the toilet","toilet","dee to-i-LET-eh"),
("der Garten","the garden","बगीचो","dehr GAR-ten"),
("der Balkon","the balcony","बालकनी","dehr bal-KOHN"),
("das Sofa","the sofa","sofa","das ZOH-fah"),
("der Tisch","the table","टेबल","dehr TISH"),
("das Bett","the bed","बिछानो","das BET"),
("der Schrank","the wardrobe / cupboard","कपाट","dehr SHRANK"),
("die Lampe","the lamp","दीवडी","dee LAM-peh"),
],
[
("Ich wohne in einer kleinen Wohnung.","I live in a small flat.","हूं नानी apartment मां रहूं छूं."),
("Meine Wohnung hat drei Zimmer.","My flat has three rooms.","मारी apartment मां त्रण ओरडा छे."),
("Das Wohnzimmer ist sehr gemütlich.","The living room is very cosy.","बेठक खंड खूब आरामदायक छे."),
("Wo ist die Toilette, bitte?","Where is the toilet, please?","toilet क्यां छे, कृपा?"),
("Ich habe einen kleinen Balkon mit Aussicht.","I have a small balcony with a view.","मारे दृश्य वाळी नानी balcony छे."),
],
[
("🇩🇪","German Housing","Most Germans rent (mieten) rather than own. A typical apartment is called 'Wohnung'. Finding one in cities like Munich or Berlin can be very competitive!","E6F4EA"),
("💡","Room Sizes","'1-Zimmer-Wohnung' = studio/one room. '3-Zimmer-Wohnung' = 3-room flat (excludes kitchen/bathroom). 'WG' (Wohngemeinschaft) = shared flat.","FFF8E1"),
("🔊","'gemütlich'","'gemütlich' = cosy, comfortable, warm, pleasant. This uniquely German word has no direct English translation. Very important cultural concept!","E8F0FE"),
],
["Describe your home in German using the vocabulary (8-10 sentences).","Draw your ideal German home and label all rooms and furniture in German.","Write 5 'es gibt' sentences about what is in your home."],
[("Translate: 'Ich wohne in einer Wohnung.'","I live in a flat."),("'das Schlafzimmer' means?","Bedroom"),("How do you say 'the kitchen'?","die Küche"),("'gemütlich' means?","Cosy / comfortable"),("Translate: 'Das Sofa ist im Wohnzimmer.'","The sofa is in the living room.")],
"Home vocabulary complete! You can describe your living situation, find accommodation in Germany, and feel at home discussing where you live — wherever that may be!"
))

# ── Day 95 ─────────────────────────────────────────────────────────────────
DAYS.append((95,
"HOUSEHOLD CHORES 🧹",
"Hausarbeit — Keeping House in German",
"Household chores are a practical everyday topic. Today you learn chore vocabulary, how to discuss responsibilities, and how to politely ask someone to help with tasks around the home.",
"Household chores practical everyday topic छे. आज chore vocabulary, जवाबदारी बद्दल वात करवी, अने घर ना कामों मां मदद politely मागवी शीखशो.",
[
("die Hausarbeit","household chore","घर नूं काम","dee HOWS-ar-byt"),
("aufräumen","to tidy up (sep.)","साफ-सुथरूं करवूं","OWF-roy-men"),
("putzen","to clean","सफाई करवी","POOT-sen"),
("saugen","to vacuum (sep. with 'Staub')","vacuum करवूं","ZOW-gen"),
("Staub saugen","to vacuum","vacuum करवूं","SHTOWP ZOW-gen"),
("wischen","to mop / wipe","पोछो करवो","VI-shen"),
("waschen","to wash","धोवूं","VA-shen"),
("bügeln","to iron","इस्त्री करवी","Bü-geln"),
("kochen","to cook","रांधवूं","KO-khen"),
("spülen / abwaschen","to do the dishes","वासण धोवां","SHPÜ-len / AB-va-shen"),
("den Müll rausbringen","to take out the rubbish (sep.)","कचरो बहार काढवो","dayn MÜL ROWS-bring-en"),
("einkaufen gehen","to go shopping (sep.)","खरीदी करवा जवूं","EYN-kow-fen GAY-en"),
("der Staubsauger","the vacuum cleaner","vacuum cleaner","dehr SHTOWP-zow-ger"),
("das Waschmittel","washing powder/detergent","साबु","das VASH-mit-el"),
("die Spülmaschine","dishwasher","dishwasher","dee SHPÜL-ma-shee-neh"),
],
[
("Kannst du bitte das Zimmer aufräumen?","Can you please tidy up the room?","शूं तूं room साफ करी शकशे?"),
("Ich muss heute die Wäsche waschen.","I have to do the laundry today.","मारे आज कपडा धोवाना छे."),
("Wer macht heute den Abwasch?","Who does the dishes today?","आज वासण कोण धोशे?"),
("Ich sauge jeden Samstag Staub.","I vacuum every Saturday.","हूं दर शनिवारे vacuum करूं छूं."),
("Kannst du den Müll rausbringen?","Can you take out the rubbish?","शूं तूं कचरो बहार काढशे?"),
],
[
("🇩🇪","German Cleanliness","Germans take home cleanliness very seriously! In a WG (shared flat), chores are often scheduled on a Putzplan (cleaning schedule). Tidy house = happy Germans!","E6F4EA"),
("💡","Separable Chores","aufräumen, Staub saugen, rausbringen, abwaschen — all separable. In present tense the prefix goes to the end: 'Ich räume das Zimmer auf.'","FFF8E1"),
("📌","Making Requests","'Kannst du bitte…?' (Can you please…?) is the polite way to ask someone to do a chore. Always add 'bitte' to soften requests!","E8F0FE"),
],
["Write a weekly chore schedule (Putzplan) in German.","Write 5 requests using 'Kannst du bitte…?'","Conjugate 3 separable chore verbs in present and past (Perfekt) tense."],
[("Translate: 'Ich räume mein Zimmer auf.'","I tidy up my room."),("'bügeln' means?","To iron"),("How do you say 'vacuum cleaner'?","der Staubsauger"),("Translate: 'Wer macht die Hausarbeit?'","Who does the housework?"),("'die Spülmaschine' means?","Dishwasher")],
"Household chore vocabulary mastered! You can now manage a German household, communicate with flatmates, and keep your German home sparkling clean!"
))

print("Days 91-95 appended!")

# ── Day 96 ─────────────────────────────────────────────────────────────────
DAYS.append((96,
"PHONE CALLS & MESSAGES 📱",
"Am Telefon — Communicating in German",
"Phone calls in a foreign language can be nerve-wracking! Today you learn telephone vocabulary, how to introduce yourself on the phone, leave a message, and send written messages in German.",
"विदेशी भाषा मां phone call ने nerves ना काम! आज telephone vocabulary, phone पर परिचय, message छोडवो, अने German मां written messages शीखशो.",
[
("das Telefon","the telephone","telephone","das teh-leh-FOHN"),
("das Smartphone","the smartphone","smartphone","das SMART-fohn"),
("anrufen","to call (sep.)","phone करवूं","AN-roo-fen"),
("abheben","to answer (phone) (sep.)","phone उपाडवो","AB-hay-ben"),
("auflegen","to hang up (sep.)","phone मूकवो","OWF-lay-gen"),
("eine Nachricht hinterlassen","to leave a message","message छोडवो","EY-neh NAKH-rikht HIN-ter-las-en"),
("Wer spricht dort?","Who is speaking?","कोण बोले छे?","vehr SHPRIKHT dort"),
("Einen Moment, bitte.","One moment, please.","एक क्षण, कृपा.","EY-nen moh-MENT BIT-eh"),
("Ich rufe später zurück.","I'll call back later.","हूं पछी call करूं छूं.","ikh ROO-feh SHPÄ-ter tsoo-RÜK"),
("die SMS","the text message","SMS","dee ES-em-ES"),
("die WhatsApp-Nachricht","the WhatsApp message","WhatsApp message","dee VATS-ap-NAKH-rikht"),
("die E-Mail","the email","email","dee EE-mail"),
("schicken","to send","मोकलवूं","SHIK-en"),
("empfangen","to receive","मळवूं","em-PFANG-en"),
("das Passwort","the password","password","das PAS-vort"),
],
[
("Guten Tag, hier spricht Müller.","Good day, this is Müller speaking.","नमस्ते, हूं Müller बोलूं छूं."),
("Kann ich bitte mit Frau Schmidt sprechen?","Can I speak with Ms Schmidt please?","शूं मने Frau Schmidt सात वात कराववी?"),
("Ich bin gerade nicht erreichbar.","I am not reachable right now.","हूं अत्यारे उपलब्ध नथी."),
("Bitte hinterlassen Sie eine Nachricht.","Please leave a message.","कृपा message छोडो."),
("Ich schicke dir die Information per WhatsApp.","I'll send you the information via WhatsApp.","हूं तने WhatsApp पर information मोकलूं छूं."),
],
[
("🇩🇪","German Phone Culture","Germans are private about phone numbers. Do NOT call before 8am or after 9pm (except emergencies). Texting first is increasingly common.","E6F4EA"),
("💡","On the Phone","Standard opening: 'Guten Tag / Guten Morgen, hier spricht [your name].' Always identify yourself immediately!","FFF8E1"),
("📌","Voicemail German","'Sie haben die Mailbox von [name] erreicht' = You've reached [name]'s voicemail. 'Bitte hinterlassen Sie eine Nachricht' = Please leave a message.","E8F0FE"),
],
["Write a phone call script: calling to make a doctor's appointment.","Write 3 WhatsApp messages in German to different people.","Practise: say your phone number in German and spell your name using the German alphabet."],
[("How do you identify yourself on the phone?","'Hier spricht [name].' (This is [name] speaking.)"),("Translate: 'Ich rufe später zurück.'","I'll call back later."),("'auflegen' means?","To hang up"),("Translate: 'Bitte schicken Sie mir eine E-Mail.'","Please send me an email."),("'empfangen' means?","To receive")],
"Phone and messaging vocabulary complete! You can handle calls, leave voicemails, and write messages in German — essential for daily communication in Germany!"
))

# ── Day 97 ─────────────────────────────────────────────────────────────────
DAYS.append((97,
"MAKING APPOINTMENTS 📅",
"Termine vereinbaren — Scheduling in German",
"In Germany, everything runs on appointments! Doctors, hairdressers, offices, even friends — you need a Termin (appointment). Today you learn how to schedule, confirm, change, and cancel appointments politely.",
"Germany मां बधूं appointments पर चाले! Doctor, hairdresser, office, even दोस्त — Termin (appointment) जोईए. आज schedule, confirm, change, अने politely cancel करवूं शीखशो.",
[
("der Termin","the appointment","appointment","dehr tehr-MEEN"),
("vereinbaren","to arrange / schedule","ठराववूं","fehr-EYN-bah-ren"),
("bestätigen","to confirm","पुष्टि करवी","beh-SHTÄ-tikh-en"),
("absagen","to cancel (sep.)","रद्द करवूं","AB-zah-gen"),
("verschieben","to postpone","मुलतवी रखवूं","fehr-SHEE-ben"),
("passen","to suit","અनुकूळ होवूं","PAS-en"),
("der Kalender","the calendar","calendar","dehr kah-LEN-der"),
("die Uhrzeit","the time","समय","dee OOR-tsyte"),
("frei","free / available","उपलब्ध","FRY"),
("besetzt","busy / occupied","व्यस्त","beh-ZETST"),
("vorschlagen","to suggest (sep.)","सूचन करवूं","FOR-shlah-gen"),
("Wann passt es Ihnen?","When suits you?","तमने क्यारे ফ़ायदाकारक छे?","van PAST es EE-nen"),
("Das geht leider nicht.","Unfortunately that won't work.","दुर्भाग्यथी नहीं थाय.","das GAYT LY-der nikht"),
("Ich halte den Termin.","I'll keep the appointment.","हूं appointment मां आवीश.","ikh HAL-teh dayn tehr-MEEN"),
("Können wir es auf Donnerstag verschieben?","Can we postpone to Thursday?","शूं आपणे गुरुवारे ले जई शकीए?","KÖN-en veer es owf DON-ers-tahg fehr-SHEE-ben"),
],
[
("Ich möchte einen Termin vereinbaren.","I would like to make an appointment.","मारे appointment लेवी छे."),
("Wann haben Sie einen freien Termin?","When do you have a free appointment?","तमारे क्यारे खाली appointment छे?"),
("Dienstag um 15 Uhr würde passen.","Tuesday at 3 pm would suit.","मंगळवारे 3 वाग्ये अनुकूळ रहेशे."),
("Leider muss ich den Termin absagen.","Unfortunately I have to cancel the appointment.","दुर्भाग्यथी मारे appointment रद्द करवी पडशे."),
("Können wir einen neuen Termin finden?","Can we find a new appointment?","शूं आपणे नवी appointment शोधी शकीए?"),
],
[
("🇩🇪","German Punctuality","Pünktlichkeit (punctuality) is sacred in Germany! Arrive 5 minutes EARLY for appointments. Arriving late (without calling) is very rude.","E6F4EA"),
("💡","Making Appointments","When calling: state your name, state why you're calling, ask for an appointment: 'Ich möchte einen Termin beim Arzt vereinbaren.'","FFF8E1"),
("📌","Days of the Week","Montag, Dienstag, Mittwoch, Donnerstag, Freitag, Samstag, Sonntag. Always capitalised in German!","E8F0FE"),
],
["Write a phone dialogue arranging a dentist appointment.","Write a message changing an appointment to a different day and time.","Write your own weekly schedule in German using days and times."],
[("Translate: 'Ich muss den Termin leider absagen.'","I unfortunately have to cancel the appointment."),("'vereinbaren' means?","To arrange / schedule"),("How do you ask 'When suits you?'","Wann passt es Ihnen?"),("'besetzt' means?","Busy / occupied"),("Translate: 'Mittwoch um 10 Uhr passt mir gut.'","Wednesday at 10 o'clock suits me well.")],
"Appointment vocabulary mastered! You can now schedule, confirm, change and cancel appointments in German — critical for life in Germany's punctual, appointment-driven culture!"
))

# ── Day 98 ─────────────────────────────────────────────────────────────────
DAYS.append((98,
"SCHOOL & EDUCATION 📚",
"Schule und Bildung — Learning in German",
"Education is a key topic for students, parents, and anyone integrating into German society. Today you learn school vocabulary, describe your education, and discuss learning in German.",
"Education students, parents, अने German society मां integrate थवा माटे key topic छे. आज school vocabulary, education describe करवी, अने German मां learning बद्दल discuss करवूं शीखशो.",
[
("die Schule","the school","शाळा","dee SHOO-leh"),
("die Universität","the university","university","dee oo-ni-vehr-zi-TÄT"),
("die Ausbildung","vocational training","व्यावसायिक तालीम","dee OWS-bil-dung"),
("das Studium","university studies","university ना अभ्यास","das SHTOO-dee-oom"),
("der Student / die Studentin","student (uni)","university student","dehr shtoo-DENT"),
("der Schüler / die Schülerin","school pupil","शाळाना विद्यार्थी","dehr SHÜL-er"),
("der Lehrer / die Lehrerin","teacher","शिक्षक","dehr LAY-rer"),
("das Fach","the school subject","विषय","das FAKH"),
("die Note","the grade","ग्रेड","dee NOH-teh"),
("die Prüfung","the exam","परीक्षा","dee PRÜ-fung"),
("bestehen","to pass (an exam)","pass करवूं","beh-SHTAY-en"),
("durchfallen","to fail (an exam)","fail थवूं","DOOKH-fal-en"),
("das Zeugnis","school report","report card","das TSOYG-nis"),
("der Abschluss","the qualification / degree","degree","dehr AP-shloos"),
("lernen","to study / learn","शीखवूं","LEHR-nen"),
],
[
("Ich studiere Informatik an der Universität.","I am studying computer science at university.","हूं university मां computer science भणूं छूं."),
("Welche Fächer magst du am liebsten?","Which subjects do you like best?","तने कया विषय सौथी वधु गमे छे?"),
("Ich habe die Prüfung mit einer Eins bestanden.","I passed the exam with a grade 1.","में परीक्षा ग्रेड 1 सात pass करी."),
("In Deutschland ist die Schulpflicht bis 18 Jahre.","In Germany schooling is compulsory until 18.","Germany मां 18 वर्ष सुधी शाळा अनिवार्य छे."),
("Ich mache eine Ausbildung als Elektriker.","I am doing vocational training as an electrician.","हूं electrician तरीके vocational training करूं छूं."),
],
[
("🇩🇪","German School System","Germany has different school tracks: Hauptschule, Realschule, Gymnasium. Gymnasium leads to Abitur (A-level equivalent) needed for university.","E6F4EA"),
("💡","German Grades","1=sehr gut (excellent), 2=gut (good), 3=befriedigend (satisfactory), 4=ausreichend (sufficient), 5=mangelhaft (poor), 6=ungenügend (fail). 1 is best!","FFF8E1"),
("📌","'Studium' vs 'Lernen'","'lernen' = to learn (general / school). 'studieren' = to study at university. Don't mix them up!","FFE0E0"),
],
["Write about your own education history in German.","Describe your favourite school subject and why you like it in German.","Write 5 questions you might ask a German student about their studies."],
[("'die Prüfung' means?","The exam"),("What is grade 1 in the German system?","sehr gut (excellent)"),("Translate: 'Ich studiere Medizin.'","I study medicine."),("'bestehen' means (in context of exams)?","To pass"),("What is the Abitur?","The German school leaving certificate / university entrance qualification")],
"Education vocabulary complete! You can discuss school, exams, subjects, and university in German. Whether you're a student or parent, this vocabulary is invaluable in Germany!"
))

# ── Day 99 ─────────────────────────────────────────────────────────────────
DAYS.append((99,
"GERMAN CULTURE & FESTIVALS 🎉",
"Deutsche Kultur und Feste — Celebrating Germany",
"Understanding German culture makes you a better communicator and helps you integrate into German society. Today you learn about major German festivals, cultural traditions, and social customs.",
"German culture समजवी तमने better communicator बनावे अने German society मां integrate थवा मदद करे. आज major German festivals, cultural traditions, अने social customs शीखशो.",
[
("das Oktoberfest","Oktoberfest beer festival","October festival","das OK-toh-ber-fest"),
("das Weihnachten","Christmas","ख्रिस्मस","das VY-nakh-ten"),
("Weihnachtsmarkt","Christmas market","Christmas market","VY-nakhts-markt"),
("Ostern","Easter","ईस्टर","OH-stern"),
("Karneval / Fasching","Carnival / Mardi Gras","Carnival","kar-neh-VAL / FA-shing"),
("der Tag der Deutschen Einheit","German Unity Day","German Unity Day (Oct 3)","dehr TAHG dehr DOY-tsh-en EYN-hite"),
("Silvester","New Year's Eve","नव वर्ष ની ઉꬁᵀ","zil-VES-ter"),
("Neujahr","New Year","नवूं वर्ष","NOY-yahr"),
("die Tradition","the tradition","परंपरा","dee tra-dit-ZYOHN"),
("feiern","to celebrate","ઉꬁᵀ ઉ꬀","FY-ern"),
("der Brauch","the custom","રिवाज","dehr BROWKH"),
("schenken","to give (a gift)","ভेट आपवी","SHENK-en"),
("das Geschenk","the gift","ভेट","das geh-SHENK"),
("Frohe Weihnachten!","Merry Christmas!","ख्रिस्मस नी शुभकामना!","FROH-eh VY-nakh-ten"),
("Prosit Neujahr!","Happy New Year!","नवा वर्षनी शुभकामना!","PROH-zit NOY-yahr"),
],
[
("Das Oktoberfest findet jedes Jahr in München statt.","Oktoberfest takes place every year in Munich.","Oktoberfest दर वर्षे Munich मां थाय छे."),
("Zu Weihnachten schenken sich die Deutschen Geschenke.","At Christmas Germans give each other gifts.","Christmas मां Germans एकबीजाने ভेट आपे छे."),
("Karneval ist besonders in Köln und Düsseldorf beliebt.","Carnival is especially popular in Cologne and Düsseldorf.","Carnival Köln अने Düsseldorf मां खास popular छे."),
("Am 3. Oktober feiern wir den Tag der Deutschen Einheit.","On 3rd October we celebrate German Unity Day.","3 October एं आपणे German Unity Day उꬁᵀ कर्वाय."),
("Auf dem Weihnachtsmarkt gibt es Glühwein und Lebkuchen.","At the Christmas market there is mulled wine and gingerbread.","Christmas market पर Glühwein अने Lebkuchen मळे."),
],
[
("🇩🇪","Oktoberfest","The world's biggest folk festival! 6 million visitors, traditional Dirndl and Lederhosen clothing, beer halls (Bierzelte), and Bavarian food. Prost! (Cheers!)","E6F4EA"),
("🎄","Weihnachten","Christmas is the biggest German celebration. Advent starts 4 Sundays before Christmas. Weihnachtsmarkt (Christmas markets) are magical!","FFF8E1"),
("💡","Greeting Cards","'Frohe Weihnachten!' (Merry Christmas) | 'Frohe Ostern!' (Happy Easter) | 'Alles Gute zum Geburtstag!' (Happy Birthday!)","E8F0FE"),
],
["Research and write a short paragraph about Oktoberfest in German.","Learn the dates of Germany's 5 public holidays.","Write a Weihnachtskarte (Christmas card) in German to a friend."],
[("When is German Unity Day?","3rd October (3. Oktober)"),("Translate: 'Frohe Weihnachten!'","Merry Christmas!"),("What is 'Glühwein'?","Mulled wine (hot spiced wine served at Christmas markets)"),("'feiern' means?","To celebrate"),("Where is Oktoberfest held?","Munich (München)")],
"German culture and festivals vocabulary complete! You now understand the heart of German social life. Join the celebrations — Prost! Frohe Feste! (Cheers! Happy celebrations!)"
))

# ── Day 100 ─────────────────────────────────────────────────────────────────
DAYS.append((100,
"100 DAYS! 🎊 GRAND CELEBRATION TEST",
"Das große 100-Tage-Fest — You Made It!",
"INCREDIBLE! 100 days of German learning! You have come an enormous distance from Day 1 greetings. Today is a celebration AND a comprehensive review. Test yourself on everything from Days 1-100 and see how far you have come. You should be proud!",
"અવিश्वसनीय! 100 दिवस German शिखवानां! Day 1 ना greetings थी खूब आगे आव्यां. आज celebration AND comprehensive review बंने छे. Days 1-100 ना बधां पर खुद ने test करो अने जुओ केटले दूर आव्यां. Tame proud थवो जोईए!",
[
("das Jubiläum","the jubilee / anniversary","उत्सव","das yoo-bi-LÄ-oom"),
("der Meilenstein","the milestone","milestone","dehr MYL-en-shtyne"),
("die Leistung","achievement / performance","उपलब्धि","dee LYS-tung"),
("stolz","proud","गर्वित","SHTOLTS"),
("der Durchhaltewille","perseverance","दृढता","dehr DOOKH-hal-teh-vil-eh"),
("der Wortschatz","vocabulary","शब्दभंडोळ","dehr VORT-shats"),
("die Grammatik","grammar","व्याकरण","dee gra-MAH-tik"),
("die Aussprache","pronunciation","उच्चार","dee OWS-spra-kheh"),
("das Selbstvertrauen","self-confidence","आत्मविश्वास","das ZELPST-fehr-trow-en"),
("weitermachen","to keep going (sep.)","आगे वधvun","VY-ter-makh-en"),
("die Herausforderung","the challenge","पडकार","dee heh-ROWS-for-deh-rung"),
("meistern","to master","महारत मेळववी","MYS-tern"),
("auf dem richtigen Weg","on the right path","सही रस्ते","owf daym RIKH-tikh-en VAYK"),
("Bravo!","Bravo! Well done!","शाबाश!","BRAH-voh"),
("Weiter so!","Keep it up!","यही चालू रखो!","VY-ter zo"),
],
[
("Ich habe 100 Tage Deutsch gelernt — ich bin so stolz!","I have learned German for 100 days — I am so proud!","में 100 दिवस German शीख्यूं — हूं खूब गर्वित छूं!"),
("Mein Wortschatz hat sich sehr verbessert.","My vocabulary has improved a lot.","मारो शब्दभंडोळ ઘणो सुधर्यो."),
("Die größte Herausforderung war die Grammatik.","The biggest challenge was the grammar.","सौथी मोटो पडकार grammar हतो."),
("Jetzt fühle ich mich sicherer auf Deutsch.","Now I feel more confident in German.","हवे मने German मां confidence वधारे लागे छे."),
("Ich mache weiter bis zum Ende!","I'll keep going until the end!","हूं अंत सुधी आगे वधीश!"),
],
[
("🏆","100-Day Achievement","You know: Dative & Accusative cases, 6 modal verbs, separable verbs, Perfekt & Präteritum tenses, 50+ A2 topics, 1000+ German words!","E8F0FE"),
("🎯","What's Ahead","Days 101-120 will teach you: future tense, comparative/superlatives, relative clauses, adjective endings, and prepare you for B1!","FFF8E1"),
("💪","Motivation","Every language expert started exactly where you started. You've proven you have the discipline. The fluency is coming. Weiter so!","E6F4EA"),
],
["Write a 10-sentence reflection on your 100-day German journey.","Test yourself: translate 20 random sentences from Days 1-100.","Write your German goals for the next 100 days."],
[("How do you say 'I am proud' in German?","Ich bin stolz."),("Translate: 'Weiter so!'","Keep it up!"),("What is the Perfekt of 'lernen'?","gelernt (Ich habe gelernt)"),("How do you say 'vocabulary' in German?","der Wortschatz"),("Translate: 'Ich mache weiter!'","I keep going!")],
"DAY 100 COMPLETE! Herzlichen Glückwunsch! (Congratulations!) You have achieved something most people never attempt. Your German is real, growing, and powerful. The final 20 days will take you to the peak of A2. You are unstoppable!"
))

print("Days 96-100 appended!")

# ── Day 101 ─────────────────────────────────────────────────────────────────
DAYS.append((101,
"FUTURE TENSE — FUTUR I 🔮",
"Zukunft: werden + Infinitiv",
"German has a future tense formed with 'werden' + infinitive. However, Germans often use the present tense with a future time expression for near future events. 'Werden' + infinitive is used for more certain or formal future statements.",
"German future tense 'werden' + infinitive थी बनाव. पण Germans नजीक ना future माटे present tense + future time expression वापरे. 'Werden' + infinitive formal या निश्चित future माटे.",
[
("werden","to become / will (future)","बनवूं / थशे","VEHR-den"),
("ich werde","I will","हूं ...शे","ikh VEHR-deh"),
("du wirst","you will","तूं ...शे","doo VEERST"),
("er/sie/es wird","he/she/it will","ते ...शे","ehr VEERT"),
("wir werden","we will","आपणे ...शुं","veer VEHR-den"),
("ihr werdet","you all will","तमे ...शो","eer VEHR-det"),
("sie/Sie werden","they/you(formal) will","ते/आप ...शे","zee VEHR-den"),
("morgen","tomorrow","काल","MOR-gen"),
("nächste Woche","next week","अगली अठवाड","NÄKHS-teh VO-kheh"),
("nächstes Jahr","next year","अगलूं वर्ष","NÄKHS-tes YAHR"),
("bald","soon","जल्दी","BALT"),
("in der Zukunft","in the future","भविष्य मां","in dehr TSOO-koonft"),
("eines Tages","one day","एक दिवस","EY-nes TAH-ges"),
("hoffentlich","hopefully","आशा छे ते","HOF-ent-likh"),
("wahrscheinlich","probably","संभवत:","VAHR-shhyne-likh"),
],
[
("Ich werde morgen Deutsch lernen.","I will learn German tomorrow.","काले हूं German शीखीश."),
("Er wird nächstes Jahr nach Deutschland ziehen.","He will move to Germany next year.","ते अगलां वर्षे Germany shift थशे."),
("Was wirst du in der Zukunft machen?","What will you do in the future?","तूं भविष्य मां शूं करशे?"),
("Hoffentlich werde ich fließend Deutsch sprechen.","Hopefully I will speak German fluently.","आशा छे हूं fluent German बोलीश."),
("Wir werden bald ankommen.","We will arrive soon.","आपणे जल्दी पहोंचशुं."),
],
[
("��","Future with Present","For near future, present tense is common: 'Ich lerne morgen Deutsch.' (I'm learning German tomorrow.) = same meaning as Futur I!","E8F0FE"),
("💡","'werden' also means...","'werden' also means 'to become': 'Ich werde Arzt.' (I'm becoming a doctor.) Context distinguishes future auxiliary from 'to become'.","FFF8E1"),
("🔊","Conjugation","werden conjugation is irregular: ich werde, du WIRST, er WIRD — note the vowel change in 2nd and 3rd person singular!","FFE0E0"),
],
["Conjugate 'werden' for all 6 persons.","Write 5 plans for the next year using Futur I.","Write a vision statement for your German future: 'In 5 Jahren werde ich...'"],
[("Future tense auxiliary verb in German?","werden"),("Conjugate: 'er' form of werden?","er wird"),("Translate: 'Ich werde bald fließend Deutsch sprechen.'","I will soon speak German fluently."),("'hoffentlich' means?","Hopefully"),("Translate: 'Was wirst du morgen machen?'","What will you do tomorrow?")],
"Future tense mastered! You can now talk about plans, dreams, and predictions in German. The future is bright — auf Deutsch natürlich! (In German, of course!)"
))

# ── Day 102 ─────────────────────────────────────────────────────────────────
DAYS.append((102,
"COMPARATIVE & SUPERLATIVE 📊",
"Komparativ und Superlativ — Comparing Things",
"Comparisons make your German much richer and more natural. Today you learn how to compare things using comparatives (bigger, smaller, better) and superlatives (the biggest, the best, the most beautiful).",
"Comparisons तमारूं German ઘणू natural बनावे. आज comparative (मोटूं, नानूं, सारूं) अने superlative (सौथी मोटूं, सौथी सारूं) शीखशो.",
[
("der Komparativ","comparative","तुलनात्मक","dehr kom-pah-rah-TEEF"),
("der Superlativ","superlative","सर्वोच्च","dehr zoo-per-lah-TEEF"),
("größer","bigger (groß)","मोटूं","GRÖ-ser"),
("kleiner","smaller (klein)","नानूं","KLY-ner"),
("besser","better (gut)","सारूं","BES-er"),
("schlechter","worse (schlecht)","खराब","SHLEKH-ter"),
("schneller","faster (schnell)","झड़पी","SHNEL-er"),
("langsamer","slower (langsam)","ধीरूं","LANG-zah-mer"),
("teurer","more expensive (teuer)","महंगूं","TOY-rer"),
("günstiger","cheaper (günstig)","ससतूं","GÜN-stikh-er"),
("am größten","the biggest","सौथी मोटूं","am GRÖ-sten"),
("am besten","the best","सौथी सारूं","am BES-ten"),
("am schnellsten","the fastest","सौथी ज़ड़प","am SHNEL-sten"),
("als","than (comparative)","...करता","ALS"),
("genauso … wie","just as … as","एटलूं ज ... जेटलूं","geh-NOW-zo...vee"),
],
[
("Berlin ist größer als München.","Berlin is bigger than Munich.","Berlin Munich करता मोटूं छे."),
("Mein Bruder ist schneller als ich.","My brother is faster than me.","मारो भाई मारा करता झड़पो छे."),
("Das ist das beste Restaurant in der Stadt.","That is the best restaurant in the city.","तेे शहेर नूं सौथी सारूं restaurant छे."),
("Deutsch ist nicht so schwer wie man denkt.","German is not as hard as one thinks.","जर्मन एटलूं अधरूं नथी जेटलूं कोई विचारे."),
("Je mehr du übst, desto besser wirst du.","The more you practise, the better you become.","जेटलूं तूं practice करे, उतरूं तूं better बनशे."),
],
[
("📌","Comparative Formation","Add -er to adjective: schnell→schneller, groß→größer (umlaut!), teuer→teurer. Irregular: gut→besser, viel→mehr.","E8F0FE"),
("💡","Superlative","'am + adjective + sten': am schnellsten, am größten, am besten. Irregulars: gut→am besten, viel→am meisten.","FFF8E1"),
("⚠️","'als' vs 'wie'","Comparative + unequal: use 'als' (than). Equal comparison: use 'genauso...wie' (just as...as). NEVER mix them!","FFE0E0"),
],
["Make a comparison table: 5 things in your city using comparative and superlative.","Write 5 sentences comparing Germany to your home country.","Learn the irregular comparatives: gut/besser/am besten, viel/mehr/am meisten, gern/lieber/am liebsten."],
[("Comparative of 'groß'?","größer"),("Superlative of 'gut'?","am besten"),("Translate: 'Das ist teurer als das andere.'","That is more expensive than the other."),("'genauso...wie' means?","Just as...as"),("Comparative of 'viel'?","mehr")],
"Comparatives and superlatives mastered! Your German now sounds much more natural and sophisticated. You can compare anything — from food to cities to people!"
))

# ── Day 103 ─────────────────────────────────────────────────────────────────
DAYS.append((103,
"SUBORDINATING CONJUNCTIONS 🔗",
"Nebensätze: weil, dass, ob, wenn, obwohl…",
"Subordinating conjunctions connect a main clause to a subordinate clause. The KEY rule: in the subordinate clause, the VERB goes to the END. This is one of the most important grammar rules in A2 German!",
"Subordinating conjunctions main clause ने subordinate clause सात जोडे. KEY rule: subordinate clause मां VERB ने अंते जाय. आ A2 German मां सौथी महत्त्वनो grammar rule छे!",
[
("weil","because","कारण के","VILE"),
("dass","that (conjunction)","के","DAS"),
("ob","whether / if (indirect)","के शूं","OP"),
("wenn","when / if (conditional)","जो / ज्यारे","VEN"),
("obwohl","although","ছतां","op-VOHL"),
("damit","so that","जेथी","dah-MIT"),
("bevor","before","पहेलां","beh-FOR"),
("nachdem","after","पछी","nakh-DAYM"),
("seitdem","since (time)","त्यारथी","zite-DAYM"),
("während","while / during","दरम्यान","VÄH-rent"),
("sobald","as soon as","जेवूं...केवूं","zo-BALT"),
("bis","until","सुधी","BIS"),
("falls","in case / if","जो...तो","FALS"),
("trotzdem","nevertheless (adv.)","तेम छतां","TROTS-daym"),
("weder … noch","neither … nor","न...न","VAY-der...nokh"),
],
[
("Ich lerne Deutsch, weil ich in Deutschland arbeiten möchte.","I learn German because I want to work in Germany.","हूं German शीखूं छूं कारण के Germany मां काम करवूं छे."),
("Ich weiß, dass Deutsch schwer ist.","I know that German is hard.","मने खबर छे के German अधरूं छे."),
("Er fragt, ob du morgen kommst.","He asks whether you're coming tomorrow.","ते पूछे छे के शूं तूं काल आवशे."),
("Obwohl es regnet, gehe ich spazieren.","Although it is raining, I go for a walk.","वरसाद छतां हूं ચालवा जाउं छूं."),
("Ruf mich an, sobald du ankommst.","Call me as soon as you arrive.","जेवो तूं पहोंचे केवो मने call करजे."),
],
[
("🔑","VERB LAST Rule","In every subordinate clause (after weil, dass, ob, wenn etc.) the conjugated verb ALWAYS goes to the very end. This is non-negotiable!","E8F0FE"),
("⚠️","Word Order Change","Main: 'Ich lerne Deutsch.' Sub: '...weil ich Deutsch LERNE.' The verb 'lerne' jumps to end! Modal verbs: infinitive last, modal before it.","FFE0E0"),
("💡","'weil' vs 'denn'","Both mean 'because' but: 'weil' → verb to end (subordinate clause). 'denn' → normal word order (coordinating conjunction). Very different!","FFF8E1"),
],
["Write 5 'weil' sentences explaining why you like/dislike things.","Write 3 'obwohl' sentences (contrasting ideas).","Transform 5 main clauses into subordinate clauses using 'dass'."],
[("In a 'weil' clause, where does the verb go?","To the end of the sentence"),("Translate: 'Ich komme nicht, weil ich krank bin.'","I'm not coming because I am ill."),("'obwohl' means?","Although"),("Translate: 'Ich weiß, dass er kommt.'","I know that he is coming."),("What is the difference between 'weil' and 'denn'?","'weil' sends verb to end; 'denn' keeps normal word order")],
"Subordinating conjunctions mastered! This is one of the biggest grammar upgrades in A2. Your sentences are now complex, nuanced, and impressive!"
))

# ── Day 104 ─────────────────────────────────────────────────────────────────
DAYS.append((104,
"ADJECTIVE ENDINGS I 📝",
"Adjektivdeklination — After Definite Articles",
"German adjectives change their endings depending on case, gender, and article. Today you learn adjective endings after DEFINITE articles (der/die/das/die). This is the weak declension.",
"German adjectives case, gender, अने article प्रमाणे ending बदले. आज DEFINITE articles (der/die/das/die) पछी adjective endings शीखशो — weak declension.",
[
("die Adjektivendung","adjective ending","adjective ending","dee ad-YEK-teef-END-ung"),
("stark / schwach","strong / weak (declension)","strong / weak","SHTARK / SHVAKH"),
("der neue Mantel","the new coat (Nom. m)","नवो ઓvercoat","dehr NOY-eh MAN-tel"),
("die neue Jacke","the new jacket (Nom. f)","नवी jacket","dee NOY-eh YAK-eh"),
("das neue Hemd","the new shirt (Nom. n)","नवूं shirt","das NOY-eh HEMD"),
("den neuen Mantel (Acc. m)","the new coat (Acc. m)","नवो coat (acc)","dayn NOY-en MAN-tel"),
("der neuen Jacke (Dat. f)","the new jacket (Dat. f)","नवी jacket (dat)","dehr NOY-en YAK-eh"),
("dem neuen Mantel (Dat. m)","the new coat (Dat. m)","नवा coat ने","daym NOY-en MAN-tel"),
("das schöne Haus","the beautiful house","सुंदर घर","das shö-neh HOWS"),
("die rote Rose","the red rose","लाल गुलाब","dee ROH-teh ROH-zeh"),
("der alte Mann","the old man","वृद्ध पुरुष","dehr AL-teh MAN"),
("das kleine Kind","the small child","नानूं बाळक","das KLY-neh KINT"),
("die langen Haare (pl)","the long hair","लांबा वाळ","dee LANG-en HAH-reh"),
("-e ending","nominative (m/f/n/pl definite)","Nom. ending","-eh"),
("-en ending","accusative/dative (all genders)","Acc/Dat ending","-en"),
],
[
("Ich sehe den alten Mann.","I see the old man.","हूं वृद्ध पुरुषने जोउं छूं."),
("Das rote Kleid ist sehr schön.","The red dress is very beautiful.","लाल dress खूब सुंदर छे."),
("Ich gebe dem kleinen Kind ein Buch.","I give the small child a book.","हूं नाना बाळकने पुस्तक आपूं छूं."),
("Die neuen Schuhe sind sehr bequem.","The new shoes are very comfortable.","नवा जूता खूब आरामदायक छे."),
("Er wohnt in dem großen Haus.","He lives in the large house.","ते मोटा घर मां रहे छे."),
],
[
("📌","Weak Ending Summary","After der/die/das: Nominative = -e. Accusative (m) = -en. ALL other cases/genders = -en. So: mostly -en, with a few -e in Nominative.","E8F0FE"),
("💡","Quick Rule","After definite articles, most adjective endings are -en. Only Nominative (and Acc f/n) takes -e. When in doubt, use -en!","FFF8E1"),
("⚠️","Why It Matters","Adjective endings signal case and gender. Without them, sentences sound unnatural. Germans notice wrong endings immediately!","FFE0E0"),
],
["Make a table: adjective 'alt' (old) in all cases after definite articles.","Write 5 sentences using adjectives in different cases.","Go back through Days 70-95 and add adjectives to 10 sentences you wrote."],
[("'der alte Mann' in Accusative?","den alten Mann"),("Adjective ending after 'die' in Nominative?","-e (die rote Rose)"),("Adjective ending in Dative?","-en (dem alten Mann)"),("Translate: 'Das kleine Kind schläft.'","The small child is sleeping."),("'die neuen Schuhe' — why -en?","Plural Nominative after definite article takes -en")],
"Adjective endings after definite articles — done! This is a complex area but you've tackled it head-on. Keep practising with real sentences and it will become automatic!"
))

# ── Day 105 ─────────────────────────────────────────────────────────────────
DAYS.append((105,
"ADJECTIVE ENDINGS II 📝",
"Adjektivdeklination — After Indefinite Articles",
"Yesterday you learned adjective endings after definite articles (der/die/das). Today: adjective endings after INDEFINITE articles (ein/eine/ein). This is called the mixed declension. The key change: the adjective must show the gender marker when 'ein' doesn't!",
"काल definite articles पछी adjective endings शीख्यां. आज INDEFINITE articles (ein/eine/ein) पछी — mixed declension. Key change: जो 'ein' gender marker नहीं दर्शावे तो adjective दर्शाववूं पडे!",
[
("ein neuer Mantel","a new coat (Nom. m)","एक नवो coat","eyn NOY-er MAN-tel"),
("eine neue Jacke","a new jacket (Nom. f)","एक नवी jacket","EY-neh NOY-eh YAK-eh"),
("ein neues Hemd","a new shirt (Nom. n)","एक नवूं shirt","eyn NOY-es HEMD"),
("einen neuen Mantel (Acc. m)","a new coat (Acc.)","एक नवो coat (acc)","EY-nen NOY-en MAN-tel"),
("einer neuen Jacke (Dat. f)","a new jacket (Dat.)","एक नवी jacket (dat)","EY-ner NOY-en YAK-eh"),
("einem neuen Mantel (Dat. m)","a new coat (Dat.)","एक नवा coat (dat)","EY-nem NOY-en MAN-tel"),
("kein","no / not a","कोई नहीं","KYNE"),
("keine","no (f / pl)","कोई नहीं","KY-neh"),
("mein alter Freund","my old friend","मारो जूनो मित्र","myne AL-ter FROYNT"),
("meine liebe Mutter","my dear mother","मारी प्रिय माँ","MY-neh LEE-beh MOOT-er"),
("mein kleines Baby","my little baby","मारूं नानूं baby","myne KLY-nes BAY-bee"),
("-er ending","Nom. m (no -r in article)","Nom. m ending","-ehr"),
("-es ending","Nom./Acc. n (no -s in article)","Nom./Acc. n ending","-es"),
("-en ending","Acc. m, Dat. all, Gen. all","most oblique cases","-en"),
("-e ending","Nom./Acc. f, Nom./Acc. n-ish","Nom. f, Acc. f","-eh"),
],
[
("Ich habe einen neuen Job.","I have a new job.","मारे नवी job छे."),
("Das ist ein interessantes Buch.","That is an interesting book.","तेे एक interesting पुस्तक छे."),
("Er wohnt in einer kleinen Wohnung.","He lives in a small flat.","ते नानी apartment मां रहे छे."),
("Ich habe keine Zeit heute.","I have no time today.","आज मारी पासे समय नथी."),
("Mein bester Freund heißt Marco.","My best friend is called Marco.","मारा सौथी सारा मित्र नूं नाम Marco छे."),
],
[
("🔑","The Pattern Logic","'ein' has no ending in m/n Nom. and n Acc. In these slots the ADJECTIVE steps in: ein NEUER Mantel (m Nom.), ein NEUES Hemd (n Nom./Acc.)","E8F0FE"),
("💡","Possessives Same Pattern","mein, dein, sein, ihr, unser, euer, ihr, Ihr all follow the same pattern as 'ein'. Master 'ein' → master all possessives!","FFF8E1"),
("📌","Rule of Thumb","If 'ein' ends in a vowel (eine, einem, einer, einen) → adjective gets -en. If 'ein' has NO ending (ein...) → adjective shows the gender: -er(m), -es(n).","E8F0FE"),
],
["Make a table: adjective 'groß' (big) in all cases after indefinite article.","Write 5 sentences using 'ein/eine/ein' + adjective + noun.","Compare: der neue Mantel (def.) vs ein neuer Mantel (indef.) in all cases."],
[("'ein neuer Mann' — why -er?","Nominative masculine, 'ein' has no ending so adjective shows -er"),("Translate: 'Ich kaufe einen roten Apfel.'","I buy a red apple."),("'kein' is used for?","Expressing 'no' / 'not a' — same endings as 'ein'"),("Translate: 'Sie hat eine kleine Katze.'","She has a small cat."),("Adjective ending in 'einem' (dat. m/n)?","-en (einem neuen...)") ],
"Adjective endings after indefinite articles — complete! Combined with yesterday, you now have a full picture of German adjective declension. This is advanced grammar territory — well done!"
))

print("Days 101-105 appended!")

# ── Day 106 ─────────────────────────────────────────────────────────────────
DAYS.append((106,
"EXPRESSING OPINIONS 💭",
"Meinungen äußern — Speaking Your Mind",
"Being able to express opinions, agree, disagree, and give reasons is essential for conversations, discussions, and integration into German society. Today you learn opinion phrases at A2 level.",
"Opinions express करवी, agree/disagree करवूं, अने reasons आपवा conversation, discussions, अने German society integration माटे essential छे. आज A2 level ना opinion phrases शीखशो.",
[
("die Meinung","opinion","मत / राय","dee MY-nung"),
("Ich meine…","I think / In my opinion…","मारी राय मां...","ikh MY-neh"),
("Ich denke, dass…","I think that…","मारा मते...","ikh DENK-eh das"),
("Ich finde…","I find / I think…","मने लागे...","ikh FIN-deh"),
("Meiner Meinung nach…","In my opinion…","मारी राय प्रमाणे...","MY-ner MY-nung nakh"),
("Ich stimme zu.","I agree.","हूं सहमत छूं.","ikh SHTIM-eh TSOO"),
("Ich stimme nicht zu.","I disagree.","हूं असहमत छूं.","ikh SHTIM-eh nikht TSOO"),
("Das stimmt.","That's right / correct.","सही छे.","das SHTIMMT"),
("Das stimmt nicht.","That's not right.","सही नथी.","das SHTIMMT nikht"),
("Ich bin anderer Meinung.","I have a different opinion.","मारी अलग राय छे.","ikh bin AN-deh-rer MY-nung"),
("Einerseits…andererseits…","On the one hand…on the other hand…","एक तरफ...बीजी तरफ...","EY-ner-zyts...AN-deh-rer-zyts"),
("Das ist interessant.","That is interesting.","तेे रोचक छे.","das ist in-teh-res-ANT"),
("Ich bin der Meinung, dass…","I am of the opinion that…","मारी राय छे के...","ikh bin dehr MY-nung das"),
("Was denken Sie?","What do you think?","तमे शूं विचारो?","vas DENK-en zee"),
("Wie sehen Sie das?","How do you see this?","तमे आ केवूं जुओ?","vee ZAY-en zee das"),
],
[
("Ich finde, dass Deutsch sehr logisch ist.","I think that German is very logical.","मने लागे के German खूब logical छे."),
("Meiner Meinung nach ist Lesen sehr wichtig.","In my opinion, reading is very important.","मारी राय प्रमाणे वांचवूं खूब जरूरी छे."),
("Ich stimme nicht zu — das ist zu teuer.","I disagree — that is too expensive.","हूं असहमत छूं — तेे खूब महंगूं छे."),
("Einerseits mag ich die Stadt, andererseits vermisse ich die Natur.","On the one hand I like the city, on the other I miss nature.","एक तरफ शहेर गमे, बीजी तरफ कुदरत याद आवे."),
("Was denken Sie über diese Idee?","What do you think about this idea?","तमे आ idea बद्दल शूं विचारो?"),
],
[
("💡","Soft Opinion","Use 'Ich finde' or 'Ich meine' for personal opinions. They sound softer than 'Ich denke' and are very natural in conversation.","FFF8E1"),
("🇩🇪","German Discussion Style","Germans appreciate directness and logical reasoning. Don't be afraid to share your opinion clearly. Use 'weil' + reason to back up your views!","E6F4EA"),
("📌","Subordinate After Opinion","'Ich denke, dass + VERB LAST.' 'Ich finde, dass Deutsch SCHWER IST.' Remember: 'dass' triggers verb-to-end rule!","E8F0FE"),
],
["Write your opinion on 5 different topics in German using the phrases above.","Write a mini-debate: give pro and contra for living in Germany.","Practise agreeing and disagreeing with 5 statements about Germany."],
[("Translate: 'Meiner Meinung nach ist das falsch.'","In my opinion, that is wrong."),("'Ich stimme zu' means?","I agree."),("Translate: 'Was denkst du darüber?'","What do you think about it?"),("After 'dass', where does the verb go?","To the end"),("'einerseits...andererseits' means?","On the one hand...on the other hand")],
"Opinion vocabulary mastered! You can now participate in German discussions, express your views, agree and disagree politely — the hallmark of real A2-level communication!"
))

# ── Day 107 ─────────────────────────────────────────────────────────────────
DAYS.append((107,
"DESCRIBING PEOPLE 👤",
"Personen beschreiben — Appearance & Character",
"Describing people's appearance and character is useful for countless situations — identifying someone, writing a description, meeting new people, or even writing a dating profile! Today you learn appearance and personality vocabulary.",
"लोकोना appearance अने character describe करवूं अनेक situations मां useful छे — कोईने ओळखाववां, description लखवी, नवा लोको मळवा, या dating profile! आज appearance अने personality vocabulary शीखशो.",
[
("das Aussehen","appearance","देखाव","das OWS-zay-en"),
("der Charakter","character/personality","સ्वभाव","dehr ka-RAK-ter"),
("groß","tall","ऊंचूं","GROHS"),
("klein","short","नानूं","KLYNE"),
("schlank","slim","पातळूं","SHLANK"),
("kräftig","strong / muscular","तगडूं","KREF-tikh"),
("das Haar","hair","वाळ","das HAHR"),
("blond","blonde","सुनेरी वाळ","BLONT"),
("braun","brown (hair/eyes)","ভूरूं","BROWN"),
("die Augenfarbe","eye colour","आंखोनो रंग","dee OW-gen-far-beh"),
("freundlich","friendly","मिळनसार","FROYNT-likh"),
("lustig","funny","मजाकी","LOOS-tikh"),
("ruhig","calm","शांत","ROO-ikh"),
("fleißig","hardworking","मेhनती","FLY-sikh"),
("neugierig","curious","जिज्ञासु","noy-GHEE-rikh"),
],
[
("Er ist groß und schlank mit braunen Haaren.","He is tall and slim with brown hair.","ते ऊंचो अने पातळो छे अने ભूरा वाळ छे."),
("Sie hat blaue Augen und ist sehr freundlich.","She has blue eyes and is very friendly.","तेने वादळी आंखें छे अने खूब मिळनसार छे."),
("Mein bester Freund ist immer lustig und ruhig.","My best friend is always funny and calm.","मारो सौथी सारो मित्र हंमेशां मजाकी अने शांत छे."),
("Wie siehst du aus?","What do you look like?","तूं केवो देखाय?"),
("Er hat kurze, lockige Haare.","He has short, curly hair.","तेने ट़ूंका, कुंडाळां वाळ छे."),
],
[
("💡","Adjective Endings","When describing people, adjective endings apply! 'Er ist ein FREUNDLICHER Mann.' (mixed decl.) vs 'Der freundliche Mann' (weak decl.)","FFF8E1"),
("🇩🇪","Description Phrases","'Er/Sie sieht gut aus.' (He/She looks good.) 'Er/Sie ist ungefähr 30 Jahre alt.' (He/She is about 30 years old.)","E6F4EA"),
("📌","Hair + Colour","Adjective + Haare: 'Er hat blonde/braune/schwarze/rote/graue Haare.' Augen: 'Sie hat blaue/grüne/braune Augen.'","E8F0FE"),
],
["Describe yourself in German (5-8 sentences, appearance + character).","Write a description of a famous person without naming them — see if a friend can guess!","Write descriptions of 3 people you know well in German."],
[("Translate: 'Sie ist sehr freundlich.'","She is very friendly."),("'fleißig' means?","Hardworking"),("How do you say 'He has blue eyes'?","Er hat blaue Augen."),("'neugierig' means?","Curious"),("Translate: 'Er ist groß und hat schwarze Haare.'","He is tall and has black hair.")],
"Description vocabulary complete! You can now paint a vivid picture of any person in German — their looks, personality, and character. A truly expressive achievement!"
))

# ── Day 108 ─────────────────────────────────────────────────────────────────
DAYS.append((108,
"REFLEXIVE VERBS ��",
"Reflexive Verben — Actions on Yourself",
"Reflexive verbs describe actions you do TO yourself. They always come with a reflexive pronoun (mich, dich, sich…). Many daily activities are reflexive in German: getting dressed, washing, feeling, remembering.",
"Reflexive verbs ते action describe करे जे तमे खुद पर करो. तेमनी सात reflexive pronoun (mich, dich, sich...) आवे. German मां ઘणी daily activities reflexive छे: कपडा पहेरवां, नाहवूं, feel करवूं, याद करवूं.",
[
("sich freuen","to be happy / look forward","खुश थवूं","zikh FROY-en"),
("sich ärgern","to be annoyed","ग़़ुस्सो आवोन","zikh ÄR-gern"),
("sich fühlen","to feel","feel करवूं","zikh FÜ-len"),
("sich erinnern","to remember","याद करवूं","zikh ehr-IN-ern"),
("sich interessieren","to be interested","रस होवो","zikh in-teh-res-EE-ren"),
("sich entschuldigen","to apologise","माफी मागवी","zikh ent-SHOOL-di-gen"),
("sich beeilen","to hurry","ઉतावळ करवी","zikh beh-EYE-len"),
("sich vorstellen","to imagine / introduce","कल्पना करवी / परिचय","zikh FOR-shtel-en"),
("sich setzen","to sit down","बेसवूं","zikh ZET-sen"),
("sich legen","to lie down","सूई जवूं","zikh LAY-gen"),
("mich","myself (Acc. ich)","मने","MIKH"),
("dich","yourself (Acc. du)","तने","DIKH"),
("sich","himself/herself (Acc.)","तेने/तेणे","ZIKH"),
("uns","ourselves (Acc. wir)","आपणने","OONS"),
("euch","yourselves (Acc. ihr)","तमने","OYKH"),
],
[
("Ich freue mich auf den Urlaub.","I am looking forward to the holiday.","मने vacation नी राह छे."),
("Er fühlt sich heute nicht gut.","He doesn't feel well today.","ते आज ठीक feel नथी करतो."),
("Beeile dich! Wir sind spät.","Hurry up! We are late.","ઉतावળ कर! आपणे મोडા छीए."),
("Ich erinnere mich an meinen ersten Schultag.","I remember my first day at school.","मने मारो पहेलो school ना दिवस याद छे."),
("Entschuldigen Sie sich bitte.","Please apologise.","कृपा माफी मागो."),
],
[
("📌","Reflexive Pronoun","Reflexive pronouns: mich (ich), dich (du), sich (er/sie/es/sie/Sie), uns (wir), euch (ihr). Note: 'sich' covers 3rd person AND formal Sie!","E8F0FE"),
("💡","'sich freuen auf'","'sich freuen auf + Accusative' = to look forward to something. 'Ich freue mich auf Weihnachten.' Very natural and common!","FFF8E1"),
("⚠️","Dative Reflexive","Some reflexive verbs take Dative: 'Ich wasche mir die Hände.' (I wash my hands.) Here 'mir' is Dative reflexive. Check each verb!","FFE0E0"),
],
["Conjugate 5 reflexive verbs for all 6 persons.","Write a daily routine using 5 reflexive verbs.","Write 5 sentences using 'sich freuen auf', 'sich fühlen', 'sich erinnern an'."],
[("Reflexive pronoun for 'wir'?","uns"),("Translate: 'Ich fühle mich gut.'","I feel well."),("'sich beeilen' means?","To hurry"),("Translate: 'Er entschuldigt sich.'","He apologises."),("'sich freuen auf' takes which case?","Accusative")],
"Reflexive verbs mastered! Daily actions, emotions, and social situations are now fully expressible in German. Your conversations are becoming genuinely natural!"
))

# ── Day 109 ─────────────────────────────────────────────────────────────────
DAYS.append((109,
"ENVIRONMENT & TECHNOLOGY 🌍💻",
"Umwelt und Technologie — Modern German Topics",
"Germany is a world leader in environmental technology (Grüne Energie) and digital innovation. Today you learn vocabulary for environmental issues and technology — topics you'll encounter in modern German conversations.",
"Germany environmental technology (Green energy) अने digital innovation मां world leader छे. आज environment अने technology vocabulary शीखशो — modern German conversations मां आ topics common छे.",
[
("die Umwelt","the environment","पर्यावरण","dee OOM-velt"),
("der Klimawandel","climate change","जळवायु परिवर्तन","dehr KLEE-mah-van-del"),
("die Energie","energy","ऊर्जा","dee eh-nehr-GHEE"),
("die Solarenergie","solar energy","सौर ऊर्जा","dee zo-LAR-eh-nehr-ghee"),
("recyceln","to recycle","recycle करवूं","reh-SY-keln"),
("der Müll","rubbish / waste","कचरो","dehr MÜL"),
("die Mülltrennung","waste separation","कचराना प्रकार अलग करवा","dee MÜL-tren-nung"),
("das Fahrrad","bicycle","साइकल","das FAR-raht"),
("nachhaltig","sustainable","टकाऊ","NAKH-hal-tikh"),
("die Technik","technology / engineering","technology","dee TEKH-nik"),
("das Internet","the internet","internet","das IN-ter-net"),
("die App","the app","app","dee EP"),
("das Passwort","the password","password","das PAS-vort"),
("herunterladen","to download (sep.)","download करवूं","heh-ROON-ter-lah-den"),
("hochladen","to upload (sep.)","upload करवूं","HOKH-lah-den"),
],
[
("Deutschland ist führend in der Solarenergie.","Germany is a leader in solar energy.","Germany solar energy मां अग्रणी छे."),
("Ich trenne meinen Müll sorgfältig.","I carefully separate my rubbish.","हूं ध्यानपूर्वक कचरो अलग करूं छूं."),
("Wir müssen nachhaltiger leben.","We must live more sustainably.","आपणे वधु sustainable रीते जीववूं जोईए."),
("Kannst du mir diese App empfehlen?","Can you recommend this app to me?","शूं तूं मने आ app BI suggest करशे?"),
("Ich lade das Dokument jetzt herunter.","I am downloading the document now.","हूं अत्यारे document download करूं छूं."),
],
[
("🇩🇪","Mülltrennung","Germany is world-famous for waste separation! Separate bins: Gelber Sack (yellow = packaging), Papiertonne (paper), Biotonne (organic), Restmüll (general waste), Glascontainer (glass).","E6F4EA"),
("💡","Green Germany","Germany aims for carbon neutrality by 2045. Terms like Energiewende (energy transition), Windkraft (wind power), and Elektroauto are in everyday German news.","FFF8E1"),
("📌","Technology Verbs","Most technology verbs in German are borrowed from English but conjugated German-style: downloaden→herunterladen, uploaden→hochladen, googeln→googeln.","E8F0FE"),
],
["Write 5 things you personally do for the environment in German.","Describe your smartphone/computer use habits in German.","Write a short paragraph about climate change in German using the vocabulary."],
[("'der Klimawandel' means?","Climate change"),("'recyceln' means?","To recycle"),("What is 'Mülltrennung'?","Waste separation"),("Translate: 'Ich lade das Video herunter.'","I am downloading the video."),("'nachhaltig' means?","Sustainable")],
"Environment and technology vocabulary complete! You can now discuss modern German topics — from green energy to social media — with confidence and depth!"
))

# ── Day 110 ─────────────────────────────────────────────────────────────────
DAYS.append((110,
"RELATIVE CLAUSES — INTRODUCTION 📎",
"Relativsätze — Adding Detail to Sentences",
"Relative clauses add detail to nouns: 'the man WHO lives next door', 'the book THAT I read'. In German, relative pronouns (der, die, das, die) match the GENDER of the noun they refer to, and the verb goes to the END of the relative clause.",
"Relative clauses nouns ने detail ઉमेरे: 'ते माणस जे बाजु मां रहे'. German मां relative pronouns (der, die, das, die) referent noun ना GENDER सात match थाय, अने verb relative clause ना END मां जाय.",
[
("das Relativpronomen","relative pronoun","relative pronoun","das reh-lah-TEEF-proh-noh-men"),
("der (rel. m)","who/that/which (m)","जे (m)","dehr"),
("die (rel. f)","who/that/which (f)","जे (f)","dee"),
("das (rel. n)","who/that/which (n)","जे (n)","das"),
("die (rel. pl)","who/that/which (pl)","जे (pl)","dee"),
("dem (rel. Dat. m/n)","whom (Dat. m/n)","जेने (Dat.)","daym"),
("der (rel. Dat. f)","whom (Dat. f)","जेने (Dat. f)","dehr"),
("den (rel. Acc. m)","whom (Acc. m)","जेने (Acc.)","dayn"),
("Der Mann, der…","The man who…","ते माणस जे...","dehr MAN, dehr"),
("Die Frau, die…","The woman who…","ते स्त्री जे...","dee FROW, dee"),
("Das Buch, das…","The book that…","ते पुस्तक जे...","das BOOKH, das"),
("..., der … ist","…who is…","...जे ...छे","...dehr...ist"),
("..., die … hat","…who has…","...जेनी पासे ...छे","...dee...hat"),
("Ich kenne jemanden, der…","I know someone who…","मने कोई ओळखाय छे जे...","ikh KEN-eh YEH-man-den dehr"),
("Das ist das Haus, das…","That is the house that…","तेे घर छे जे...","das ist das HOWS das"),
],
[
("Der Mann, der dort steht, ist mein Bruder.","The man who stands there is my brother.","ते माणस जे त्यां ऊभो छे, ते मारो भाई छे."),
("Das Buch, das ich lese, ist sehr interessant.","The book that I am reading is very interesting.","ते पुस्तक जे हूं वांचूं छूं, खूब रोचक छे."),
("Die Frau, die ich kenne, wohnt in Berlin.","The woman whom I know lives in Berlin.","ते स्त्री जेने हूं ओळखूं, Berlin मां रहे छे."),
("Das ist das Hotel, in dem wir gewohnt haben.","That is the hotel in which we stayed.","तेे hotel छे जेमां आपणे रह्यां."),
("Ich suche jemanden, der Deutsch spricht.","I am looking for someone who speaks German.","हूं कोईने शोधूं छूं जे German बोले."),
],
[
("🔑","Relative Pronoun = Article","Relative pronouns look almost exactly like definite articles! Only Dat. pl (denen), Gen. m/n (dessen), Gen. f/pl (deren) differ significantly.","E8F0FE"),
("📌","Comma Rule","ALWAYS put a comma before the relative clause: 'Der Mann[,] der dort steht...' The comma is mandatory in German writing!","FFE0E0"),
("💡","Verb to End","Inside the relative clause, verb always goes last: 'das Buch, das ich LESE' — 'lese' goes to the end. Same rule as subordinate clauses!","FFF8E1"),
],
["Write 5 relative clause sentences about people you know.","Expand 5 simple nouns with relative clauses (das Buch, der Lehrer, die Stadt…).","Translate 5 English relative clause sentences into German."],
[("What determines the gender of a relative pronoun?","The gender of the noun it refers to"),("Translate: 'Die Stadt, in der ich wohne, ist schön.'","The city in which I live is beautiful."),("Where does the verb go in a relative clause?","To the end"),("Which punctuation is required before a relative clause?","A comma"),("'Das Buch, das ich lese' — why 'das'?","Because 'Buch' is neuter (das Buch)")],
"Relative clauses — introduced and understood! This is B1-level territory that you've now cracked at A2. Your German sentences are becoming truly sophisticated. Excellent work!"
))

print("Days 106-110 appended!")

# ── Day 111 ─────────────────────────────────────────────────────────────────
DAYS.append((111,
"AT THE HOTEL 🏨",
"Im Hotel — Check-In to Check-Out",
"Hotel stays are an important practical situation. Today you learn comprehensive hotel vocabulary — check-in, room requests, problems, amenities, and check-out. Perfect for travel anywhere in the German-speaking world.",
"Hotel stays important practical situation छे. आज comprehensive hotel vocabulary — check-in, room requests, problems, amenities, अने check-out शीखशो.",
[
("einchecken","to check in (sep.)","check-in करवूं","EYN-chek-en"),
("auschecken","to check out (sep.)","check-out करवूं","OWS-chek-en"),
("die Rezeption","the reception","reception","dee reh-tsep-TSIOHN"),
("der Empfang","the reception desk","स्वागत काउन्टर","dehr em-PFANG"),
("das Einzelzimmer","single room","single room","das EYN-tsel-tsim-er"),
("das Doppelzimmer","double room","double room","das DOP-el-tsim-er"),
("das Frühstück ist inbegriffen","breakfast included","breakfast included","...ist IN-beh-grif-en"),
("der Aufzug / der Fahrstuhl","the lift / elevator","lift","dehr OWF-tsoog / FAR-shtool"),
("das WLAN","WiFi","WiFi","das VAY-lan"),
("der Safe","the safe","safe","dehr SAYF"),
("der Schlüssel","the key","चाबी","dehr SHLÜS-el"),
("das Handtuch","the towel","ટ‍ওয়েল","das HANT-tookh"),
("der Zimmerdienst","room service","room service","dehr TSIM-er-deenst"),
("die Minibar","the minibar","minibar","dee MEE-nee-bar"),
("abreisen","to depart (hotel) (sep.)","hotel छोडवूं","AB-ry-zen"),
],
[
("Ich habe eine Reservierung auf den Namen Patel.","I have a reservation in the name of Patel.","मारे Patel नाम पर reservation छे."),
("Könnte ich bitte ein ruhigeres Zimmer haben?","Could I please have a quieter room?","शूं मने please शांत ओरडो मळशे?"),
("Wann ist der Check-out?","When is check-out?","check-out क्यारे छे?"),
("Das Zimmer ist nicht sauber.","The room is not clean.","ओरडो साफ नथी."),
("Bitte wecken Sie mich um 7 Uhr.","Please wake me at 7 o'clock.","कृपा मने 7 वागे जगाडो."),
],
[
("🇩🇪","Hotel Classification","Germany uses 1-5 star (Sterne) classification. Even budget hotels (Pension/Hostel) are usually very clean. Always check: 'Ist das Frühstück inbegriffen?'","E6F4EA"),
("💡","Politeness Level","Use 'Könnten Sie…?' (Could you…?) for polite hotel requests — more polite than 'Können Sie'. Perfect for hotel situations!","FFF8E1"),
("📌","Common Problems","'Die Heizung funktioniert nicht.' (The heating doesn't work.) 'Es gibt kein warmes Wasser.' (There's no hot water.) 'Das Fenster lässt sich nicht öffnen.' (The window won't open.)","E8F0FE"),
],
["Write a complete hotel check-in conversation in German.","Write an email complaining about a hotel room problem in German.","Write 5 requests you might make to hotel reception."],
[("Translate: 'Ich möchte einchecken.'","I would like to check in."),("'das Einzelzimmer' means?","Single room"),("How do you ask 'Is breakfast included?'","Ist das Frühstück inbegriffen?"),("'der Schlüssel' means?","The key"),("Translate: 'Das Zimmer ist zu laut.'","The room is too loud.")],
"Hotel vocabulary complete! You can now navigate any German-speaking hotel from check-in to check-out with confidence and politeness!"
))

# ── Day 112 ─────────────────────────────────────────────────────────────────
DAYS.append((112,
"AT THE BANK & POST OFFICE 🏦",
"Bank und Post — Essential Services",
"Banking and postal services require specific vocabulary. Today you learn how to open an account, send packages, understand bills, and use German administrative services — essential for living in Germany.",
"Banking अने postal services specific vocabulary मांगे. आज account खोलवूं, packages मोकलवां, bills समजवां, अने German administrative services वापरवी शीखशो.",
[
("die Bank","the bank","bank","dee BANK"),
("das Konto","the (bank) account","account","das KON-toh"),
("ein Konto eröffnen","to open an account","account खोलवूं","eyn KON-toh ehr-ÖF-nen"),
("überweisen","to transfer (money)","transfer करवूं","Ü-ber-vye-zen"),
("die Überweisung","bank transfer","bank transfer","dee Ü-ber-vye-zung"),
("abheben","to withdraw (money) (sep.)","पैसा उपाडवा","AB-hay-ben"),
("einzahlen","to deposit (sep.)","पैसा जमा करवा","EYN-tsah-len"),
("der Kontostand","account balance","account balance","dehr KON-toh-shtant"),
("die Post","post office / mail","post office","dee POST"),
("das Paket","the parcel","parcel","das pa-KAY T"),
("der Brief","the letter","पत्र","dehr BREEF"),
("die Briefmarke","the stamp","stamp (postal)","dee BREEF-mar-keh"),
("einschreiben","registered post","registered mail","EYN-shry-ben"),
("die Adresse","the address","address","dee ah-DRES-eh"),
("der Absender","the sender","मोकलनार","dehr AP-zen-der"),
],
[
("Ich möchte ein Konto bei Ihrer Bank eröffnen.","I would like to open an account at your bank.","मारे तमारी bank मां account खोलवूं छे."),
("Ich möchte 200 Euro auf mein Konto einzahlen.","I would like to deposit 200 euros into my account.","मारे मारा account मां 200 euro जमा करवा छे."),
("Ich möchte dieses Paket nach Indien schicken.","I would like to send this parcel to India.","मारे आ parcel India मोकलवूं छे."),
("Was kostet es, diesen Brief zu versenden?","How much does it cost to send this letter?","आ पत्ر મોকalvāno केटलो खर्ч छे?"),
("Mein Kontostand beträgt 1500 Euro.","My account balance is 1500 euros.","मारो account balance 1500 euro छे."),
],
[
("🇩🇪","German Banking","Sparkasse and Volksbank are local German banks. Deutsche Bank and Commerzbank are national. Opening an account requires: passport, Anmeldung (registration), and TIN number.","E6F4EA"),
("💡","Anmeldung","Before banking, register your address at the Einwohnermeldeamt (registration office). The Anmeldebestätigung (confirmation) is needed for almost everything in Germany!","FFF8E1"),
("📌","Post Tracking","DHL is Germany's main postal service. 'Einschreiben' (registered) gives you a tracking number. 'Expresszustellung' = express delivery.","E8F0FE"),
],
["Write an email asking a German bank about opening an account.","Write a postal form for sending a package to your home country.","Write 5 banking transactions in German sentences."],
[("'überweisen' means?","To transfer money"),("How do you say 'stamp' (postal)?","die Briefmarke"),("Translate: 'Ich möchte Geld abheben.'","I would like to withdraw money."),("What is needed to open a German bank account?","Passport and Anmeldung (address registration)"),("'das Paket' means?","The parcel")],
"Banking and postal vocabulary mastered! These are essential services you'll use from day one in Germany. You can now manage your money and send packages in German!"
))

# ── Day 113 ─────────────────────────────────────────────────────────────────
DAYS.append((113,
"COORDINATING CONJUNCTIONS 🔗",
"Nebenordnende Konjunktionen: und, aber, oder, denn, sondern",
"Coordinating conjunctions join two main clauses with NORMAL word order (no verb-to-end rule!). The five most important are: und (and), aber (but), oder (or), denn (because), sondern (but rather). These are essential for building natural German sentences.",
"Coordinating conjunctions बे main clauses ने NORMAL word order सात join करे (verb-to-end rule नहीं!). पांच सौथी important: und (अने), aber (पण), oder (या), denn (कारण), sondern (पण बलके). Natural German sentences माटे essential.",
[
("und","and","अने","OONT"),
("aber","but","पण","AH-ber"),
("oder","or","या","OH-der"),
("denn","because (coord.)","कारण (main clause)","DEN"),
("sondern","but rather (after negative)","पण बलके","ZON-dern"),
("sowohl … als auch","both … and","...पण ...पण","zo-VOHL...als OWKH"),
("weder … noch","neither … nor","ना...ना","VAY-der...nokh"),
("entweder … oder","either … or","ક્યां...ک্যां","ent-VAY-der...OH-der"),
("nicht … sondern","not … but rather","...नहीं ... पण","nikht...ZON-dern"),
("deshalb","therefore","इसलिए","DES-halb"),
("trotzdem","nevertheless","तेम छतां","TROTS-daym"),
("außerdem","moreover / in addition","उपरांत","OWS-er-daym"),
("deswegen","that's why / therefore","तेथी","DES-vay-gen"),
("zwar … aber","admittedly … but","जरूर...पण","TSVAHR...AH-ber"),
("jedoch","however (formal)","परंतु","yeh-DOHKH"),
],
[
("Ich lerne Deutsch und ich übe täglich.","I learn German and I practise daily.","हूं German शीखूं छूं अने रोज practice करूं छूं."),
("Ich mag Kaffee, aber ich trinke lieber Tee.","I like coffee, but I prefer to drink tea.","मने coffee गमे, पण चाय वधु पीवी गमे."),
("Komm bitte pünktlich, denn ich habe wenig Zeit.","Please come on time, because I have little time.","Please समय पर आव्यो, कारण मारे ओछो समय छे."),
("Das ist kein Kaffee, sondern Tee.","That is not coffee, but rather tea.","तेे coffee नहीं, पण चाय छे."),
("Entweder du übst, oder du wirst es nicht lernen.","Either you practise, or you won't learn it.","ক্যां तूं practice कर, या नहीं शीखशे."),
],
[
("🔑","vs Subordinating","Coordinating conjunctions (und, aber, oder, denn, sondern) = NORMAL word order. Subordinating conjunctions (weil, dass, ob) = VERB TO END. Crucial difference!","E8F0FE"),
("⚠️","'denn' vs 'weil'","Both = 'because'. 'denn' (coord.) → normal order: 'Ich komme, denn ich will.' 'weil' (sub.) → verb last: 'Ich komme, weil ich WILL.'","FFE0E0"),
("💡","'sondern' Rule","'sondern' is only used after a negative statement: 'Nicht A, sondern B' = Not A, but B. Example: 'Ich trinke nicht Kaffee, sondern Tee.'","FFF8E1"),
],
["Write 5 sentences using each of the 5 main conjunctions (25 total).","Write a paragraph with 3 contrasting ideas using aber/sondern/jedoch.","Create 5 either/or sentences about daily choices using 'entweder…oder'."],
[("What is the word order rule after coordinating conjunctions?","Normal word order (no change)"),("Translate: 'Ich bin müde, aber ich lerne weiter.'","I am tired, but I continue learning."),("When is 'sondern' used?","After a negative statement (not A, but rather B)"),("Difference between 'denn' and 'weil'?","denn = normal order; weil = verb to end"),("Translate: 'Ich lerne Deutsch, deshalb werde ich nach Deutschland ziehen.'","I learn German, therefore I will move to Germany.")],
"Coordinating conjunctions mastered! Combined with subordinating conjunctions from Day 103, you can now build truly complex and natural German sentences. Impressive progress!"
))

# ── Day 114 ─────────────────────────────────────────────────────────────────
DAYS.append((114,
"GERMAN LETTER & EMAIL WRITING ✉️",
"Briefe und E-Mails — Formal and Informal Writing",
"Written communication is crucial for job applications, official letters, and everyday correspondence in Germany. Today you learn the structure of formal and informal letters/emails in German.",
"Written communication job applications, official letters, अने everyday correspondence माटे crucial छे. आज formal अने informal letters/emails नी structure शीखशो.",
[
("der Brief","the letter","पत्र","dehr BREEF"),
("die E-Mail","the email","email","dee EE-mail"),
("formell","formal","औपचारिक","for-MEL"),
("informell","informal","अनौपचारिक","in-for-MEL"),
("Sehr geehrte Damen und Herren","Dear Sir/Madam (formal opening)","आदरणीय (formal opening)","zayr geh-EHR-teh DAH-men oont HEH-ren"),
("Sehr geehrter Herr…","Dear Mr… (formal)","आदरणीय श्री...","zayr geh-EHR-ter hehr"),
("Sehr geehrte Frau…","Dear Ms… (formal)","आदरणीय श्रीमती...","zayr geh-EHR-teh frow"),
("Lieber…/ Liebe…","Dear… (informal m/f)","प्रिय... (informal)","LEE-ber / LEE-beh"),
("Mit freundlichen Grüßen","Yours sincerely (formal closing)","आदरसहित (formal closing)","mit FROYNT-likh-en GRÜ-sen"),
("Viele Grüße","Many greetings (informal closing)","ઘणा শुভेच्छा (informal)","FEE-leh GRÜ-seh"),
("ich schreibe Ihnen bezüglich","I am writing to you regarding","...बाबत तमने लखूं छूं","ikh SHRY-beh EE-nen beh-TSÜ G-likh"),
("im Anhang senden","to send as attachment","attachment मां मोकलvun","im AN-hang ZEN-den"),
("Ich freue mich auf Ihre Antwort.","I look forward to your reply.","तमारा जवाब नी राह छे.","ikh FROY-eh mikh owf EE-reh ANT-vort"),
("Bitte teilen Sie mir mit…","Please let me know…","कृपा मने जणावो...","BIT-eh TY-len zee meer mit"),
("hochachtungsvoll","yours faithfully (very formal)","अत्यंत आदरसहित","hok-AKH-tungs-fol"),
],
[
("Sehr geehrte Frau Müller, ich schreibe Ihnen bezüglich meiner Bewerbung.","Dear Ms Müller, I am writing to you regarding my application.","आदरणीय Frau Müller, हूं मारी application बाबत लखूं छूं."),
("Liebe Sarah, ich hoffe, es geht dir gut!","Dear Sarah, I hope you are well!","प्रिय Sarah, आशा छे तूं ठीक छे!"),
("Im Anhang sende ich Ihnen meinen Lebenslauf.","In the attachment I am sending you my CV.","attachment मां हूं तमने मारो CV मोकलूं छूं."),
("Ich freue mich auf Ihre Rückmeldung.","I look forward to your feedback.","तमारा feedback नी राह छे."),
("Mit freundlichen Grüßen, Aniket Patel","Yours sincerely, Aniket Patel","आदरसहित, Aniket Patel"),
],
[
("📌","Letter Structure","Opening greeting → purpose paragraph → details → closing request → Closing formula → Signature. This structure is standard in German formal letters!","E8F0FE"),
("🇩🇪","German Formality","Germany is more formal in writing than many cultures. Always use 'Sie' (formal) with people you don't know well. First-name basis only when invited!","E6F4EA"),
("💡","Email Subject Line","Always write a clear Betreff (subject line): 'Betreff: Bewerbung als Softwareentwickler' or 'Betreff: Frage zu meiner Bestellung'.","FFF8E1"),
],
["Write a formal email applying for a German language course.","Write an informal email to a German friend about your weekend.","Write both formal and informal versions of the same message (a request for information)."],
[("Formal letter opening to unknown recipient?","Sehr geehrte Damen und Herren"),("Informal letter opening to a male friend?","Lieber [Name]"),("Formal closing?","Mit freundlichen Grüßen"),("Informal closing?","Viele Grüße"),("Translate: 'Ich freue mich auf Ihre Antwort.'","I look forward to your reply.")],
"Letter and email writing vocabulary complete! Professional and personal written communication in German is now within your reach. You can apply for jobs, write to authorities, and stay in touch with German friends!"
))

# ── Day 115 ─────────────────────────────────────────────────────────────────
DAYS.append((115,
"GENITIVE CASE — INTRODUCTION 🔠",
"Der Genitiv — Showing Possession",
"German has four cases. You know Nominative, Accusative, and Dative. The fourth is Genitive — used to show possession. In everyday spoken German it is often replaced by 'von + Dative', but you must recognise it in written German.",
"German ना 4 cases. Nominative, Accusative, अने Dative जाणो. चोथो Genitive — possession दर्शावे. Spoken German मां घणे वखत 'von + Dative' वपराय, पण written German मां Genitive ओळखवो जरूरी छे.",
[
("der Genitiv","genitive case","ચоth case","dehr geh-neh-TEEF"),
("des Mannes","of the man (m/n Gen.)","माणस नो/नी/नूं","des MAN-es"),
("der Frau","of the woman (f Gen.)","स्त्री नो/नी/नूं","dehr FROW"),
("des Kindes","of the child (n Gen.)","बाळक नो/नी/नूं","des KIN-des"),
("der Kinder","of the children (pl Gen.)","बाळको नो/नी/नूं","dehr KIN-der"),
("Das Auto des Mannes","the man's car","माणस नी कार","das OW-toh des MAN-es"),
("Die Tasche der Frau","the woman's bag","स्त्री नी थेली","dee TA-sheh dehr FROW"),
("Das Spielzeug des Kindes","the child's toy","बाळक नूं रमकडूं","das SHPEEL-tsoyg des KIN-des"),
("wegen (+ Genitive)","because of","ना कारणे","VAY-gen"),
("während (+ Genitive)","during","दरम्यान","VÄH-rent"),
("trotz (+ Genitive)","despite","छतां","TROTS"),
("statt / anstatt","instead of","ने बदले","SHTAT / AN-shtat"),
("außerhalb","outside of","ना बहार","OWS-er-halb"),
("innerhalb","within / inside of","ना अंदर","IN-er-halb"),
("von + Dative (colloquial)","of (spoken replacement)","नो/नी/नूं (spoken)","FON..."),
],
[
("Das ist das Auto meines Vaters.","That is my father's car.","तेे मारा पिता नी car छे."),
("Wegen des Regens blieb ich zu Hause.","Because of the rain I stayed home.","वरसाद ना कारणे हूं घेर रह्यो/रही."),
("Während des Sommers fahren wir ans Meer.","During the summer we go to the sea.","ऊनाळा दरम्यान आपणे समुद्र किनारे जईए."),
("Trotz des schlechten Wetters gingen wir spazieren.","Despite the bad weather we went for a walk.","ख़राब हवामान छतां आपणे ফালवा गया."),
("Das Haus meiner Schwester ist sehr groß.","My sister's house is very big.","मारी बहेन नूं घर खूब मोटूं छे."),
],
[
("📌","Genitive Articles","m/n: des (+ noun gets -s/-es). f: der. pl: der. So: des Mannes, der Frau, des Kindes, der Kinder.","E8F0FE"),
("💡","Colloquial Alternative","In spoken German: 'das Auto von meinem Vater' instead of 'das Auto meines Vaters'. Both correct — written uses Genitive, spoken uses 'von + Dative'.","FFF8E1"),
("🇩🇪","Genitive Prepositions","wegen, trotz, während, statt, außerhalb, innerhalb all take Genitive. In spoken German 'wegen mir' (Dative) is becoming common but technically incorrect.","FFE0E0"),
],
["Write 5 Genitive possession sentences (Das Buch..., Die Tasche...).","Use 3 Genitive prepositions in sentences (wegen, trotz, während).","Find 5 examples of Genitive in German texts (news articles, signs)."],
[("Genitive article for masculine nouns?","des (e.g., des Mannes)"),("Translate: 'Das ist das Haus meiner Mutter.'","That is my mother's house."),("'trotz' takes which case?","Genitive"),("Colloquial alternative for Genitive?","von + Dative"),("Translate: 'wegen des Wetters'","because of the weather")],
"Genitive case introduced! You can now understand all four German cases. You're reading authentic German texts, understanding signs, and following complex grammar. Almost at the A2 finish line!"
))

print("Days 111-115 appended!")

# ── Day 116 ─────────────────────────────────────────────────────────────────
DAYS.append((116,
"LOCATION VERBS & PLACEMENT 📦",
"Stehen, Liegen, Hängen vs Stellen, Legen, Hängen",
"German uses different verbs for WHERE something IS (location) and WHERE you PUT something (direction). This pairs perfectly with the two-way prepositions from Day 66. Master these six verbs and your German becomes precise and native-sounding.",
"German मां WHERE something IS (location) अने WHERE you PUT it (direction) माटे अलग verbs छे. Day 66 ना two-way prepositions सात perfectly pair थाय. आ 6 verbs master करो अने German precise बने.",
[
("stehen","to stand / to be standing","ઊभूं छे","SHTAY-en"),
("stellen","to place upright (dir.)","ઊभूं कर","SHTEL-en"),
("liegen","to lie / to be lying","सूतूं/पड्यूं छे","LEE-gen"),
("legen","to lay / to place flat","सपाट मूकवूं","LAY-gen"),
("hängen","to hang (location)","टàng्यूं छे","HENG-en"),
("hängen","to hang (direction, sep)","टांगवूं","HENG-en"),
("sitzen","to sit / to be sitting","बेठूं छे","ZIT-sen"),
("setzen","to seat / to place (dir.)","बेसाडवूं","ZET-sen"),
("stecken","to stick / be inserted","ठोकेलूं छे","SHTEK-en"),
("stecken","to put / to stick in","ठोकेलवूं","SHTEK-en"),
("Das Buch steht im Regal.","The book stands in the shelf.","पुस्तक shelf मां ऊभूं छे.","..."),
("Ich stelle das Buch ins Regal.","I put the book in the shelf.","हूं पुस्तक shelf मां मूकूं छूं.","..."),
("Das Buch liegt auf dem Tisch.","The book lies on the table.","पुस्तक ट़ेबल पर पड्यूं छे.","..."),
("Ich lege das Buch auf den Tisch.","I lay the book on the table.","हूं पुस्तक टेबल पर मूकूं छूं.","..."),
("Das Bild hängt an der Wand.","The picture hangs on the wall.","ছবि दीवाल पर टàng्यूं छे.","..."),
],
[
("Wo ist mein Schlüssel? Er liegt auf dem Tisch.","Where is my key? It is lying on the table.","मारी चाबी क्यां? टेबल पर पडी छे."),
("Ich stelle die Vase auf das Regal.","I place the vase on the shelf.","हूं vase shelf पर मूकूं छूं."),
("Die Kinder sitzen auf dem Sofa.","The children are sitting on the sofa.","बाळको sofa पर बेठां छे."),
("Er hängt das Bild an die Wand.","He hangs the picture on the wall.","ते ছबि दीवाल पर टांगे छे."),
("Ich lege mich ins Bett.","I lie down in the bed.","हूं बिछाने मां सूई जाउं छूं."),
],
[
("🔑","Location vs Direction","Location (Wo?) → Dative: Das Buch LIEGT auf dem Tisch. Direction (Wohin?) → Accusative: Ich LEGE das Buch auf den Tisch.","E8F0FE"),
("💡","Pairs to Know","stehen/stellen | liegen/legen | hängen/hängen | sitzen/setzen. Each pair = location verb / direction verb. Learn them as pairs!","FFF8E1"),
("🔊","Pronunciation","'stehen' = SHTAY-en. 'stellen' = SHTEL-en. 'liegen' = LEE-gen. 'legen' = LAY-gen. The vowel change signals location vs direction!","E6F4EA"),
],
["Describe your room using 5 location sentences (stehen/liegen/hängen).","Write 5 direction sentences: where you put things (stellen/legen/hängen).","Make a comparison table: location verb | direction verb | example."],
[("'Das Buch liegt auf dem Tisch' uses which case?","Dative (location: Wo?)"),("'Ich lege das Buch auf den Tisch' uses which case?","Accusative (direction: Wohin?)"),("'stellen' is the direction verb for…?","stehen (standing objects)"),("Translate: 'Sie hängt das Bild an die Wand.'","She hangs the picture on the wall."),("'liegen' corresponds to which direction verb?","legen")],
"Location and direction verb pairs mastered! This is one of the most uniquely German grammar concepts. Your spatial descriptions in German are now native-level precise!"
))

# ── Day 117 ─────────────────────────────────────────────────────────────────
DAYS.append((117,
"GERMAN SOCIAL SITUATIONS 🤝",
"Soziale Situationen — Everyday German Life",
"Day-to-day social situations in Germany require specific phrases. Today you learn how to navigate parties, introduce people to each other, make small talk, show sympathy, congratulate, and comfort someone in German.",
"Germany मां day-to-day social situations specific phrases मांगे. आज parties, एकबीजाने introduce करवां, small talk करवूं, सहानुभूति दर्शाववी, congratulate करवूं, अने German मां comfort करवूं शीखशो.",
[
("die Einladung","the invitation","आमंत्रण","dee EYN-lah-dung"),
("einladen","to invite (sep.)","आमंत्रण आपवूं","EYN-lah-den"),
("vorstellen","to introduce (sep.)","परिचय कराववो","FOR-shtel-en"),
("Darf ich vorstellen? Das ist…","May I introduce? This is…","शूं परिचय करावूं? आ...","DARF ikh FOR-shtel-en? das ist"),
("sich wohl fühlen","to feel comfortable","ठीक/comfortable feel करवूं","zikh VOHL FÜ-len"),
("Herzlichen Glückwunsch!","Congratulations!","अभिनंदन!","HERTS-likh-en GLÜK-voonsh"),
("Alles Gute!","All the best!","शुभकामनाएं!","AL-es GOO-teh"),
("Das tut mir leid.","I'm sorry (sympathy).","मने खेद छे.","das toot meer LIDE"),
("Gute Besserung!","Get well soon!","जल्दी ठीक थाओ!","GOO-teh BES-eh-rung"),
("Zum Wohl!","To your health! (toast)","आपका स्वास्थ्य!","tsoom VOHL"),
("Prost!","Cheers!","jai ho!","PROHST"),
("Viel Erfolg!","Good luck! (much success)","ઘणी सफळता!","feel ehr-FOLK"),
("Viel Spaß!","Have fun!","ઘણી मजा!","feel SHPAS"),
("Wie schön!","How lovely!","केटलूं सुंदर!","vee SHÖN"),
("Na dann, viel Spaß!","Well then, have fun!","तो पछी, मजा करजे!","nah DAN, feel SHPAS"),
],
[
("Herzlichen Glückwunsch zu deiner Beförderung!","Congratulations on your promotion!","तारी promotion बद्दल अभिनंदन!"),
("Darf ich vorstellen? Das ist mein Kollege Thomas.","May I introduce? This is my colleague Thomas.","शूं परिचय करावूं? आ मारो colleague Thomas छे."),
("Das tut mir wirklich leid — ich hoffe, es geht dir bald besser.","I'm really sorry — I hope you get better soon.","मने सच मां खेद छे — आशा छे तूं जल्दी ठीक थशे."),
("Zum Wohl! Auf eine gute Zusammenarbeit!","To your health! To good collaboration!","आपका स्वास्थ्य! सारी collaboration माटे!"),
("Ich fühle mich hier sehr wohl.","I feel very comfortable here.","मने अहीं खूब comfortable feel थाय छे."),
],
[
("🇩🇪","German Parties","Germans tend to be punctual even at parties! Bring a small gift (wine, chocolate, flowers) when invited to someone's home. Never arrive empty-handed!","E6F4EA"),
("💡","'Leid' Expressions","'Das tut mir leid' = I'm sorry (sympathy). 'Es tut mir leid' = I'm sorry (apology). Similar structure, slightly different meaning!","FFF8E1"),
("📌","Toast Etiquette","When toasting in Germany: make EYE CONTACT with everyone at the table! Not making eye contact is considered bad luck (7 years bad luck, they say!).","E8F0FE"),
],
["Write a script for a German dinner party conversation.","Practise: write congratulation messages for 5 different occasions.","Write how you would comfort a friend going through a difficult time in German."],
[("How do you say 'Congratulations!'?","Herzlichen Glückwunsch!"),("Translate: 'Gute Besserung!'","Get well soon!"),("How do you propose a toast?","Prost! / Zum Wohl!"),("Translate: 'Das tut mir wirklich leid.'","I am really sorry."),("'Viel Erfolg!' means?","Good luck! / Much success!")],
"Social German vocabulary complete! You can navigate any German social situation with warmth, politeness, and cultural awareness. You're truly ready for German social life!"
))

# ── Day 118 ─────────────────────────────────────────────────────────────────
DAYS.append((118,
"TELLING STORIES IN GERMAN 📖",
"Eine Geschichte erzählen — Narrative German",
"Storytelling is the heart of real communication. Today you learn narrative phrases, sequence words, and how to tell a story (what happened, who was there, how it ended) using all your past tense knowledge.",
"Storytelling real communication नो heart छे. आज narrative phrases, sequence words, अने story कहेवी (शूं थयूं, कोण हतूं, केवूं पूरूं थयूं) — past tense ना बधां knowledge वापरीने शीखशो.",
[
("erzählen","to tell / narrate","कहेवूं","ehr-TSÄH-len"),
("die Geschichte","the story","कहाणी","dee geh-SHIKH-teh"),
("zuerst","first of all","पहेलां","tsoo-ERST"),
("dann","then","पछी","DAN"),
("danach","after that","त्यार पछी","dah-NAKH"),
("schließlich","finally","अंते","SHLES-likh"),
("plötzlich","suddenly","અचानक","PLÖTS-likh"),
("leider","unfortunately","दुर्भाग्यथी","LY-der"),
("zum Glück","luckily","सारूं थयूं","tsoom GLÜK"),
("das Ergebnis","the result","परिणाम","das ehr-GAIP-nis"),
("am Ende","at the end","अंते","am EN-deh"),
("übrigens","by the way","वैसे","Ü-bri-gens"),
("Es war einmal…","Once upon a time…","एक वखत...","es VAHR EYN-mahl"),
("Stell dir vor…","Imagine…","कल्पना कर...","SHTEL deer FOR"),
("Kannst du dir vorstellen…?","Can you imagine…?","तूं imagine करी शकशे?","KANST doo deer FOR-shtel-en"),
],
[
("Zuerst sind wir in den Park gegangen.","First we went to the park.","पहेलां आपणे park गया."),
("Dann haben wir ein Picknick gemacht.","Then we had a picnic.","पछी आपणे picnic किया."),
("Plötzlich hat es angefangen zu regnen!","Suddenly it started to rain!","અचानक वरसाद शरू थई गयो!"),
("Zum Glück hatten wir einen Regenschirm dabei.","Luckily we had an umbrella with us.","सारूं थयूं के आपणे छत्री साथे रखी हती."),
("Am Ende sind wir alle lachend nach Hause gegangen.","At the end we all went home laughing.","अंते आपणे बधां हसतां हसतां घेर गया."),
],
[
("💡","Story Template","Zuerst… → Dann… → Danach… → Plötzlich… → Zum Glück / Leider… → Schließlich… → Am Ende…. Use this framework for any story!","FFF8E1"),
("📌","Past Tenses in Stories","Use Perfekt for spoken stories. Use war/hatte/konnte/musste (Präteritum) freely. Mix them naturally — just like native speakers do!","E8F0FE"),
("🇩🇪","German Storytelling","Germans appreciate detailed, logical stories. Don't rush to the punchline — build up with context, characters, and sequence words. They love a well-told story!","E6F4EA"),
],
["Tell the story of your most memorable day in German (8-10 sentences).","Write a fairy-tale opening: 'Es war einmal…' and continue for 5 sentences.","Re-tell a movie plot in German using past tense and narrative words."],
[("Translate: 'Zuerst bin ich aufgestanden.'","First I got up."),("'plötzlich' means?","Suddenly"),("'zum Glück' means?","Luckily"),("Translate: 'Am Ende haben wir gewonnen.'","At the end we won."),("How do you begin a fairy tale in German?","Es war einmal… (Once upon a time…)")],
"Storytelling German mastered! You can now narrate events, experiences, and adventures in German with natural flow and drama. Your German has become a complete communication tool!"
))

# ── Day 119 ─────────────────────────────────────────────────────────────────
DAYS.append((119,
"A2 GRAND REVIEW 🏆",
"Zusammenfassung — Everything in A2",
"You are ONE day away from completing A2! Today is your comprehensive review of all A2 grammar and vocabulary. Test yourself on cases, tenses, modal verbs, prepositions, and everything you've learned from Days 61 to 118.",
"A2 complete थवामां ONE day बाकी! आज सभी A2 grammar अने vocabulary नो comprehensive review. Days 61-118 थी cases, tenses, modal verbs, prepositions — बधां पर खुद ने test करो.",
[
("der Nominativ","nominative case (subject)","प्रथम कारक","dehr noh-mi-nah-TEEF"),
("der Akkusativ","accusative (direct obj.)","द्वितीय कारक","dehr ak-koo-zah-TEEF"),
("der Dativ","dative (indirect obj.)","तृतीय कारक","dehr dah-TEEF"),
("der Genitiv","genitive (possession)","चतुर्थ कारक","dehr geh-neh-TEEF"),
("das Präsens","present tense","वर्तमान काळ","das PRÄ-zens"),
("das Perfekt","present perfect","Perfekt tense","das pehr-FEKT"),
("das Präteritum","simple past","सामान्य भूतकाळ","das prä-teh-REE-toom"),
("das Futur I","future tense","भविष्य काळ","das foo-TOOR EYNS"),
("der Konjunktiv II","subjunctive II (würde)","शर्तार्थ","dehr kon-YOONK-teef TSVY"),
("die Modalverben","modal verbs","modal verbs","dee moh-DAHL-fehr-ben"),
("trennbare Verben","separable verbs","separable verbs","TREN-bah-reh FEHR-ben"),
("reflexive Verben","reflexive verbs","reflexive verbs","reh-FLEX-ee-veh FEHR-ben"),
("der Nebensatz","subordinate clause","subordinate clause","dehr NAY-ben-zats"),
("der Hauptsatz","main clause","main clause","dehr HOWPT-zats"),
("der Relativsatz","relative clause","relative clause","dehr reh-lah-TEEF-zats"),
],
[
("Ich habe A2 fast abgeschlossen!","I have almost completed A2!","में A2 लगभग पूर्ण कर्यूं!"),
("Mein Deutsch hat sich sehr verbessert.","My German has improved a lot.","मारूं German घणूं सुधर्यूं छे."),
("Ich kann jetzt über die Vergangenheit sprechen.","I can now speak about the past.","हवे हूं भूतकाळ बद्दल बोली शकूं छूं."),
("Ich verstehe Deutsche Texte viel besser.","I understand German texts much better.","हवे हूं German texts घणा सारी रीते समजूं छूं."),
("Morgen beende ich offiziell mein A2-Niveau.","Tomorrow I officially complete my A2 level.","काले हूं officially A2 level पूर्ण करूं छूं."),
],
[
("📊","A2 Grammar Summary","4 cases ✓ | 6 modal verbs ✓ | Perfekt (haben+sein) ✓ | Präteritum (sein/haben/modals) ✓ | Futur I ✓ | Separable verbs ✓ | Reflexive verbs ✓ | Sub+Coord conjunctions ✓ | Adjective endings ✓ | Relative clauses ✓","E8F0FE"),
("📚","A2 Vocabulary","Topics: shopping, travel, weather, body, health, work, home, hobbies, transport, culture, technology, environment, social situations, letters. 1500+ words!","E6F4EA"),
("🎯","Tomorrow","Day 120 is your A2 Grand Finale with a comprehensive final test and your 'What's Next?' roadmap to B1!","FFF8E1"),
],
["Write a 15-sentence German paragraph using as many A2 grammar structures as possible.","Make a 'what I know' list: grammar points and vocabulary topics you have mastered.","Translate a paragraph from Gujarati/English to German using all your A2 knowledge."],
[("Name all 6 modal verbs.","können, müssen, wollen, sollen, dürfen, mögen/möchten"),("Perfekt of 'fahren' (uses sein)?","ist gefahren"),("Translate: 'Obwohl es regnet, gehe ich spazieren.'","Although it rains, I go for a walk."),("What is the Futur I formula?","werden + infinitive"),("Name 3 Dative-only prepositions.","aus, bei, mit, nach, seit, von, zu (any 3)")],
"A2 Grand Review complete! You have proven your mastery of A2 German. One final day remains — the Grand Finale celebration and your B1 roadmap. Du schaffst das! (You can do it!)"
))

# ── Day 120 ─────────────────────────────────────────────────────────────────
DAYS.append((120,
"A2 GRAND FINALE & WHAT'S NEXT! 🎓",
"A2 Abschluss — Your Journey Continues",
"TODAY IS THE DAY! You have completed 120 days of German learning. You have gone from complete beginner (A1) to solid intermediate foundation (A2). Today we celebrate your achievement, complete a final comprehensive test, and map out your journey to B1!",
"આ दिवस आव्यो! 120 दिवस German शिखवाना पूर्ण. Complete beginner (A1) थी solid intermediate foundation (A2) सुधी आव्यां. आज achievement celebrate करीए, final test, अने B1 journey नो roadmap!",
[
("der Abschluss","completion / graduation","पूर्णता","dehr AP-shloos"),
("das Niveau","level / standard","स्तर","das nee-VOH"),
("die Zertifizierung","certification","प्रमाणपत्र","dee tsehr-tee-fee-TSEE-rung"),
("das Goethe-Institut","Goethe Institute","Goethe Institute","das GÖ-teh-in-sti-TOOT"),
("das Sprachdiplom","language diploma","भाषा diploma","das SHPRAKHT-di-plohm"),
("fließend","fluent","fluent","FLEE-sent"),
("fortgeschritten","advanced","उन्नत","FORT-geh-shrit-en"),
("die Kompetenz","competence","क्षमता","dee kom-peh-TENTS"),
("anwenden","to apply (knowledge)","ज्ञान वापरवूं","AN-ven-den"),
("das Selbststudium","self-study","स्व-अध्ययन","das ZELPST-shtoo-dee-oom"),
("der Sprachkurs","language course","language course","dehr SHPRAKHT-koors"),
("Herzlichen Glückwunsch!","Congratulations!","अभिनंदन!","HERTS-likh-en GLÜK-voonsh"),
("Ich bin stolz auf mich.","I am proud of myself.","हूं मारी जाते पर गर्वित छूं.","ikh bin SHTOLTS owf mikh"),
("weitermachen","to continue (sep.)","आगे वधवूं","VY-ter-makh-en"),
("Ich schaffe das!","I can do it!","हूं करी शकूं!","ikh SHAF-eh das"),
],
[
("Ich habe A2 erfolgreich abgeschlossen — ich bin so stolz auf mich!","I have successfully completed A2 — I am so proud of myself!","में सफળतापूर्वक A2 पूर्ण कर्यूं — हूं मारी जाते पर घणो/ઘणी गर्वित छूं!"),
("Mein Deutsch ist jetzt viel besser als vor 120 Tagen.","My German is now much better than 120 days ago.","120 दिवस पहेलांथी मारूं German घणूं सारूं छे."),
("Ich werde bald mit dem B1-Kurs beginnen.","I will soon begin the B1 course.","हूं जल्दी B1 course शरू करूं छूं."),
("Das Goethe-Institut bietet offizielle Sprachprüfungen an.","The Goethe Institute offers official language exams.","Goethe Institute official language exams offer करे छे."),
("Auf Wiedersehen, A2! Willkommen, B1!","Goodbye, A2! Welcome, B1!","Alvida, A2! Svāgat, B1!"),
],
[
("🏆","A2 Achievement Summary","COMPLETED: 4 German cases, 6 modal verbs, Perfekt + Präteritum + Futur I tenses, separable & reflexive verbs, adjective endings, relative clauses, subordinate/coordinating conjunctions, 1500+ vocabulary words across 30+ life topics!","E8F0FE"),
("🎯","What's Next — B1 Topics","B1 covers: Konjunktiv II (würde…), passive voice, infinitive constructions (um…zu, ohne…zu), two-way prepositions mastery, complex reading, email writing, and preparing for the official Goethe B1 exam!","E6F4EA"),
("🌟","Recommended Next Steps","1. Review all A2 days weekly. 2. Register for Goethe-Institut A2 exam (certificate proof). 3. Start watching German TV with subtitles. 4. Find a German conversation partner (Tandem partner). 5. Start B1 in this book series!","FFF8E1"),
],
["Write a reflection: 'In 120 Tagen habe ich...' (10+ sentences about your learning journey).","Take the A2 Final Test on the next pages.","Write your 3 B1 goals in German."],
[("What are the 4 German cases?","Nominativ, Akkusativ, Dativ, Genitiv"),("Futur I: 'He will come' in German?","Er wird kommen."),("Translate: 'Ich bin stolz auf mich.'","I am proud of myself."),("Where can you take an official German exam?","Goethe-Institut"),("Translate: 'Ich schaffe das!'","I can do it!")],
"🎓 HERZLICHEN GLÜCKWUNSCH! You have completed the A2 German Course — 120 days, 1500+ words, all major grammar topics, real-life vocabulary. You are no longer a beginner. You are on your way to fluency. The German-speaking world is opening up to you. Weiter so — auf nach B1! (Keep going — onwards to B1!)"
))

print("Days 116-120 appended! All 60 days done!")

# ═══════════════════════ LEVEL TESTS ════════════════════════════════════════

TESTS = []

TESTS.append((1,"GRAMMAR FOUNDATIONS TEST","Days 61–70: Cases, Prepositions & Modal Verbs",[
("What are the four German cases?","Nominativ (subject), Akkusativ (direct object), Dativ (indirect object), Genitiv (possession)"),
("Give the Dative article for: der Mann / die Frau / das Kind.","dem Mann / der Frau / dem Kind"),
("Translate: 'Ich helfe dem Mann.'","I help the man."),
("Name 5 Dative-only prepositions.","aus, bei, mit, nach, seit, von, zu, gegenüber, außer (any 5)"),
("Name 4 Accusative-only prepositions.","durch, für, gegen, ohne, um, bis, entlang (any 4)"),
("What question word finds the Dative?","Wem? (To whom?)"),
("What question word finds the Accusative?","Wen? / Was? (Whom? / What?)"),
("Translate: 'Ich kaufe das Buch für meine Mutter.'","I buy the book for my mother."),
("Two-way prepositions: what case after 'Wo?' and 'Wohin?'","Wo? → Dativ | Wohin? → Akkusativ"),
("Translate: 'Das Buch liegt auf dem Tisch.'","The book is lying on the table."),
("What are the 6 German modal verbs?","können, müssen, wollen, sollen, dürfen, mögen/möchten"),
("Translate: 'Ich kann Deutsch sprechen.'","I can speak German."),
("In a modal verb sentence, where does the infinitive go?","At the end of the sentence"),
("Translate: 'Du musst mehr üben.'","You must practise more."),
("What is the polite form of 'wollen' in a shop?","möchten — 'Ich möchte…'"),
("Name 5 common separable verb prefixes.","auf-, an-, ab-, ein-, aus-, mit-, vor-, zurück- (any 5)"),
("Translate: 'Ich stehe um 6 Uhr auf.'","I get up at 6 o'clock."),
("Translate: 'Wir kaufen heute ein.'","We go shopping today."),
("'Was kostet das?' means?","How much does this cost?"),
("Translate: 'Ich möchte einen Tisch reservieren.'","I would like to reserve a table."),
]))

TESTS.append((2,"SHOPPING, FOOD & TRANSPORT TEST","Days 71–80: Real-Life German",[
("How do you say 'How much does this cost?' in German?","Was kostet das?"),
("Translate: 'Ich suche eine Jacke in Größe M.'","I am looking for a jacket in size M."),
("How do you say 'receipt' in German?","der Kassenbon / die Quittung"),
("'günstig' vs 'teuer' — what do they mean?","günstig = cheap/affordable; teuer = expensive"),
("How do you ask for 500g of cheese at a market?","Ich hätte gerne 500 Gramm Käse."),
("Translate: 'Sind die Tomaten frisch?'","Are the tomatoes fresh?"),
("How do you say 'the fridge' in German?","der Kühlschrank"),
("Translate: 'Das schmeckt sehr lecker!'","That tastes very delicious!"),
("Name 4 German taste adjectives.","süß (sweet), salzig (salty), scharf (spicy), sauer (sour), bitter (any 4)"),
("How do you ask for the bill in a restaurant?","Die Rechnung, bitte."),
("How do you politely call a waiter's attention?","Entschuldigung!"),
("'Ich bin Vegetarier' — what does this mean?","I am vegetarian."),
("How do you say 'cash' in German?","Bargeld"),
("What does 'Stimmt so' mean when paying?","Keep the change."),
("How do you say '1 million' in German?","eine Million"),
("Name the 3 main German rail types.","ICE, IC, Regional (or S-Bahn, U-Bahn, Straßenbahn)"),
("Translate: 'Muss ich umsteigen?'","Do I have to change trains?"),
("What does 'Verspätung' mean?","Delay"),
("Translate: 'Gehen Sie geradeaus, dann rechts.'","Go straight ahead, then turn right."),
("'die Kreuzung' means?","The crossroads / intersection"),
]))

TESTS.append((3,"PAST TENSE & PROFESSIONS TEST","Days 81–90: Travel, Health, Work & Perfekt",[
("How do you form the Perfekt tense?","haben/sein (conjugated) + Partizip II at the end"),
("Partizip II of 'machen'?","gemacht"),
("Partizip II of 'spielen'?","gespielt"),
("Partizip II of 'essen'?","gegessen"),
("Partizip II of 'trinken'?","getrunken"),
("Which verbs use 'sein' in Perfekt?","Verbs of movement (gehen, fahren, fliegen) and change of state (aufwachen, werden)"),
("Partizip II of 'gehen'?","gegangen"),
("Translate: 'Ich bin gestern nach Hause gegangen.'","I went home yesterday."),
("Präteritum of 'sein' (ich)?","ich war"),
("Präteritum of 'haben' (er)?","er hatte"),
("Präteritum of 'können' (sie pl.)?","sie konnten"),
("Translate: 'Wir waren sehr müde.'","We were very tired."),
("How do you say 'I have had a fever for 3 days'?","Ich habe seit drei Tagen Fieber."),
("What is the emergency number in Germany?","112"),
("'die Apotheke' means?","The pharmacy"),
("'Was sind Sie von Beruf?' means?","What is your profession?"),
("Translate: 'Ich arbeite als Ingenieur.'","I work as an engineer."),
("'Pünktlichkeit' means?","Punctuality"),
("Describe Mülltrennung in one sentence.","Waste separation — sorting rubbish into different bins for recycling in Germany."),
("Translate: 'Ich bin seit zwei Jahren in Deutschland.'","I have been in Germany for two years."),
]))

TESTS.append((4,"DAILY LIFE & CULTURE TEST","Days 91–100: Routines, Home, Hobbies & Culture",[
("Name 5 reflexive verbs for daily routine.","sich waschen, sich anziehen, sich duschen, sich setzen, sich fühlen (any 5)"),
("What is the reflexive pronoun for 'wir'?","uns"),
("Translate: 'Ich freue mich auf den Urlaub.'","I am looking forward to the holiday."),
("What is a 'Wohngemeinschaft (WG)'?","A shared flat"),
("Translate: 'Das Wohnzimmer ist sehr gemütlich.'","The living room is very cosy."),
("'gemütlich' means?","Cosy / comfortable"),
("'Staub saugen' means?","To vacuum"),
("Translate: 'Kannst du bitte aufräumen?'","Can you please tidy up?"),
("How do you express interest in a hobby?","'Ich interessiere mich für…' + noun in Accusative"),
("'gerne' after a verb means?","'I like to' + verb (e.g., 'Ich lese gerne' = I like to read)"),
("Name 3 popular sports in Germany.","Fußball, Tennis, Schwimmen, Radfahren, Wandern (any 3)"),
("Translate: 'Deutschland hat gewonnen!'","Germany won!"),
("What is the Bundesliga?","Germany's top football league"),
("When is Oktoberfest held, and where?","September-October in Munich (München)"),
("'Frohe Weihnachten!' means?","Merry Christmas!"),
("What is the 'Weihnachtsmarkt'?","The Christmas market"),
("When is German Unity Day?","3rd October (3. Oktober)"),
("Translate: 'Ich bin so stolz auf mich!'","I am so proud of myself!"),
("What does 'Weiter so!' mean?","Keep it up!"),
("Translate: 'Ich mache weiter bis zum Ende!'","I'll keep going until the end!"),
]))

TESTS.append((5,"ADVANCED GRAMMAR TEST","Days 101–110: Future, Comparatives, Conjunctions & Clauses",[
("Futur I formula?","werden (conjugated) + infinitive at the end"),
("Translate: 'Er wird morgen kommen.'","He will come tomorrow."),
("Comparative of 'gut'?","besser"),
("Superlative of 'gut'?","am besten"),
("Comparative of 'groß'?","größer"),
("Translate: 'Das ist teurer als das andere.'","That is more expensive than the other one."),
("'genauso…wie' means?","Just as…as (equal comparison)"),
("In a 'weil' clause, where does the verb go?","To the end"),
("Translate: 'Ich komme nicht, weil ich krank bin.'","I am not coming because I am ill."),
("Difference between 'weil' and 'denn'?","weil = verb to end (subordinate); denn = normal word order (coordinating)"),
("'obwohl' means?","Although"),
("What case do adjectives take after 'die' (Nominative)?","-e ending (die rote Rose)"),
("What ending do adjectives usually take after 'einen' (Accusative m)?","-en (einen neuen Mantel)"),
("Translate: 'Das kleine Kind schläft.'","The small child is sleeping."),
("What determines the gender of a relative pronoun?","The gender of the noun it refers to"),
("Where does the verb go in a relative clause?","To the end"),
("Translate: 'Das ist das Buch, das ich lese.'","That is the book that I am reading."),
("'hoffentlich' means?","Hopefully"),
("Translate: 'Was wirst du in Zukunft machen?'","What will you do in the future?"),
("Name 3 subordinating conjunctions.","weil, dass, ob, wenn, obwohl, damit, bevor, nachdem (any 3)"),
]))

TESTS.append((6,"A2 FINAL COMPREHENSIVE TEST","Days 61–120: The Complete A2 Challenge",[
("State the rule for verb position in subordinate clauses.","The verb goes to the END of the subordinate clause."),
("Translate: 'Obwohl es kalt ist, gehe ich spazieren.'","Although it is cold, I am going for a walk."),
("Conjugate 'werden' for: ich, du, er.","ich werde, du wirst, er wird"),
("Partizip II of 'schreiben'?","geschrieben"),
("'Ich habe das Buch gelesen' — translate.","I have read the book."),
("'Sie ist nach Berlin geflogen' — translate.","She has flown to Berlin."),
("Präteritum of 'haben' (wir)?","wir hatten"),
("Give the Dative forms: der Mann / die Frau / das Kind.","dem Mann / der Frau / dem Kind"),
("Name all 6 modal verbs with their meanings.","können (can), müssen (must), wollen (want), sollen (should), dürfen (may), mögen (like)"),
("Two-way prepositions: 'Wo?' uses which case?","Dative"),
("'in + Accusative' is used for?","Direction: movement into a place"),
("'in + Dative' is used for?","Location: being inside a place"),
("Translate: 'Ich stelle das Buch ins Regal.'","I put the book in the shelf."),
("Translate: 'Das Buch steht im Regal.'","The book is (standing) in the shelf."),
("Reflexive pronoun for 'er/sie/es'?","sich"),
("Translate: 'Ich fühle mich heute nicht gut.'","I don't feel well today."),
("Adjective ending after definite article in Dative?","-en (dem alten Mann)"),
("Translate: 'Das ist der Mann, der Deutsch spricht.'","That is the man who speaks German."),
("How do you say 'I am proud of myself'?","Ich bin stolz auf mich."),
("Translate: 'Herzlichen Glückwunsch! Du hast A2 abgeschlossen!'","Congratulations! You have completed A2!"),
]))

print("Tests defined!")

# ═══════════════════════ MAIN FUNCTION ════════════════════════════════════

def build_book():
    doc = Document()

    # ── Page size A4 ──────────────────────────────────────────────────────
    from docx.shared import Cm
    for section in doc.sections:
        section.page_width  = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # ── TITLE PAGE ────────────────────────────────────────────────────────
    add_para(doc, "🇩🇪", size_pt=60, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=40, space_after=8)
    add_para(doc, "DEUTSCH LERNEN", size_pt=32, bold=True, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "A2 Course — Days 61 to 120", size_pt=18, bold=True, color=YELLOW, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "ગુજરાતી ભાષકો માટે A2 જર્મન અભ્યાસ", size_pt=14, bold=True, color=DARK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "For Gujarati Speakers | English + Gujarati Explanations", size_pt=12, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, "60 Days · 6 Level Tests · Real-Life German · What's Next Guide", size_pt=12, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)
    doc.add_paragraph()

    # ── HOW TO USE ────────────────────────────────────────────────────────
    add_para(doc, "📖 How to Use This Book", size_pt=14, bold=True, color=BLUE, space_after=4)
    tips = [
        "�� Study ONE day at a time — consistency beats speed.",
        "✍️  Write vocabulary in a notebook every day.",
        "🗣️  Say every German word OUT LOUD — pronunciation is key!",
        "🔁  Review the previous day's words before starting the new day.",
        "🎯  Complete every Practice Task and Mini Test honestly.",
        "🇩🇪  Imagine yourself in Germany while reading examples — it makes it stick!",
    ]
    for t in tips:
        add_bullet(doc, t)
    doc.add_paragraph()

    # ── A2 LEVEL INTRO ────────────────────────────────────────────────────
    doc.add_page_break()
    add_para(doc, "🟩  A2 LEVEL  —  THE REAL START", size_pt=18, bold=True, color=GREEN, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=20, space_after=4)
    add_para(doc, "Days 61 – 120", size_pt=14, bold=True, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    add_para(doc, "A2 transforms you from a cautious beginner into a confident communicator. After 60 days you will:", size_pt=11, space_after=4)
    achievements = [
        "✅ Use all four German cases (Nominativ, Akkusativ, Dativ, Genitiv)",
        "✅ Speak in three tenses: Present, Perfect, and Future",
        "✅ Use all 6 modal verbs and separable/reflexive verbs naturally",
        "✅ Shop, travel, order food, make appointments, and handle emergencies",
        "✅ Write formal and informal letters/emails in German",
        "✅ Talk about work, home, hobbies, health, and culture",
        "✅ Understand German texts and hold real conversations",
    ]
    for a in achievements:
        add_bullet(doc, a)
    add_para(doc, "🎯 Motivation: A2 is where German starts to feel REAL. Every day brings you closer to fluency!", size_pt=11, bold=True, color=BLUE, space_before=6, space_after=4)
    doc.add_paragraph()

    # ── ALL 60 DAYS ───────────────────────────────────────────────────────
    for entry in DAYS:
        add_day(doc, *entry)

    # ── 6 LEVEL TESTS ─────────────────────────────────────────────────────
    for test_num, title, days_covered, questions in TESTS:
        add_level_test(doc, test_num, title, days_covered, questions)

    # ── WHAT'S NEXT PAGE ─────────────────────────────────────────────────
    doc.add_page_break()
    add_para(doc, "🚀  WHAT'S NEXT?", size_pt=28, bold=True, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=20, space_after=4)
    add_para(doc, "Your Journey to B1 and Beyond", size_pt=16, bold=True, color=YELLOW, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)

    add_para(doc, "🎓 You Have Completed A2!", size_pt=13, bold=True, color=GREEN, space_after=3)
    add_para(doc, (
        "Finishing A2 means you can: communicate in everyday situations, understand simple texts, "
        "write basic emails, describe past and future events, express opinions, and navigate life "
        "in Germany with confidence. This is a REAL language skill — be proud!"
    ), size_pt=11, space_after=8)

    add_para(doc, "📚 B1 — What You Will Learn", size_pt=13, bold=True, color=BLUE, space_after=3)
    b1_topics = [
        "Konjunktiv II — 'würde' constructions (I would do / I could do…)",
        "Passive Voice (Das Buch wird gelesen — The book is being read)",
        "Infinitive Constructions: um…zu (in order to), ohne…zu (without…-ing), anstatt…zu",
        "Relative clauses with Genitive pronouns (dessen/deren)",
        "Extended written German — essays, formal applications, complex arguments",
        "Reading authentic German texts: newspapers, novels, instructions",
        "Complex vocabulary: politics, philosophy, science, emotions",
        "Preparing for the official Goethe-Institut B1 exam",
    ]
    for t in b1_topics:
        add_bullet(doc, f"✅ {t}")
    doc.add_paragraph()

    add_para(doc, "🗓️ Recommended Study Plan for B1", size_pt=13, bold=True, color=BLUE, space_after=3)
    plan = [
        "Days 121–150: Konjunktiv II, Passive Voice, Infinitive constructions",
        "Days 151–180: Extended reading, complex grammar, debate language",
        "Days 181–210: Authentic texts, advanced vocabulary, writing skills",
        "Days 211–240: Exam preparation, full practice tests, revision",
    ]
    for p in plan:
        add_bullet(doc, p)
    doc.add_paragraph()

    add_para(doc, "🌟 5 Steps to Accelerate Your German Right Now", size_pt=13, bold=True, color=BLUE, space_after=3)
    steps = [
        "1️⃣  REVIEW — Go through A2 vocabulary weekly. Spaced repetition = long-term memory.",
        "2️⃣  CONSUME — Watch German TV shows (Dark, Babylon Berlin) with German subtitles.",
        "3️⃣  SPEAK — Find a Tandem partner (language exchange). Apps: Tandem, HelloTalk.",
        "4️⃣  CERTIFY — Take the Goethe-Institut A2 exam. It proves your level officially.",
        "5️⃣  CONTINUE — Start Day 121 in the B1 book. Don't stop the momentum!",
    ]
    for s in steps:
        add_bullet(doc, s)
    doc.add_paragraph()

    add_para(doc, "📱 Recommended German Learning Resources", size_pt=13, bold=True, color=BLUE, space_after=3)
    resources = [
        "Apps: Duolingo (vocab), Anki (flashcards), Deutsche Welle (DW Learn German — free!)",
        "YouTube: Easy German, Deutsch mit Marija, Learn German with Anja",
        "Podcasts: Slow German, Coffee Break German, GermanPod101",
        "Websites: dw.com/deutsch-lernen, goethe.de, Deutsche Bahn (for reading practice!)",
        "Books: Menschen A2, Netzwerk A2 (Klett Verlag)",
    ]
    for r in resources:
        add_bullet(doc, f"🔗 {r}")
    doc.add_paragraph()

    add_para(doc, "🇩🇪 Final Message from Your German Learning Journey", size_pt=13, bold=True, color=BLUE, space_after=3)
    add_para(doc, (
        "You have done something extraordinary. When most people say 'I want to learn German' "
        "they give up within two weeks. You didn't. You showed up every single day for 120 days. "
        "You wrestled with Dativ and Akkusativ, you memorised irregular Partizip II, you practised "
        "at restaurants, markets, and doctor's offices — all in your head, all building real skills.\n\n"
        "The German language has a word for the feeling of joy from the success of others: Mitfreude. "
        "We feel Mitfreude for YOUR achievement today.\n\n"
        "Germany is waiting for you. The language is yours. Keep going.\n\n"
        "Auf Wiedersehen — und bis bald auf B1!"
    ), size_pt=11, space_after=6)

    add_para(doc, "🏆  Herzlichen Glückwunsch zum A2-Abschluss!", size_pt=16, bold=True, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10, space_after=4)
    add_para(doc, "Weiter so — Du schaffst das!", size_pt=14, bold=True, color=YELLOW, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_para(doc, "(Keep it up — You can do it!)", size_pt=11, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)

    # ── SAVE ──────────────────────────────────────────────────────────────
    out = "German_Learning_Book_A2_Gujarati.docx"
    doc.save(out)
    print(f"✅  Saved: {out}")
    print(f"   Days: {len(DAYS)}")
    print(f"   Tests: {len(TESTS)}")

build_book()
