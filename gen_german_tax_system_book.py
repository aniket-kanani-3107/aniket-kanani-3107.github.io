"""
Generate a full-length Word (.docx) book on the German tax system.
This script uses only the Python standard library.
"""

from __future__ import annotations

from xml.sax.saxutils import escape
import zipfile


def para(text: str, *, bold: bool = False, size: int = 22) -> str:
    text = escape(text)
    rpr = []
    if bold:
        rpr.append("<w:b/>")
    if size:
        # half-points in OOXML
        rpr.append(f'<w:sz w:val="{size * 2}"/>')
        rpr.append(f'<w:szCs w:val="{size * 2}"/>')
    rpr_xml = f"<w:rPr>{''.join(rpr)}</w:rPr>" if rpr else ""
    return (
        "<w:p><w:r>"
        f"{rpr_xml}"
        f"<w:t xml:space=\"preserve\">{text}</w:t>"
        "</w:r></w:p>"
    )


def bullet(text: str) -> str:
    text = escape(text)
    return (
        "<w:p>"
        "<w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>"
        "<w:r><w:rPr><w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/></w:rPr>"
        f"<w:t xml:space=\"preserve\">{text}</w:t></w:r>"
        "</w:p>"
    )


def heading_1(text: str) -> str:
    return para(text, bold=True, size=18)


def heading_2(text: str) -> str:
    return para(text, bold=True, size=15)


def build_book_blocks() -> list[str]:
    blocks: list[str] = []
    blocks.append(para("The German Tax System: Full-Length Beginner-Friendly Book (English + Key German Terms)", bold=True, size=24))
    blocks.append(para("Edition: 2026 educational reference | Focus: Germany only | Structure: Introduction -> Direct Taxes -> Indirect Taxes", size=11))

    blocks.append(heading_1("Part I — Introduction to the German Tax System"))
    blocks.append(
        para(
            "Welcome to this full-length guide to Germany’s tax system. This book is written in "
            "simple English so even a non-commerce learner can understand the basic structure and "
            "the practical steps. Wherever useful, important German words are added in brackets."
        )
    )
    blocks.append(
        para(
            "Germany (Deutschland) has a detailed rule-based tax system. Taxes fund schools, "
            "hospitals, roads, police, courts, pensions, and public services."
        )
    )
    blocks.append(para("Main groups of taxes:"))
    blocks.append(bullet("Direct taxes (Direkte Steuern): charged directly on income, profit, or property."))
    blocks.append(bullet("Indirect taxes (Indirekte Steuern): charged on spending, sales, and consumption."))
    blocks.append(
        para(
            "Important note: Laws can change through new acts, court decisions, and tax-office guidance. "
            "This book is educational and practical, but not personal legal or tax advice."
        )
    )

    blocks.append(heading_2("1.1 Core institutions and key German words"))
    for x in [
        "Bundesministerium der Finanzen (BMF): Federal Ministry of Finance.",
        "Finanzamt: Local tax office.",
        "Steuer-ID (Steueridentifikationsnummer): Personal tax ID.",
        "Steuernummer: Filing-specific tax number.",
        "ELSTER: Official online filing portal.",
        "Abgabenordnung (AO): Core procedural tax code.",
        "Einkommensteuergesetz (EStG): Income Tax Act.",
        "Umsatzsteuergesetz (UStG): VAT Act.",
    ]:
        blocks.append(bullet(x))

    blocks.append(heading_2("1.2 How tax collection works"))
    for x in [
        "Taxable event happens (income, sale, transfer, import, etc.).",
        "Tax base is calculated under legal rules.",
        "Return/declaration is filed or withholding is made at source.",
        "Tax authority issues notice (Steuerbescheid).",
        "Tax is paid, refunded, or adjusted through appeals (Einspruch).",
    ]:
        blocks.append(bullet(x))

    blocks.append(heading_1("Part II — Direct Taxes in Germany (Direkte Steuern)"))
    blocks.append(
        para(
            "Direct taxes are paid by the person or company who legally owes the tax. "
            "If you earn income, own property, or receive certain transfers, you may owe direct tax."
        )
    )

    direct_taxes = [
        (
            "2.1 Personal Income Tax (Einkommensteuer)",
            [
                "Who pays: Individuals with taxable income in Germany.",
                "What is taxed: Total taxable income after deductions/allowances.",
                "Rates: Progressive (higher income can face higher rates).",
                "Filing: Annual income tax return (Einkommensteuererklärung), often via ELSTER.",
                "Simple example: Two people earning different amounts can pay different effective rates.",
                "Important words: Einkünfte, Freibetrag, Sonderausgaben, Werbungskosten.",
            ],
        ),
        (
            "2.2 Wage Tax (Lohnsteuer)",
            [
                "Who pays: Employees; employer withholds from salary.",
                "What is taxed: Salary and wages.",
                "Rates: Based on payroll tables and Steuerklasse (tax class).",
                "Filing: Employer remits monthly; annual payroll certificate is issued.",
                "Simple example: Gross salary minus withholding equals net pay.",
                "Important words: Steuerklasse, Lohnabrechnung, Lohnsteuerbescheinigung.",
            ],
        ),
        (
            "2.3 Corporate Income Tax (Körperschaftsteuer)",
            [
                "Who pays: Corporations such as GmbH and AG.",
                "What is taxed: Corporate profit after allowed expenses.",
                "Rates: Flat statutory framework, plus related surcharge effects.",
                "Filing: Annual corporate return; advance payments may apply.",
                "Important words: Gewinn, Betriebsausgaben, Steuerbilanz.",
            ],
        ),
        (
            "2.4 Trade Tax (Gewerbesteuer)",
            [
                "Who pays: Commercial businesses in municipalities.",
                "What is taxed: Trade income (Gewerbeertrag) with legal adjustments.",
                "Rates: Base amount multiplied by local Hebesatz.",
                "Filing: Return processed with municipality factors.",
                "Important words: Hebesatz, Gemeinde, Gewerbeertrag.",
            ],
        ),
        (
            "2.5 Solidarity Surcharge (Solidaritätszuschlag)",
            [
                "Nature: Surcharge linked to specific tax liabilities.",
                "Application: Depends on legal thresholds/exemptions.",
                "Important words: Zuschlag, Freigrenze, Steuerlast.",
            ],
        ),
        (
            "2.6 Church Tax (Kirchensteuer)",
            [
                "Who pays: Members of qualifying religious communities.",
                "How calculated: Percentage linked to income tax amount.",
                "Regionality: Rules vary by Bundesland.",
                "Important words: Religionsgemeinschaft, Kirchenaustritt, Bundesland.",
            ],
        ),
        (
            "2.7 Capital Income Withholding (Kapitalertragsteuer / Abgeltungsteuer)",
            [
                "Who pays: Investors receiving dividends, interest, certain gains.",
                "Mechanism: Banks commonly withhold tax at source.",
                "Adjustment: Annual filing may optimize final result.",
                "Important words: Kapitalerträge, Sparer-Pauschbetrag, Freistellungsauftrag.",
            ],
        ),
        (
            "2.8 Inheritance Tax (Erbschaftsteuer)",
            [
                "Who pays: Heirs receiving estate assets.",
                "Tax base: Asset value minus exemptions and allowances.",
                "Rates: Depend on relationship class and taxable value.",
                "Important words: Erbe, Nachlass, Steuerklasse.",
            ],
        ),
        (
            "2.9 Gift Tax (Schenkungsteuer)",
            [
                "Who pays: Recipient of taxable gifts during lifetime transfers.",
                "Tax base: Value transferred minus exemptions.",
                "Relation: Works similarly to inheritance-tax logic in many areas.",
                "Important words: Schenker, Beschenkter, Anzeigepflicht.",
            ],
        ),
        (
            "2.10 Property Tax (Grundsteuer)",
            [
                "Who pays: Property owners.",
                "Tax base: Official valuation framework and municipal multiplier logic.",
                "Filing: Declarations and notices under current valuation periods.",
                "Important words: Grundstück, Messbetrag, Grundsteuerbescheid.",
            ],
        ),
    ]
    for title, points in direct_taxes:
        blocks.append(heading_2(title))
        for p in points:
            blocks.append(bullet(p))
        blocks.append(para("Beginner tip: Keep all documents (Belege) safely and in date order."))

    blocks.append(heading_1("Part III — Indirect Taxes in Germany (Indirekte Steuern)"))
    blocks.append(
        para(
            "Indirect taxes are usually included in prices. Businesses collect the tax from "
            "customers and pass it to the tax authority."
        )
    )

    indirect_taxes = [
        (
            "3.1 Value Added Tax (Umsatzsteuer / Mehrwertsteuer)",
            [
                "Final burden: Usually carried by end consumer.",
                "Business role: Charge output VAT and claim input VAT (Vorsteuer) where eligible.",
                "Rates: Standard and reduced categories under law.",
                "Filing: Periodic VAT returns and annual return.",
                "Words: Vorsteuer, USt-IdNr., Rechnung, Umsatzsteuervoranmeldung.",
            ],
        ),
        (
            "3.2 Import VAT (Einfuhrumsatzsteuer)",
            [
                "When: Goods are imported into Germany from relevant non-EU customs context.",
                "Base: Customs value plus relevant additions.",
                "Words: Einfuhr, Zollanmeldung, Abfertigung.",
            ],
        ),
        (
            "3.3 Customs Duty (Zoll)",
            [
                "Depends on tariff classification, origin, and customs value.",
                "Different products may have different duty rates.",
                "Words: Warennummer, Ursprung, Zollsatz.",
            ],
        ),
        (
            "3.4 Insurance Tax (Versicherungsteuer)",
            [
                "Collected in insurance premiums under taxable lines.",
                "Usually handled by insurer; burden often included in total cost.",
                "Words: Versicherungsprämie, Steuerschuldner, Anmeldung.",
            ],
        ),
        (
            "3.5 Energy Tax (Energiesteuer)",
            [
                "Applied to selected energy products under excise framework.",
                "Cost can be reflected in fuel and energy prices.",
                "Words: Energieerzeugnisse, Steuerlager, Entlastung.",
            ],
        ),
        (
            "3.6 Electricity Tax (Stromsteuer)",
            [
                "Applied to electricity consumption under legal scope.",
                "Collected through supply chain, with exemptions/reliefs in specific cases.",
                "Words: Stromverbrauch, Entnahme, Befreiung.",
            ],
        ),
        (
            "3.7 Tobacco Tax (Tabaksteuer)",
            [
                "Integrated into product pricing for tobacco goods.",
                "Words: Tabakwaren, Steuerzeichen, Inverkehrbringen.",
            ],
        ),
        (
            "3.8 Alcohol and Beer Taxes (Alkoholsteuer / Biersteuer)",
            [
                "Excise tax based on product class and legal criteria.",
                "Words: Verbrauchsteuer, Steueraussetzung, Hersteller.",
            ],
        ),
        (
            "3.9 Coffee Tax (Kaffeesteuer)",
            [
                "Excise on defined coffee products.",
                "Words: Röstkaffee, Löslicher Kaffee, Steuerentstehung.",
            ],
        ),
        (
            "3.10 Air Transport Tax (Luftverkehrsteuer)",
            [
                "Linked to passenger departures and destination zones.",
                "May be reflected in ticket prices.",
                "Words: Abflug, Beförderer, Steuersatz.",
            ],
        ),
        (
            "3.11 Real Estate Transfer Tax (Grunderwerbsteuer)",
            [
                "Triggered by qualifying real estate transfers.",
                "Rates differ by federal state (Bundesland).",
                "Words: Kaufvertrag, Notar, Grundbuch.",
            ],
        ),
    ]
    for title, points in indirect_taxes:
        blocks.append(heading_2(title))
        for p in points:
            blocks.append(bullet(p))
        blocks.append(para("Beginner tip: In indirect tax, who pays legally and who bears cost can differ."))

    blocks.append(heading_1("Part IV — A to Z Rulebook (Germany Only)"))
    blocks.append(
        para(
            "This A to Z section summarizes key practical rules and concepts in plain English "
            "with important German terms."
        )
    )

    az = {
        "A": ["Abgabenordnung (AO): Core procedure code.", "Aufbewahrungspflicht: Keep records for legal periods.", "Anmeldung: Tax declaration filing."],
        "B": ["Betriebsausgaben: Business expenses.", "Bescheid: Official tax notice.", "BZSt: Federal central tax authority roles."],
        "C": ["Compliance matters: file, pay, and document on time.", "Correct invoices are central for VAT rights.", "Cross-border cases need extra checks."],
        "D": ["DBA (tax treaties) reduce double taxation.", "Deadlines (Fristen) must be tracked carefully.", "Documentation quality protects during audits."],
        "E": ["Einkünftearten: income categories.", "ELSTER: e-filing portal.", "Einspruch: appeal against tax notice."],
        "F": ["Finanzamt: local authority.", "Freibetrag: allowance.", "Fristversäumnis can cause penalties."],
        "G": ["Gewerbesteuer: trade tax.", "Grundsteuer: property tax.", "Grunderwerbsteuer: transfer tax on property deals."],
        "H": ["Hebesatz changes local burdens.", "Haushaltsnahe Dienstleistungen may produce relief.", "Hinzurechnungen can affect trade tax base."],
        "I": ["Input VAT (Vorsteuer) needs proper invoice.", "International residency drives tax scope.", "Invoice integrity is essential."],
        "J": ["Jahreserklärung: annual return.", "Job expenses (Werbungskosten) may reduce tax.", "Joint assessment can apply for spouses."],
        "K": ["Kirchensteuer: church tax.", "Kapitalertragsteuer: investment withholding.", "Körperschaftsteuer: corporate tax."],
        "L": ["Lohnsteuer withheld by employer.", "Lohnsteuerklasse influences withholding.", "Late filing can trigger Zuschläge."],
        "M": ["Mehrwertsteuer = VAT.", "Mitwirkungspflichten: duty to cooperate.", "Meldepflichten: reporting duties."],
        "N": ["Nachzahlung: extra payment due.", "Nebenleistungen: ancillary charges.", "NV-Bescheinigung may apply in specific cases."],
        "O": ["Ordnungsmäßige bookkeeping is mandatory.", "Online filing is standard.", "Offsetting may apply to overpayments."],
        "P": ["Progression increases rate with higher income.", "Prüfung (audit) can review records.", "Property taxes include annual and transfer taxes."],
        "Q": ["Quality records reduce risk.", "Quick response to authority letters helps.", "Quick correction is often safer than delay."],
        "R": ["Rechnung rules drive VAT compliance.", "Rechtsbehelf means legal remedy.", "Rückstellungen can influence taxable profit."],
        "S": ["Solidaritätszuschlag as surcharge.", "Schenkungsteuer on gifts.", "Steuerhinterziehung is a serious offense."],
        "T": ["Tax residency status is fundamental.", "Trade and corporate tax may both apply.", "Timely payment avoids surcharge."],
        "U": ["Umsatzsteuer-ID for intra-EU transactions.", "USt advance returns may be periodic.", "Unclear cases should be documented deeply."],
        "V": ["Veranlagung: assessment process.", "Vorauszahlungen: advance payments.", "Verjährung: limitation periods."],
        "W": ["Werbungskosten reduce employment income.", "Withholding taxes may be creditable.", "Wohnsitz affects residency status."],
        "X": ["X-factor cases: digital/cross-border need extra care.", "XML is common in digital tax exchange.", "eXact recordkeeping is protective."],
        "Y": ["Year-end planning improves cash flow.", "Yield taxation often includes withholding.", "Your reference numbers must match filings."],
        "Z": ["Zinsen can apply in late-payment contexts.", "Zoll applies to imports.", "Zusammenveranlagung: joint assessment option."],
    }

    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        blocks.append(heading_2(f"4.{letter} — {letter} Rules"))
        for item in az[letter]:
            blocks.append(bullet(item))
        blocks.append(para("Simple takeaway: know the term, know the form, know the deadline, keep records."))

    blocks.append(heading_1("Part V — Easy Summary for Beginners"))
    for x in [
        "Step 1: Identify tax type (direct or indirect).",
        "Step 2: Identify who files/pays.",
        "Step 3: Identify what amount is taxed.",
        "Step 4: File correctly and on time.",
        "Step 5: Keep documents and notices safely.",
    ]:
        blocks.append(bullet(x))
    blocks.append(
        para(
            "Final note: This book is only about Germany’s tax system and uses easy English with "
            "important German words. For real personal/business decisions, always verify the latest law."
        )
    )
    return blocks


def document_xml(paragraphs: list[str]) -> str:
    body = "".join(paragraphs)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{body}<w:sectPr/></w:body>"
        "</w:document>"
    )


CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
"""

ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOC_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
"""

NUMBERING_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
"""


def generate_book(output_path: str) -> None:
    paragraphs = build_book_blocks()
    doc_xml = document_xml(paragraphs)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        zf.writestr("_rels/.rels", ROOT_RELS_XML)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS_XML)
        zf.writestr("word/numbering.xml", NUMBERING_XML)


if __name__ == "__main__":
    out = "German_Tax_System_Full_Length_Book_EN.docx"
    generate_book(out)
    print(f"Generated: {out}")
