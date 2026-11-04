"""Build js-ugd-template.docx, the Word version of template.tex. Needs: pip install python-docx."""
import datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
FONT = "Arial"
M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
ID_HEAD, ID_BULLET, ID_REF = 90, 91, 92


def set_font(style, size, bold=None, italic=None):
    rpr = style.element.get_or_add_rPr()
    for tag in ("w:rFonts", "w:color"):
        for old in rpr.findall(qn(tag)):
            rpr.remove(old)
    rf = OxmlElement("w:rFonts")
    for a in ("ascii", "hAnsi", "cs", "eastAsia"):
        rf.set(qn("w:" + a), FONT)
    rpr.insert(0, rf)
    style.font.size, style.font.bold, style.font.italic = Pt(size), bold, italic
    style.font.color.rgb = RGBColor(0, 0, 0)


def para_style(doc, name, size, before, after, align, bold=None, italic=None, keep=False, base="Normal"):
    try:
        st = doc.styles[name]
    except KeyError:
        st = doc.styles.add_style(name, 1)
    st.base_style = doc.styles[base] if name != "Normal" else None
    st.quick_style, st.hidden, st.unhide_when_used = True, False, False
    set_font(st, size, bold, italic)
    pf = st.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(before), Pt(after), 1.0
    pf.alignment, pf.first_line_indent, pf.left_indent, pf.keep_with_next = align, None, None, keep
    return st


def link_numbering(style, num_id, ilvl=0):
    numpr = style.element.get_or_add_pPr().get_or_add_numPr()
    numpr.get_or_add_ilvl().val = ilvl
    numpr.get_or_add_numId().val = num_id


def setup_numbering(doc):
    """Automatic numbers for headings and references, and the bullet at 0.7 cm with the text at 1.2 cm."""
    np = doc.part.numbering_part.element
    w = nsdecls("w")
    abstract = [
        f'''<w:abstractNum {w} w:abstractNumId="{ID_HEAD}"><w:multiLevelType w:val="multilevel"/>
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/><w:suff w:val="space"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>
        <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading2"/><w:suff w:val="space"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl>
        </w:abstractNum>''',
        f'''<w:abstractNum {w} w:abstractNumId="{ID_BULLET}"><w:multiLevelType w:val="hybridMultilevel"/>
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:pStyle w:val="JEBullet"/><w:lvlText w:val="&#8226;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="680" w:hanging="283"/></w:pPr><w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}"/></w:rPr></w:lvl>
        </w:abstractNum>''',
        f'''<w:abstractNum {w} w:abstractNumId="{ID_REF}"><w:multiLevelType w:val="hybridMultilevel"/>
        <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="JEReference"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="284" w:hanging="284"/></w:pPr><w:rPr><w:b/></w:rPr></w:lvl>
        </w:abstractNum>''',
    ]
    for i, x in enumerate(abstract):
        np.insert(i, parse_xml(x))
    for n in (ID_HEAD, ID_BULLET, ID_REF):
        np.append(parse_xml(f'<w:num {w} w:numId="{n}"><w:abstractNumId w:val="{n}"/></w:num>'))


def setup_styles(doc):
    rf = doc.styles.element.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr")).find(qn("w:rFonts"))
    for k in list(rf.attrib):
        del rf.attrib[k]
    for a in ("ascii", "hAnsi", "cs", "eastAsia"):
        rf.set(qn("w:" + a), FONT)
    para_style(doc, "Normal", 11, 0, 2, AL.JUSTIFY)
    para_style(doc, "Heading 1", 14, 20, 10, AL.LEFT, bold=True, italic=False, keep=True)
    para_style(doc, "Heading 2", 12, 20, 10, AL.LEFT, bold=True, italic=True, keep=True)
    link_numbering(doc.styles["Heading 1"], ID_HEAD, 0)
    link_numbering(doc.styles["Heading 2"], ID_HEAD, 1)
    st = para_style(doc, "JE Heading Unnumbered", 14, 20, 10, AL.LEFT, bold=True, italic=False, keep=True, base="Heading 1")
    link_numbering(st, 0, 0)
    para_style(doc, "JE Title", 20, 0, 15, AL.LEFT, bold=True)
    para_style(doc, "JE Authors", 16, 0, 10, AL.LEFT, italic=True)
    para_style(doc, "JE Affiliation", 11, 0, 10, AL.LEFT, italic=True)
    para_style(doc, "JE Abstract", 11, 20, 10, AL.JUSTIFY, bold=True)
    para_style(doc, "JE Keywords Heading", 12, 20, 2, AL.LEFT, bold=True, keep=True)
    para_style(doc, "JE Table Caption", 10, 10, 10, AL.CENTER, keep=True)
    para_style(doc, "JE Table Cell", 10, 0, 0, AL.LEFT)
    para_style(doc, "JE Table Note", 8.5, 3, 8, AL.JUSTIFY)
    para_style(doc, "JE Figure", 11, 6, 0, AL.CENTER, keep=True)
    para_style(doc, "JE Figure Caption", 10, 5, 10, AL.CENTER)
    link_numbering(para_style(doc, "JE Bullet", 11, 0, 2, AL.LEFT), ID_BULLET)
    link_numbering(para_style(doc, "JE Reference", 10, 0, 0, AL.LEFT), ID_REF)
    para_style(doc, "Footer", 11, 0, 0, AL.CENTER, bold=True)
    # Only horizontal rules, centred
    doc.styles.element.append(parse_xml(
        f'''<w:style {nsdecls("w")} w:type="table" w:customStyle="1" w:styleId="JETable"><w:name w:val="JE Table"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="99"/><w:qFormat/>
        <w:tblPr><w:jc w:val="center"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr></w:style>'''))


def page_setup(doc):
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Cm(2.5)
    s.header_distance, s.footer_distance, s.gutter = Cm(1.25), Cm(1), 0
    fp = s.footer.paragraphs[0]
    fp.style = doc.styles["Footer"]
    w = nsdecls("w")
    for x in (f'<w:fldChar {w} w:fldCharType="begin"/>', f'<w:instrText {w} xml:space="preserve"> PAGE </w:instrText>', f'<w:fldChar {w} w:fldCharType="end"/>'):
        fp.add_run()._r.append(parse_xml(x))


def link(p, url, text, italic=None):
    rid = p.part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), rid)
    i = "<w:i/>" if italic else ""
    h.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr>{i}<w:u w:val="single"/></w:rPr><w:t>{text}</w:t></w:r>'))
    p._p.append(h)


def seq(p, kind, n):
    """Bold 'Table 1' or 'Figure 1' with a number Word updates by itself."""
    p.add_run(kind + " ").bold = True
    p._p.append(parse_xml(
        f'<w:fldSimple {nsdecls("w")} w:instr=" SEQ {kind} \\* ARABIC "><w:r><w:rPr><w:b/></w:rPr><w:t>{n}</w:t></w:r></w:fldSimple>'))
    p.add_run(" ")


def rich(doc, parts, style="Normal"):
    """parts: strings, or ('i', text) for italics."""
    p = doc.add_paragraph(style=style)
    for x in parts if isinstance(parts, list) else [parts]:
        if isinstance(x, tuple):
            p.add_run(x[1]).italic = True
        else:
            p.add_run(x)
    return p


def clean_table(tbl, widths, style=None):
    pr = tbl._tbl.tblPr
    for el in pr.findall(qn("w:tblStyle")):
        pr.remove(el)
    if style:
        pr.insert(0, parse_xml(f'<w:tblStyle {nsdecls("w")} w:val="{style}"/>'))
    for col, wd in zip(tbl._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        col.set(qn("w:w"), str(wd))
    for row in tbl.rows:
        for cell, wd in zip(row.cells, widths):
            tcpr = cell._tc.get_or_add_tcPr()
            for el in tcpr.findall(qn("w:tcW")):
                tcpr.remove(el)
            tcpr.insert(0, parse_xml(f'<w:tcW {nsdecls("w")} w:w="{wd}" w:type="dxa"/>'))


def margins_table(doc):
    rows = [("Margin", "Cm"), ("Top", "2.5"), ("Bottom", "2.5"), ("Left", "2.5"), ("Right", "2.5"), ("Header", "1.25"), ("Footer", "1")]
    tbl = doc.add_table(rows=len(rows), cols=2)
    clean_table(tbl, [2381, 2381], "JETable")
    look = tbl._tbl.tblPr.find(qn("w:tblLook"))
    jc = parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>')
    look.addprevious(jc) if look is not None else tbl._tbl.tblPr.append(jc)
    for i, row in enumerate(rows):
        trpr = tbl.rows[i]._tr.get_or_add_trPr()
        trpr.append(OxmlElement("w:cantSplit"))
        if i == 0:
            trpr.append(OxmlElement("w:tblHeader"))
        for j, t in enumerate(row):
            p = tbl.cell(i, j).paragraphs[0]
            p.style = doc.styles["JE Table Cell"]
            p.add_run(t).bold = True if i == 0 else None
            p.paragraph_format.keep_with_next = i < len(rows) - 1
    return tbl


def equation(doc, number):
    """Centred equation with its number at the right, in a borderless layout table."""
    tbl = doc.add_table(rows=1, cols=2)
    clean_table(tbl, [8220, 850])
    left, right = tbl.rows[0].cells
    p = left.paragraphs[0]
    rpr = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:color w:val="000000"/></w:rPr>'
    def mrun(t):
        sty = '<m:rPr><m:sty m:val="p"/></m:rPr>' if t in "=+" else ""
        return f"<m:r>{sty}{rpr}<m:t>{t}</m:t></m:r>"
    runs = "".join(mrun(t) for t in ("Y", "=", "\u03b1", "+", "\u03b2", "X", "+", "\u03b5"))
    p._p.append(parse_xml(
        f'<m:oMathPara {M} {nsdecls("w")}><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr><m:oMath>{runs}</m:oMath></m:oMathPara>'))
    q = right.paragraphs[0]
    q.alignment = AL.RIGHT
    q.add_run(f"({number})")
    for cell in (left, right):
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))


def reference(doc, parts, url=None):
    p = rich(doc, parts, "JE Reference")
    if url:
        link(p, url, url)


def build():
    doc = Document()
    for el in list(doc.element.body):
        if el.tag != qn("w:sectPr"):
            doc.element.body.remove(el)
    setup_numbering(doc)
    setup_styles(doc)
    page_setup(doc)

    # Word opens the file in its current mode, not Compatibility Mode
    for cs in doc.settings.element.iter(qn("w:compatSetting")):
        if cs.get(qn("w:name")) == "compatibilityMode":
            cs.set(qn("w:val"), "15")

    def sup(p, t):
        p.add_run(t).font.superscript = True

    doc.add_paragraph("Title of the Paper in Title Case", style="JE Title")
    a = doc.add_paragraph(style="JE Authors")
    a.add_run("Firstname Lastname"); sup(a, "1"); a.add_run(", Firstname Lastname"); sup(a, "2")
    for n, mail in (("1", "first.author@example.com"), ("2", "second.author@example.com")):
        a = doc.add_paragraph(style="JE Affiliation")
        sup(a, n)
        a.add_run("Faculty, University, Street and number, City, Country, ")
        link(a, "mailto:" + mail, mail, italic=True)
    doc.add_paragraph("Replace this text with the abstract. The abstract should be at least 250 and at most 500 words, including up to five keywords and keyphrases, and must fit on the first page of the paper.", style="JE Abstract")
    doc.add_paragraph("Keywords", style="JE Keywords Heading")
    doc.add_paragraph("First keyword, Second keyword, Third keyword, Fourth keyword, Fifth keyword")

    H1, H2 = "Heading 1", "Heading 2"
    doc.add_paragraph("Introduction", style=H1)
    doc.add_paragraph("Replace this text with your paper. Papers are submitted electronically through the journal’s system, in Microsoft Word and PDF format. They should be from 8 to 12 pages long.")

    doc.add_paragraph("Formatting Guidelines", style=H1)
    doc.add_paragraph("Number sections 1, 2, 3 and subsections 1.1, 1.2, 1.3. Use a capital for the initial letter of every word in section and subsection titles. In this file the numbers are automatic: use the styles Heading 1 and Heading 2. All the formats are saved as styles starting with “JE” in the Styles gallery.")

    doc.add_paragraph("Text", style=H2)
    rich(doc, ["Write paragraphs in Arial 11, single-spaced, with no indent and fully justified. For emphasis on words or key terms, use italics, for example ", ("i", "this is an emphasised text"), ". Do not use bold, underlining or another typeface."])

    doc.add_paragraph("Tables", style=H2)
    doc.add_paragraph("Tables contain only horizontal rules. Each table needs a short descriptive title above it, column headings that clearly define the data, and all units of measurement. Put any footnotes below the table, and cite every table in the text. Number tables 1, 2, 3 and centre them with their captions. Place them in the text at the appropriate point, not at the end. For example, the margins of the page are set as in Table 1.")
    p = doc.add_paragraph(style="JE Table Caption")
    seq(p, "Table", 1)
    p.add_run("Setup of margins for the paper")
    margins_table(doc)

    doc.add_paragraph("Figures", style=H2)
    doc.add_paragraph("Centre figures and images, and place the caption after them. Number them 1, 2, 3, as in Figure 1.")
    p = doc.add_paragraph(style="JE Figure")
    p.add_run().add_picture(str(ROOT / "figure-placeholder.png"), width=Cm(13.6))
    for dp in p._p.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"):
        dp.set("descr", "Placeholder: replace with your figure"); dp.set("name", "Figure 1")
    p = doc.add_paragraph(style="JE Figure Caption")
    seq(p, "Figure", 1)
    p.add_run("Title of the figure")

    doc.add_paragraph("Bullets", style=H2)
    doc.add_paragraph("A bullet list looks as follows:")
    doc.add_paragraph("Bullet item one, some text;", style="JE Bullet")
    doc.add_paragraph("Bullet item two, some other text.", style="JE Bullet")

    doc.add_paragraph("Equations", style=H2)
    doc.add_paragraph("Centre equations and number them at the right, as in equation (1):")
    equation(doc, 1)

    doc.add_paragraph("Abbreviations", style=H2)
    doc.add_paragraph("As far as possible, avoid initials, except for terms in common use. Spell abbreviations out the first time they are mentioned, with the abbreviation in brackets, as in Goce Delcev University (UGD). List them in alphabetical order before the references.")

    doc.add_paragraph("Citations and References", style=H2)
    doc.add_paragraph("Use APA style in the text. A parenthetical citation has the surname and the year: (Taylor, 1993) for one author, (Diebold & Li, 2006) for two, and (Clarida et al., 2000) for three or more. A narrative citation has the surname and the year in brackets: Taylor (1993).")
    doc.add_paragraph("In the reference list, references are numbered, listed in APA style and alphabetized by the last name of the first author. Every reference in the text must be in the list, with its URL where available. The list below has a journal article (Veselinova & Petrov, 2025), a book (Hamilton, 1994), a working paper (Svensson, 1994) and a webpage (Journal of Economics, n.d.).")

    doc.add_paragraph("Conclusions", style=H1)
    doc.add_paragraph("Replace this text with your conclusions.")

    doc.add_paragraph("Abbreviations", style="JE Heading Unnumbered")
    doc.add_paragraph("APA: American Psychological Association; UGD: Goce Delcev University.")

    doc.add_paragraph("References", style="JE Heading Unnumbered")
    i = lambda t: ("i", t)
    reference(doc, ["Clarida, R., Galí, J., & Gertler, M. (2000). Monetary policy rules and macroeconomic stability: Evidence and some theory. ", i("Quarterly Journal of Economics"), ", ", i("115"), "(1), 147–180. "], "https://doi.org/10.1162/003355300554692")
    reference(doc, ["Diebold, F. X., & Li, C. (2006). Forecasting the term structure of government bond yields. ", i("Journal of Econometrics"), ", ", i("130"), "(2), 337–364. "], "https://doi.org/10.1016/j.jeconom.2005.03.005")
    reference(doc, ["Hamilton, J. D. (1994). ", i("Time series analysis"), ". Princeton University Press."])
    reference(doc, ["Journal of Economics. (n.d.). ", i("Submissions"), ". "], "https://js.ugd.edu.mk/index.php/JE/about/submissions")
    reference(doc, ["Svensson, L. E. O. (1994). ", i("Estimating and interpreting forward interest rates: Sweden 1992–1994"), " (NBER Working Paper No. 4871). National Bureau of Economic Research. "], "https://doi.org/10.3386/w4871")
    reference(doc, ["Taylor, J. B. (1993). Discretion versus policy rules in practice. ", i("Carnegie-Rochester Conference Series on Public Policy"), ", ", i("39"), ", 195–214. "], "https://doi.org/10.1016/0167-2231(93)90009-L")
    reference(doc, ["Veselinova, E., & Petrov, R. (2025). Is MBI10 the suitable index for the Macedonian stock exchange – an overview of MSE. ", i("Journal of Economics"), ", ", i("10"), "(1), 12–34. "], "https://doi.org/10.46763/JOE2510112v")

    cp = doc.core_properties
    cp.title, cp.author, cp.last_modified_by, cp.comments, cp.subject, cp.keywords = "Journal of Economics (UGD): paper template", "", "", "", "", ""
    cp.created = cp.modified = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None, microsecond=0)
    doc.save(ROOT / "js-ugd-template.docx")


if __name__ == "__main__":
    build()
