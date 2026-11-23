"""JSON 명세 -> 규칙을 만족하는 DOCX 보고서.
사용: python build_docx.py spec.json output.docx
규칙(표지 분리, 목차 필드, Heading 스타일, 쪽 번호, 하이퍼링크, 캡션, 글꼴, 줄간격, 2색)을 코드로 강제한다.
"""
import json, re, sys, datetime
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor

NAVY = RGBColor(0x1F, 0x38, 0x64)
GRAY = RGBColor(0x59, 0x59, 0x59)
INLINE = re.compile(r"(\[[^\]]+\]\(https?://[^)\s]+\)|\*\*[^*]+\*\*)")
LINK = re.compile(r"\[([^\]]+)\]\((https?://[^)\s]+)\)")


def set_font(el_with_rpr, name):
    rpr = el_with_rpr.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for k in list(rf.attrib):
        del rf.attrib[k]
    for k in ("ascii", "hAnsi", "eastAsia", "cs"):
        rf.set(qn("w:" + k), name)


def style_font(style, name, size, color=None, bold=None):
    style.font.name = name
    set_font(style.element, name)
    style.font.size = Pt(size)
    if color is not None:
        style.font.color.rgb = color
    if bold is not None:
        style.font.bold = bold


def run_font(run, name):
    run.font.name = name
    set_font(run._r, name)


def add_field(par, instr, result, font):
    """복합 필드(begin/instr/separate/result/end). 결과 run을 반환."""
    def mk(kind):
        r = par.add_run()
        run_font(r, font)
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), kind)
        r._r.append(fc)
    mk("begin")
    r = par.add_run()
    run_font(r, font)
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r._r.append(it)
    mk("separate")
    rr = par.add_run(result)
    run_font(rr, font)
    mk("end")
    return rr


def add_hyperlink(par, text, url, font):
    rid = par.part.relate_to(url, RT.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), rid)
    h.set(qn("w:history"), "1")
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    for k in ("ascii", "hAnsi", "eastAsia", "cs"):
        rf.set(qn("w:" + k), font)
    col = OxmlElement("w:color")
    col.set(qn("w:val"), "1F3864")
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    for e in (rf, col, u):
        rpr.append(e)
    r.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    h.append(r)
    par._p.append(h)


def add_rich(par, text, font, size=None):
    """[텍스트](URL) -> 하이퍼링크, **굵게** -> bold."""
    for tok in INLINE.split(text):
        if not tok:
            continue
        m = LINK.fullmatch(tok)
        if m:
            add_hyperlink(par, m.group(1), m.group(2), font)
            continue
        bold = tok.startswith("**") and tok.endswith("**")
        r = par.add_run(tok[2:-2] if bold else tok)
        r.bold = bold or None
        run_font(r, font)
        if size:
            r.font.size = Pt(size)


def caption(doc, label, n, text, font, above):
    p = doc.add_paragraph(style="Caption")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(label + " ")
    run_font(r, font)
    add_field(p, f"SEQ {label} \\* ARABIC", str(n), font)
    r = p.add_run(". " + text)
    run_font(r, font)
    p.paragraph_format.keep_with_next = above
    return p


def page_number_footer(section, font):
    section.footer.is_linked_to_previous = False
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_field(p, "PAGE", "1", font)
    pg = OxmlElement("w:pgNumType")
    pg.set(qn("w:start"), "1")
    section._sectPr.append(pg)


def validate(spec):
    errs = []
    for k in ("title", "author", "blocks"):
        if not spec.get(k):
            errs.append(f"필수 항목 누락: {k}")
    h1 = [b["text"] for b in spec.get("blocks", []) if b.get("type") == "h1"]
    pos = []
    for need in ("서론", "본론", "결론"):
        idx = next((i for i, t in enumerate(h1) if need in t), None)
        if idx is None:
            errs.append(f"Heading 1 '{need}' 없음")
        pos.append(idx)
    if None not in pos and pos != sorted(pos):
        errs.append("서론 -> 본론 -> 결론 순서가 아님")
    refs = spec.get("references", [])
    if not refs:
        errs.append("references가 비어 있음 (출처 없는 보고서는 만들지 않는다)")
    for i, r in enumerate(refs, 1):
        for k in ("title", "org", "accessed", "url"):
            if not r.get(k):
                errs.append(f"references[{i}] '{k}' 누락")
    ref_urls = {r.get("url") for r in refs}
    for b in spec.get("blocks", []):
        texts = [b.get("text", ""), b.get("note", "")] + list(b.get("items", []))
        for row in b.get("rows", []):
            texts += [str(c) for c in row]
        for field in texts:
            for _, u in LINK.findall(field):
                if u not in ref_urls:
                    errs.append(f"본문 링크가 참고문헌에 없음: {u}")
    return errs


def build(spec, out):
    font = spec.get("font", "맑은 고딕")
    date = spec.get("date") or datetime.date.today().isoformat()
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.5)
    sec.top_margin = sec.bottom_margin = Cm(2.5)

    st = doc.styles
    style_font(st["Normal"], font, 11)
    st["Normal"].paragraph_format.line_spacing = 1.6
    st["Normal"].paragraph_format.space_after = Pt(4)
    for name, size in (("Heading 1", 18), ("Heading 2", 14), ("Heading 3", 12)):
        style_font(st[name], font, size, NAVY, True)
        st[name].paragraph_format.line_spacing = 1.6
        st[name].paragraph_format.space_before = Pt(12)
        st[name].paragraph_format.space_after = Pt(6)
    ppr = st["Normal"].element.get_or_add_pPr()
    for tag in ("w:autoSpaceDE", "w:autoSpaceDN"):  # 한글과 영문·숫자 사이 자동 간격 끄기
        e = OxmlElement(tag)
        e.set(qn("w:val"), "0")
        ppr.append(e)
    from docx.enum.style import WD_STYLE_TYPE
    for lvl, indent in ((1, 0), (2, 0.6), (3, 1.2)):  # 목차 줄간격을 좁게(본문 160%를 상속하면 목차가 2쪽이 된다)
        ts = st.add_style(f"toc {lvl}", WD_STYLE_TYPE.PARAGRAPH)
        ts.element.set(qn("w:styleId"), f"TOC{lvl}")
        ts.element.attrib.pop(qn("w:customStyle"), None)  # 내장 스타일로 인식시켜야 Word 목차가 이 서식을 쓴다  # Word 내장 목차 스타일 ID와 일치시켜야 적용된다
        ts.base_style = st["Normal"]
        ts.paragraph_format.line_spacing = 1.2
        ts.paragraph_format.space_before = Pt(0)
        ts.paragraph_format.space_after = Pt(2)
        ts.paragraph_format.left_indent = Cm(indent)
        ts.font.size = Pt(10.5)
        ts.font.bold = True if lvl == 1 else None
    style_font(st["Caption"], font, 10, GRAY, True)
    st["Caption"].paragraph_format.line_spacing = 1.2
    style_font(st["List Bullet"], font, 11)

    # 표지 (Heading 스타일 미사용 -> 목차에 안 들어감).
    # 빈 문단 대신 space_before를 쓴다(빈 문단은 줄간격 160%로 늘어나 표지가 2쪽으로 넘칠 수 있다).
    def cover_par(text, size, color, bold=False, before=0):
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.line_spacing = 1.0
        cp.paragraph_format.space_before = Pt(before)
        cp.paragraph_format.space_after = Pt(6)
        r = cp.add_run(text)
        run_font(r, font)
        r.font.size = Pt(size)
        r.bold = bold or None
        r.font.color.rgb = color
    cover_par(spec["title"], 24, NAVY, True, before=230)
    if spec.get("subtitle"):
        cover_par(spec["subtitle"], 14, GRAY, before=14)
    cover_par(f"작성자: {spec['author']}", 12, GRAY, before=200)
    cover_par(f"작성일: {date}", 12, GRAY)

    # 본문 섹션: 표지와 분리. 쪽 번호는 여기서부터 1쪽으로 시작. 표지 푸터는 비움.
    body_sec = doc.add_section(WD_SECTION.NEW_PAGE)
    page_number_footer(body_sec, font)

    # 목차
    p = doc.add_paragraph()
    r = p.add_run("목차")
    run_font(r, font)
    r.font.size = Pt(18)
    r.bold = True
    r.font.color.rgb = NAVY
    heads = [b["text"] for b in spec["blocks"] if b["type"] in ("h1", "h2", "h3")] + ["참고문헌"]
    toc = doc.add_paragraph()
    res_run = add_field(toc, 'TOC \\o "1-3" \\h \\z \\u', "  /  ".join(heads), font)
    toc.add_run().add_break(WD_BREAK.PAGE)

    t_no = f_no = 0
    for b in spec["blocks"]:
        ty = b["type"]
        if ty in ("h1", "h2", "h3"):
            h = doc.add_heading(b["text"], level=int(ty[1]))
            for r_ in h.runs:
                run_font(r_, font)
        elif ty == "p":
            add_rich(doc.add_paragraph(), b["text"], font)
        elif ty == "bullets":
            for it in b["items"]:
                add_rich(doc.add_paragraph(style="List Bullet"), it, font)
        elif ty == "table":
            t_no += 1
            caption(doc, "표", t_no, b["caption"], font, above=True)
            hdr, rows = b["header"], b["rows"]
            tb = doc.add_table(rows=1 + len(rows), cols=len(hdr))
            tb.style = "Table Grid"
            tb.alignment = WD_TABLE_ALIGNMENT.CENTER
            for ri, row in enumerate([hdr] + rows):
                for ci, val in enumerate(row):
                    par = tb.cell(ri, ci).paragraphs[0]
                    par.paragraph_format.line_spacing = 1.2
                    par.paragraph_format.space_after = Pt(2)
                    add_rich(par, str(val), font, 10)
                    if ri == 0:
                        for r_ in par.runs:
                            r_.bold = True
                            r_.font.color.rgb = NAVY
            for ri, row_ in enumerate(tb.rows):  # 행이 쪽 사이에서 잘리지 않게, 머리글 행은 쪽마다 반복
                trpr = row_._tr.get_or_add_trPr()
                trpr.append(OxmlElement("w:cantSplit"))
                if ri == 0:
                    trpr.append(OxmlElement("w:tblHeader"))
            note = doc.add_paragraph()
            if b.get("note"):
                add_rich(note, b["note"], font, 9)
        elif ty == "figure":
            f_no += 1
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.keep_with_next = True
            p.add_run().add_picture(b["path"], width=Cm(b.get("width_cm", 14)))
            caption(doc, "그림", f_no, b["caption"], font, above=False)
        else:
            raise ValueError(f"알 수 없는 블록: {ty}")

    # 참고문헌 (자동 생성)
    h = doc.add_heading("참고문헌", level=1)
    for r_ in h.runs:
        run_font(r_, font)
    for i, ref in enumerate(spec["references"], 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.first_line_indent = Cm(-0.8)
        r = p.add_run(f"[{i}] {ref['title']}. {ref['org']}. 접속일 {ref['accessed']}. ")
        run_font(r, font)
        add_hyperlink(p, ref["url"], ref["url"], font)

    # Word가 열 때 목차(쪽 번호 포함)를 갱신하도록 표시
    uf = OxmlElement("w:updateFields")
    uf.set(qn("w:val"), "true")
    doc.settings.element.append(uf)
    doc.core_properties.title = spec["title"]
    doc.core_properties.author = spec["author"]
    doc.save(out)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("사용: python build_docx.py spec.json output.docx")
    with open(sys.argv[1], encoding="utf-8-sig") as f:
        spec = json.load(f)
    errs = validate(spec)
    if errs:
        print("명세 오류:")
        for e in errs:
            print(" -", e)
        sys.exit(1)
    build(spec, sys.argv[2])
    print("생성:", sys.argv[2])
