"""DOCX 보고서 자동 점검. 사용: python check_docx.py report.docx  (실패 항목이 있으면 종료코드 1)"""
import re, sys, zipfile
from docx import Document
from docx.shared import Pt


def main(path):
    doc = Document(path)
    z = zipfile.ZipFile(path)
    body = z.read("word/document.xml").decode("utf8")
    parts = {n: z.read(n).decode("utf8") for n in z.namelist() if n.startswith("word/") and n.endswith(".xml")}
    rels = z.read("word/_rels/document.xml.rels").decode("utf8")
    res = []

    def chk(name, ok, detail=""):
        res.append(ok)
        print("PASS" if ok else "FAIL", name, detail)

    paras = doc.paragraphs
    heads = [p for p in paras if p.style.name.startswith("Heading")]
    chk("Heading 스타일 사용", len(heads) > 0, f"{len(heads)}개")
    h1 = [p.text for p in heads if p.style.name == "Heading 1"]
    order = [next((i for i, t in enumerate(h1) if k in t), None) for k in ("서론", "본론", "결론", "참고문헌")]
    chk("구조 순서(서론-본론-결론-참고문헌)", None not in order and order == sorted(order), str(h1))
    chk("목차(TOC) 필드", bool(re.search(r"TOC\s+\\", body)))
    m = re.search(r"TOC\s+\\.*?fldCharType=\"end\"", body, re.S)
    toc_text = re.sub(r"<[^>]+>", "", m.group(0)) if m else ""
    missing = [p.text for p in heads if p.text not in toc_text]
    chk("목차 항목 = 실제 제목", bool(m) and not missing, f"누락 {missing}" if missing else "")
    links = re.findall(r'Target="(https?://[^"]+)"', rels)
    hl = len(re.findall(r"<w:hyperlink\b", body))
    chk("하이퍼링크(실제 w:hyperlink)", hl > 0 and len(links) > 0, f"{hl}개")
    bad = [u for u in links if not re.match(r"https?://[^\s/]+\.[^\s]+", u)]
    chk("URL 형식", not bad, str(bad) if bad else "")
    i = next((k for k, p in enumerate(paras) if p.style.name == "Heading 1" and "참고문헌" in p.text), None)
    refs = [p for p in (paras[i + 1:] if i is not None else []) if p.text.strip()]
    ok = bool(refs) and all("접속일" in p.text and re.search(r"https?://", p.text) for p in refs)
    chk("참고문헌 항목(제목·기관·접속일·URL)", ok, f"{len(refs)}건")
    foot = [v for k, v in parts.items() if "footer" in k]
    chk("쪽 번호 필드(푸터)", any("PAGE" in f for f in foot))
    chk("표지 분리(섹션 2개 이상)", len(re.findall(r"<w:sectPr", body)) >= 2)
    s0 = doc.sections[0]
    chk("표지 푸터에 쪽 번호 없음", "PAGE" not in s0.footer._element.xml)
    caps = [p.text for p in paras if re.match(r"^(표|그림)\s*\d+\.", p.text.strip())]
    n_obj = len(doc.tables) + body.count("<w:drawing")
    chk("표·그림 캡션", len(caps) == n_obj, f"캡션 {len(caps)} / 개체 {n_obj}")
    n = doc.styles["Normal"]
    chk("본문 11pt", n.font.size == Pt(11))
    chk("줄간격 160%", n.paragraph_format.line_spacing == 1.6)
    styles = parts.get("word/styles.xml", "")
    used = set(re.findall(r'<w:(?:pStyle|rStyle|tblStyle) w:val="([^"]+)"', body)) | {"Normal"}
    used_styles = "".join(x for x in re.findall(r"<w:style .*?</w:style>", styles, re.S)
                          if re.search(r'w:styleId="([^"]+)"', x).group(1) in used)
    fonts = set(re.findall(r'<w:rFonts[^>]*w:eastAsia="([^"]+)"', used_styles + body))
    chk("글꼴 1종 통일", len(fonts) == 1, ", ".join(sorted(fonts)))
    colors = {c.upper() for c in re.findall(r'w:color w:val="([0-9A-Fa-f]{6})"', body + used_styles + "".join(foot))}
    chk("색 2가지 이내(검정 제외)", len(colors - {"000000"}) <= 2, ", ".join(sorted(colors)))
    txt = "\n".join(p.text for p in paras)
    print(f"INFO '확인 필요' 표시 {txt.count('확인 필요')}건 -> 사용자에게 목록으로 알릴 것")
    sys.exit(0 if all(res) else 1)


if __name__ == "__main__":
    main(sys.argv[1])
