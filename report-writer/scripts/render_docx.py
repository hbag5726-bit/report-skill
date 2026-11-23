"""DOCX를 Word로 열어 필드(목차·쪽 번호)를 갱신하고 PDF/PNG로 변환한 뒤 쪽 단위로 점검한다.
사용: python render_docx.py report.docx outdir [--soffice]
필요: Windows + Microsoft Word, pywin32, pymupdf.  (LibreOffice가 있으면 --soffice 로 대체 가능)
산출: outdir/report.pdf, outdir/page-01.png ..., 점검 결과 출력.
"""
import os, re, sys, shutil, subprocess


def word_export(docx, pdf):
    import pythoncom
    import win32com.client as wc
    pythoncom.CoInitialize()
    app = wc.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        d = app.Documents.Open(os.path.abspath(docx), ReadOnly=False, AddToRecentFiles=False)
        d.Repaginate()
        for toc in d.TablesOfContents:
            toc.Update()
        d.Fields.Update()
        for toc in d.TablesOfContents:
            toc.UpdatePageNumbers()
        d.Repaginate()
        d.ExportAsFixedFormat(os.path.abspath(pdf), 17)
        n = d.ComputeStatistics(2)  # wdStatisticPages
        d.Close(False)  # 저장하지 않는다. Word로 다시 저장하면 서식이 바뀐다
        return n
    finally:
        app.Quit()


def soffice_export(docx, outdir):
    exe = shutil.which("soffice") or r"C:\Program Files\LibreOffice\program\soffice.exe"
    subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", outdir, docx], check=True)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    docx, outdir = args[0], args[1]
    os.makedirs(outdir, exist_ok=True)
    pdf = os.path.join(outdir, "report.pdf")
    if "--soffice" in sys.argv:
        soffice_export(docx, outdir)
        os.replace(os.path.join(outdir, os.path.splitext(os.path.basename(docx))[0] + ".pdf"), pdf)
    else:
        word_export(docx, pdf)
    import pymupdf
    doc = pymupdf.open(pdf)
    print(f"총 {len(doc)}쪽 (표지 포함)")
    ok = True
    footer_nums = {}
    link_pages = {}
    for i, pg in enumerate(doc):
        pg.get_pixmap(dpi=80).save(os.path.join(outdir, f"page-{i + 1:02d}.png"))
        h = pg.rect.height
        foot = [b[4].strip() for b in pg.get_text("blocks") if b[1] > h - 90]
        nums = [t for t in foot if re.fullmatch(r"\d+", t)]
        footer_nums[i + 1] = nums[0] if nums else None
        link_pages[i + 1] = [l["uri"] for l in pg.get_links() if l.get("uri")]
    # 쪽 번호: 표지(1쪽)는 없음, 이후는 1부터 연속
    cover_ok = footer_nums[1] is None
    seq_ok = all(footer_nums[p] == str(p - 1) for p in range(2, len(doc) + 1))
    print("PASS" if cover_ok else "FAIL", "표지에 쪽 번호 없음", footer_nums[1])
    print("PASS" if seq_ok else "FAIL", "2쪽부터 쪽 번호 1,2,3... 연속",
          {p: v for p, v in footer_nums.items() if p > 1 and v != str(p - 1)} or "")
    total_links = sum(len(v) for v in link_pages.values())
    print("PASS" if total_links else "FAIL", f"PDF 안의 링크 주석 {total_links}개 (본문 {sum(len(v) for p, v in link_pages.items() if p < len(doc))}개 포함)")
    # 목차 vs 실제 쪽수: 목차 쪽의 '제목 ... 숫자' 줄을 읽어 실제 제목 위치와 비교
    toc_re = re.compile(r"^(.+?)[\s.·…]*\t?\s*(\d+)\s*$")
    toc = {}
    for i in range(1, min(4, len(doc))):
        txt = doc[i].get_text()
        if "목차" not in txt and not toc:
            continue
        for ln in txt.splitlines():
            m = toc_re.match(ln.strip())
            if m and len(m.group(1)) > 1:
                toc.setdefault(m.group(1).strip(" .·…"), int(m.group(2)))
        if toc:
            break
    bad = []
    for title, pno in toc.items():
        found = None
        for i in range(2, len(doc)):
            if any(ln.strip() == title for ln in doc[i].get_text().splitlines()):
                found = footer_nums.get(i + 1)  # 그 쪽 푸터에 찍힌 번호
                break
        if found is None or str(found) != str(pno):
            bad.append((title, pno, found))
    print("PASS" if toc and not bad else "FAIL", f"목차 {len(toc)}항목 vs 실제 쪽", bad[:8] if bad else "")
    print("본문 쪽수(표지 제외):", len(doc) - 1)
    print("PNG:", os.path.join(outdir, "page-NN.png"))


if __name__ == "__main__":
    main()
