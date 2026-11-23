# report-skill

조사·비교·분석 **보고서를 일정한 규칙으로 작성**하는 Claude Code 스킬 `report-writer`입니다.
표지·목차·서론·본론·결론·참고문헌 구조, 클릭 가능한 출처 링크, 표지 제외 쪽 번호, 표·그림 캡션 같은 규칙을 DOCX 생성 스크립트가 코드로 강제하고, 완성 뒤에는 자동 점검과 렌더링 점검으로 확인합니다.

## 한눈에 보기
- **주력 형식:** DOCX (스크립트로 생성·점검). HWPX·PPTX는 규칙만 적용하고, 생성은 각각 `hwpx-skill`, `anthropic-skills:pptx` 스킬에 맡깁니다.
- **사실 규칙:** 수치와 사실은 웹 검색으로 확인한 것만 씁니다. 출처를 못 찾으면 쓰지 않고 "확인 필요"로 표시합니다.
- **점검:** 구조 점검(`check_docx.py`)과, Word로 열어 PDF·이미지로 변환하는 렌더링 점검(`render_docx.py`)을 둘 다 통과해야 완료로 봅니다.

## 설치
`report-writer/` 폴더를 `~/.claude/skills/`에 복사합니다.

```bash
# macOS / Linux / Git Bash
cp -r report-writer ~/.claude/skills/

# Windows PowerShell
Copy-Item report-writer "$HOME\.claude\skills\report-writer" -Recurse -Force
```

저장소의 파일을 고치면 다시 복사해야 합니다. 복사한 뒤 **새 Claude Code 세션**에서 `/report-writer`가 목록에 나타납니다.

### 필요한 것
| 용도 | 필요 |
|---|---|
| DOCX 생성·구조 점검 | Python 3, `pip install python-docx` |
| 렌더링 점검 | Windows + Microsoft Word, `pip install pywin32 pymupdf` (Word가 없으면 LibreOffice를 설치하고 `--soffice` 옵션 사용) |
| HWPX 보고서 | `hwpx-skill` (별도 스킬, 이 저장소에 포함되지 않음) |
| PPTX 보고서 | `anthropic-skills:pptx` |

## 사용 예
Claude Code에서 스킬을 직접 부릅니다.

```
/report-writer 생활 속 해열진통제 성분 비교 보고서를 DOCX로 만들어 줘. 분량은 본문 15~20쪽, 출력은 out 폴더에 저장해 줘.
```

"보고서 써줘", "조사 보고서", "docx 보고서" 같은 요청이면 스킬이 자동으로 선택되기도 합니다.

스크립트를 직접 실행할 수도 있습니다.

```bash
# 1) JSON 명세 -> DOCX  (명세 예시: report-writer/templates/report_spec.example.json)
python report-writer/scripts/build_docx.py spec.json out/report.docx

# 2) 구조 점검 (실패 항목이 있으면 종료 코드 1)
python report-writer/scripts/check_docx.py out/report.docx

# 3) 렌더링 점검: Word로 열어 PDF/PNG 변환 후 쪽 번호·링크·목차 쪽수 확인
python -W ignore report-writer/scripts/render_docx.py out/report.docx render/
```

## 규칙 요약
| 구분 | 규칙 |
|---|---|
| 구조 | 표지(제목·작성자·날짜) → 목차 → 서론 → 본론 → 결론 → 참고문헌 |
| 제목·목차 | Heading 1/2/3 스타일, 자동 목차(TOC 필드) |
| 출처 | 본문 인용은 모두 클릭 가능한 하이퍼링크. 참고문헌에 제목·기관·접속일·URL |
| 사실 | 웹 검색으로 확인한 것만. 못 찾으면 쓰지 않고 "확인 필요" |
| 쪽 번호 | 모든 페이지 하단, **표지 제외**, 본문부터 1쪽 |
| 글꼴·간격 | 맑은 고딕(HWPX는 함초롬바탕), 본문 11pt, 줄간격 160% |
| 표·그림 | 번호와 캡션 (`표 1. ○○ 비교`) |
| 색 | 2가지 이내 (남색 `#1F3864`, 회색 `#595959`, 본문은 검정) |
| 완료 전 | 구조 점검 + 렌더링 점검 + 수동 체크리스트. 분량은 렌더링으로 센 쪽수로 보고 |

자세한 작업 순서와 체크리스트는 [`report-writer/SKILL.md`](report-writer/SKILL.md)에 있습니다.

## 구성
```
report-skill/
  README.md
  .gitignore
  report-writer/
    SKILL.md                             스킬 본문(규칙·작업 순서·체크리스트)
    scripts/build_docx.py                JSON 명세 -> 규칙을 적용한 DOCX
    scripts/check_docx.py                DOCX 구조 자동 점검
    scripts/render_docx.py               Word로 PDF/PNG 변환 후 쪽 번호·링크·목차 쪽수 점검
    templates/report_spec.example.json   명세 예시(내용은 가짜)
  out/
    cosmetics_report.docx                예시 결과물
    cosmetics_report.pdf                 위 DOCX를 Word로 변환한 PDF
    cosmetics_report.spec.json           위 보고서의 명세
```

## 예시 결과물
`out/`의 「화장품 성분 비교·분석 보고서」(스킨케어·선크림·베이스·아이 메이크업, 본문 19쪽, 참고문헌 45건)는 이 스킬로 만든 실제 출력입니다. 알아 둘 점:

- 사실과 수치는 **2026-11-24 기준** 웹 검색 결과이며, 규제는 바뀔 수 있습니다. 이용 전에 원문으로 확인하세요.
- 법령 원문이나 논문 본문을 열지 못한 항목은 초록·요약·2차 자료로 확인했고, 본문과 결론의 표 9에 "확인 필요"로 정리했습니다.
- EU 규정 일부는 영국 정부가 보존한 사본을 열람한 것이라 현행 EU 원문과 다를 수 있습니다.
- 일반 정보 제공용이며 의학적 조언이나 특정 제품의 안전성 판정이 아닙니다.

## 한계
- 자동 생성·점검은 DOCX만 지원합니다.
- DOCX를 Word로 열면 "필드를 업데이트할까요?" 창이 뜹니다. '예'를 누르면 목차에 쪽 번호가 채워집니다.
- 줄간격 160%는 Word의 "배수 1.6"입니다. 맑은 고딕에서는 한글(HWP)의 160%보다 넓게 보입니다.
- 사실의 정확성과 출처가 주장을 뒷받침하는지는 코드가 확인하지 못하므로 사람이 확인해야 합니다.
