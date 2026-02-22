# python-hwpx HWP 자동화 에이전트

이 CLAUDE.md가 오케스트레이터이며, `.claude/commands/hwpx-*.md` skill 파일들을 필요 시 로드하여 작업을 수행한다.
python-hwpx(순수 Python, OWPML/OPC 직접 파싱) 사용. 한/글 설치 없이, OS 무관하게 HWPX 문서 생성/편집.

---

## ⛔ 치명적 규칙

1. **리소스 해제**: 작업 완료 시 `doc.close()` 호출. with문 또는 try/finally 패턴 필수.
2. **동시 문서 관리**: `HwpxDocument` 인스턴스는 1개만 유지. 기존 인스턴스 close 후 새로 생성.
3. **인코딩**: 파일 경로에 한글 포함 시 `os.path.abspath()` 사용. UTF-8 기본.
4. **임시 파일 관리**: 실행용 .py 스크립트는 반드시 `temp/` 디렉토리에 작성. 작업 완료 시 `temp/` 내 모든 파일 삭제.
5. **한글 출력 인코딩**: 스크립트 실행 시 stdout 한글이 깨질 수 있음. 결과를 `temp/*.txt`에 UTF-8로 저장 후 Read 도구로 읽어야 함.
   ```python
   with open("temp/result.txt", "w", encoding="utf-8") as f:
       f.write("\n".join(out))
   ```
6. **한글 파일명 복사 금지**: `shutil.copy2("한글명.hwpx", "다른한글명.hwpx")` → `SameFileError` 위험. 복사본은 반드시 `temp/` 내 영문명으로 생성.
   ```python
   shutil.copy2("한글문서.hwpx", "temp/_work.hwpx")  # ✓
   ```
7. **lxml/ET 호환 패치**: python-hwpx는 lxml과 stdlib ET를 혼합 사용. `ET.SubElement(lxml_parent, ...)` 호출 시 TypeError 발생 가능. **스크립트 상단에 패치 1회 적용 필수.**
   ```python
   import xml.etree.ElementTree as ET
   _original_sub_element = ET.SubElement
   def _safe_sub_element(parent, tag, attrib=None, **extra):
       attrs = {} if attrib is None else dict(attrib)
       attrs.update(extra)
       try:
           return _original_sub_element(parent, tag, attrs)
       except TypeError:
           if hasattr(parent, "makeelement"):
               child = parent.makeelement(tag, attrs)
               parent.append(child)
               return child
           raise
   ET.SubElement = _safe_sub_element
   ```
8. **저장 시 format**: `doc.save_to_path("output.hwpx")` 사용. HWPX 형식 고정 (HWP 레거시 저장 불가).
9. **API 불일치 대응**: references/ 파일이 현재 python-hwpx 버전과 다를 수 있음. 에러 시 `dir(doc)`, `help(doc.메서드명)` 참조.
10. **XML 수정 후 dirty 마킹**: section/header XML을 직접 수정한 후 반드시 `section.mark_dirty()` 또는 `header.mark_dirty()` 호출.
11. **스타일 적용 시 paraPr/charPr 동시 지정 필수**: `add_paragraph(style_id_ref=...)` 만으로는 스타일 서식이 적용되지 않음. 한/글은 문단에 명시된 `paraPrIDRef`/`charPrIDRef`가 스타일의 것보다 우선하므로, **스타일이 참조하는 paraPr/charPr ID를 문단과 런에도 함께 설정**해야 한다. 아래 '공통 헬퍼 함수'의 `add_styled_paragraph()` 사용 필수.
12. **표 너비**: `add_table(width=...)` 하드코딩 금지. 반드시 `get_body_width(doc)`로 본문 영역 너비를 계산하여 사용. 셀 내부 문단에도 적절한 paraPr/charPr 적용 필요.
13. **자동 기호 판별**: 기존 문서의 스타일에 글머리 기호(heading.type=BULLET)가 내장된 경우, 텍스트에 기호를 수동 포함하면 중복 출력됨. XML 분석으로 자동/수동 판별 필수 (`/hwpx-style` A절 참조).

---

## 공통 헬퍼 함수

모든 스크립트에 포함해야 하는 핵심 헬퍼. 스타일 적용, 표 생성, 너비 계산을 안전하게 처리.

```python
HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"

def _style_ids(doc, name):
    """스타일 이름 → (styleID, paraPrID, charPrID) 조회"""
    for sid, s in doc.styles.items():
        if s.name == name:
            return sid, str(s.para_pr_id_ref or "0"), str(s.char_pr_id_ref or "0")
    # 폴백: 바탕글
    for sid, s in doc.styles.items():
        if s.name == "바탕글":
            return sid, str(s.para_pr_id_ref or "0"), str(s.char_pr_id_ref or "0")
    return "0", "0", "0"

def add_styled_paragraph(doc, text, style_name, *, pageBreak="0"):
    """스타일 이름으로 문단 추가. paraPr/charPr를 스타일에서 자동 해석."""
    sid, pp, cp = _style_ids(doc, style_name)
    return doc.add_paragraph(text, style_id_ref=sid, para_pr_id_ref=pp,
                              char_pr_id_ref=cp, pageBreak=pageBreak)

def get_body_width(doc, section_index=0):
    """본문 영역 너비 (HwpUnit) 계산"""
    props = doc.sections[section_index].properties
    ps = props.page_size
    pm = props.page_margins
    return ps.width - pm.left - pm.right

def add_styled_table(doc, rows, cols, *, container_style="바탕글", cell_style="바탕글"):
    """표 추가 + 컨테이너/셀 내부 문단에 스타일 적용 + 본문 너비 자동 계산.
    container_style: 표를 감싸는 문단의 스타일 (표위치 등 가운데 정렬 권장)
    cell_style: 셀 내부 문단의 스타일 (표내용 등 가운데 정렬 권장)
    """
    width = get_body_width(doc)
    c_sid, c_pp, c_cp = _style_ids(doc, container_style)
    tbl = doc.add_table(rows=rows, cols=cols, width=width,
                         style_id_ref=c_sid, para_pr_id_ref=c_pp, char_pr_id_ref=c_cp)
    # 셀 내부 문단에 cell_style 적용
    _, cell_pp, cell_cp = _style_ids(doc, cell_style)
    for pos in tbl.iter_grid():
        for p_elem in pos.cell.element.iter(f"{HP}p"):
            p_elem.set("paraPrIDRef", cell_pp)
            for run_elem in p_elem.iter(f"{HP}run"):
                run_elem.set("charPrIDRef", cell_cp)
    tbl.mark_dirty()
    return tbl

def mm_to_hu(mm):
    """mm → HwpUnit 변환 (±1~3 오차 가능 — 실용적으로 무방)"""
    return round(mm * 7200 / 25.4)
```

> **⚠ 기존 문서 기반 작업 시**: `container_style`과 `cell_style`은 해당 문서의 표 관련 스타일명으로 변경.
> 예: `container_style="표위치"`, `cell_style="표내용"` (스타일.hwpx 기준)

---

## 초기화 패턴

### 환경 설정

```bash
# conda 환경 확인
conda env list
# hwpx 환경 활성화 (없으면 생성)
conda activate hwpx
# 또는 새 환경: conda create -n hwpx python=3.11 && conda activate hwpx
pip install python-hwpx
```

### 기본 코드 구조

```python
from hwpx import HwpxDocument
import os

# lxml/ET 호환 패치 (⛔ 필수)
import xml.etree.ElementTree as ET
_original_sub_element = ET.SubElement
def _safe_sub_element(parent, tag, attrib=None, **extra):
    attrs = {} if attrib is None else dict(attrib)
    attrs.update(extra)
    try:
        return _original_sub_element(parent, tag, attrs)
    except TypeError:
        if hasattr(parent, "makeelement"):
            child = parent.makeelement(tag, attrs)
            parent.append(child)
            return child
        raise
ET.SubElement = _safe_sub_element

doc = HwpxDocument.open(os.path.abspath("input.hwpx"))
try:
    # === 작업 수행 ===
    # ... 편집 작업 ...
    doc.save_to_path(os.path.abspath("output.hwpx"))
finally:
    doc.close()
```

### 새 문서 작성

```python
doc = HwpxDocument.new()
try:
    # 바로 작업 시작 (기본 빈 문서)
    doc.add_paragraph("문서 내용")
    doc.save_to_path(os.path.abspath("new_doc.hwpx"))
finally:
    doc.close()
```

### Context Manager 방식 (권장)

```python
with HwpxDocument.open("input.hwpx") as doc:
    # ... 편집 ...
    doc.save_to_path("output.hwpx")
# 자동 close
```

### 기존 문서 편집

```python
with HwpxDocument.open(os.path.abspath("existing.hwpx")) as doc:
    # 문서 상태 확인
    print(f"섹션 수: {len(doc.sections)}")
    print(f"문단 수: {len(doc.paragraphs)}")
    # ... 편집 ...
    doc.save_to_path(os.path.abspath("existing.hwpx"))
```

---

## 기본 문서 조작

### 파일 열기/저장/닫기

```python
# 열기
doc = HwpxDocument.open("report.hwpx")
doc = HwpxDocument.open(bytes_data)       # bytes에서 열기
doc = HwpxDocument.new()                  # 새 빈 문서

# 저장
doc.save_to_path("output.hwpx")           # 파일로 저장 (권장)
doc.save_to_stream(binary_io)             # 스트림으로 저장
raw = doc.to_bytes()                      # bytes로 직렬화

# 닫기
doc.close()
```

### 지원 파일 형식

| 동작 | 형식 | 지원 |
|------|------|------|
| 열기/편집/저장 | HWPX | ✓ |
| 읽기 전용 | HWP (레거시) | ✓ (텍스트 추출만) |
| PDF 출력 | PDF | ✗ (미지원) |

---

## 문단 추가/편집

### 기본 문단 추가

```python
# 헬퍼 사용 (★ 권장 — 스타일의 paraPr/charPr 자동 해석)
add_styled_paragraph(doc, "제1장 서론", "개요 1")
add_styled_paragraph(doc, "본문 내용.", "바탕글")

# 저수준 API 직접 사용 (스타일의 paraPr/charPr를 수동 지정해야 함)
para = doc.add_paragraph(
    "제1장 서론",
    style_id_ref="1",        # 스타일 ID
    para_pr_id_ref="2",      # ⚠ 스타일이 참조하는 paraPr ID (필수)
    char_pr_id_ref="0",      # ⚠ 스타일이 참조하는 charPr ID (필수)
)
```

### 쪽 나누기

```python
# pageBreak="1" 속성으로 쪽 나누기 효과
para = doc.add_paragraph("새 페이지 시작", pageBreak="1")
```

### 여러 문단 연속 입력

```python
P = add_styled_paragraph  # 축약
P(doc, "제1장 서론", "개요 1")
P(doc, "본문 내용을 한 단락 전체 한 번에 입력한다.", "바탕글")
P(doc, "1.1 배경", "개요 2")
P(doc, "배경 설명 내용을 한 단락으로 입력한다.", "바탕글")
```

### 고속 입력 패턴 (스타일 + 텍스트 반복)

`add_styled_paragraph()` 헬퍼가 스타일명 → paraPr/charPr 자동 해석을 처리하므로 ID를 외울 필요 없음.

```python
P = add_styled_paragraph
P(doc, "제1장 서론", "개요 1")
P(doc, "본문 내용.", "바탕글")
P(doc, "1.1 배경", "개요 2")
```

### 기호 입력 규칙 (프리셋 A 사용 시)

**⚠ 자동 기호 판별 필수** — 기존 문서의 스타일에 글머리 기호가 내장(heading.type=BULLET)된 경우, 텍스트에 기호를 수동 포함하면 중복 출력됨. `/hwpx-style` A절에서 XML 분석으로 판별 후 사용.

| 구분 | 자동 기호 (BULLET) | 수동 기호 (NONE) |
|------|------------------|----------------|
| 텍스트 | 기호 **제외** (내용만) | 기호 **포함** |
| 예시 (자동) | `P(doc, "내용", "❍ 내용1")` → ❍ 내용 | |
| 예시 (수동) | | `P(doc, "□ 내용", "□")` → □ 내용 |

```python
P = add_styled_paragraph

# 기존 문서 스타일에 글머리 기호가 내장된 경우 (heading.type=BULLET):
# ❍ - · ※ → 기호 자동 출력 → 텍스트에 기호 제외
P(doc, "반도체 시장은 TSMC, 삼성전자가 공급을 주도", "❍ 내용1")   # ❍ 자동
P(doc, "(점유율) TSMC 54.4%, 삼성전자 16.5%", "- 내용2")         # - 자동
P(doc, "2024년 기준, IDM 자체생산분 제외", "※ 내용4")              # ※ 자동

# 글머리 기호가 없는 스타일 (heading.type=NONE):
# □ 가. → 기호를 텍스트에 직접 포함
P(doc, "□ 글로벌 반도체 시장 현황", "□")                          # □ 수동
P(doc, "가. TSMC 공급망 전략", "가.")                              # 가. 수동
P(doc, "나. 삼성전자 파운드리 현황", "가.")                         # 나. 수동

# 새 문서 (스타일 없음) → ensure_run_style + 기호 직접 포함
bold_id = doc.ensure_run_style(bold=True)
normal_id = doc.ensure_run_style(bold=False)
doc.add_paragraph("□ 글로벌 반도체 시장 현황", char_pr_id_ref=bold_id)
doc.add_paragraph("○ 세부 내용", char_pr_id_ref=normal_id)
```

### 문단 속성 수정 (기존 문단)

```python
para = doc.paragraphs[0]
para.text = "수정된 텍스트"
para.style_id_ref = "1"         # 스타일 변경
para.para_pr_id_ref = "2"       # 문단속성 변경
```

### 런(Run) 추가/편집

```python
para = doc.paragraphs[0]

# 런 추가 (문단 내 서식 분리)
run = para.add_run("굵은 텍스트", bold=True)
run = para.add_run("일반 텍스트", bold=False)

# 기존 런 수정
for run in para.runs:
    run.bold = True
    run.text = run.text.upper()
```

---

## 텍스트 추출

```python
from hwpx.tools.text_extractor import TextExtractor

# 문서 전체 텍스트
with TextExtractor("document.hwpx") as te:
    full_text = te.extract_text()
    print(full_text)

# 섹션별/문단별 순회
with TextExtractor("document.hwpx") as te:
    for para in te.iter_document_paragraphs():
        text = para.text()
        if text.strip():
            print(f"[섹션{para.section.index} 문단{para.index}] {text[:80]}")
```

### 문단 텍스트 직접 접근

```python
# 파이썬 객체를 통한 직접 접근 (TextExtractor 없이)
with HwpxDocument.open("document.hwpx") as doc:
    for para in doc.paragraphs:
        print(para.text)
```

---

## 찾기/바꾸기

```python
# 전체 문서에서 텍스트 치환
count = doc.replace_text_in_runs("구텍스트", "신텍스트")
print(f"{count}건 치환됨")

# 스타일 조건부 치환 (빨간 글자만)
count = doc.replace_text_in_runs("TODO", "DONE", text_color="#FF0000")

# 런 단위 치환
for run in doc.iter_runs():
    if "검색어" in run.text:
        run.replace_text("검색어", "대체어")
```

---

## 문서 속성 확인

```python
print(f"섹션 수: {len(doc.sections)}")
print(f"문단 수: {len(doc.paragraphs)}")
print(f"스타일 수: {len(doc.styles)}")

# 용지 설정
ps = doc.sections[0].properties.page_size
print(f"용지: {ps.width}x{ps.height} ({ps.orientation})")

# 여백
pm = doc.sections[0].properties.page_margins
print(f"여백: 좌{pm.left} 우{pm.right} 상{pm.top} 하{pm.bottom}")
```

---

## 단위 변환

```python
# mm → HwpUnit 변환
def mm_to_hu(mm):
    return round(mm * 7200 / 25.4)   # ≈ mm × 283

# HwpUnit → mm 변환
def hu_to_mm(hu):
    return round(hu * 25.4 / 7200, 1)

# pt → charPr Height 변환
# charPr height 속성값 = pt × 100 (HwpUnit)
# 예: 12pt → height="1200"

# 색상: "#RRGGBB" 형식 직접 사용
red = "#FF0000"
blue = "#0000FF"
```

→ 상세: `references/python-hwpx-api.md`

---

## Skill 라우팅 규칙

### 의사결정 트리

```
사용자 요청 분석
├─ 스타일 (조회/생성/적용/프리셋)
│   → /hwpx-style 로드
├─ 표 (생성/편집/셀조작/데이터입력)
│   → /hwpx-table 로드
├─ 글자/문단 직접 서식 (charPr, paraPr, 런 서식)
│   → /hwpx-format 로드
├─ 페이지 레이아웃 (용지/머리말/꼬리말/구역)
│   → /hwpx-page 로드
├─ 개체 삽입 (도형/캡션)
│   → /hwpx-insert 로드
├─ 문서 전체 자동 생성 (보고서/공문서 작성 워크플로우)
│   → /hwpx-write 로드 (내부에서 다른 skill 추가 로드)
└─ 기본 조작 (열기/저장/문단추가/찾기바꾸기)
    → 이 CLAUDE.md 내장 기능으로 직접 수행
```

### 트리거 테이블

| 키워드 | Skill | 설명 |
|--------|-------|------|
| 스타일, 개요, 글머리표, 번호, 프리셋 | `/hwpx-style` | 스타일 관리 |
| 표, 테이블, 셀, 행, 열, DataFrame | `/hwpx-table` | 표 조작 |
| 글꼴, 폰트, 굵게, 기울임, 색상, 줄간격, 정렬, 서식 | `/hwpx-format` | 직접 서식 |
| 용지, 여백, 머리말, 꼬리말, 구역, 단 | `/hwpx-page` | 페이지 레이아웃 |
| 도형, 캡션 | `/hwpx-insert` | 개체 삽입 |
| 보고서 작성, 문서 생성, 공문서, 워크플로우 | `/hwpx-write` | 통합 작성 |

### 복합 작업 오케스트레이션 ★

**⚠ 요청 1개 = skill 1개가 아니다.** 대부분의 실제 작업은 여러 skill이 필요하다.

- 작업을 단계별로 분해하고, **각 단계에서 필요한 skill을 로드**한다.
- 1단계(기획안)에서 **모든 작업 단계를 먼저 도출**하고, 각 단계마다 어떤 skill이 필요한지 `[skill: xxx]`로 태깅한다.
- CLAUDE.md 내장 기능(문단추가/찾기바꾸기/저장)과 외부 skill 파일을 자유롭게 조합한다.

**예시: "보고서 작성" 요청 시:**

```
1단계 기획안에서 작업 항목 도출:
  1. 용지 설정 → [skill: hwpx-page]
  2. 스타일 분석/생성 → [skill: hwpx-style]
  3. 표지 작성 → [CLAUDE.md 내장] (add_paragraph)
  4. 본문 입력 → [CLAUDE.md 내장] (add_paragraph + style_id_ref)
  5. 표 생성 → [skill: hwpx-table] (캡션+데이터+출처)
  6. 머리말/꼬리말 → [skill: hwpx-page]
  7. 최종 검토 + 저장 → [CLAUDE.md 내장]

2단계 실행:
  각 항목 진입 시 해당 skill 로드 → 수행 → G1~G3 검증 게이트 통과 → 다음 항목
```

---

## 스타일 프리셋 요약 (기획안 시 참조)

기획안의 "산출물 설계 > 스타일" 항목 결정 시 이 테이블을 참조한다.
세부 코드/설정값은 실행 단계에서 `/hwpx-style` 로드 후 참조.

### 프리셋 선택 기준

| 프리셋 | 적합 문서 유형 | 수준 수 |
|--------|-------------|--------|
| **A: 실전 보고서** | 정책보고서, 연구보고서, 분석보고서 | 9수준 |
| **B: 간단 보고서** | 업무보고, 내부문서, 검토보고 | 4수준 |
| **C: 공문서** | 공문, 품의서, 기안문 | 3수준 |
| **없음 (기본)** | 단순 문서, 메모, 1~2쪽 | 기본 그대로 |

### 프리셋 A: 실전 보고서 — 기호 계위 ★

한국 정책·연구보고서의 번호+기호 혼합 계위.
**개요 1~3**: `style_id_ref` 지정 (기본 내장 스타일).
**□ ○ - · ※ 가.**: 직접 서식(charPr/paraPr) + 기호 직접 포함.

| 기호 | 계위 위치 | 용도 |
|------|---------|------|
| 제N장 | 최상위 | 보고서 대단위 구분 |
| N.1 / N.1.1 | 절/소절 | 장 아래 논리 단위 |
| □ | 절 아래 | 독립 논점·핵심 주제. 굵게 처리로 시각적 강조 |
| ○ | □ 아래 | 세부 논거·근거·사실 |
| - | ○ 아래 | 구체 수치·사례·데이터 |
| · | - 아래 | 부연 설명 (꼭 필요한 경우만) |
| ※ | 계층 무관 | 참고사항·주석·출처·단서 |
| 가. | □/○ 아래 | 병렬 나열 (가.나.다. 순) |

**문단 작성 규칙**:
- `style_id_ref` — 개요 스타일은 ID 직접 지정 (3번 연타 불필요)
- □ ○ - · ※: 기호를 텍스트에 직접 포함, charPr/paraPr로 서식 지정
- 가. 나.: 수동 번호 직접 입력

```python
P = add_styled_paragraph

# 개요 스타일 (스타일명으로 직접 적용)
P(doc, "제1장 서론", "개요 1")
P(doc, "1.1 해외 시장 동향", "개요 2")

# 기존 문서 스타일에 기호 계위가 있는 경우 (자동/수동 판별 후 사용):
P(doc, "□ 글로벌 반도체 시장 현황", "□")                 # □: heading.type=NONE → 기호 수동 포함
P(doc, "반도체 시장은 TSMC, 삼성전자가 공급을 주도", "❍ 내용1")  # ❍: BULLET → 기호 자동
P(doc, "(점유율) TSMC 54.4%, 삼성전자 16.5%", "- 내용2")       # -: BULLET → 기호 자동
P(doc, "2024년 기준, IDM 자체생산분 제외", "※ 내용4")            # ※: BULLET → 기호 자동

P(doc, "가. TSMC 공급망 전략", "가.")                      # 가.: heading.type=NONE → 수동 포함
P(doc, "나. 삼성전자 파운드리 현황", "가.")
```

### 프리셋 B: 간단 보고서

```
개요1  제목 (18pt/굵게/가운데)
 개요2  소제목 (14pt/굵게)
  개요3  세부제목 (12pt/굵게)
   바탕글  본문 (11pt/양쪽/170%)
```

### 프리셋 C: 공문서

```
개요1  제목 (16pt/굵게/가운데)
 개요2  항목 (12pt/굵게)
  바탕글  내용 (10pt/양쪽/160%)
```

### 표/그림 캡션 형식 (공통)

| 구분 | 형식 | 예시 |
|------|------|------|
| 표 캡션 | `<표 장번호-순번> 제목` | `<표 2-1> 시장 규모` |
| 그림 캡션 | `[그림 장번호-순번] 제목` | `[그림 2-1] 추이 그래프` |
| 단위 | `(단위: 억 원, %)` | 표 캡션 우측 또는 바로 아래 |
| 주석 | `주: 내용` | 표/그림 바로 아래 |
| 출처 | `자료: 기관명(연도), 자료명` | 주석 아래 |

---

## ⛔ 요청 처리 절차 (필수 준수)

**모든 사용자 요청에 대해 반드시 다음 절차를 따른다.**

### 1단계: 기획안 출력 (실행 전 필수)

**스타일이 필요한 작업이면 기획 전에 다음을 먼저 확인한다:**

1. 기존 .hwpx 파일(스타일 정의된 파일) 보유 여부 → 사용자에게 질문
2. .hwpx 파일이 있으면 → **기획안 출력 전에** `/hwpx-style` A절로 스타일 분석 + 자동 기호 판별 실행
   - 스타일 목록, 서식 매핑, 자동 기호(heading.type=BULLET) 여부를 `temp/style_analysis.txt`에 저장 후 Read → 기획안에 포함
   - `heading.type=BULLET` → 텍스트에 기호 **제외** (스타일이 자동 출력)
   - `heading.type=NONE` → 텍스트에 기호 **포함** (수동 입력)
   - ⚠ 어떤 기호가 자동/수동인지는 문서마다 다름 → **반드시 `detect_auto_symbols()` 결과 기반**
3. .hwpx 파일이 없으면 → `ensure_run_style()` + 기호를 텍스트에 직접 포함 (기획안에 명시)

아래 3가지를 **사용자에게 텍스트로 출력**한다:

1. **목적**: 최종 결과물을 한 줄로 정의
2. **산출물 설계**: 문서 구조, 내용 초안, 스타일/서식 계획을 마크다운으로 구체 기획
   - 새 문서: 문서 구조(장/절), 스타일 및 기호 계층(분석 결과 포함), 표/그림 구성
   - 기존 문서 수정: 변경 범위, 영향 분석
   - **표 포함 시 기획안에 명시**: 헤더 볼드 기본 적용, 행/열 크기 조정 필요 여부 → 사용자 확인
   - **표/그림 캡션 포함 시**: 수동 텍스트 입력 방식 (HWP 캡션 기능 `<hp:caption>` 미지원). 기존 문서에 `표 제목`/`그림 제목` 스타일이 있으면 활용
   - **가. 나. 번호 포함 시**: 자동 기호 여부 확인 후 수동/자동 결정
3. **작업 항목**: TODO 목록. 각 항목에 필요한 skill을 `[skill: xxx]` 형태로 명시

```
예시:
## 기획안

**목적**: A4 세로 10페이지 분량의 반도체 시장 분석 보고서 생성

**산출물 설계**:
- 용지: A4 세로, 공문서 여백(위20/아래15/좌우20)
- 스타일: 프리셋 A (실전 보고서 9수준)
- 구조:
  - 표지 (제목 + 작성자/날짜)
  - 제1장 서론 (배경, 목적)
  - 제2장 시장 현황 (<표 2-1> 시장 규모 포함)
  - 제3장 전망
- 표: 2개 (시장 규모, 점유율) — 기본: 헤더 볼드. 열 너비·행 높이 조정 필요 시 알려주세요.
- 가. 나. 번호: 수동 입력 방식
- 머리말: 보고서 제목 / 꼬리말: 텍스트

**작업 항목**:
1. 용지 설정 [skill: hwpx-page]
2. 스타일 분석/생성 [skill: hwpx-style]
3. 표지 작성 [CLAUDE.md 내장]
4. 제1장 본문 입력 [CLAUDE.md 내장]
5. 제2장 본문 + 표 생성 [skill: hwpx-table]
6. 제3장 본문 입력 [CLAUDE.md 내장]
7. 머리말/꼬리말 [skill: hwpx-page]
8. 최종 검토 + 저장 [CLAUDE.md 내장]
```

**기획안 출력 시 다른 방향의 대안도 함께 제시한다.**

**⚠ 사용자가 기획안을 확인하고 승인/수정 지시를 내린 후 2단계로 진행한다. 사용자 확인 없이 실행 금지.**

### 2단계: 실행

승인된 기획안의 작업 항목을 순차 수행.
- 각 항목 진입 시 명시된 skill을 로드(`/hwpx-*`)한 후 진행
- CLAUDE.md 내장 기능으로 충분한 항목은 skill 로드 없이 직접 수행
- 스타일/서식은 온디맨드 (처음 필요할 때만 설정)
- 단락 전체를 한 번에 입력 (문장 단위 분할 금지)
- **장(章) 시작 시**: 첫 번째 장은 그대로, **두 번째 장부터 반드시 `pageBreak="1"` 속성 사용**
  ```python
  doc.add_paragraph("제2장 시장 동향", style_id_ref="1", pageBreak="1")
  ```

### 3단계: 검증 및 보고

각 단계 결과 및 최종 결과물이 목적에 부합하는지 확인. 미달 시 보정.

**구간별 검증 게이트:**

| 게이트 | 시점 | 검증 방법 |
|--------|------|----------|
| G1 | 스타일 설정 후 | `doc.styles` → 적용 확인 |
| G2 | 각 섹션 완료 시 | `doc.paragraphs` → 텍스트 확인 |
| G3 | 문서 완료 시 | 전체 문단 수 + 텍스트 확인 + 저장 |

> python-hwpx는 XML 직접 조작이므로 pyhwpx의 "표 안 갇힘" 문제(G3 FATAL)가 발생하지 않는다.

완료 후 **결과 간략 보고**: 문단 수, 주요 내용 요약, 저장 경로

---

## 공통 에러 처리

### 패키지 오류

```python
try:
    doc = HwpxDocument.open("file.hwpx")
except Exception as e:
    print(f"문서 열기 실패: {e}")
    print("1. python-hwpx 설치 확인: pip install python-hwpx")
    print("2. 파일 경로/형식 확인 (.hwpx만 편집 가능)")
    print("3. 파일 손상 여부 확인")
```

### XML 구조 오류

```python
# lxml/ET 혼합 에러 시 패치 확인
# TypeError: Argument 'parent' has unexpected type '...' 발생 시
# → 스크립트 상단에 lxml/ET 호환 패치 적용 여부 확인 (치명적 규칙 7번)
```

### 안전 종료

```python
doc = None
try:
    doc = HwpxDocument.open(os.path.abspath("input.hwpx"))
    # ... 작업 ...
except Exception as e:
    print(f"에러 발생: {e}")
    if doc:
        try:
            doc.save_to_path("temp/_emergency_save.hwpx")
        except:
            pass
finally:
    if doc:
        try:
            doc.close()
        except:
            pass
```

### API 변경 대응

```python
# 메서드 존재 확인
if hasattr(doc, 'some_method'):
    doc.some_method()
else:
    similar = [m for m in dir(doc) if 'some' in m.lower()]
    print(f"유사 메서드: {similar}")

# python-hwpx 버전 확인
import subprocess
result = subprocess.run(["pip", "show", "python-hwpx"], capture_output=True, text=True)
print(result.stdout)
```

---

## 기존 문서 읽기/파싱/편집 워크플로우

기존 HWPX 문서를 열어 내용 파악 → 분석 → 편집하는 패턴.

### 문서 구조 파악

```python
from hwpx import HwpxDocument
from hwpx.tools.text_extractor import TextExtractor
import os

with HwpxDocument.open(os.path.abspath("existing.hwpx")) as doc:
    # 1. 문서 기본 정보
    print(f"섹션 수: {len(doc.sections)}")
    print(f"문단 수: {len(doc.paragraphs)}")

    # 2. 전체 텍스트 추출
    with TextExtractor("existing.hwpx") as te:
        full_text = te.extract_text()
        print(full_text[:500])

    # 3. 스타일 목록
    for sid, s in doc.styles.items():
        print(f"  [{sid}] {s.name} (type={s.type}, paraPr={s.para_pr_id_ref}, charPr={s.char_pr_id_ref})")

    # 4. 문단별 텍스트 + 스타일
    for i, para in enumerate(doc.paragraphs[:20]):
        style = doc.style(para.style_id_ref)
        style_name = style.name if style else "?"
        print(f"  [{i}] [{style_name}] {para.text[:60]}")
```

### 문단별 스캔 (스타일 + 서식 동시 파악)

```python
with HwpxDocument.open("existing.hwpx") as doc:
    for para in doc.paragraphs:
        style = doc.style(para.style_id_ref)
        style_name = style.name if style else "N/A"
        text = para.text.strip()
        if not text:
            continue

        # 첫 번째 런의 서식 확인
        if para.runs:
            run = para.runs[0]
            cp = doc.char_property(run.char_pr_id_ref)
            height = cp.attributes.get("height", "?") if cp else "?"
            bold = run.bold
            print(f"[{style_name}] {repr(text[:40])} (height={height}, bold={bold})")
```

### 객체 검색 (ObjectFinder)

```python
from hwpx.tools.object_finder import ObjectFinder

finder = ObjectFinder("existing.hwpx")

# 모든 표 찾기
tables = finder.find_all(tag="tbl")
print(f"표 {len(tables)}개 발견")

# 모든 그림 찾기
pics = finder.find_all(tag="pic")
print(f"그림 {len(pics)}개 발견")

# 속성 조건으로 검색
results = finder.find_all(tag="tbl", attrs={"width": lambda v: int(v) > 30000})
```

### 선택적 편집 패턴

```python
with HwpxDocument.open("existing.hwpx") as doc:
    # 1. 전체 텍스트 치환
    count = doc.replace_text_in_runs("구버전 텍스트", "신버전 텍스트")
    print(f"{count}건 치환")

    # 2. 특정 스타일 문단 찾아서 수정
    for para in doc.paragraphs:
        style = doc.style(para.style_id_ref)
        if style and style.name == "개요 1":
            print(f"개요 1 문단: {para.text}")
            # 스타일 변경
            para.style_id_ref = "2"   # 개요 2로 변경

    # 3. 런 서식 일괄 변경
    bold_id = doc.ensure_run_style(bold=True)
    for run in doc.iter_runs():
        if "중요" in run.text:
            run.bold = True

    doc.save_to_path("modified.hwpx")
```

### 기존 문서 스타일 분석 패턴

```python
with HwpxDocument.open("source.hwpx") as doc:
    # 스타일 목록
    print("=== 스타일 목록 ===")
    for sid, s in doc.styles.items():
        print(f"  [{sid}] {s.name} (eng: {s.eng_name})")

    # 문단별 스타일 + 텍스트 동시 파악
    print("\n=== 문단별 분석 ===")
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style = doc.style(para.style_id_ref)
        style_name = style.name if style else "?"
        print(f"  [{style_name}] {repr(text[:40])}")

    # 문자속성 분석
    print("\n=== 문자속성 ===")
    for cid, cp in doc.char_properties.items():
        height = cp.attributes.get("height", "?")
        bold = cp.attributes.get("bold", "0")
        color = cp.text_color() or "#000000"
        print(f"  charPr[{cid}] height={height} bold={bold} color={color}")
```

---

## python-hwpx 미지원 기능 (포기 목록)

| 기능 | 사유 |
|------|------|
| HWP(레거시) 파일 편집 | 읽기 전용만 가능 |
| PDF 출력 | 렌더링 엔진 없음 |
| 이미지 삽입 | XML 구조 복잡 (renderingInfo 등), 불안정 |
| 쪽 번호 자동 삽입 | XML 구조 파악됨이나 API 없음, 불안정 |
| 캡션 자동 번호 | 번호 렌더링이 한/글 엔진 의존 |
| 각주/미주 삽입 | 본체+참조 연결 복잡 |
| 필드/누름틀 생성 | API 없음 |
| 매크로/액션 실행 | COM API 전용 |
| 커서 기반 실시간 편집 | XML 직접 조작 방식 |
| 페이지 이미지 생성 | 렌더링 엔진 없음 |
| 메모 삽입/수정/삭제 | lxml/ET 호환 버그 (append 시 타입 불일치) |
| 셀 분할 (split_merged_cell) | lxml/ET 호환 버그 |

---

## 참조 파일 목록

### API 레퍼런스 (references/)

| 파일 | 내용 |
|------|------|
| `references/python-hwpx-api.md` | HwpxDocument 전체 API 레퍼런스 |
| `references/python-hwpx-oxml.md` | 저수준 XML 조작 API (Section, Paragraph, Run, Table, Header 등) |
| `references/python-hwpx-tools.md` | TextExtractor, ObjectFinder API |
| `references/hwpx-xml-schema.md` | OWPML 스키마 핵심 요소 정리 |

### Skill 파일 (.claude/commands/)

| Skill | 파일 | 트리거 |
|-------|------|--------|
| `/hwpx-style` | `hwpx-style.md` | 스타일 관리 |
| `/hwpx-table` | `hwpx-table.md` | 표 조작 |
| `/hwpx-format` | `hwpx-format.md` | 글자/문단 서식 |
| `/hwpx-page` | `hwpx-page.md` | 페이지 레이아웃 |
| `/hwpx-insert` | `hwpx-insert.md` | 개체/메모 삽입 |
| `/hwpx-write` | `hwpx-write.md` | 문서 작성 통합 |

### 공식 문서

- python-hwpx GitHub: https://github.com/airmang/python-hwpx
- python-hwpx PyPI: https://pypi.org/project/python-hwpx/
