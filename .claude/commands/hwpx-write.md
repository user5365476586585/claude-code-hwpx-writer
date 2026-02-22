# HWP 문서 작성 통합 워크플로우 (/hwpx-write)

python-hwpx로 문서를 자동 작성하는 통합 워크플로우 skill.
CLAUDE.md 내장 기능과 다른 skill(style, table, format, page, insert)을 조합한다.

사용자 요청: $ARGUMENTS

---

## ⛔ 실행 전 절차 (CLAUDE.md 요청 처리 절차 준수)

이 skill이 로드되었다는 것은 1단계(기획안)에서 문서 작성 작업이 승인되었다는 의미.
그러나 **기획안이 아직 출력되지 않았다면** 반드시 먼저 기획안을 출력한다:

1. **목적**: 최종 결과물 한 줄 정의
2. **산출물 설계**: 문서 구조, 스타일 프리셋, 표 구성, 예상 분량
3. **작업 항목**: TODO + `[skill: xxx]` 태깅

**⚠ 사용자 확인 없이 본문 작성에 들어가지 않는다.**

---

## ★ 온디맨드 워크플로우 원칙

### 핵심: 본문 작성 즉시 시작, 설정은 필요할 때만

```
[잘못된 접근 — 순차 설정]
용지 설정 → 스타일 전체 생성 → 머리말/꼬리말 → 그제서야 본문 시작
→ 설정 단계에서 시간 낭비, 불필요한 설정까지 수행

[올바른 접근 — 온디맨드]
본문 작성 시작 → 제목 스타일 필요? → 그때 해당 스타일만 확인
→ 표 필요? → 그때 표 생성 → 꼬리말 필요? → 본문 완성 후 추가
```

### 설정별 수행 시점

| 설정 | 수행 시점 | 건너뛰는 조건 |
|------|----------|-------------|
| 편집 용지 | 용지/여백 변경 필요 시 | A4 세로 기본값이면 건너뜀 |
| 스타일 확인 | 해당 스타일을 처음 사용할 때 | 기본값이 적합하면 건너뜀 |
| charPr 생성 | 커스텀 서식이 필요할 때 | 기존 charPr로 충분하면 건너뜀 |
| 머리말/꼬리말 | 본문 완성 후 | 불필요하면 건너뜀 |

---

## 레시피 1: 기존 문서 분석

기존 문서에 추가/수정 시 먼저 분석.

```python
from hwpx import HwpxDocument
from hwpx.tools.text_extractor import TextExtractor

with HwpxDocument.open("existing.hwpx") as doc:
    # 1. 전체 텍스트 확인
    with TextExtractor("existing.hwpx") as te:
        full_text = te.extract_text()
        print(full_text[:2000])

    # 2. 사용된 스타일 확인
    used_ids = set(p.style_id_ref for p in doc.paragraphs)
    for sid in used_ids:
        s = doc.style(sid)
        if s:
            print(f"  [{sid}] {s.name}")

    # 3. 페이지 설정 확인
    ps = doc.sections[0].properties.page_size
    pm = doc.sections[0].properties.page_margins
    print(f"용지: {ps.width}x{ps.height} ({ps.orientation})")
    print(f"여백: L={pm.left} R={pm.right} T={pm.top} B={pm.bottom}")

    # 4. 표/그림 확인
    for para in doc.paragraphs:
        for tbl in para.tables:
            print(f"  표: {tbl.row_count}×{tbl.column_count}")

    # 5. 번호 체계 파악
    # → 텍스트에서 패턴 확인: 장(제N장), 절(N.1), □○-· 등
```

**번호 체계 파악 체크리스트:**
- 장(제N장), 절(N.1), 소절(N.1.1) 구조 여부
- 기호 계층: □ → ○ → - → · → ※ 또는 다른 체계
- 표/그림 캡션 형식: `<표 N-M>` 또는 `[표 N-M]`
- 출처 표기: `자료:` / `주:` / `출처:`
- 어미: ~임, ~이다, ~합니다, ~함

---

## 레시피 2: 용지 설정 (필요할 때만)

A4 세로 기본값이면 **건너뛴다**.

```python
props = doc.sections[0].properties
ps = props.page_size

def mm_to_hu(mm):
    return round(mm * 7200 / 25.4)

if ps.width != 59528 or ps.height != 84188:
    # 기본 A4가 아닌 경우에만 변경
    props.set_page_size(width=59528, height=84188, orientation="PORTRAIT")

# 여백 변경 (공문서 기준)
props.set_page_margins(
    left=mm_to_hu(20), right=mm_to_hu(20),
    top=mm_to_hu(20), bottom=mm_to_hu(15),
    header=mm_to_hu(15), footer=mm_to_hu(15),
    gutter=0,
)
```

→ 상세: `/hwpx-page`

---

## 레시피 3: 스타일 온디맨드 설정

**원칙**: 스타일은 처음 사용할 때 확인하고, 부적합하면 그때 생성.

```python
# 스타일 ID 매핑
style_map = {s.name: sid for sid, s in doc.styles.items()}

# "개요 1"이 있는지 확인
if "개요 1" in style_map:
    outline1_id = style_map["개요 1"]
else:
    # 없으면 생성 → /hwpx-style C절 패턴 사용
    pass

# charPr 준비
bold_id = doc.ensure_run_style(bold=True)
normal_id = doc.ensure_run_style(bold=False)
```

→ 상세: `/hwpx-style`

---

## 레시피 4: 본문 작성 (고속 입력)

### 기본 패턴: add_styled_paragraph 헬퍼 반복

```python
P = add_styled_paragraph  # CLAUDE.md 공통 헬퍼

# 간단 보고서
P(doc, "보고서 제목", "개요 1")
P(doc, "본문 첫 단락 전체를 한 번에 입력한다.", "바탕글")
P(doc, "1.1 배경", "개요 2")
P(doc, "배경 설명 내용을 한 단락으로 입력한다.", "바탕글")
```

### 실전 보고서 패턴 (기존 문서 스타일 사용)

**⚠ 자동 기호 판별 먼저 수행** (`/hwpx-style` A절 `detect_auto_symbols()`) → 자동 기호 스타일은 텍스트에서 기호 제외.

```python
P = add_styled_paragraph

# 절 제목
P(doc, "3.1 해외 시장·산업동향", "1.1")

# 기호 수준 — 자동/수동 판별 결과에 따라 기호 포함 여부 결정
P(doc, "□ 글로벌 반도체 시장 현황", "□")                    # heading.type=NONE → 기호 수동 포함
P(doc, "반도체 시장은 TSMC, 삼성전자가 공급을 주도", "❍ 내용1")  # heading.type=BULLET → 기호 자동
P(doc, "(점유율) TSMC 54.4%, 삼성전자 16.5%", "- 내용2")      # heading.type=BULLET → 기호 자동
P(doc, "상기 점유율은 2024년 기준이며 IDM 자체 생산분 제외", "※ 내용4")  # BULLET → 자동

P(doc, "가. TSMC 공급망 전략", "가.")                        # heading.type=NONE → 수동 포함
P(doc, "나. 삼성전자 파운드리 현황", "가.")
```

### 속도 최적화 원칙

- **단락 전체를 한 번의 add_paragraph()로 입력** (문장마다 나누지 않음)
- 500자 초과 시에만 분할 고려
- add_paragraph()만 반복 → 중간 검증 최소화
- **장(章) 시작 시**: 첫 번째 장은 그대로, **두 번째 장부터 `pageBreak="1"` 사용**
  ```python
  doc.add_paragraph("제2장 시장 동향", style_id_ref=outline1_id, pageBreak="1")
  ```

---

## 레시피 5: 장 표지 생성

```python
# 쪽 나누기 + 빈 줄 + 제목
doc.add_paragraph("", pageBreak="1")

# 빈 줄 (페이지 중간)
for _ in range(10):
    doc.add_paragraph("")

# 장 제목 (가운데 정렬 + 큰 글씨 — charPr/paraPr 필요)
# center_para_id, title_char_id는 /hwpx-format C절로 사전 생성
doc.add_paragraph("제2장 시장 동향", char_pr_id_ref=title_char_id, para_pr_id_ref=center_para_id)

# 부제 (선택)
doc.add_paragraph("Market Trends", char_pr_id_ref=subtitle_char_id, para_pr_id_ref=center_para_id)

# 다음 페이지
doc.add_paragraph("", pageBreak="1")
```

→ 상세: `/hwpx-page` E절

---

## 레시피 6: 표 삽입 + 캡션/출처

```python
P = add_styled_paragraph

# 1. 캡션 (기존 문서의 '표 제목' 스타일 사용, 없으면 바탕글)
P(doc, "<표 1-1> 글로벌 반도체 시장 규모", "표 제목")

# 2. 단위 (필요시)
P(doc, "(단위: 억 달러, %)", "표,그림의 단위")

# 3. 표 생성 (add_styled_table 헬퍼 — 너비 자동 + 셀 paraPr 적용)
table = add_styled_table(doc, 4, 4, container_style="표위치", cell_style="표내용")

# 4. 헤더 + 데이터 입력
headers = ["구분", "2022", "2023", "증감률"]
for c, h in enumerate(headers):
    table.set_cell_text(0, c, h)

data = [
    ["메모리", "1,580", "1,298", "-17.8"],
    ["로직", "1,875", "1,920", "+2.4"],
    ["합계", "3,455", "3,218", "-6.9"],
]
for r, row in enumerate(data):
    for c, val in enumerate(row):
        table.set_cell_text(r + 1, c, val)

# 5. 출처
doc.add_paragraph("자료: WSTS(2024), World Semiconductor Trade Statistics",
                   char_pr_id_ref=normal_id)

# 6. 본문 복귀
doc.add_paragraph("", style_id_ref=style_map.get("바탕글", "0"))
```

→ 상세: `/hwpx-table`

---

## 레시피 7: 머리말/꼬리말 (본문 완성 후)

본문 작성이 끝난 후 필요시 추가.

```python
# 머리말 설정
doc.set_header_text("보고서 제목", page_type="BOTH")

# 꼬리말 설정 (텍스트만 가능, 자동 쪽번호 미지원)
doc.set_footer_text("기밀문서", page_type="BOTH")
```

→ 상세: `/hwpx-page` C절

---

## 레시피 8: 최종 검토 + 저장

```python
from hwpx.tools.text_extractor import TextExtractor

# 1. 전체 문단 확인
print(f"총 문단: {len(doc.paragraphs)}")
for i, para in enumerate(doc.paragraphs[:30]):
    style = doc.style(para.style_id_ref)
    style_name = style.name if style else "?"
    text = para.text[:60]
    print(f"  [{i}] [{style_name}] {text}")

# 2. TextExtractor로 전체 텍스트 확인
doc.save_to_path("temp/_review.hwpx")
with TextExtractor("temp/_review.hwpx") as te:
    full_text = te.extract_text()
    print(f"총 글자 수: {len(full_text)}")

# 3. 누락 확인
print(f"섹션 수: {len(doc.sections)}")
print(f"스타일 수: {len(doc.styles)}")

# 4. 저장
import os
doc.save_to_path(os.path.abspath("output/report.hwpx"))
print("저장 완료")

# 5. temp 정리
import glob
for f in glob.glob("temp/*"):
    os.remove(f)
```

---

## 스타일 프리셋 (문서 유형별)

### 프리셋 A: 실전 보고서 (9수준)

| 수준 | 크기 | 굵기 | 정렬 | 줄간격 | 형식 |
|------|------|------|------|--------|------|
| 개요 1 (장) | 18pt | ○ | 가운데 | 160% | 제1장 서론 |
| 개요 2 (절) | 14pt | ○ | 왼쪽 | 160% | 1.1 배경 |
| 개요 3 (소절) | 12pt | ○ | 왼쪽 | 160% | 1.1.1 현황 |
| □ 대주제 | 11pt | ○ | 양쪽 | 170% | □ 내용 |
| ○ 중주제 | 11pt | × | 양쪽 | 170% | ○ 내용 |
| - 소항목 | 10pt | × | 양쪽 | 170% | - 내용 |
| · 세부 | 10pt | × | 양쪽 | 170% | · 내용 |
| ※ 참고 | 9pt | × | 양쪽 | 160% | ※ 내용 |
| 가. 나열 | 10pt | × | 양쪽 | 170% | 가. 내용 |

→ 설정 코드: `/hwpx-style` F절 + G절

### 프리셋 B: 간단 보고서

개요 1(18pt/가운데) + 개요 2(14pt) + 개요 3(12pt) + 바탕글(11pt/170%)

### 프리셋 C: 공문서

개요 1(16pt/가운데) + 개요 2(12pt) + 바탕글(10pt/160%)

---

## 문서 유형별 온디맨드 흐름

### 유형 1: 실전 보고서 ★

```
본문 시작: 제목 입력 → 개요 1 스타일 확인
    ↓
본문 계속: □ ○ 기호 필요? → charPr 생성
    ↓
표 필요 시점: 캡션 → 표 생성 → 데이터 → 출처
    ↓
본문 완성 후: 머리말/꼬리말 추가
    ↓
검토 + 저장
```

### 유형 2: 간단 보고서

```
개요 1 → 제목 → 바탕글 → 본문 → 저장
대부분 설정 건너뜀
```

### 유형 3: 표 중심 문서

```
(용지: 표가 넓으면 가로로 변경)
    ↓
캡션 → 표 생성 → 데이터 → 출처
    ↓
저장
```

---

## 구간별 검증 게이트

| 게이트 | 시점 | 검증 방법 |
|--------|------|----------|
| G1 | 스타일 설정 후 | `doc.styles` → 적용 확인 |
| G2 | 섹션 완료 시 | 문단 텍스트 확인 (`doc.paragraphs`) |
| G3 | 문서 완료 시 | 전체 문단 수 + 텍스트 확인 + 저장 |

> python-hwpx는 XML 직접 조작이므로 "표 안 갇힘" 문제 없음. pyhwpx의 G3 FATAL은 해당 없음.

---

## 에러 처리

### 스타일이 없음
```python
style_map = {s.name: sid for sid, s in doc.styles.items()}
target = "개요 1"

if target in style_map:
    doc.add_paragraph("제1장 서론", style_id_ref=style_map[target])
else:
    # 스타일 없음 → 직접 서식 대체
    bold_id = doc.ensure_run_style(bold=True)
    doc.add_paragraph("제1장 서론", char_pr_id_ref=bold_id)
```

### 문서 저장 실패
```python
try:
    doc.save_to_path("output/report.hwpx")
except Exception as e:
    print(f"저장 실패: {e}")
    # 비상 저장
    doc.save_to_path("temp/_emergency.hwpx")
```

---

## 참조

- 다른 skill: `/hwpx-style`, `/hwpx-table`, `/hwpx-format`, `/hwpx-page`, `/hwpx-insert`
- CLAUDE.md: 기본 조작 (열기/저장/문단추가/찾기바꾸기)
- 스타일 프리셋 상세: `/hwpx-style` G절
