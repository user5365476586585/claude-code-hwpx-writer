# HWP 개체 삽입 (/hwpx-insert)

python-hwpx로 도형/캡션을 삽입하는 skill.

사용자 요청: $ARGUMENTS

---

## ⛔ 핵심 규칙

1. **도형/컨트롤**: `add_shape()` / `add_control()` API 사용 (제한적).
2. **캡션**: 수동 텍스트 입력 방식 (자동 번호 미지원).

---

## 미지원 기능 (포기 목록)

| 기능 | 사유 |
|------|------|
| 이미지 삽입 | XML 구조 복잡 (renderingInfo 등), 불안정 |
| 메모 삽입/수정/삭제 | lxml/ET 호환 버그 (append 시 타입 불일치) |
| 각주/미주 | 본체+참조 연결 복잡, API 없음 |
| 필드/누름틀 | API 없음 |
| 캡션 자동 번호 | 번호 렌더링이 한/글 엔진 의존 |
| 쪽 번호 자동 삽입 | API 없음 |
| 하이퍼링크 | API 없음 |
| 특수문자/겹침글자 | COM API 전용 |
| 셀 분할 (split_merged_cell) | lxml/ET 호환 버그 |

---

## A. 도형/컨트롤 (제한적)

### 도형 삽입

```python
shape = doc.add_shape(
    shape_type="rect",          # 사각형
    attributes=None,
)

# 속성 접근
shape.get_attribute("width")
shape.set_attribute("width", "12000")
```

### 컨트롤 삽입

```python
ctrl = doc.add_control(
    control_type=None,
    attributes=None,
)

ctrl.get_attribute("name")
ctrl.set_attribute("name", "TextBox1")
```

> **제한**: 도형/컨트롤 삽입은 기본적인 XML 요소만 생성. 복잡한 도형(그라데이션, 그룹화 등)은 미지원.

---

## B. 캡션

### 수동 캡션 입력 (텍스트 기반)

python-hwpx에서 캡션 자동 번호(`ShapeObjInsertCaptionNum`)는 미지원.
수동 텍스트로 캡션을 입력한다.

```python
# 표 캡션 (표 위에 배치)
bold_id = doc.ensure_run_style(bold=True)
doc.add_paragraph("<표 1-1> 글로벌 반도체 시장 규모",
                   char_pr_id_ref=bold_id,
                   para_pr_id_ref=center_para_id)

# 표 생성 ...

# 출처 (표 아래)
normal_id = doc.ensure_run_style(bold=False)
doc.add_paragraph("자료: WSTS(2024)", char_pr_id_ref=normal_id)
```

### 캡션 형식 참조

| 구분 | 형식 | 예시 |
|------|------|------|
| 표 캡션 | `<표 장번호-순번> 제목` | `<표 2-1> 시장 규모` |
| 그림 캡션 | `[그림 장번호-순번] 제목` | `[그림 2-1] 추이 그래프` |
| 단위 | `(단위: 억 원, %)` | 표 캡션 우측 또는 바로 아래 |
| 주석 | `주: 내용` | 표/그림 바로 아래 |
| 출처 | `자료: 기관명(연도), 자료명` | 주석 아래 |

### 캡션 + 표 + 출처 전체 패턴

```python
# → /hwpx-table F절 참조 (통합 패턴)
```

---

## C. 객체 검색 (ObjectFinder)

문서 내 기존 객체를 검색하여 분석/수정.

```python
from hwpx.tools.object_finder import ObjectFinder

finder = ObjectFinder("document.hwpx")

# 모든 표 찾기
tables = finder.find_all(tag="tbl")
for t in tables:
    print(f"표: {t.path}")
    print(f"  width={t.get('width')}")

# 모든 그림 찾기
pics = finder.find_all(tag="pic")
for p in pics:
    print(f"그림: {p.path}")

# 속성 조건 검색
import re
results = finder.find_all(attrs={"name": re.compile(r"^IMG")})

# 주석 검색 (형광펜, 각주, 하이퍼링크 등)
for ann in finder.iter_annotations(kinds=["highlight", "footnote"]):
    print(f"{ann.kind}: {ann.value}")
```

---

## D. 컨트롤 조회 (문서 내 기존 객체)

```python
# 문단 내 표/도형 등 조회
for para in doc.paragraphs:
    # 표
    for tbl in para.tables:
        print(f"표: {tbl.row_count}×{tbl.column_count}")

    # 런 정보
    for run in para.runs:
        print(f"  run: {run.text[:30]}, charPr={run.char_pr_id_ref}")
```

---

## 참조

- ObjectFinder: `references/python-hwpx-tools.md`
- OWPML 스키마: `references/hwpx-xml-schema.md`
