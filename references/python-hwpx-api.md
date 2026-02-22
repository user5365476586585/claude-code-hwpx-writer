# python-hwpx API 레퍼런스

`pip install python-hwpx` → `from hwpx import HwpxDocument`

---

## HwpxDocument (메인 API)

### 생성/열기/닫기

```python
from hwpx import HwpxDocument

# 새 문서
doc = HwpxDocument.new()

# 파일 열기
doc = HwpxDocument.open("input.hwpx")         # 경로
doc = HwpxDocument.open(path_or_bytes_or_io)   # bytes, BinaryIO 가능

# Context manager
with HwpxDocument.open("input.hwpx") as doc:
    ...  # 자동 close

# 닫기
doc.close()
```

### 저장

```python
doc.save_to_path("output.hwpx")       # 파일로 저장 (권장)
doc.save_to_stream(binary_io)          # 스트림으로 저장
raw = doc.to_bytes()                   # bytes로 직렬화
doc.save("output.hwpx")               # deprecated — save_to_path 사용
```

### 문서 속성 (읽기 전용)

| 속성 | 타입 | 설명 |
|------|------|------|
| `doc.package` | `HwpxPackage` | 하위 OPC 패키지 |
| `doc.oxml` | `HwpxOxmlDocument` | 저수준 XML 객체 트리 |
| `doc.sections` | `list[HwpxOxmlSection]` | 모든 구역 |
| `doc.paragraphs` | `list[HwpxOxmlParagraph]` | 전체 문단 목록 |
| `doc.headers` | `list[HwpxOxmlHeader]` | 헤더 파트 |
| `doc.version` | `HwpxOxmlVersion \| None` | 버전 메타데이터 |

### 스타일/속성 딕셔너리

| 속성 | 타입 | 설명 |
|------|------|------|
| `doc.styles` | `dict[str, Style]` | 스타일 정의 (id → Style) |
| `doc.char_properties` | `dict[str, RunStyle]` | 문자속성 정의 |
| `doc.paragraph_properties` | `dict[str, ParagraphProperty]` | 문단속성 정의 |
| `doc.border_fills` | `dict[str, GenericElement]` | 테두리채우기 정의 |
| `doc.bullets` | `dict[str, Bullet]` | 글머리표 정의 |
| `doc.memo_shapes` | `dict[str, MemoShape]` | 메모 모양 정의 |
| `doc.track_changes` | `dict[str, TrackChange]` | 변경 추적 |
| `doc.track_change_authors` | `dict[str, TrackChangeAuthor]` | 변경 추적 작성자 |
| `doc.memos` | `list[HwpxOxmlMemo]` | 모든 메모 |

### ID 조회

```python
doc.style(style_id_ref)              # → Style | None
doc.char_property(char_pr_id_ref)    # → RunStyle | None
doc.paragraph_property(para_pr_id_ref)  # → ParagraphProperty | None
doc.border_fill(border_fill_id_ref)  # → GenericElement | None
doc.bullet(bullet_id_ref)           # → Bullet | None
doc.memo_shape(memo_shape_id_ref)   # → MemoShape | None
doc.track_change(change_id_ref)     # → TrackChange | None
doc.track_change_author(author_id_ref) # → TrackChangeAuthor | None
```

---

## 문단 추가/편집

### add_paragraph

```python
para = doc.add_paragraph(
    text="문단 텍스트",
    section=None,           # HwpxOxmlSection 또는 None(마지막 섹션)
    section_index=None,     # 섹션 인덱스 (section 미지정 시)
    para_pr_id_ref=None,    # 문단속성 ID
    style_id_ref=None,      # 스타일 ID
    char_pr_id_ref=None,    # 문자속성 ID
    run_attributes=None,    # 런 추가 속성 dict
    include_run=True,       # False면 빈 문단 (런 없음)
    pageBreak="0",          # "1"이면 쪽 나누기
    columnBreak="0",        # "1"이면 단 나누기
)
# 반환: HwpxOxmlParagraph
```

### 문단 속성 접근

```python
para.text               # 전체 텍스트 (get/set)
para.runs               # list[HwpxOxmlRun]
para.tables             # list[HwpxOxmlTable]
para.para_pr_id_ref     # 문단속성 ID (get/set)
para.style_id_ref       # 스타일 ID (get/set)
para.char_pr_id_ref     # 문자속성 ID (get/set)
```

### 런(Run) 추가/편집

```python
# 런 추가
run = para.add_run(
    text="텍스트",
    char_pr_id_ref=None,
    bold=False,
    italic=False,
    underline=False,
    attributes=None,
)

# 런 속성
run.text                # 텍스트 (get/set)
run.char_pr_id_ref      # 문자속성 ID (get/set)
run.bold                # bool (get/set)
run.italic              # bool (get/set)
run.underline           # bool (get/set)
run.style               # RunStyle | None

# 런 텍스트 치환
count = run.replace_text("old", "new", count=None)

# 런 삭제
run.remove()
```

---

## 문서 전체 런 순회/검색/치환

```python
# 모든 런 순회
for run in doc.iter_runs():
    print(run.text, run.bold)

# 스타일 조건으로 런 검색
runs = doc.find_runs_by_style(
    text_color="#FF0000",       # 빨간 글자
    underline_type="Bottom",
    underline_color=None,
    char_pr_id_ref=None,
)

# 조건부 텍스트 치환
count = doc.replace_text_in_runs(
    search="TODO",
    replacement="DONE",
    text_color="#FF0000",
    underline_type=None,
    underline_color=None,
    char_pr_id_ref=None,
    limit=None,
)
```

---

## 런 스타일 생성/조회

### ensure_run_style

기존 charPr 중 조건에 맞는 것을 찾거나, 없으면 새로 생성.

```python
char_pr_id = doc.ensure_run_style(
    bold=True,
    italic=False,
    underline=True,
    base_char_pr_id=None,    # 기반 charPr ID
)
# 반환: charPrIDRef 문자열 (예: "12")
```

### RunStyle 속성

```python
rs = doc.char_property("5")   # RunStyle
rs.id                         # "5"
rs.attributes                 # dict
rs.child_attributes           # dict (fontRef 등)
rs.text_color()               # "#000000" | None
rs.underline_type()           # "Bottom" | None
rs.underline_color()          # "#0000FF" | None
rs.matches(text_color="#FF0000", underline_type="Bottom")  # bool
```

---

## 표 (Table)

### 생성

```python
table = doc.add_table(
    rows=3,
    cols=4,
    section=None,
    section_index=None,
    width=None,           # 총 너비 (HwpUnit). None이면 기본값
    height=None,          # 행 높이
    border_fill_id_ref=None,
    para_pr_id_ref=None,
    style_id_ref=None,
    char_pr_id_ref=None,
    run_attributes=None,
)
# 반환: HwpxOxmlTable
```

### 표 속성/접근

```python
table.row_count           # 행 수
table.column_count        # 열 수
table.rows                # list[HwpxOxmlTableRow]

# 셀 접근
cell = table.cell(row_index, col_index)

# 셀 텍스트 설정 (⚠ logical/split_merged 주의)
table.set_cell_text(row, col, "텍스트",
                    logical=False,         # True: 병합 고려 논리 좌표
                    split_merged=False)     # True: 병합 셀 자동 분할

# 그리드 순회
for pos in table.iter_grid():
    print(pos.row, pos.column, pos.cell.text)
    # pos.anchor, pos.span 도 접근 가능

# 셀 맵 (2D 배열)
cell_map = table.get_cell_map()
# cell_map[row][col] → HwpxTableGridPosition
```

### 셀 조작

```python
cell.text               # 셀 텍스트 (get/set)
cell.address            # (row, col) 튜플
cell.span               # (row_span, col_span) 튜플
cell.width              # 셀 너비
cell.height             # 셀 높이

cell.set_span(row_span, col_span)
cell.set_size(width=7200, height=3600)
cell.remove()
```

### 셀 병합/분할

```python
# 병합 (start_row, start_col, end_row, end_col 포함)
merged_cell = table.merge_cells(0, 0, 0, 2)   # 첫 행 3열 병합

# 분할
split_cell = table.split_merged_cell(row, col)
```

---

## 머리말/꼬리말

```python
# 머리말 설정 (page_type: "BOTH", "ODD", "EVEN")
hf = doc.set_header_text("머리말 텍스트", section=None, section_index=None, page_type="BOTH")

# 꼬리말 설정
hf = doc.set_footer_text("꼬리말 텍스트", section=None, section_index=None, page_type="BOTH")

# 제거
doc.remove_header(section=None, section_index=None, page_type="BOTH")
doc.remove_footer(section=None, section_index=None, page_type="BOTH")
```

---

## 페이지 설정

### 구역(Section) 속성 접근

```python
section = doc.sections[0]
props = section.properties    # HwpxOxmlSectionProperties
```

### 용지 크기

```python
# 조회
ps = props.page_size
# ps.width, ps.height (HwpUnit), ps.orientation ("PORTRAIT"/"LANDSCAPE"), ps.gutter_type

# 설정
props.set_page_size(
    width=59528,        # A4 가로 210mm (HwpUnit = round(mm * 7200 / 25.4))
    height=84188,       # A4 세로 297mm
    orientation="PORTRAIT",
    gutter_type="LEFT_ONLY",
)
```

### 여백

```python
# 조회
pm = props.page_margins
# pm.left, pm.right, pm.top, pm.bottom, pm.header, pm.footer, pm.gutter (HwpUnit)

# 설정
props.set_page_margins(
    left=1800,
    right=1800,
    top=2000,
    bottom=2000,
    header=1000,
    footer=1000,
    gutter=0,
)
```

### 번호 매기기 시작값 (구역별)

```python
# 조회
sn = props.start_numbering
# sn.page_starts_on, sn.page, sn.picture, sn.table, sn.equation

# 설정
props.set_start_numbering(
    page_starts_on="BOTH",
    page=1,
    picture=1,
    table=1,
    equation=1,
)
```

---

## 문서 전체 번호 매기기 (헤더)

```python
header = doc.headers[0]

# 조회
bn = header.begin_numbering
# bn.page, bn.footnote, bn.endnote, bn.picture, bn.table, bn.equation

# 설정
header.set_begin_numbering(
    page=1,
    footnote=1,
    endnote=1,
    picture=1,
    table=1,
    equation=1,
)
```

---

## 메모 (⚠ lxml/ET 호환 버그로 사용 불가)

> **미지원**: `add_memo`, `add_memo_with_anchor`, `remove_memo` 등 메모 API는
> lxml/ET 타입 불일치 버그로 실행 시 `TypeError` 발생. 사용하지 않는다.

```python
# 메모 추가 (별도 영역에 생성)
memo = doc.add_memo(
    text="메모 내용",
    section=None,
    section_index=None,
    memo_shape_id_ref=None,
    memo_id=None,
    char_pr_id_ref=None,
    attributes=None,
)

# 메모 + 앵커 필드 동시 생성 (문단에 메모 마커 삽입)
memo, para, field_value = doc.add_memo_with_anchor(
    text="메모 내용",
    paragraph=None,          # 기존 문단 또는 None(새 문단 생성)
    section=None,
    section_index=None,
    paragraph_text=None,
    memo_shape_id_ref=None,
    memo_id=None,
    char_pr_id_ref=None,
    attributes=None,
    field_id=None,
    author=None,
    created=None,
    number=1,
    anchor_char_pr_id_ref=None,
)

# 메모 삭제
doc.remove_memo(memo)

# 메모 필드 부착 (기존 메모를 기존 문단에 연결)
field_val = doc.attach_memo_field(paragraph, memo, field_id=None, author=None, created=None, number=1)
```

---

## 도형/컨트롤

```python
# 도형 삽입 (제한적)
shape = doc.add_shape(
    shape_type="rect",
    section=None,
    section_index=None,
    attributes=None,
    para_pr_id_ref=None,
    style_id_ref=None,
    char_pr_id_ref=None,
    run_attributes=None,
)

# 컨트롤 삽입 (제한적)
ctrl = doc.add_control(
    section=None,
    section_index=None,
    attributes=None,
    control_type=None,
    para_pr_id_ref=None,
    style_id_ref=None,
    char_pr_id_ref=None,
    run_attributes=None,
)

# 속성 접근
shape.get_attribute("width")
shape.set_attribute("width", "12000")
```

---

## 패키지 저수준 접근

```python
pkg = doc.package

# 파일 목록
pkg.files()             # list[str]
pkg.part_names()        # 동일

# 파트 읽기/쓰기
data = pkg.read("BinData/image001.png")
pkg.write("BinData/new_image.png", image_bytes)
pkg.delete("BinData/old.png")

# XML 파트 접근
elem = pkg.get_xml("Contents/header.xml")
pkg.set_xml("Contents/header.xml", elem)

# 텍스트 파트
text = pkg.get_text("version.xml")

# 매니페스트
manifest = pkg.manifest_tree()

# 경로 목록
pkg.section_paths()     # 섹션 XML 경로 리스트
pkg.header_paths()      # 헤더 XML 경로 리스트
pkg.master_page_paths()
pkg.history_paths()
```

---

## 단위 참조

| 단위 | 설명 | 변환 |
|------|------|------|
| HwpUnit | 한/글 내부 단위 | 1mm ≈ 283 HwpUnit (= 7200 / 25.4) |
| pt | 포인트 | 1pt ≈ 100 HwpUnit |

### 주요 용지 크기 (HwpUnit)

> **참고**: `mm_to_hu()` 계산값과 python-hwpx 기본값이 ±1~3 차이날 수 있음 (≈0.01mm, 실용적으로 무방).

| 용지 | width | height |
|------|-------|--------|
| A4 세로 | 59528 | 84188 |
| A4 가로 | 84188 | 59528 |
| Letter 세로 | 61236 | 79092 |
| B5 세로 | 49892 | 70868 |

### 일반 여백 (HwpUnit)

| 여백 유형 | 좁게 | 보통 | 넓게 |
|-----------|------|------|------|
| 좌우 | 1134 | 2268 | 3402 |
| 상하 | 1134 | 1701 | 2835 |

> **참고**: python-hwpx의 `set_page_size()`, `set_page_margins()`에 전달하는 값은 모두 HwpUnit 정수.
> mm → HwpUnit 변환: `round(mm * 7200 / 25.4)` 또는 대략 `mm * 283`
