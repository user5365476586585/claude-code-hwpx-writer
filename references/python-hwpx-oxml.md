# python-hwpx 저수준 XML 조작 API

python-hwpx의 OXML 래퍼 클래스들. `HwpxDocument`의 고수준 API로 부족할 때 사용.

---

## HwpxOxmlSection (구역)

```python
section = doc.sections[0]

section.element         # ET.Element (원시 XML)
section.properties      # HwpxOxmlSectionProperties
section.paragraphs      # list[HwpxOxmlParagraph]
section.memo_group      # HwpxOxmlMemoGroup | None
section.memos           # list[HwpxOxmlMemo]
section.dirty           # bool (수정 플래그)

# 문단 추가
para = section.add_paragraph(text, **kwargs)

# 메모 추가
memo = section.add_memo(text, ...)

# dirty 마킹 (XML 수동 수정 후 반드시 호출)
section.mark_dirty()

# 직렬화
xml_bytes = section.to_bytes()
```

---

## HwpxOxmlSectionProperties (구역 설정)

```python
props = section.properties

# 용지 크기
ps = props.page_size    # PageSize(width, height, orientation, gutter_type)
props.set_page_size(width=59528, height=84188, orientation="PORTRAIT", gutter_type="LEFT_ONLY")

# 여백
pm = props.page_margins  # PageMargins(left, right, top, bottom, header, footer, gutter)
props.set_page_margins(left=2268, right=2268, top=1701, bottom=1701, header=1134, footer=1134, gutter=0)

# 번호 매기기 시작값
sn = props.start_numbering  # SectionStartNumbering(page_starts_on, page, picture, table, equation)
props.set_start_numbering(page_starts_on="BOTH", page=1, picture=1, table=1, equation=1)

# 머리말/꼬리말
hdr = props.get_header("BOTH")         # HwpxOxmlSectionHeaderFooter | None
ftr = props.get_footer("ODD")
props.set_header_text("텍스트", "BOTH")
props.set_footer_text("텍스트", "BOTH")
props.remove_header("BOTH")
props.remove_footer("BOTH")
```

---

## HwpxOxmlParagraph (문단)

```python
para = doc.paragraphs[0]

# 속성
para.element            # ET.Element
para.section            # HwpxOxmlSection
para.text               # 전체 텍스트 (get/set)
para.runs               # list[HwpxOxmlRun]
para.tables             # list[HwpxOxmlTable]

# ID 참조 (get/set)
para.para_pr_id_ref     # 문단속성 ID
para.style_id_ref       # 스타일 ID
para.char_pr_id_ref     # 문자속성 ID (문단 기본 런 속성)

# 런 추가
run = para.add_run(text="텍스트", char_pr_id_ref=None, bold=False, italic=False, underline=False, attributes=None)

# 표 추가
table = para.add_table(rows, cols, width=None, height=None, border_fill_id_ref=None)

# 도형 추가
shape = para.add_shape(shape_type, attributes=None)

# 컨트롤 추가
ctrl = para.add_control(attributes=None, control_type=None)

# Body 모델 변환 (lxml 기반 구조적 편집용)
model = para.to_model()      # body.Paragraph 데이터클래스
para.apply_model(model)      # 모델 → XML 역반영
```

---

## HwpxOxmlRun (런)

```python
run = para.runs[0]

# 속성
run.element             # ET.Element
run.paragraph           # HwpxOxmlParagraph
run.text                # 텍스트 (get/set)
run.char_pr_id_ref      # 문자속성 ID (get/set)
run.bold                # bool (get/set)
run.italic              # bool (get/set)
run.underline           # bool (get/set)
run.style               # RunStyle | None

# 텍스트 치환
count = run.replace_text("old", "new", count=None)

# 런 삭제
run.remove()

# Body 모델 변환
model = run.to_model()       # body.Run 데이터클래스
run.apply_model(model)
```

---

## HwpxOxmlTable (표)

```python
table = para.tables[0]  # 또는 doc.add_table(...)

# 속성
table.element           # ET.Element
table.paragraph         # HwpxOxmlParagraph
table.row_count         # int
table.column_count      # int
table.rows              # list[HwpxOxmlTableRow]

# 셀 접근 (물리 좌표)
cell = table.cell(0, 0)

# 셀 텍스트 설정
table.set_cell_text(row, col, "텍스트", logical=False, split_merged=False)
# logical=True: 병합 셀 건너뛴 논리 좌표 사용
# split_merged=True: 병합 셀 자동 분할

# 그리드 순회 (병합 고려)
for pos in table.iter_grid():
    # pos: HwpxTableGridPosition(row, column, cell, anchor, span)
    print(f"({pos.row},{pos.column}) text={pos.cell.text}")

# 셀 맵 (2D 배열, 병합 영역은 같은 cell 객체 공유)
cell_map = table.get_cell_map()

# 셀 병합 (start_row, start_col, end_row, end_col — 모두 포함)
merged = table.merge_cells(0, 0, 0, 2)

# 셀 분할 (⚠ lxml/ET 호환 버그로 사용 불가)
# split = table.split_merged_cell(0, 0)  # TypeError 발생

# dirty 마킹
table.mark_dirty()

# XML 생성 (클래스 메서드)
elem = HwpxOxmlTable.create(rows=3, cols=4, width=54000, height=3600, border_fill_id_ref="1")
```

---

## HwpxOxmlTableCell (셀)

```python
cell = table.cell(0, 0)

cell.element            # ET.Element
cell.table              # HwpxOxmlTable
cell.address            # (row, col) 물리 좌표
cell.span               # (row_span, col_span)
cell.width              # 셀 너비 (HwpUnit)
cell.height             # 셀 높이 (HwpUnit)
cell.text               # 셀 텍스트 (get/set)

cell.set_span(row_span, col_span)
cell.set_size(width=7200, height=3600)
cell.remove()
```

---

## HwpxOxmlHeader (헤더)

문서 메타데이터, 스타일/속성 정의, 번호 매기기 등.

```python
header = doc.headers[0]

header.element          # ET.Element
header.dirty            # bool

# 번호 매기기
bn = header.begin_numbering  # DocumentNumbering(page, footnote, endnote, picture, table, equation)
header.set_begin_numbering(page=1, footnote=1, endnote=1, picture=1, table=1, equation=1)

# 속성 컬렉션 (딕셔너리, id → 객체)
header.border_fills     # dict[str, GenericElement]
header.memo_shapes      # dict[str, MemoShape]
header.bullets          # dict[str, Bullet]
header.paragraph_properties  # dict[str, ParagraphProperty]
header.styles           # dict[str, Style]
header.track_changes    # dict[str, TrackChange]
header.track_change_authors  # dict[str, TrackChangeAuthor]

# ID 조회
header.border_fill(id)
header.memo_shape(id)
header.bullet(id)
header.paragraph_property(id)
header.style(id)
header.track_change(id)
header.track_change_author(id)

# 문자속성 생성/조회 (★ 고급 서식의 핵심)
char_pr_elem = header.ensure_char_property(
    predicate=lambda elem: ...,     # 기존 속성 검색 조건
    modifier=lambda elem: ...,      # 새 속성 수정 함수
    base_char_pr_id=None,           # 기반 속성 ID
    preferred_id=None,              # 선호 ID
)
# 반환: ET.Element (<hh:charPr> 노드)

# 기본 테두리채우기 ID
bf_id = header.find_basic_border_fill_id()
bf_id = header.ensure_basic_border_fill()

# dirty 관리
header.mark_dirty()
header.reset_dirty()
```

---

## 헤더 데이터 모델 (oxml.header 모듈)

### Style

```python
style = doc.style("0")     # 또는 header.style("0")

style.id              # 정규화된 ID 문자열
style.raw_id          # 원본 ID
style.type            # 스타일 타입 ("para" 등)
style.name            # 한글 이름
style.eng_name        # 영문 이름
style.para_pr_id_ref  # 연결된 문단속성 ID
style.char_pr_id_ref  # 연결된 문자속성 ID
style.next_style_id_ref  # 다음 스타일 ID
style.lang_id         # 언어 ID
style.lock_form       # 잠금 여부
style.attributes      # 기타 속성 dict
```

### ParagraphProperty

```python
pp = doc.paragraph_property("0")

pp.id                 # ID
pp.raw_id             # 원본 ID
pp.tab_pr_id_ref      # 탭속성 ID
pp.condense           # 응축 여부
pp.font_line_height   # 글꼴 줄 높이
pp.snap_to_grid       # 격자에 맞춤
pp.align              # ParagraphAlignment(horizontal, vertical)
pp.heading            # ParagraphHeading(type, id_ref, level)
pp.break_setting      # ParagraphBreakSetting(...)
pp.margin             # ParagraphMargin(intent, left, right, prev, next)
pp.line_spacing       # ParagraphLineSpacing(spacing_type, value, unit)
pp.border             # ParagraphBorder(border_fill_id_ref, offsets...)
pp.auto_spacing       # ParagraphAutoSpacing(e_asian_eng, e_asian_num)
pp.attributes         # 기타 속성 dict
pp.other_children     # 기타 자식 요소
```

### CharProperty (RunStyle 원본)

```python
cp = doc.char_property("0")   # RunStyle 래퍼
cp.id                 # ID
cp.attributes         # dict (Height, Bold, Italic, textColor 등)
cp.child_attributes   # dict (fontRef, underline 등 하위 요소)
cp.text_color()       # "#000000" | None
cp.underline_type()   # "Bottom" | None
cp.underline_color()  # "#0000FF" | None
cp.matches(text_color="#FF0000")  # bool
```

### 기타 모델

```python
# MemoShape
ms = doc.memo_shape("0")
ms.id, ms.width, ms.line_width, ms.line_type, ms.line_color, ms.fill_color, ms.active_color, ms.memo_type

# Bullet
bullet = doc.bullet("0")
bullet.id, bullet.raw_id, bullet.char, bullet.checked_char, bullet.use_image
bullet.para_head  # BulletParaHead(text, level, start, align, ...)

# TrackChange
tc = doc.track_change("0")
tc.id, tc.raw_id, tc.change_type, tc.date, tc.author_id, tc.hide
```

---

## Body 데이터 모델 (oxml.body 모듈)

`para.to_model()` / `run.to_model()`로 접근하는 구조적 모델.

### Paragraph

```python
model = para.to_model()
model.tag              # 태그명
model.id               # 문단 ID
model.para_pr_id_ref   # 문단속성 ID
model.style_id_ref     # 스타일 ID
model.page_break       # "0" | "1"
model.column_break     # "0" | "1"
model.merged           # "0" | "1"
model.runs             # list[Run]
model.attributes       # dict
model.other_children   # list[GenericElement]
model.content          # list[ParagraphChild] (원래 순서 유지)
```

### Run

```python
run_model = model.runs[0]
run_model.tag              # 태그명
run_model.char_pr_id_ref   # 문자속성 ID
run_model.section_properties  # 구역 속성 (있으면)
run_model.controls         # list[Control]
run_model.tables           # list[Table]
run_model.inline_objects   # list[InlineObject]
run_model.text_spans       # list[TextSpan]
run_model.attributes       # dict
run_model.content          # list[RunChild]
```

### TextSpan

```python
ts = run_model.text_spans[0]
ts.tag                 # 태그명
ts.leading_text        # 선행 텍스트
ts.marks               # list[TextMarkup]
ts.text                # 전체 텍스트 (get/set)
ts.attributes          # dict
```

---

## XML 직접 조작 패턴

### 네임스페이스

```python
HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"   # 문단/런/텍스트
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"        # 헤더 (스타일/속성)
HC = "{http://www.hancom.co.kr/hwpml/2011/core}"         # 코어
HA = "{http://www.hancom.co.kr/hwpml/2011/app}"          # 앱
```

### section XML 직접 수정

```python
import xml.etree.ElementTree as ET

section = doc.sections[0]
root = section.element

# 문단 요소 접근
for p_elem in root.iter(f"{HP}p"):
    para_id = p_elem.get("id")
    style_ref = p_elem.get("styleIDRef")

# 수정 후 반드시 dirty 마킹
section.mark_dirty()
```

### 헤더 XML 직접 수정

```python
header = doc.headers[0]
root = header.element

# 스타일 목록 접근
styles_elem = root.find(f".//{HH}styles")
for style_elem in styles_elem.findall(f"{HH}style"):
    name = style_elem.get("name")
    style_id = style_elem.get("id")

# 새 스타일 추가
from copy import deepcopy
last_style = styles_elem.findall(f"{HH}style")[-1]
new_style = deepcopy(last_style)
new_id = str(max(int(s.get("id", "0")) for s in styles_elem.findall(f"{HH}style")) + 1)
new_style.set("id", new_id)
new_style.set("name", "사용자정의")
new_style.set("engName", "UserDefined")
styles_elem.append(new_style)
styles_elem.set("itemCnt", str(len(styles_elem.findall(f"{HH}style"))))

header.mark_dirty()
```

### lxml/ET 혼합 패치 (⛔ 필수)

python-hwpx는 lxml과 stdlib ET를 혼합 사용. `ET.SubElement(lxml_parent, ...)` 호출 시 TypeError 발생 가능.

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

이 패치는 스크립트 초기화 시 1회 적용.

---

## 런 분할 서식 적용 패턴

특정 문자 범위에만 서식을 적용하려면 런을 분할해야 한다.

```python
from copy import deepcopy

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"

def clone_run_after(paragraph, source_run, text):
    """source_run을 복제하고 text로 교체하여 바로 뒤에 삽입"""
    source_elem = source_run.element
    parent = source_elem.getparent() if hasattr(source_elem, "getparent") else paragraph.element
    new_elem = deepcopy(source_elem)
    # 기존 텍스트 노드 제거
    for child in list(new_elem):
        if child.tag == f"{HP}t":
            new_elem.remove(child)
    # 새 텍스트 노드
    text_node = ET.SubElement(new_elem, f"{HP}t")
    text_node.text = text
    # source 뒤에 삽입
    siblings = list(parent)
    insert_at = siblings.index(source_elem) + 1 if source_elem in siblings else len(siblings)
    parent.insert(insert_at, new_elem)
    return new_elem

def format_text_range(doc, para_index, start, end, *, bold=None, italic=None, underline=None):
    """문단 내 [start, end) 문자 범위에 서식 적용 (런 분할 방식)"""
    para = doc.paragraphs[para_index]
    cursor = 0
    for run in list(para.runs):
        run_len = len(run.text)
        run_start = cursor
        run_end = cursor + run_len
        cursor = run_end

        overlap_start = max(start, run_start)
        overlap_end = min(end, run_end)
        if overlap_start >= overlap_end:
            continue

        # 왼쪽 잔여 분리
        if overlap_start > run_start:
            left_text = run.text[:overlap_start - run_start]
            rest_text = run.text[overlap_start - run_start:]
            run.text = left_text
            new_elem = clone_run_after(para, run, rest_text)
            # 새 런에서 다시 처리
            continue

        # 오른쪽 잔여 분리
        if overlap_end < run_end:
            keep_text = run.text[:overlap_end - run_start]
            tail_text = run.text[overlap_end - run_start:]
            run.text = keep_text
            clone_run_after(para, run, tail_text)

        # 대상 런에 서식 적용
        if bold is not None:
            run.bold = bold
        if italic is not None:
            run.italic = italic
        if underline is not None:
            run.underline = underline

    para.section.mark_dirty()
```

---

## Cross-Run 텍스트 치환 패턴

하나의 검색어가 여러 런에 걸쳐 분할된 경우의 치환.

```python
def replace_across_runs(runs, find_text, replace_text):
    """여러 런에 걸친 텍스트를 치환"""
    merged = "".join(r.text for r in runs)
    if find_text not in merged:
        return 0
    new_merged = merged.replace(find_text, replace_text)
    # 원래 런 길이 비율로 재분배
    total = len(new_merged)
    weights = [len(r.text) for r in runs]
    weight_sum = sum(weights)
    if weight_sum == 0:
        return 0
    pos = 0
    for i, run in enumerate(runs):
        if i == len(runs) - 1:
            share = total - pos
        else:
            share = total * weights[i] // weight_sum
        run.text = new_merged[pos:pos + share]
        pos += share
    return merged.count(find_text)
```

---

## 상수

```python
# 기본 셀 크기
_DEFAULT_CELL_WIDTH = 7200
_DEFAULT_CELL_HEIGHT = 3600

# 기본 문단 속성
_DEFAULT_PARAGRAPH_ATTRS = {
    "paraPrIDRef": "0",
    "styleIDRef": "0",
    "pageBreak": "0",
    "columnBreak": "0",
    "merged": "0",
}
```
