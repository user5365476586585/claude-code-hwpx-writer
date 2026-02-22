# HWP 표 조작 (/hwpx-table)

python-hwpx로 HWP 문서의 표를 생성/편집/서식/데이터입력하는 skill.

사용자 요청: $ARGUMENTS

---

## ⛔ 핵심 규칙

1. **XML 직접 조작**: python-hwpx는 표를 XML로 직접 조작하므로 "표 안 갇힘" 문제 없음
2. **셀 좌표 0-based**: `table.cell(row, col)` — 행/열 모두 0부터 시작
3. **표 너비**: `width` 하드코딩 금지. `get_body_width(doc)` 또는 `add_styled_table()` 헬퍼 사용 (CLAUDE.md 공통 헬퍼 참조)
4. **셀 내부 paraPr**: 표 생성 시 셀 내부 문단의 `paraPrIDRef`가 기본값("0")으로 설정됨. 기존 문서에서 paraPr[0]이 RIGHT 정렬인 경우 표 내용이 오른쪽 정렬됨 → `add_styled_table(cell_style=...)` 헬퍼로 셀 내부 paraPr 일괄 적용 필수

---

## A. 표 생성

### 기본 표 생성

```python
# CLAUDE.md 공통 헬퍼 사용 (★ 권장)
# container_style: 표 컨테이너 문단 스타일 (가운데 정렬 권장)
# cell_style: 셀 내부 문단 스타일 (가운데 정렬 권장)
table = add_styled_table(doc, 3, 4, container_style="표위치", cell_style="표내용")

# 새 문서 (표 관련 스타일 없는 경우)
table = add_styled_table(doc, 3, 4)  # 기본 바탕글 스타일 사용

# 저수준 API (width 수동 지정 필요)
body_w = get_body_width(doc)
table = doc.add_table(rows=3, cols=4, width=body_w)
```

### 데이터로 표 생성

```python
# 데이터 정의
headers = ["구분", "2023", "2024", "증감"]
data = [
    ["매출", "100", "120", "+20%"],
    ["영업이익", "15", "18", "+20%"],
]

# 표 생성 (헤더 + 데이터 행)
rows = 1 + len(data)
cols = len(headers)
table = doc.add_table(rows=rows, cols=cols, width=54000)

# 헤더 입력
for c, h in enumerate(headers):
    table.set_cell_text(0, c, h)

# 데이터 입력
for r, row in enumerate(data):
    for c, val in enumerate(row):
        table.set_cell_text(r + 1, c, val)
```

### DataFrame으로 표 생성

```python
import pandas as pd

df = pd.DataFrame({
    "구분": ["매출", "영업이익"],
    "2023": [100, 15],
    "2024": [120, 18],
})

# DataFrame → 표
cols = list(df.columns)
rows = len(df) + 1  # 헤더 포함
table = doc.add_table(rows=rows, cols=len(cols), width=54000)

# 헤더
for c, col_name in enumerate(cols):
    table.set_cell_text(0, c, str(col_name))

# 데이터
for r in range(len(df)):
    for c in range(len(cols)):
        table.set_cell_text(r + 1, c, str(df.iloc[r, c]))
```

---

## A-1. 표 기본 스타일 (기획 단계에서 사용자에게 안내)

**기본 적용 항목** (기획 단계에서 안내, 오버라이드 가능):

| 항목 | 기본값 | 오버라이드 |
|------|--------|---------|
| 헤더 볼드 | 헤더 행 런에 bold 적용 | 생략 가능 |
| 행 높이 균등화 | **요청 시만** | `cell.set_size()` |
| 열 너비 비율 | **요청 시만** | `cell.set_size()` |

### 표준 생성 + 헤더 서식 패턴

```python
HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"

# 표 생성
table = doc.add_table(rows=4, cols=4, width=54000)

# 헤더 입력 + 볼드
headers = ["구분", "2022", "2023", "증감률"]
bold_id = doc.ensure_run_style(bold=True)
for c, h in enumerate(headers):
    table.set_cell_text(0, c, h)
    # 셀의 문단 런에 bold 적용
    cell = table.cell(0, c)
    for p_elem in cell.element.iter(f"{HP}p"):
        for run_elem in p_elem.iter(f"{HP}run"):
            run_elem.set("charPrIDRef", bold_id)

# 데이터 입력
data = [
    ["메모리", "1,580", "1,298", "-17.8"],
    ["로직", "1,875", "1,920", "+2.4"],
    ["합계", "3,455", "3,218", "-6.9"],
]
for r, row in enumerate(data):
    for c, val in enumerate(row):
        table.set_cell_text(r + 1, c, val)
```

---

## B. 셀 접근/내용 입력

### 셀 접근

```python
# 물리 좌표 (0-based)
cell = table.cell(0, 0)         # 1행 1열
cell = table.cell(2, 3)         # 3행 4열

# 셀 텍스트 설정 (⚠ logical/split_merged 옵션 주의)
table.set_cell_text(0, 0, "텍스트")
table.set_cell_text(0, 0, "텍스트", logical=True)        # 병합 고려 논리 좌표
table.set_cell_text(0, 0, "텍스트", split_merged=True)   # 병합 셀 자동 분할
```

### 셀 속성

```python
cell = table.cell(0, 0)

cell.text                # 셀 텍스트 (get/set)
cell.address             # (row, col) 물리 좌표
cell.span                # (row_span, col_span)
cell.width               # 셀 너비 (HwpUnit)
cell.height              # 셀 높이 (HwpUnit)
```

### 순차 입력 패턴

```python
data = [
    ["구분", "2023", "2024", "증감"],
    ["매출", "100", "120", "+20%"],
    ["영업이익", "15", "18", "+20%"],
]

for r, row in enumerate(data):
    for c, val in enumerate(row):
        table.set_cell_text(r, c, str(val))
```

---

## C. 그리드 순회 (병합 고려)

### iter_grid

```python
for pos in table.iter_grid():
    print(f"({pos.row},{pos.column}) text={pos.cell.text}")
    print(f"  anchor={pos.anchor}, span={pos.span}")
```

### get_cell_map (2D 배열)

```python
cell_map = table.get_cell_map()
# cell_map[row][col] → HwpxTableGridPosition
# 병합 영역은 같은 cell 객체 공유

for r, row in enumerate(cell_map):
    for c, pos in enumerate(row):
        print(f"  [{r},{c}] text={pos.cell.text} span={pos.span}")
```

---

## D. 행/열 조작

### 셀 합치기 (★ pyhwpx 한계 해결)

```python
# 셀 병합 (start_row, start_col, end_row, end_col — 모두 포함)
merged_cell = table.merge_cells(0, 0, 0, 2)    # 첫 행 3열 병합
print(f"병합 후: span={merged_cell.span}")
```

> **⚠ 셀 분할 (`split_merged_cell`)은 lxml/ET 호환 버그로 사용 불가.**
> 병합은 가능하나 분할은 불가. 병합 전에 구조를 확정할 것.

### 셀 크기 조정

```python
def mm_to_hu(mm):
    return round(mm * 7200 / 25.4)

# 셀 크기 설정
cell = table.cell(0, 0)
cell.set_size(width=14400, height=3600)    # HwpUnit

# 행 높이 균등화 (수동)
target_height = mm_to_hu(8)    # 8mm
for row in table.rows:
    for cell in row.cells:
        cell.set_size(height=target_height)
```

### 열 너비 자동 조정 (텍스트 길이 기반)

```python
def auto_fit_columns(table):
    """열 너비를 텍스트 길이에 비례하여 조정"""
    CHAR_UNIT = 720
    PADDING = 2
    MIN_WIDTH = 3600
    MAX_WIDTH = 36000

    col_count = table.column_count
    max_lens = [0] * col_count

    for pos in table.iter_grid():
        if pos.span[1] == 1:  # colspan=1인 셀만
            text_len = sum(2 if ord(c) > 127 else 1 for c in pos.cell.text)
            max_lens[pos.column] = max(max_lens[pos.column], text_len)

    # 각 열 너비 계산
    widths = []
    for length in max_lens:
        w = (length + PADDING) * CHAR_UNIT
        w = max(MIN_WIDTH, min(MAX_WIDTH, w))
        widths.append(w)

    # 적용
    for pos in table.iter_grid():
        if pos.span[1] >= 1:
            cell = table.cell(pos.row, pos.column)
            cell.set_size(width=widths[pos.column])

    table.mark_dirty()
    return widths

# 사용
auto_fit_columns(table)
```

---

## E. 표 → 데이터 추출

### 표 텍스트 추출

```python
# 표 찾기
for para in doc.paragraphs:
    for tbl in para.tables:
        print(f"표: {tbl.row_count}행 × {tbl.column_count}열")
        for pos in tbl.iter_grid():
            print(f"  ({pos.row},{pos.column}) {pos.cell.text}")
```

### 표 → 2D 리스트

```python
def table_to_list(table):
    """표를 2D 리스트로 변환"""
    cell_map = table.get_cell_map()
    result = []
    for row in cell_map:
        result.append([pos.cell.text for pos in row])
    return result

# 사용
data = table_to_list(table)
for row in data:
    print(row)
```

### 표 → DataFrame

```python
import pandas as pd

def table_to_df(table, header_row=0):
    """표를 DataFrame으로 변환"""
    data = table_to_list(table)
    if not data:
        return pd.DataFrame()
    headers = data[header_row]
    rows = data[header_row + 1:]
    return pd.DataFrame(rows, columns=headers)

df = table_to_df(table)
print(df)
```

---

## F. 캡션 + 출처 통합 패턴

보고서에서 표 + 캡션 + 출처를 작성하는 전체 패턴.

```python
P = add_styled_paragraph  # CLAUDE.md 공통 헬퍼

# 1. 캡션 입력 (표 위) — 기존 문서에 '표 제목' 스타일 있으면 사용
P(doc, "<표 1-1> 글로벌 반도체 시장 규모", "표 제목")

# 2. (단위 표기 - 필요시) — '표,그림의 단위' 스타일 사용
P(doc, "(단위: 억 달러, %)", "표,그림의 단위")

# 3. 표 생성 — add_styled_table 사용 (너비 자동 + 셀 paraPr 적용)
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

# 5. 출처 입력
doc.add_paragraph("자료: WSTS(2024), World Semiconductor Trade Statistics",
                   char_pr_id_ref=normal_id)

# 6. 본문 복귀
doc.add_paragraph("", style_id_ref="0")   # 바탕글
```

> **python-hwpx 장점**: 표 안 갇힘(is_cell) 문제가 없으므로 G3 검증 게이트 불필요.

---

## G. 셀 서식

### 셀 텍스트 서식

```python
HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"

# 셀 내부 런의 charPrIDRef 변경
cell = table.cell(0, 0)
for p_elem in cell.element.iter(f"{HP}p"):
    for run_elem in p_elem.iter(f"{HP}run"):
        run_elem.set("charPrIDRef", bold_id)

table.mark_dirty()
```

### 셀 배경색 (borderFill XML 수정)

```python
from copy import deepcopy

HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"

def create_fill_border(doc, fill_color="#D9D9D9"):
    """배경색 있는 borderFill 생성"""
    header = doc.headers[0]
    bf_list = header.element.find(f".//{HH}borderFills")
    existing = bf_list.findall(f"{HH}borderFill")

    new_bf = deepcopy(existing[0])
    new_id = str(max(int(e.get("id", "0")) for e in existing) + 1)
    new_bf.set("id", new_id)

    # fillBrush 추가
    fill_brush = new_bf.find(f"{HH}fillBrush")
    if fill_brush is None:
        import xml.etree.ElementTree as ET
        fill_brush = ET.SubElement(new_bf, f"{HH}fillBrush")
    solid = fill_brush.find(f"{HH}solidBrush")
    if solid is None:
        import xml.etree.ElementTree as ET
        solid = ET.SubElement(fill_brush, f"{HH}solidBrush")
    solid.set("color", fill_color)

    bf_list.append(new_bf)
    bf_list.set("itemCnt", str(len(bf_list.findall(f"{HH}borderFill"))))
    header.mark_dirty()
    return new_id

# 헤더 행 배경색 적용
fill_bf_id = create_fill_border(doc, "#D9D9D9")
for c in range(table.column_count):
    cell = table.cell(0, c)
    cell_bf = cell.element.find(f"{HP}cellBorderFill")
    if cell_bf is not None:
        cell_bf.set("borderFillIDRef", fill_bf_id)
table.mark_dirty()
```

---

## H. 표 속성

```python
table.row_count           # 행 수
table.column_count        # 열 수
table.rows                # list[HwpxOxmlTableRow]
table.element             # ET.Element (원시 XML)

# 행별 셀 접근
for row in table.rows:
    for cell in row.cells:
        print(f"  {cell.address}: {cell.text}")
```

---

## I. 블록 계산식 대안

HWP 블록 계산식 대신 Python으로 계산 후 결과 입력.

```python
# Python으로 계산
values = [100, 120, 150, 80]
total = sum(values)
avg = total / len(values)

# 결과 입력
table.set_cell_text(4, 1, f"{total:,}")      # 합계 셀
table.set_cell_text(5, 1, f"{avg:,.1f}")      # 평균 셀
```

---

## 참조

- HwpxOxmlTable API: `references/python-hwpx-oxml.md`
- 표 XML 구조: `references/hwpx-xml-schema.md`
- HwpxDocument.add_table(): `references/python-hwpx-api.md`
