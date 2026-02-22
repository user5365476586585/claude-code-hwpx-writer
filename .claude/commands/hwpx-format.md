# HWP 글자/문단 직접 서식 (/hwpx-format)

python-hwpx로 글자 모양/문단 모양을 직접 설정하는 skill.
스타일 기반 서식이 우선이나, 특정 텍스트만 서식 변경할 때 직접 서식 사용.

사용자 요청: $ARGUMENTS

---

## 핵심 원칙

1. **스타일 우선**: 반복 사용할 서식은 스타일로 정의 → `/hwpx-style`
2. **직접 서식은 예외적 사용**: 특정 텍스트만 굵게/색상 변경 등 국소적 서식에 사용
3. **charPrIDRef 기반**: python-hwpx에서 글자 서식은 charPr ID를 통해 적용

---

## A. 글자 모양 (Run 서식)

### 기본 사용 — 런 단위 서식

```python
# 문단 추가 시 런 서식 지정
para = doc.add_paragraph("")
run = para.add_run("굵은 텍스트", bold=True)
run = para.add_run("기울임 텍스트", italic=True)
run = para.add_run("밑줄 텍스트", underline=True)
run = para.add_run("일반 텍스트")
```

### 기존 런 서식 변경

```python
para = doc.paragraphs[0]
for run in para.runs:
    run.bold = True           # 굵게
    run.italic = False        # 기울임 해제
    run.underline = True      # 밑줄
```

### charPrIDRef 기반 서식 (고급)

```python
# ensure_run_style: 조건에 맞는 charPr ID를 찾거나 새로 생성
bold_id = doc.ensure_run_style(bold=True)
italic_id = doc.ensure_run_style(italic=True)
bold_underline_id = doc.ensure_run_style(bold=True, underline=True)

# 문단 추가 시 charPr ID 지정
doc.add_paragraph("굵은 단락 전체", char_pr_id_ref=bold_id)

# 런 추가 시 charPr ID 지정
para = doc.add_paragraph("")
run = para.add_run("서식 텍스트", char_pr_id_ref=bold_id)
```

### Height(글자 크기) — charPr 속성

python-hwpx에서 글자 크기는 charPr의 `height` 속성 (HwpUnit, pt×100).

| pt | height 값 | pt | height 값 |
|----|-----------|-----|-----------|
| 8 | 800 | 16 | 1600 |
| 9 | 900 | 18 | 1800 |
| 10 | 1000 | 20 | 2000 |
| 11 | 1100 | 22 | 2200 |
| 12 | 1200 | 24 | 2400 |
| 14 | 1400 | 28 | 2800 |

---

## B. 글자 서식 — 헤더 XML 직접 수정 (글꼴/크기/색상)

python-hwpx의 `ensure_run_style()`은 bold/italic/underline만 지원.
글꼴, 크기, 색상 변경은 헤더 XML의 charPr 요소를 직접 수정해야 한다.

### ensure_char_property 패턴 (★ 고급 서식의 핵심)

```python
header = doc.headers[0]
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"

# 기존 charPr 중 조건에 맞는 것 검색, 없으면 새로 생성
char_pr_elem = header.ensure_char_property(
    predicate=lambda elem: (
        elem.get("height") == "1400" and
        elem.get("bold") == "1" and
        elem.get("textColor") == "#000000"
    ),
    modifier=lambda elem: (
        elem.set("height", "1400"),      # 14pt
        elem.set("bold", "1"),
        elem.set("textColor", "#000000"),
    ),
    base_char_pr_id=None,    # 복제 기반 charPr ID (None이면 첫 번째)
)
char_pr_id = char_pr_elem.get("id")
header.mark_dirty()

# 사용
doc.add_paragraph("14pt 굵은 텍스트", char_pr_id_ref=char_pr_id)
```

### 색상 변경 패턴

```python
header = doc.headers[0]

# 빨간색 굵은 charPr 생성
char_pr = header.ensure_char_property(
    predicate=lambda e: e.get("textColor") == "#FF0000" and e.get("bold") == "1",
    modifier=lambda e: (e.set("textColor", "#FF0000"), e.set("bold", "1")),
)
red_bold_id = char_pr.get("id")
header.mark_dirty()

# 사용
doc.add_paragraph("빨간 굵은 텍스트", char_pr_id_ref=red_bold_id)
```

### 보고서 자주 쓰는 색상

| 색상 | 코드 | 용도 |
|------|------|------|
| 검정 | `#000000` | 기본 |
| 진한 남색 | `#002060` | 제목 강조 |
| 진한 빨강 | `#C00000` | 경고/중요 |
| 진한 파랑 | `#004684` | 부제 |
| 회색 | `#808080` | 보조 텍스트 |

---

## C. 문단 모양 (paraPr)

### 문단속성 조회

```python
# 문단속성 목록
for pid, pp in doc.paragraph_properties.items():
    align = pp.align.horizontal if pp.align else "?"
    ls = pp.line_spacing.value if pp.line_spacing else "?"
    margin_l = pp.margin.left if pp.margin else 0
    print(f"  paraPr[{pid}] align={align} lineSpacing={ls} marginL={margin_l}")
```

### paraPrIDRef로 문단 서식 적용

```python
# 기존 paraPr ID를 문단에 적용
doc.add_paragraph("양쪽 정렬 170% 문단", para_pr_id_ref="2")

# 기존 문단의 paraPr 변경
para = doc.paragraphs[0]
para.para_pr_id_ref = "3"    # 다른 문단속성 적용
```

### 문단 서식 XML 직접 생성

```python
from copy import deepcopy

header = doc.headers[0]
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"

# 기존 paraPr 복제 + 수정
para_props = header.element.find(f".//{HH}paraProperties")
existing = para_props.findall(f"{HH}paraPr")
new_pp = deepcopy(existing[0])

# 새 ID 할당
new_id = str(max(int(e.get("id", "0")) for e in existing) + 1)
new_pp.set("id", new_id)

# 정렬 변경
align = new_pp.find(f"{HH}align")
if align is not None:
    align.set("horizontal", "Center")    # Center/Justify/Left/Right

# 줄간격 변경
ls = new_pp.find(f"{HH}lineSpacing")
if ls is not None:
    ls.set("type", "Percent")
    ls.set("value", "170")

# 여백 변경 (HwpUnit)
margin = new_pp.find(f"{HH}margin")
if margin is not None:
    margin.set("left", "1417")     # 약 5mm (5 × 283)
    margin.set("right", "0")
    margin.set("intent", "0")

para_props.append(new_pp)
para_props.set("itemCnt", str(len(para_props.findall(f"{HH}paraPr"))))
header.mark_dirty()

# 사용
doc.add_paragraph("가운데 170% 문단", para_pr_id_ref=new_id)
```

### AlignType 값

| 값 | 정렬 |
|-----|------|
| `"Justify"` | 양쪽 정렬 (기본) |
| `"Left"` | 왼쪽 정렬 |
| `"Right"` | 오른쪽 정렬 |
| `"Center"` | 가운데 정렬 |
| `"Distribute"` | 배분 정렬 |

### 줄간격 설정

| type | 의미 | value 예시 |
|------|------|-----------|
| `"Percent"` | 비율 (기본) | "160" = 160% |
| `"Fixed"` | 고정값 (HwpUnit) | "1800" = 약 18pt |
| `"BetweenLines"` | 줄 사이 | "500" |
| `"AtLeast"` | 최소 | "1500" |

---

## D. 서식 복사 패턴

특정 문단/런의 서식을 다른 문단/런에 복사.

```python
# 원본 문단의 서식 ID 저장
source_para = doc.paragraphs[0]
source_style = source_para.style_id_ref
source_para_pr = source_para.para_pr_id_ref

# 원본 런의 서식 ID 저장
source_run = source_para.runs[0]
source_char_pr = source_run.char_pr_id_ref

# 대상 문단에 복사
target_para = doc.paragraphs[5]
target_para.style_id_ref = source_style
target_para.para_pr_id_ref = source_para_pr
for run in target_para.runs:
    run.char_pr_id_ref = source_char_pr
```

---

## E. 런 분할 서식 적용 (범위 서식)

문단 내 특정 문자 범위에만 서식을 적용하려면 런을 분할해야 한다.

```python
from copy import deepcopy

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"

def clone_run_after(paragraph, source_run, text):
    """source_run을 복제하고 text로 교체하여 바로 뒤에 삽입"""
    source_elem = source_run.element
    parent = source_elem.getparent() if hasattr(source_elem, "getparent") else paragraph.element
    new_elem = deepcopy(source_elem)
    for child in list(new_elem):
        if child.tag == f"{HP}t":
            new_elem.remove(child)
    import xml.etree.ElementTree as ET
    text_node = ET.SubElement(new_elem, f"{HP}t")
    text_node.text = text
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
        if overlap_start > run_start:
            left_text = run.text[:overlap_start - run_start]
            rest_text = run.text[overlap_start - run_start:]
            run.text = left_text
            clone_run_after(para, run, rest_text)
            continue
        if overlap_end < run_end:
            keep_text = run.text[:overlap_end - run_start]
            tail_text = run.text[overlap_end - run_start:]
            run.text = keep_text
            clone_run_after(para, run, tail_text)
        if bold is not None:
            run.bold = bold
        if italic is not None:
            run.italic = italic
        if underline is not None:
            run.underline = underline
    para.section.mark_dirty()
```

---

## F. 특정 텍스트만 서식 변경 패턴

```python
# "중요" 단어를 찾아서 굵게 처리
for run in doc.iter_runs():
    if "중요" in run.text:
        run.bold = True

# 스타일 조건으로 런 검색
red_runs = doc.find_runs_by_style(text_color="#FF0000")
for run in red_runs:
    print(f"빨간 텍스트: {run.text}")
```

---

## G. 문자속성(charPr) 조회

```python
# 전체 문자속성 목록
for cid, cp in doc.char_properties.items():
    height = cp.attributes.get("height", "?")
    bold = cp.attributes.get("bold", "0")
    color = cp.text_color() or "#000000"
    ul = cp.underline_type() or "None"
    print(f"  charPr[{cid}] height={height} bold={bold} color={color} underline={ul}")

# 특정 charPr 조회
cp = doc.char_property("5")
if cp:
    print(f"  ID={cp.id}, color={cp.text_color()}, bold={cp.attributes.get('bold')}")
    # 스타일 매칭 확인
    print(f"  빨간색 여부: {cp.matches(text_color='#FF0000')}")
```

---

## 참조

- charPr/paraPr XML 구조: `references/hwpx-xml-schema.md`
- ensure_run_style / ensure_char_property: `references/python-hwpx-oxml.md`
- RunStyle 속성: `references/python-hwpx-api.md`
