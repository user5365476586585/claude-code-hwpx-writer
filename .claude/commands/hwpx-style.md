# HWP 스타일 관리 (/hwpx-style)

python-hwpx로 HWP 문서의 스타일을 관리하는 skill.
스타일 조회/생성/적용/삭제 및 보고서 표준 프리셋 제공.

사용자 요청: $ARGUMENTS

---

## 핵심 원칙

1. **스타일 우선**: 직접 서식(charPr, paraPr) 대신 스타일 기반 서식 사용 권장
2. **온디맨드**: 스타일은 처음 필요해지는 시점에 생성/수정 (사전 일괄 설정 지양)
3. **기존 스타일 존중**: 기존 문서의 스타일을 먼저 확인하고, 부적합할 때만 수정
4. **style_id_ref 직접 설정**: pyhwpx의 `set_style()` 3번 연타 버그 없음 — 직접 ID 지정

---

## A. 스타일 조회

### 전체 스타일 목록

```python
# 딕셔너리 (id → Style)
for sid, s in doc.styles.items():
    print(f"  [{sid}] {s.name} (eng: {s.eng_name}, type: {s.type})")
    print(f"       paraPrIDRef={s.para_pr_id_ref}, charPrIDRef={s.char_pr_id_ref}")
    print(f"       nextStyleIDRef={s.next_style_id_ref}")
```

### 특정 스타일 조회

```python
style = doc.style("0")    # ID로 조회
if style:
    print(f"이름: {style.name}, 영문: {style.eng_name}")
    print(f"문단속성: {style.para_pr_id_ref}, 문자속성: {style.char_pr_id_ref}")
```

### 스타일 갭 분석 패턴

```python
# 현재 문서에 필요한 스타일이 있는지 확인
needed = {"바탕글", "개요 1", "개요 2", "개요 3"}
existing = {s.name: sid for sid, s in doc.styles.items()}

for name in needed:
    if name in existing:
        print(f"  [존재] {name} (ID: {existing[name]})")
    else:
        print(f"  [없음] {name} → 생성 필요 또는 직접 서식 대체")
```

### 스타일별 서식 분석 (XML 기반)

pyhwpx의 probe 방식 대신, XML을 직접 파싱하여 스타일 서식을 분석한다.

```python
def analyze_styles(doc):
    """문서의 모든 스타일과 연결된 서식 정보를 분석"""
    results = []
    for sid, style in doc.styles.items():
        info = {
            'id': sid,
            'name': style.name,
            'eng_name': style.eng_name,
            'type': style.type,
        }
        # 연결된 문자속성
        cp = doc.char_property(style.char_pr_id_ref)
        if cp:
            info['height'] = cp.attributes.get('height', '?')
            info['bold'] = cp.attributes.get('bold', '0')
            info['text_color'] = cp.text_color() or '#000000'
        # 연결된 문단속성
        pp = doc.paragraph_property(style.para_pr_id_ref)
        if pp:
            info['align'] = pp.align.horizontal if pp.align else '?'
            info['line_spacing'] = pp.line_spacing.value if pp.line_spacing else '?'
            info['margin_left'] = pp.margin.left if pp.margin else 0
        results.append(info)
    return results

# 사용
analysis = analyze_styles(doc)
for info in analysis:
    h = int(info.get('height', 0)) // 100 if info.get('height', '?') != '?' else '?'
    bold = '○' if info.get('bold') == '1' else '×'
    print(f"  [{info['id']}] {info['name']}: {h}pt, 굵게={bold}, 정렬={info.get('align','?')}")
```

### 자동 기호(글머리) 판별 — XML heading 분석 (★ 핵심)

pyhwpx의 probe 방식 대신 XML을 직접 파싱. heading.type=BULLET이면 자동 기호.

```python
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"

def detect_auto_symbols(doc):
    """각 스타일의 자동 기호 여부를 XML heading 분석으로 판별.
    Returns: {style_name: {'auto': bool, 'bullet_char': str|None}}
    """
    header = doc.headers[0]
    pp_list = header.element.find(f".//{HH}paraProperties")

    # paraPr별 heading 정보 수집
    heading_map = {}
    for pp in pp_list.findall(f"{HH}paraPr"):
        pid = pp.get("id", "0")
        h = pp.find(f"{HH}heading")
        if h is not None:
            heading_map[pid] = h.get("type", "NONE")

    # bullet 정의 수집
    bullet_chars = {}
    for bid, bullet in doc.bullets.items():
        bullet_chars[bid] = bullet.char

    results = {}
    for sid, s in doc.styles.items():
        pp_id = str(s.para_pr_id_ref or "0")
        htype = heading_map.get(pp_id, "NONE")
        is_auto = htype == "BULLET"

        # heading의 bulletIdRef에서 기호 문자 조회
        bullet_char = None
        if is_auto:
            pp_elem = None
            for pp in pp_list.findall(f"{HH}paraPr"):
                if pp.get("id") == pp_id:
                    pp_elem = pp
                    break
            if pp_elem is not None:
                h = pp_elem.find(f"{HH}heading")
                if h is not None:
                    bid = h.get("idRef", "0")
                    bullet_char = bullet_chars.get(bid)

        results[s.name] = {'auto': is_auto, 'bullet_char': bullet_char}

    return results

# 사용
auto_info = detect_auto_symbols(doc)
for name, info in auto_info.items():
    if info['auto']:
        char_repr = repr(info['bullet_char']) if info['bullet_char'] else "?"
        print(f"  [자동] {name:<15} bullet={char_repr}")
    # 수동인 경우 생략 (기본이 수동이므로)
```

**판별 결과 활용**:
- `auto=True` → `add_styled_paragraph(doc, "내용만", style_name)` (기호 제외)
- `auto=False` → `add_styled_paragraph(doc, "□ 내용", style_name)` (기호 포함)

### 문서 내 실제 사용된 스타일 파악

```python
used_styles = set()
for para in doc.paragraphs:
    used_styles.add(para.style_id_ref)

print("사용된 스타일 ID:", used_styles)
for sid in used_styles:
    s = doc.style(sid)
    if s:
        print(f"  [{sid}] {s.name}")
```

---

## B. 스타일 적용

### style_id_ref로 적용 (★ 핵심)

```python
# 문단 생성 시 스타일 지정
doc.add_paragraph("이 단락은 개요 1 스타일", style_id_ref="1")
doc.add_paragraph("본문 내용", style_id_ref="0")

# 기존 문단의 스타일 변경
para = doc.paragraphs[0]
para.style_id_ref = "1"     # 개요 1로 변경
```

### 스타일 이름으로 ID 찾기

```python
def find_style_id(doc, name):
    """스타일 이름으로 ID 찾기"""
    for sid, s in doc.styles.items():
        if s.name == name:
            return sid
    return None

# 사용
outline1_id = find_style_id(doc, "개요 1")
if outline1_id:
    doc.add_paragraph("제1장 서론", style_id_ref=outline1_id)
```

---

## C. 스타일 생성 (★ pyhwpx 한계 해결)

python-hwpx에서는 헤더 XML에 `<hh:style>` 노드를 직접 추가하여 새 스타일을 생성할 수 있다.

```python
from copy import deepcopy

HH = "{http://www.hancom.co.kr/hwpml/2011/head}"

def create_style(doc, name, eng_name, *, para_pr_id_ref=None, char_pr_id_ref=None):
    """새 named 스타일 생성"""
    header = doc.headers[0]
    styles_elem = header.element.find(f".//{HH}styles")
    existing = styles_elem.findall(f"{HH}style")

    # 중복 확인
    for e in existing:
        if e.get("name") == name:
            return e.get("id")  # 이미 존재

    # 마지막 스타일 복제
    new_style = deepcopy(existing[-1])

    # 새 ID 할당
    new_id = str(max(int(e.get("id", "0")) for e in existing) + 1)
    new_style.set("id", new_id)
    new_style.set("name", name)
    new_style.set("engName", eng_name)
    new_style.set("type", "para")
    new_style.set("nextStyleIDRef", new_id)

    if para_pr_id_ref is not None:
        new_style.set("paraPrIDRef", str(para_pr_id_ref))
    if char_pr_id_ref is not None:
        new_style.set("charPrIDRef", str(char_pr_id_ref))

    styles_elem.append(new_style)
    styles_elem.set("itemCnt", str(len(styles_elem.findall(f"{HH}style"))))

    header.mark_dirty()
    return new_id

# 사용
new_id = create_style(doc, "□ 대주제", "BoxTitle",
                       para_pr_id_ref="2", char_pr_id_ref="3")
doc.add_paragraph("□ 글로벌 시장 현황", style_id_ref=new_id)
```

### 스타일 + charPr + paraPr 동시 생성

```python
from copy import deepcopy

HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
header = doc.headers[0]

# 1. charPr 생성 (14pt 굵게)
char_pr = header.ensure_char_property(
    predicate=lambda e: e.get("height") == "1400" and e.get("bold") == "1",
    modifier=lambda e: (e.set("height", "1400"), e.set("bold", "1")),
)
new_char_id = char_pr.get("id")

# 2. paraPr 생성 (가운데 160%) — /hwpx-format C절 패턴
para_props = header.element.find(f".//{HH}paraProperties")
existing_pp = para_props.findall(f"{HH}paraPr")
new_pp = deepcopy(existing_pp[0])
new_para_id = str(max(int(e.get("id", "0")) for e in existing_pp) + 1)
new_pp.set("id", new_para_id)
align = new_pp.find(f"{HH}align")
if align is not None:
    align.set("horizontal", "Center")
ls = new_pp.find(f"{HH}lineSpacing")
if ls is not None:
    ls.set("value", "160")
para_props.append(new_pp)
para_props.set("itemCnt", str(len(para_props.findall(f"{HH}paraPr"))))

# 3. 스타일 생성
new_style_id = create_style(doc, "장 제목", "ChapterTitle",
                             para_pr_id_ref=new_para_id,
                             char_pr_id_ref=new_char_id)

header.mark_dirty()
```

---

## D. 글머리표/번호 조회

```python
# 글머리표 목록
for bid, bullet in doc.bullets.items():
    print(f"  bullet[{bid}] char={bullet.char} checked={bullet.checked_char}")
    if bullet.para_head:
        print(f"    level={bullet.para_head.level} start={bullet.para_head.start}")

# 번호 매기기 (numberings)는 헤더 XML에서 직접 조회
header = doc.headers[0]
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
numberings = header.element.find(f".//{HH}numberings")
if numberings is not None:
    for num in numberings:
        print(f"  numbering: id={num.get('id')}")
```

---

## E. 스타일 삭제/정리

```python
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"

header = doc.headers[0]
styles_elem = header.element.find(f".//{HH}styles")

# 특정 스타일 삭제
for style_elem in styles_elem.findall(f"{HH}style"):
    if style_elem.get("name") == "삭제할스타일":
        styles_elem.remove(style_elem)
        break

# itemCnt 업데이트
styles_elem.set("itemCnt", str(len(styles_elem.findall(f"{HH}style"))))
header.mark_dirty()
```

---

## F. 배치 스타일 적용 패턴

### 방법 1: 기존 스타일 ID로 직접 적용

```python
# 스타일 ID 매핑 (문서마다 다를 수 있음)
style_map = {}
for sid, s in doc.styles.items():
    style_map[s.name] = sid

# 장 제목
doc.add_paragraph("제1장 서론", style_id_ref=style_map.get("개요 1", "0"))

# 본문
doc.add_paragraph("본문 내용", style_id_ref=style_map.get("바탕글", "0"))
```

### 방법 2: charPr/paraPr 직접 서식 (스타일 없이)

```python
# charPr 준비
bold_id = doc.ensure_run_style(bold=True)
normal_id = doc.ensure_run_style(bold=False)

# □ 수준 (굵게)
doc.add_paragraph("□ 글로벌 시장 현황", char_pr_id_ref=bold_id)

# ○ 수준 (일반)
doc.add_paragraph("○ 세부 내용", char_pr_id_ref=normal_id)

# - 수준
doc.add_paragraph("- 구체 데이터", char_pr_id_ref=normal_id)
```

---

## G. 한글 보고서 표준 스타일 프리셋

### 프리셋 A: 실전 보고서 (장/절 + □○-·※가. 체계) ★

한국 정책·연구보고서에서 흔히 쓰는 번호+기호 혼합 계위.

| 기호 | 계위 의미 | 용도 |
|------|---------|------|
| 제N장 | 대단위 구분 | 보고서 최상위 구조 분절 |
| N.1 / N.1.1 | 절/소절 | 장 아래 논리 단위 |
| □ | 절 아래 독립 논점 | 굵게 처리로 시각적 강조 |
| ○ | □ 아래 세부 논거 | □의 뒷받침 |
| - | ○ 아래 구체 수치 | 조사 결과, 통계, 인용 |
| · | - 아래 부연 | 세분화 필요 시만 |
| ※ | 계층 무관, 보충 | 참고사항, 단서, 출처 |
| 가. | □/○ 아래 병렬 | 가.나.다. 순 열거 |

**개요 1~3**: `style_id_ref` 적용 (기본 내장 스타일).
**□ ○ - · ※ 가.**: charPr/paraPr + 기호 직접 포함.

```python
# 스타일 ID 확인
style_map = {s.name: sid for sid, s in doc.styles.items()}

# 개요 스타일
doc.add_paragraph("제1장 서론", style_id_ref=style_map.get("개요 1", "0"))
doc.add_paragraph("1.1 배경", style_id_ref=style_map.get("개요 2", "0"))

# □ ○ - · ※ 기호 수준: charPr 사용
bold_id = doc.ensure_run_style(bold=True)
normal_id = doc.ensure_run_style(bold=False)

doc.add_paragraph("□ 글로벌 반도체 시장 현황", char_pr_id_ref=bold_id)
doc.add_paragraph("○ TSMC, 삼성전자가 공급을 주도", char_pr_id_ref=normal_id)
doc.add_paragraph("- (점유율) TSMC 54.4%, 삼성전자 16.5%", char_pr_id_ref=normal_id)
doc.add_paragraph("※ 2024년 기준, IDM 자체생산분 제외", char_pr_id_ref=normal_id)

# 가. 나열: 수동 번호 직접 입력
doc.add_paragraph("가. TSMC 공급망 전략", char_pr_id_ref=normal_id)
doc.add_paragraph("나. 삼성전자 파운드리 현황", char_pr_id_ref=normal_id)
```

### 프리셋 B: 간단 보고서

| 스타일 | 크기 | 굵기 | 정렬 | 줄간격 |
|--------|------|------|------|--------|
| 개요 1 | 18pt | ○ | 가운데 | 160% |
| 개요 2 | 14pt | ○ | 왼쪽 | 160% |
| 개요 3 | 12pt | ○ | 왼쪽 | 160% |
| 바탕글 | 11pt | × | 양쪽 | 170% |

### 프리셋 C: 공문서

| 스타일 | 크기 | 굵기 | 정렬 | 줄간격 | 비고 |
|--------|------|------|------|--------|------|
| 개요 1 | 16pt | ○ | 가운데 | 160% | 제목 |
| 개요 2 | 12pt | ○ | 왼쪽 | 160% | 소제목 |
| 바탕글 | 10pt | × | 양쪽 | 160% | 본문 |

---

## H. 주의사항

### 스타일 존재 확인 후 적용

```python
style_map = {s.name: sid for sid, s in doc.styles.items()}
target_name = "개요 1"

if target_name in style_map:
    doc.add_paragraph("제1장 서론", style_id_ref=style_map[target_name])
else:
    # 스타일 없음 → 직접 서식으로 대체
    bold_id = doc.ensure_run_style(bold=True)
    doc.add_paragraph("제1장 서론", char_pr_id_ref=bold_id)
```

### pyhwpx 대비 해결된 한계

| pyhwpx 한계 | python-hwpx 해결 |
|-------------|-----------------|
| `set_style()` 3번 연타 필요 | `style_id_ref` 직접 설정 — 버그 없음 |
| `export_style()`/`import_style()` COM 에러 | 헤더 XML `<hh:styles>` 직접 복사 |
| 새 named 스타일 생성 불가 | 헤더 XML에 `<hh:style>` 노드 추가 (C절) |

---

## 참조

- 스타일 XML 구조: `references/hwpx-xml-schema.md`
- charPr/paraPr 조작: `references/python-hwpx-oxml.md`
- HwpxDocument 스타일 API: `references/python-hwpx-api.md`
