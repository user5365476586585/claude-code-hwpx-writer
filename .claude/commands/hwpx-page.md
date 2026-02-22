# HWP 페이지 레이아웃 (/hwpx-page)

python-hwpx로 편집 용지/머리말/꼬리말/구역을 설정하는 skill.

사용자 요청: $ARGUMENTS

---

## A. 편집 용지 (페이지 설정)

### 현재 쪽 설정 조회

```python
section = doc.sections[0]
props = section.properties

# 용지 크기
ps = props.page_size
print(f"크기: {ps.width}x{ps.height} (orientation: {ps.orientation})")

# 여백
pm = props.page_margins
print(f"여백: 좌={pm.left} 우={pm.right} 상={pm.top} 하={pm.bottom}")
print(f"머리말={pm.header} 꼬리말={pm.footer} 제본={pm.gutter}")
```

### 쪽 설정 변경

```python
def mm_to_hu(mm):
    return round(mm * 7200 / 25.4)

section = doc.sections[0]
props = section.properties

# 여백 변경 (HwpUnit)
props.set_page_margins(
    left=mm_to_hu(20),      # 20mm
    right=mm_to_hu(20),
    top=mm_to_hu(20),
    bottom=mm_to_hu(15),
    header=mm_to_hu(15),
    footer=mm_to_hu(15),
    gutter=0,
)
```

### 용지 크기/방향 변경

```python
props = doc.sections[0].properties

# A4 세로
props.set_page_size(
    width=59528,      # A4 가로 (HwpUnit)
    height=84188,     # A4 세로 (HwpUnit)
    orientation="PORTRAIT",
    gutter_type="LEFT_ONLY",
)

# A4 가로 (landscape)
props.set_page_size(
    width=84188,
    height=59528,
    orientation="LANDSCAPE",
    gutter_type="LEFT_ONLY",
)
```

### 용지 크기 참조 (HwpUnit)

| 용지 | width | height | mm (가로×세로) |
|------|-------|--------|---------------|
| A4 세로 | 59528 | 84188 | 210×297 |
| A4 가로 | 84188 | 59528 | 297×210 |
| A3 세로 | 84188 | 119056 | 297×420 |
| B5 세로 | 49892 | 70868 | 176×250 |
| B4 세로 | 70868 | 100344 | 250×354 |
| Letter 세로 | 61236 | 79092 | 216×279 |

### 공문서 기본 여백 (mm → HwpUnit)

| 항목 | mm | HwpUnit |
|------|----|---------|
| TopMargin | 20 | 5669 |
| BottomMargin | 15 | 4252 |
| LeftMargin | 20 | 5669 |
| RightMargin | 20 | 5669 |
| HeaderLen | 15 | 4252 |
| FooterLen | 15 | 4252 |
| GutterLen | 0 | 0 |

### 일반 보고서 여백 참고값 (mm)

| 항목 | 값 |
|------|-----|
| TopMargin | 25~30 |
| BottomMargin | 15~20 |
| LeftMargin | 25~30 |
| RightMargin | 20~25 |

---

## B. 쪽/구역 나누기

```python
# 쪽 나누기 — pageBreak="1" 속성 사용
doc.add_paragraph("새 페이지 시작", pageBreak="1")

# 빈 쪽 나누기 (텍스트 없이)
doc.add_paragraph("", pageBreak="1", include_run=False)
```

> python-hwpx에는 `BreakSection()`, `BreakColumn()` 에 대응하는 고수준 API가 없다.
> 구역 분리가 필요하면 HWPX 파일 구조상 별도의 section XML 파일을 생성해야 하며,
> 이는 매우 복잡하므로 현재는 단일 섹션 내에서 작업하는 것을 권장한다.

---

## C. 머리말/꼬리말 (★ pyhwpx 한계 해결)

python-hwpx에서는 `set_header_text()` / `set_footer_text()` API가 제공된다.
pyhwpx의 `HAction.Execute("InsertHeader/Footer")` 비활성화 문제가 해결됨.

### 머리말 설정

```python
# 양쪽 페이지 동일 머리말
doc.set_header_text("보고서 제목", page_type="BOTH")

# 섹션 지정
doc.set_header_text("보고서 제목", section_index=0, page_type="BOTH")

# 홀수/짝수 페이지 구분
doc.set_header_text("홀수 페이지 머리말", page_type="ODD")
doc.set_header_text("짝수 페이지 머리말", page_type="EVEN")
```

### 꼬리말 설정

```python
# 양쪽 페이지 동일 꼬리말
doc.set_footer_text("기밀문서", page_type="BOTH")

# 홀수/짝수 구분
doc.set_footer_text("보고서명 | 홀수쪽", page_type="ODD")
doc.set_footer_text("작성일 | 짝수쪽", page_type="EVEN")
```

### 머리말/꼬리말 제거

```python
doc.remove_header(page_type="BOTH")
doc.remove_footer(page_type="BOTH")

# 섹션 지정
doc.remove_header(section_index=0, page_type="BOTH")
```

### 머리말/꼬리말 조회

```python
props = doc.sections[0].properties

# 머리말 목록
for hdr in props.headers:
    print(f"머리말: page_type={hdr.apply_page_type}, text={hdr.text}")

# 꼬리말 목록
for ftr in props.footers:
    print(f"꼬리말: page_type={ftr.apply_page_type}, text={ftr.text}")
```

### page_type 값

| 값 | 설명 |
|----|------|
| `"BOTH"` | 양쪽 페이지 (기본) |
| `"ODD"` | 홀수 페이지 |
| `"EVEN"` | 짝수 페이지 |

---

## D. 쪽 번호

> **⚠ 주의**: python-hwpx에서 자동 쪽 번호 삽입(`<hp:pageNum>`, `<hp:autoNum>`)은 API 미지원.
> 현재 가능한 대안:
> - 꼬리말에 텍스트로 고정 번호 입력 (자동 아님)
> - 번호 매기기 시작값 설정 (아래 참조)

### 번호 매기기 시작값 (구역별)

```python
props = doc.sections[0].properties

# 조회
sn = props.start_numbering
print(f"page_starts_on={sn.page_starts_on}, page={sn.page}")

# 설정
props.set_start_numbering(
    page_starts_on="BOTH",
    page=1,
    picture=1,
    table=1,
    equation=1,
)
```

### 문서 전체 번호 매기기 (헤더)

```python
header = doc.headers[0]

# 조회
bn = header.begin_numbering
print(f"page={bn.page}, table={bn.table}, picture={bn.picture}")

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

## E. 장 표지 생성 패턴

```python
# 쪽 나누기 + 빈 줄 + 제목 + 쪽 나누기

# 빈 줄로 위치 조정 (페이지 중간 정도)
doc.add_paragraph("", pageBreak="1")
for _ in range(10):
    doc.add_paragraph("")

# 장 제목 (가운데 정렬 + 큰 글씨 — charPr/paraPr 필요)
# 가운데 정렬 paraPr가 없다면 생성 (/hwpx-format C절 참조)
doc.add_paragraph("제2장 시장 동향", char_pr_id_ref=title_char_id, para_pr_id_ref=center_para_id)

# 부제 (선택)
doc.add_paragraph("Market Trends", char_pr_id_ref=subtitle_char_id, para_pr_id_ref=center_para_id)

# 다음 페이지 시작
doc.add_paragraph("", pageBreak="1")
```

---

## F. 구역별 페이지 설정 조회

```python
for i, section in enumerate(doc.sections):
    props = section.properties
    ps = props.page_size
    pm = props.page_margins
    print(f"=== 섹션 {i} ===")
    print(f"  크기: {ps.width}x{ps.height} ({ps.orientation})")
    print(f"  여백: L={pm.left} R={pm.right} T={pm.top} B={pm.bottom}")
```

---

## 참조

- 페이지 설정 API: `references/python-hwpx-api.md` (페이지 설정 섹션)
- 용지 크기 상수: `references/python-hwpx-api.md` (단위 참조)
- secPr XML 구조: `references/hwpx-xml-schema.md`
