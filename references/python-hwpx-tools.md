# python-hwpx Tools API 레퍼런스

TextExtractor, ObjectFinder — 문서 분석/검색 도구.

---

## TextExtractor (텍스트 추출)

### 기본 사용

```python
from hwpx.tools.text_extractor import TextExtractor

# 전체 텍스트 추출
with TextExtractor("document.hwpx") as te:
    full_text = te.extract_text()
    print(full_text)
```

### 생성자

```python
te = TextExtractor(
    source="document.hwpx",   # 파일 경로, Path, 또는 ZipFile
    namespaces=None,           # dict (기본: DEFAULT_NAMESPACES)
)
```

### extract_text (전체 추출)

```python
text = te.extract_text(
    paragraph_separator="\n",    # 문단 구분자
    skip_empty=True,             # 빈 문단 건너뛰기
    include_nested=True,         # 중첩 문단 포함 (표 안 등)
    object_behavior="skip",      # "skip" | "placeholder" | "include"
    object_placeholder=None,     # placeholder 텍스트
    preserve_breaks=True,        # 줄바꿈 보존
    annotations=None,            # AnnotationOptions
)
```

### 섹션별 순회

```python
for section in te.iter_sections():
    print(f"섹션 {section.index}: {section.name}")
    for para in te.iter_paragraphs(section, include_nested=True):
        text = para.text()
        print(f"  [{para.index}] {text[:50]}")
```

### 전체 문단 순회

```python
for para in te.iter_document_paragraphs(include_nested=True):
    print(f"섹션{para.section.index} 문단{para.index}: {para.text()}")
```

### SectionInfo

```python
section.index      # int (0-based)
section.name       # str (섹션 파일명)
section.element    # ET.Element
```

### ParagraphInfo

```python
para.section       # SectionInfo
para.index         # int (섹션 내 인덱스)
para.element       # ET.Element
para.path          # XPath 유사 경로
para.hierarchy     # tuple[str, ...] (경로 분할)
para.tag           # str (로컬 태그명)
para.ancestors     # tuple[str, ...] (조상 태그)
para.is_nested     # bool (중첩 여부)

# 텍스트 추출
text = para.text(
    object_behavior="skip",
    object_placeholder=None,
    preserve_breaks=True,
    annotations=None,
)
```

### AnnotationOptions (주석 렌더링 설정)

```python
from hwpx.tools.text_extractor import AnnotationOptions

opts = AnnotationOptions(
    # 형광펜
    highlight="ignore",            # "ignore" | "markers"
    highlight_start="[HI:{color}]",
    highlight_end="[/HI]",
    highlight_summary="",

    # 각주
    footnote="ignore",             # "ignore" | "placeholder" | "inline"
    endnote="ignore",              # "ignore" | "placeholder" | "inline"
    note_inline_format="[{kind}:{text}]",
    note_placeholder="[{kind}:{inst_id}]",
    note_summary="",
    note_joiner="",

    # 하이퍼링크
    hyperlink="ignore",            # "ignore" | "placeholder" | "target"
    hyperlink_target_format="[LINK:{target}]",
    hyperlink_placeholder="[LINK]",
    hyperlink_summary="",

    # 컨트롤
    control="ignore",              # "ignore" | "placeholder" | "nested"
    control_placeholder="[{name}]",
    control_summary="",
    control_joiner="",
)

# 주석 포함 텍스트 추출
text = te.extract_text(annotations=opts)
```

### 주석 포함 추출 예시

```python
# 각주 인라인 + 형광펜 마커
opts = AnnotationOptions(
    highlight="markers",
    footnote="inline",
    hyperlink="target",
)
text = te.extract_text(annotations=opts)
```

---

## ObjectFinder (객체 검색)

### 기본 사용

> **⚠ ObjectFinder는 context manager(with문) 미지원.** 직접 인스턴스 생성하여 사용.

```python
from hwpx.tools.object_finder import ObjectFinder

finder = ObjectFinder("document.hwpx")

# 모든 표 찾기
tables = finder.find_all(tag="tbl")
for t in tables:
    print(f"표: {t.path}")

# 첫 번째 그림 찾기
pic = finder.find_first(tag="pic")
if pic:
    print(f"그림: {pic.path}")
```

### 생성자

```python
finder = ObjectFinder(
    source="document.hwpx",   # 파일 경로, Path, 또는 ZipFile
    namespaces=None,
)
```

### 검색 메서드

#### find_all

```python
results = finder.find_all(
    tag="tbl",                    # str | list[str] | None — 태그 필터
    attrs={"width": "54000"},     # dict | None — 속성 매칭
    xpath=None,                   # str | None — XPath 표현식
    section_filter=None,          # callable(SectionInfo) -> bool
    limit=None,                   # int | None — 최대 결과 수
)
# 반환: list[FoundElement]
```

#### find_first

```python
result = finder.find_first(
    tag=None,
    attrs=None,
    xpath=None,
    section_filter=None,
)
# 반환: FoundElement | None
```

#### iter (제너레이터)

```python
for elem in finder.iter(tag="pic", limit=10):
    print(elem.path)
```

### FoundElement

```python
fe = finder.find_first(tag="tbl")

fe.section       # SectionInfo
fe.path          # str (XPath 유사 경로)
fe.element       # ET.Element (원시 XML)
fe.tag           # str (로컬 태그명)
fe.hierarchy     # tuple[str, ...]
fe.text          # str | None (요소 텍스트)

fe.get("width")              # 속성값 조회
fe.get("height", default="0")
```

### 속성 매칭 타입 (AttrMatcher)

attrs 딕셔너리의 값은 다양한 타입 지원:

```python
import re

# 1. 정확한 문자열 매칭
finder.find_all(tag="tbl", attrs={"width": "54000"})

# 2. 목록 중 하나 매칭
finder.find_all(attrs={"type": ["pic", "chart"]})

# 3. 정규식 매칭
finder.find_all(attrs={"name": re.compile(r"^IMG_\d+")})

# 4. 함수 매칭
finder.find_all(attrs={"width": lambda v: int(v) > 30000})
```

### 주석 검색 (iter_annotations)

```python
from hwpx.tools.text_extractor import AnnotationOptions

for ann in finder.iter_annotations(
    kinds=["highlight", "footnote", "endnote", "hyperlink", "control"],
    options=None,          # AnnotationOptions
    section_filter=None,
    preserve_breaks=True,
):
    print(f"{ann.kind}: {ann.value}")
    print(f"  위치: {ann.element.path}")
```

### AnnotationMatch

```python
ann.kind       # str ("highlight", "footnote", "endnote", "hyperlink", "control")
ann.element    # FoundElement
ann.value      # str | None
```

---

## 유틸리티 함수

```python
from hwpx.tools.text_extractor import strip_namespace, tag_matches, build_parent_map

# 네임스페이스 제거
strip_namespace("{http://...}paragraph")  # → "paragraph"

# 태그 비교 (네임스페이스 인식)
tag_matches("hp:p", "{http://...}p", namespaces)  # → True

# 부모 맵 구축
parent_map = build_parent_map(root_element)
```

---

## DEFAULT_NAMESPACES

```python
DEFAULT_NAMESPACES = {
    "hp":  "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hp10": "http://www.hancom.co.kr/hwpml/2011/paragraph/10",
    "hs":  "http://www.hancom.co.kr/hwpml/2011/section",
    "hc":  "http://www.hancom.co.kr/hwpml/2011/core",
    "ha":  "http://www.hancom.co.kr/hwpml/2011/app",
    "hh":  "http://www.hancom.co.kr/hwpml/2011/head",
    "hhs": "http://www.hancom.co.kr/hwpml/2011/head/section",
    "hm":  "http://www.hancom.co.kr/hwpml/2011/master-page",
    "hpf": "http://www.hancom.co.kr/hwpml/2011/content",
    "dc":  "http://purl.org/dc/elements/1.1/",
    "opf": "http://www.idpf.org/2007/opf",
}
```
