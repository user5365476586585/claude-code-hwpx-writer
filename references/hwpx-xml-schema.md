# HWPX XML 스키마 핵심 요소

OWPML(Open Word-Processor Markup Language) 스키마 기반 HWPX 파일 구조.

---

## HWPX 패키지 구조

```
document.hwpx (ZIP)
├── mimetype                          # "application/hwp+zip"
├── META-INF/
│   ├── container.xml                 # 루트 파일 참조
│   └── manifest.xml                  # (선택)
├── Contents/
│   ├── content.hpf                   # 매니페스트 (스파인)
│   ├── header.xml                    # 문서 헤더 (스타일/속성 정의)
│   ├── section0.xml                  # 본문 섹션 0
│   ├── section1.xml                  # 본문 섹션 1 ...
│   └── ...
├── BinData/                          # 이미지/OLE 바이너리 데이터
│   ├── image001.png
│   └── ...
├── Preview/
│   └── PrvImage.png                  # 미리보기 이미지
└── version.xml                       # 버전 정보
```

---

## header.xml 구조

```xml
<hh:head version="1.2" secCnt="1"
    xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">

  <hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>

  <hh:refList>
    <hh:fontfaces itemCnt="7">
      <hh:fontface lang="HANGUL" fontCnt="1">
        <hh:font id="0" face="함초롱바탕" type="TTF"/>
      </hh:fontface>
      ...
    </hh:fontfaces>

    <hh:borderFills itemCnt="1">
      <hh:borderFill id="1" threeD="0" shadow="0" slash="0" backSlash="0">
        <hh:slash .../>
        <hh:backSlash .../>
        <hh:leftBorder type="None" width="0.1mm" color="#000000"/>
        <hh:rightBorder type="None" width="0.1mm" color="#000000"/>
        <hh:topBorder type="None" width="0.1mm" color="#000000"/>
        <hh:bottomBorder type="None" width="0.1mm" color="#000000"/>
        <hh:diagonal .../>
      </hh:borderFill>
    </hh:borderFills>

    <hh:charProperties itemCnt="N">
      <hh:charPr id="0" height="1000" textColor="#000000" ...>
        <hh:fontRef hangul="0" latin="0" hanja="0" .../>
        <hh:ratio hangul="100" latin="100" .../>
        <hh:spacing hangul="0" latin="0" .../>
        <hh:relSz hangul="100" latin="100" .../>
        <hh:offset hangul="0" latin="0" .../>
        <hh:underline type="None" shape="Solid" color="#000000"/>
        <hh:strikeout type="None" shape="Solid" color="#000000"/>
        <hh:outline type="None"/>
        <hh:shadow type="None" .../>
      </hh:charPr>
    </hh:charProperties>

    <hh:tabProperties itemCnt="N">
      <hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/>
    </hh:tabProperties>

    <hh:numberings itemCnt="N">
      ...
    </hh:numberings>

    <hh:paraProperties itemCnt="N">
      <hh:paraPr id="0" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="1">
        <hh:align horizontal="Justify" vertical="Baseline"/>
        <hh:heading type="None" idRef="0" level="0"/>
        <hh:breakSetting breakLatinWord="KeepWord" breakNonLatinWord="1"
                         widowOrphan="0" keepWithNext="0" keepLines="0"
                         pageBreakBefore="0" lineWrap="Break"/>
        <hh:margin intent="0" left="0" right="0" prev="0" next="0"/>
        <hh:lineSpacing type="Percent" value="160" unit="hwpunit"/>
        <hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0"
                   offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>
        <hh:autoSpacing eAsianEng="0" eAsianNum="0"/>
      </hh:paraPr>
    </hh:paraProperties>

    <hh:styles itemCnt="N">
      <hh:style id="0" type="para" name="바탕글" engName="Body"
                paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langId="1042"/>
      <hh:style id="1" type="para" name="개요 1" engName="Outline 1"
                paraPrIDRef="1" charPrIDRef="1" nextStyleIDRef="1" langId="1042"/>
      ...
    </hh:styles>

    <hh:memoProperties itemCnt="N">
      <hh:memoPr id="0" width="..." lineWidth="..." lineType="..." .../>
    </hh:memoProperties>
  </hh:refList>
</hh:head>
```

### charPr 주요 속성

| 속성 | 설명 | 예시 |
|------|------|------|
| `id` | 고유 ID | "0" |
| `height` | 글자 크기 (HwpUnit, ×100) | "1000" = 10pt |
| `textColor` | 글자 색 | "#000000" |
| `shadeColor` | 음영 색 | "#FFFFFF" |
| `bold` | 굵게 (존재 시 "1") | |
| `italic` | 기울임 (존재 시 "1") | |
| `useFontSpace` | 장평 | |
| `useKerning` | 자간 | |

### charPr 자식 요소

| 요소 | 설명 |
|------|------|
| `<hh:fontRef>` | 글꼴 참조 (hangul, latin, hanja, japanese, other, symbol, user) |
| `<hh:ratio>` | 장평 비율 |
| `<hh:spacing>` | 자간 |
| `<hh:relSz>` | 상대 크기 |
| `<hh:offset>` | 위치 오프셋 |
| `<hh:underline>` | 밑줄 (type, shape, color) |
| `<hh:strikeout>` | 취소선 |
| `<hh:outline>` | 외곽선 |
| `<hh:shadow>` | 그림자 |

---

## section XML 구조

```xml
<hs:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
         xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
         xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">

  <!-- 문단 목록 -->
  <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
    <hp:run charPrIDRef="0">
      <hp:secPr>
        <!-- 첫 문단에만 구역 설정 포함 -->
        <hp:grid .../>
        <hp:startNum pageStartsOn="Both" page="1" pic="1" tbl="1" equation="1"/>
        <hp:pageSz width="59528" height="84188" orientation="Portrait"/>
        <hp:pageMar header="4252" footer="4252" left="8504" right="8504"
                    top="5668" bottom="4252" gutter="0"/>
        <hp:pageProperty .../>
        <hp:footNoteShape .../>
        <hp:endNoteShape .../>
        <hp:pageBorderFill .../>
        <hp:masterPage/>
      </hp:secPr>
      <hp:t>첫 번째 문단 텍스트</hp:t>
    </hp:run>
  </hp:p>

  <hp:p id="1" paraPrIDRef="0" styleIDRef="0" pageBreak="0">
    <hp:run charPrIDRef="0">
      <hp:t>두 번째 문단</hp:t>
    </hp:run>
    <hp:run charPrIDRef="1">
      <hp:t>다른 서식의 텍스트</hp:t>
    </hp:run>
  </hp:p>

  <!-- 표 포함 문단 -->
  <hp:p id="2" paraPrIDRef="0" styleIDRef="0">
    <hp:run>
      <hp:tbl ...>
        <hp:tr>
          <hp:tc ...>
            <hp:cellAddr colAddr="0" rowAddr="0"/>
            <hp:cellSpan colSpan="1" rowSpan="1"/>
            <hp:cellSz width="7200" height="3600"/>
            <hp:cellMargin left="0" right="0" top="0" bottom="0"/>
            <hp:cellBorderFill borderFillIDRef="1"/>
            <hp:subList>
              <hp:p id="0" paraPrIDRef="0" styleIDRef="0">
                <hp:run charPrIDRef="0">
                  <hp:t>셀 내용</hp:t>
                </hp:run>
              </hp:p>
            </hp:subList>
          </hp:tc>
        </hp:tr>
      </hp:tbl>
    </hp:run>
  </hp:p>

  <!-- 메모 그룹 (섹션 끝) -->
  <hp:memogroup>
    <hp:memo id="MEMO_001" memoShapeIDRef="0">
      <hp:paraList>
        <hp:p ...>
          <hp:run ...><hp:t>메모 내용</hp:t></hp:run>
        </hp:p>
      </hp:paraList>
    </hp:memo>
  </hp:memogroup>
</hs:sec>
```

### 문단(p) 주요 속성

| 속성 | 설명 |
|------|------|
| `id` | 문단 ID |
| `paraPrIDRef` | 문단속성 참조 ID |
| `styleIDRef` | 스타일 참조 ID |
| `pageBreak` | "1"이면 쪽 나누기 |
| `columnBreak` | "1"이면 단 나누기 |
| `merged` | "1"이면 병합됨 |

### 런(run) 주요 속성

| 속성 | 설명 |
|------|------|
| `charPrIDRef` | 문자속성 참조 ID |

### 표(tbl) XML 구조 요약

```
<hp:tbl>
  ├── <hp:sz width="전체너비" height="행높이"/>
  ├── <hp:tr>
  │   ├── <hp:tc>
  │   │   ├── <hp:cellAddr colAddr="0" rowAddr="0"/>
  │   │   ├── <hp:cellSpan colSpan="1" rowSpan="1"/>
  │   │   ├── <hp:cellSz width="..." height="..."/>
  │   │   ├── <hp:cellMargin .../>
  │   │   ├── <hp:cellBorderFill borderFillIDRef="1"/>
  │   │   └── <hp:subList>
  │   │       └── <hp:p> ... </hp:p>
  │   │   </hp:subList>
  │   └── <hp:tc> ...
  └── <hp:tr> ...
```

### 셀 병합 시 cellSpan 규칙

| 상태 | colSpan | rowSpan |
|------|---------|---------|
| 일반 셀 | 1 | 1 |
| 앵커 셀 (병합 시작) | N (열 수) | M (행 수) |
| 비앵커 셀 (병합 내) | 0 | 0 |

---

## 구역 설정 (secPr) 내 핵심 요소

### 용지 크기

```xml
<hp:pageSz width="59528" height="84188" orientation="Portrait"/>
```

| orientation | 설명 |
|-------------|------|
| `Portrait` | 세로 |
| `Landscape` | 가로 |

### 여백

```xml
<hp:pageMar header="4252" footer="4252"
            left="8504" right="8504"
            top="5668" bottom="4252" gutter="0"/>
```

### 번호 매기기 시작

```xml
<hp:startNum pageStartsOn="Both" page="1" pic="1" tbl="1" equation="1"/>
```

---

## 머리말/꼬리말 XML 구조

구역 설정(secPr) 내에 위치:

```xml
<hp:headerFooter>
  <hp:header>
    <hp:apply applyPageType="Both" id="0"/>
    <hp:paraList>
      <hp:p paraPrIDRef="0" styleIDRef="0">
        <hp:run charPrIDRef="0">
          <hp:t>머리말 텍스트</hp:t>
        </hp:run>
      </hp:p>
    </hp:paraList>
  </hp:header>
  <hp:footer>
    <hp:apply applyPageType="Both" id="0"/>
    <hp:paraList>
      <hp:p paraPrIDRef="0" styleIDRef="0">
        <hp:run charPrIDRef="0">
          <hp:t>꼬리말 텍스트</hp:t>
        </hp:run>
      </hp:p>
    </hp:paraList>
  </hp:footer>
</hp:headerFooter>
```

---

## content.hpf (매니페스트)

```xml
<opf:package xmlns:opf="http://www.idpf.org/2007/opf"
             xmlns:hpf="http://www.hancom.co.kr/hwpml/2011/content">
  <opf:metadata>
    <dc:title>문서 제목</dc:title>
    ...
  </opf:metadata>
  <opf:manifest>
    <opf:item id="header" href="header.xml" media-type="application/xml"/>
    <opf:item id="section0" href="section0.xml" media-type="application/xml"/>
    <opf:item id="image001" href="../BinData/image001.png" media-type="image/png"/>
    ...
  </opf:manifest>
  <opf:spine>
    <opf:itemref idref="section0"/>
    ...
  </opf:spine>
</opf:package>
```

---

## 네임스페이스 URI 요약

| 접두사 | URI | 용도 |
|--------|-----|------|
| `hp` | `http://www.hancom.co.kr/hwpml/2011/paragraph` | 문단, 런, 텍스트, 표, 컨트롤 |
| `hs` | `http://www.hancom.co.kr/hwpml/2011/section` | 섹션 루트 |
| `hh` | `http://www.hancom.co.kr/hwpml/2011/head` | 헤더 (스타일, 속성 정의) |
| `hc` | `http://www.hancom.co.kr/hwpml/2011/core` | 코어 타입 |
| `ha` | `http://www.hancom.co.kr/hwpml/2011/app` | 앱 메타 |
| `hm` | `http://www.hancom.co.kr/hwpml/2011/master-page` | 마스터 페이지 |
| `hpf` | `http://www.hancom.co.kr/hwpml/2011/content` | 콘텐츠 매니페스트 |
| `opf` | `http://www.idpf.org/2007/opf` | OPF 패키지 |
| `dc` | `http://purl.org/dc/elements/1.1/` | 더블린 코어 메타 |
