# claude-code-hwpx-writer

[python-hwpx](https://github.com/airmang/python-hwpx) 기반 한/글 문서(HWPX) 자동화를 위한 [Claude Code](https://claude.ai/claude-code) 에이전트 시스템.

순수 Python으로 HWPX 파일을 직접 파싱/생성. 한/글 설치 없이, OS 무관하게 동작.

---

## 요구사항

- Python 3.10 이상
- Claude Code CLI
- OS 무관 (Windows, macOS, Linux)

테스트 환경: Claude Opus 4.6

---

## 설치

```bash
# 가상환경 생성 (conda 또는 venv)
conda create -n hwpx python=3.11
conda activate hwpx
pip install python-hwpx
```

이 저장소를 클론한 후 루트에서 Claude Code 실행 시 `CLAUDE.md` 자동 로드.

```bash
git clone https://github.com/user5365476586585/claude-code-hwpx-writer.git
cd claude-code-hwpx-writer
claude   # Claude Code 실행
```

---

## 구조

```
├── CLAUDE.md                          # 오케스트레이터 (부모 에이전트)
├── README.md
├── .claude/
│   ├── settings.local.json            # 권한 설정
│   └── commands/                      # Skill 파일 (자식 에이전트)
│       ├── hwpx-style.md              # 스타일 관리
│       ├── hwpx-table.md              # 표 조작
│       ├── hwpx-format.md             # 글자/문단 서식
│       ├── hwpx-page.md               # 페이지 레이아웃
│       ├── hwpx-insert.md             # 개체 삽입
│       └── hwpx-write.md              # 문서 작성 통합 워크플로우
├── references/                        # python-hwpx API 레퍼런스
│   ├── python-hwpx-api.md             # HwpxDocument 전체 API
│   ├── python-hwpx-oxml.md            # 저수준 XML 조작 API
│   ├── python-hwpx-tools.md           # TextExtractor, ObjectFinder
│   └── hwpx-xml-schema.md             # OWPML 스키마 핵심 요소
└── temp/                              # 임시 파일 (스크립트 실행용)
```

### 아키텍처

- **CLAUDE.md** (부모): 오케스트레이터. 요청 분석 → skill 라우팅 → 실행 조율
- **hwpx-*.md** (자식): 각 도메인 전문 skill. 부모가 필요 시 `/hwpx-*` 명령으로 로드
- **references/**: python-hwpx API 문서. 에이전트가 코드 작성 시 참조

---

## 사용

자연어로 HWPX 문서 작업 요청. 에이전트는 실행 전 기획안을 출력하고 승인 대기.

```
"반도체 시장 분석 보고서 10페이지 작성해줘"
"output.hwpx 열어서 2장 내용 수정해줘"
"3행 4열 표 만들고 데이터 넣어줘"
"기존 문서의 스타일 분석해줘"
```

### Skill 트리거

| 키워드 | Skill | 설명 |
|--------|-------|------|
| 스타일, 개요, 글머리표, 프리셋 | `/hwpx-style` | 스타일 관리 |
| 표, 테이블, 셀, 행, 열 | `/hwpx-table` | 표 조작 |
| 글꼴, 색상, 줄간격, 정렬 | `/hwpx-format` | 글자/문단 서식 |
| 용지, 여백, 머리말, 꼬리말 | `/hwpx-page` | 페이지 레이아웃 |
| 도형, 캡션 | `/hwpx-insert` | 개체 삽입 |
| 보고서 작성, 문서 생성 | `/hwpx-write` | 통합 작성 |

---

## 주요 기능

- **문서 생성/편집/저장**: 새 문서 또는 기존 HWPX 파일 편집
- **스타일 관리**: 조회/생성/적용/자동 기호 판별 (heading.type=BULLET 분석)
- **표 조작**: 생성/셀 병합/배경색/헤더 볼드/캡션+출처 패턴
- **서식**: charPr(글꼴/크기/색상) + paraPr(정렬/줄간격) 커스텀 생성
- **페이지**: 용지/여백/머리말/꼬리말/번호 매기기
- **텍스트 처리**: TextExtractor(추출) + ObjectFinder(검색) + 찾기/바꾸기
- **보고서 프리셋**: 실전 보고서(9수준) / 간단 보고서 / 공문서

---

## 미지원 기능

| 기능 | 사유 |
|------|------|
| HWP(레거시) 파일 편집 | 읽기 전용만 가능 |
| PDF 출력 | 렌더링 엔진 없음 |
| 이미지 삽입 | XML 구조 복잡, 불안정 |
| 쪽 번호 자동 삽입 | API 없음 |
| 메모 삽입/수정/삭제 | lxml/ET 호환 버그 |
| 셀 분할 | lxml/ET 호환 버그 |
| 각주/미주/필드 | API 없음 |
| 커서 기반 실시간 편집 | XML 직접 조작 방식 |

---

## 감사의 말

**[python-hwpx](https://github.com/airmang/python-hwpx)** — 순수 Python HWPX 파싱 라이브러리. 에이전트의 모든 문서 조작 기능의 기반.

---

## 라이선스

GNU General Public License v3.0.

> python-hwpx는 별도 라이선스 적용. [저장소](https://github.com/airmang/python-hwpx) 참조.
