# hwp2md

> **아래아 한글(Hancom Office) 및 MS 오피스 의존성 없이** `.hwp` 및 `.hwpx` 문서를 Markdown(`.md`), Microsoft Word(`.docx`), 또는 순수 텍스트(`.txt`)로 변환하는 독립형 문서 변환기  
> Standalone HWP and HWPX to Markdown (.md), Microsoft Word (.docx), and Plain Text (.txt) converter with complete table support.

---

## 주요 기능 (Features)

- **오피스 프로그램 무의존 (Zero Office Dependency)**:
  - 한컴오피스 / 아래아 한글 / MS Word 프로그램 설치가 전혀 필요 없습니다.
  - Windows COM(`win32com.client`) 자동화를 일절 사용하지 않으며, 순수 파이썬 파서로 구동됩니다.
  - Windows뿐 아니라 **Linux, macOS, Docker** 환경에서도 동일하게 작동합니다.
- **HWP 및 HWPX 모두 지원 (Dual Format Support)**:
  - `.hwpx` (Open Container XML 포맷): 표준 라이브러리(`zipfile`, `xml.etree`) 기반으로 초고속 파싱
  - `.hwp` (HWP 5.0 바이너리 CFBF 포맷): `pyhwp` 및 `olefile` 기반 파싱 (내부 Warning 노이즈 로그 완벽 차단)
- **다중 포맷 출력 지원 (Markdown, Word, Plain Text)**:
  - 간결하고 호환성 높은 **Markdown(`.md`)** 변환
  - 스타일, 폰트(맑은 고딕), 복합 표 및 이미지가 완벽 보존되는 **Word(`.docx`)** 변환 (`--docx` 옵션 또는 `hwp2docx`)
  - LLM 프롬프트 및 RAG 인덱싱에 최적화된 **순수 텍스트(`.txt`)** 추출 (`--text` 옵션 또는 `hwp2txt`)
- **강력한 표(Table) 변환 (Complete Table Support)**:
  - Markdown: GitHub Flavored Markdown (GFM) 파이프 표 형식(`| ... |`)으로 자동 변환
  - Word (.docx):
    - 실선 그리드 테두리 및 **행 병합(Rowspan) / 열 병합(Colspan)** 직사각형 매트릭스 완벽 반영
    - 페이지 너비 기준 **지능형 열 너비 비례 자동 배분**
    - 페이지 분할 시 행 쪼개짐 방지(`cantSplit`) 및 **다중 페이지 헤더 반복(`tblHeader`)** 적용
    - 1x1 단일 셀 안내/공지 박스를 세련된 **리포트형 콜아웃 박스(연회색 음영 + 포인트 테두리)**로 고급 렌더링
  - 셀 내부 다중 줄바꿈(`<br>`) 및 파이프 기호(`\|`) 이스케이프 완벽 처리
  - 복잡한 표 구조 보존을 위한 HTML 테이블 모드 지원 (`--table-format html`)
- **텍스트 스타일 및 서식 보존 (Formatting Preservation)**:
  - 글꼴 크기 기반 제목(`h1`, `h2`, `h3` / Word Heading) 자동 감지
  - 굵게(**Bold**), 기울임(*Italic*), 취소선(~~Strike~~), 밑줄(<u>Underline</u>) 보존
  - 클릭 가능한 하이퍼링크(Hyperlink) 변환 보존
  - 공공기관/기업 문서의 고유 글머리 기호(`▪`, `▸`, `ㅇ`, `□`, `1.`, `가.` 등) 원본 보존
- **이미지 자동 추출 및 삽입 (Embedded Image Extraction & Embedding)**:
  - 문서 내 포함된 이미지 파일들을 지정된 폴더(`[파일명]_images/`)에 자동 추출하고 마크다운 이미지 링크(`![alt](path)`)로 연결
  - Word(`.docx`) 변환 시 추출된 원본 이미지를 문서 및 표 셀에 자동 크기 조절하여 직접 삽입
- **문서 메타데이터(Metadata) 추출**:
  - 제목, 작성자, 생성일자, 파일 크기, 섹션 수, 포함된 이미지 개수 등 문서 정보를 JSON 형태로 즉시 확인 (`--meta`)

---

## 설치 (Installation)

```bash
# 로컬 개발 모드 설치
pip install -e .

# 또는 필요한 패키지 직접 설치
pip install olefile pyhwp python-docx
```

> **참고**: `.hwpx`의 Markdown/Text 변환은 파이썬 표준 라이브러리만으로 동작합니다. `.hwp` 바이너리 변환에는 `olefile`과 `pyhwp`, `.docx` 변환에는 `python-docx`가 사용됩니다.

---

## 사용법 (Usage)

### 1. 명령줄 인터페이스 (CLI)

#### Markdown (.md) 변환
```bash
# 기본 사용법: .hwpx 또는 .hwp 파일을 마크다운(.md)으로 변환
python hwp2md.py sample.hwpx
python hwp2md.py report.hwp

# 출력 파일 경로 직접 지정
python hwp2md.py report.hwpx -o custom_output.md

# 여러 파일 일괄 변환 (Globbing)
python hwp2md.py *.hwpx

# 디렉터리 내 모든 HWP/HWPX 재귀 변환
python hwp2md.py ./documents/ -r

# 단일 셀(1x1) 안내 박스를 인용구(>) 블록으로 변환
python hwp2md.py press_release.hwpx --box-style quote

# 복잡한 병합 표를 HTML <table> 태그로 변환 (모든 마크다운 뷰어 호환)
python hwp2md.py complex_table.hwpx --table-format html

# 이미지 추출 비활성화
python hwp2md.py document.hwpx --no-images

# 콘솔 표준 출력(stdout)으로 직접 마크다운 출력
python hwp2md.py document.hwpx --stdout
```

#### Word (.docx) 변환
```bash
# --docx 플래그로 Word 문서 변환
python hwp2md.py sample.hwpx --docx
python hwp2md.py report.hwp --docx

# 출력 확장자가 .docx인 경우 자동으로 docx 모드로 동작
python hwp2md.py sample.hwpx -o result.docx

# hwp2docx 전용 스크립트 실행
python hwp2docx.py sample.hwpx
python hwp2docx.py report.hwp
```

#### 순수 텍스트 (.txt) 추출 (LLM / RAG 최적화)
```bash
# --text 플래그로 깨끗한 텍스트 추출 (표는 탭으로 구분)
python hwp2md.py sample.hwpx --text

# 출력 확장자가 .txt인 경우 자동 텍스트 모드 동작
python hwp2md.py report.hwp -o prompt_context.txt

# hwp2txt 전용 스크립트 실행
python hwp2txt.py sample.hwpx
```

#### 문서 메타데이터 조회
```bash
# 문서 정보(포맷, 크기, 이미지 수, 섹션 수 등)를 JSON 형태로 조회
python hwp2md.py sample.hwpx --meta
```

#### CLI 옵션 목록

| 옵션 | 설명 |
| :--- | :--- |
| `inputs` | 변환할 `.hwp` 또는 `.hwpx` 파일 경로 / 폴더 경로 (와일드카드 지원) |
| `-o`, `--output` | 출력 파일명 또는 출력 디렉터리 (확장자 `.docx`, `.txt`, `.md` 자동 판별) |
| `--docx` | Microsoft Word(`.docx`) 형식으로 변환 |
| `--text` | 서식 없는 순수 텍스트(`.txt`)로 추출 (LLM / RAG 연동용) |
| `--meta` | 문서 메타데이터 정보를 JSON 형태로 출력하고 종료 |
| `--format` | 출력 포맷 명시: `md`, `docx`, `txt` |
| `--table-format` | 표 변환 방식: `markdown` (기본값, GFM 표) 또는 `html` (HTML `<table>`) |
| `--box-style` | 1x1 단일 셀 박스 서식: `table` (기본 표), `quote` (인용구 `>`), `plain` (일반 문단) |
| `--no-images` | 문서 내 삽입 이미지 추출 비활성화 |
| `--images-dir` | 추출 이미지를 저장할 사용자 지정 폴더 경로 |
| `-r`, `--recursive` | 디렉터리 지정 시 하위 폴더까지 재귀적으로 탐색하여 변환 |
| `--stdout` | 파일 생성 없이 텍스트/마크다운을 콘솔로 출력 |
| `-v`, `--version` | 버전 정보 표시 |

---

### 2. 파이썬 라이브러리 API (Python API)

#### 1) Word (.docx) 변환
```python
from hwp2md import convert_to_docx, convert

# HWP/HWPX를 Word(.docx)로 변환
convert_to_docx("press_release.hwpx", output_path="press_release.docx")
convert_to_docx("report.hwp", output_path="report.docx")

# convert() 함수에서 to_docx=True 또는 output_path 확장자로 자동 감지
convert("report.hwpx", output_path="report.docx")
```

#### 2) 순수 텍스트 (.txt) 추출 & 메타데이터
```python
from hwp2md import convert_to_text, get_metadata

# 1. LLM/RAG 파이프라인용 클린 텍스트 추출
text = convert_to_text("report.hwpx")
print(text)

# 파일로 저장
convert_to_text("report.hwp", output_path="report.txt")

# 2. 문서 메타데이터 검사
meta = get_metadata("report.hwpx")
print(meta)
# {'filename': 'report.hwpx', 'format': 'hwpx', 'embedded_images_count': 4, ...}
```

#### 3) Markdown (.md) 변환
```python
from hwp2md import convert

# 1. 파일에서 파일로 변환
convert("report.hwpx", output_path="report.md")

# 2. 파일 변환 후 마크다운 문자열 반환
markdown_text = convert("document.hwp")
print(markdown_text)

# 3. 다양한 옵션 지정
markdown_text = convert(
    "press_release.hwpx",
    output_path="output.md",
    extract_images=True,
    images_dir="custom_images/",
    table_format="markdown",  # 'markdown' 또는 'html'
    box_style="quote",        # 1x1 안내 박스를 > 인용구로 변환
)
```

#### 4) 포맷별 전용 변환 함수
```python
from hwp2md import (
    convert_hwpx, convert_hwpx_to_docx,
    convert_hwp, convert_hwp_to_docx
)

# HWPX 변환
md_hwpx = convert_hwpx("sample.hwpx")
docx_path1 = convert_hwpx_to_docx("sample.hwpx", "sample.docx")

# HWP 변환
md_hwp = convert_hwp("sample.hwp")
docx_path2 = convert_hwp_to_docx("sample.hwp", "sample.docx")
```

---

## 표(Table) 변환 세부 사양

### GFM Markdown 표 변환 예시

| 원본 HWPX/HWP 표 구조 | 생성된 Markdown / Docx |
| :--- | :--- |
| **행/열 병합 (Spans)** | 직사각형 매트릭스를 유지하도록 병합된 칸을 빈 셀로 채워 열 어긋남 방지 / Word 셀 병합 완벽 재현 |
| **열 너비 비율 배분** | 전각/동아시아 문자폭을 계산하여 Word 표 열 너비를 페이지 가용 폭에 맞추어 비례 자동 배분 |
| **페이지 분할 최적화** | Word 표에서 행 잘림 방지(`cantSplit`) 및 2페이지 이상 시 첫 행 헤더 반복(`tblHeader`) |
| **1x1 공지/안내 박스** | 연회색 배경 음영(`F8FAFC`)과 좌측 포인트 테두리(`2563EB`)가 적용된 세련된 리포트형 콜아웃으로 렌더링 |
| **셀 내 여러 줄 텍스트** | `<br>` 태그로 연결하여 1개의 마크다운 행으로 유지 / Word 표 셀 내 문단으로 처리 |
| **파이프(`\|`) 문자** | 마크다운 표 구분자와 충돌하지 않도록 `\\|`로 자동 이스케이프 |

---

## 테스트 실행 (Tests)

```bash
python -m unittest discover -s tests
```

---

## 라이선스 (License)

[The Unlicense](UNLICENSE) (Public Domain). 자유롭게 복제, 수정, 배포, 상업적 이용이 가능합니다.
