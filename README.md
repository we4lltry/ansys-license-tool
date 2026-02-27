# Ansys 라이선스 확인서 생성기

> **POSCO EnC 태성에스엔이 내부 전용 도구**  
> Streamlit 기반 웹앱 → Windows 단독 실행 EXE 패키징

---

## 📁 폴더 구조

```
ansys-license-tool/
├── app.py                          ← Streamlit 앱 (메인 UI / 핵심 로직)
├── launcher.py                     ← EXE 진입점 (서버 + 브라우저 자동 실행)
├── AnsysLicenseTool.spec           ← PyInstaller 빌드 설정
├── build.bat                       ← 원클릭 빌드 스크립트
├── run_dev.bat                     ← 개발 테스트용 실행 스크립트
├── 라이선스_확인서_템플릿.pptx       ← (별도 제공) PPT 슬라이드 템플릿
└── README.md
```

---

## 🏗️ 아키텍처 흐름도

### 전체 실행 흐름 (EXE → 브라우저)

```mermaid
flowchart TD
    A["👤 사용자\n.exe 더블클릭"] --> B["launcher.py\n(PyInstaller 번들 진입점)"]

    B --> C["find_free_port()\n8501번부터 빈 포트 자동 탐색"]
    C --> D["start_streamlit()\nsubprocess로 Streamlit 서버 기동\n(CREATE_NO_WINDOW, 콘솔 숨김)"]
    D --> E["wait_for_server()\n소켓 연결 확인 (최대 40초 대기)"]

    E -- "LISTEN 확인" --> F["webbrowser.open()\nhttp://127.0.0.1:PORT\n기본 브라우저 자동 오픈"]
    E -- "타임아웃" --> G["오류 출력 후 종료"]

    F --> H["app.py\nStreamlit UI 렌더링"]
    H --> I["proc.wait()\nEXE 종료 시 서버도 함께 종료"]

    style A fill:#1a2357,color:#fff
    style H fill:#b8922a,color:#000
    style G fill:#c0392b,color:#fff
```

---

### 앱 동작 흐름 (app.py 내부)

```mermaid
flowchart TD
    U["👤 사용자"] -->|".txt 파일 업로드"| S1

    subgraph S1["STEP 01 · 파일 업로드"]
        F1["st.file_uploader()\n라이선스 .txt 파일 수신"]
    end

    S1 --> S2

    subgraph S2["STEP 02 · 파싱 & 테이블"]
        P1["content.decode('utf-8')"]
        P2["re.compile()\n정규식 패턴 매칭\n번호·제품명·수량·만료일·고객번호"]
        P3["pd.DataFrame()\n결과 테이블 생성"]
        P1 --> P2 --> P3
    end

    S2 --> S3

    subgraph S3["STEP 03 · 확인서 정보 입력"]
        I1["고객명 / 고객번호 (자동 채움)"]
        I2["설치 장소 / 라이선스 기간"]
        I3["라이선스 유형 / 발행일자"]
    end

    S3 -->|"폼 제출 (submitted)"| GEN

    subgraph GEN["STEP 04 · 문서 생성"]
        direction LR
        PDF["create_pdf()\nreportlab\nA4 · 한글 폰트\n정보표 + 제품 테이블"]
        PPT["create_pptx_from_template()\npython-pptx\n플레이스홀더 idx 매핑\n테이블 행 동적 삽입"]
    end

    GEN --> DL["⬇️ PDF / PPT\n다운로드 버튼 표시"]
    DL --> U

    style S1 fill:#0e1535,color:#b8922a
    style S2 fill:#0e1535,color:#b8922a
    style S3 fill:#0e1535,color:#b8922a
    style GEN fill:#0e1535,color:#b8922a
```

---

### EXE 빌드 파이프라인 (PyInstaller)

```mermaid
flowchart LR
    SRC["소스 파일\nlauncher.py\napp.py"] --> SPEC

    subgraph SPEC["AnsysLicenseTool.spec"]
        direction TB
        CA["collect_all()\nstreamlit / altair\npandas / numpy\nreportlab / pptx"]
        FT["폰트 포함\nmalgun.ttf\nmalgunbd.ttf"]
        SI["streamlit static\n+ runtime 복사"]
        HI["hiddenimports\n누락 모듈 명시"]
    end

    SPEC --> ANA["Analysis()"]
    ANA --> PYZ["PYZ 압축\n(.pyc 모듈)"]
    PYZ --> EXE_NODE["EXE()\nconsole=False\nupx=True"]
    EXE_NODE --> COLL["COLLECT()\n바이너리 + 데이터\n한 폴더로 묶기"]

    COLL --> DIST["dist/\nAnsysLicenseTool/\n├── AnsysLicenseTool.exe\n├── _internal/\n└── 라이선스_확인서_템플릿.pptx"]

    style SRC fill:#1a2357,color:#fff
    style DIST fill:#b8922a,color:#000
```

---

## ⚡ 빠른 시작

### 개발 환경 테스트 (EXE 불필요)

```bat
run_dev.bat
```

브라우저에서 `http://localhost:8501` 자동 열림

### EXE 빌드

```bat
build.bat
```

완료 후 `dist\AnsysLicenseTool\AnsysLicenseTool.exe` 생성

---

## 📦 의존성

```bash
pip install streamlit pandas reportlab python-pptx pyinstaller
```

| 패키지 | 용도 |
|---|---|
| `streamlit` | 웹 UI 프레임워크 |
| `pandas` | 라이선스 데이터 테이블 처리 |
| `reportlab` | PDF 생성 (한글 폰트 지원) |
| `python-pptx` | PPT 템플릿 기반 슬라이드 생성 |
| `pyinstaller` | Python → Windows EXE 패키징 |

---

## 📄 PPT 템플릿 배치

PPT 다운로드 기능을 사용하려면 **`라이선스_확인서_템플릿.pptx`** 파일을 아래 위치에 복사하세요:

| 환경 | 경로 |
|---|---|
| 개발 | `app.py` 와 같은 폴더 |
| EXE 빌드 후 배포 | `dist\AnsysLicenseTool\` |

> 템플릿이 없으면 **PDF 다운로드는 정상 작동**, PPT는 오류 메시지 표시.

---

## 🚚 EXE 배포

`dist\AnsysLicenseTool\` 폴더 **전체**를 사용자에게 전달합니다.

```
dist\AnsysLicenseTool\
├── AnsysLicenseTool.exe          ← 이것만 실행하면 됨
├── _internal\                    ← 런타임 라이브러리 (건드리지 말 것)
└── 라이선스_확인서_템플릿.pptx    ← 반드시 함께 배포
```

> ⚠️ **EXE 단독 복사 시 동작하지 않습니다. 폴더 전체 배포 필수.**

---

## 📝 라이선스 파일 패턴 예시

앱이 인식하는 `.txt` 파일 패턴 (정규식 추출):

```
# 1. Ansys Mechanical Enterprise: 5 task(s) expiring 31-Dec-2026
#    Customer # 1213401
```

| 필드 | 예시 값 |
|---|---|
| No | `1` |
| Software (제품명) | `Ansys Mechanical Enterprise` |
| QTY (수량) | `5` |
| 만료일 (ExpireDate) | `31-Dec-2026` |
| 고객번호 (CustomerNo) | `1213401` |

---

## 🔧 트러블슈팅

| 증상 | 원인 | 해결 |
|---|---|---|
| PPT 오류 메시지 | 템플릿 파일 없음 | `.pptx` 파일을 같은 폴더에 배치 |
| 한글 깨짐 (PDF) | 폰트 경로 없음 | `C:\Windows\Fonts\malgun.ttf` 확인 |
| 라이선스 항목 0개 | txt 파일 형식 불일치 | 위 패턴 예시와 비교 |
| EXE 실행 후 브라우저 미열림 | 방화벽 차단 | `127.0.0.1:8501` 로컬 포트 허용 |
| EXE 실행 후 서버 시작 실패 | 40초 타임아웃 초과 | 백신 예외 처리 또는 재실행 |

---

**태성에스엔이 내부 사용 전용**
