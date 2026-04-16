# 캡처툴

알캡처 스타일의 Windows 화면 캡처 · 편집 도구입니다.

## 주요 기능

### 캡처
| 모드 | 단축키 | 설명 |
|------|--------|------|
| 사각형 캡처 | `F1` | 마우스로 영역 직접 지정 |
| 단위영역 캡처 | `F2` | 창/요소 자동 감지 후 클릭 |
| 전체화면 캡처 | `F3` | 모니터 전체 캡처 |
| 스크롤 캡처 | `F4` | 스크롤 내리며 긴 화면 자동 합성 |
| 지정크기 캡처 | `F5` | 너비·높이 직접 입력 |

### 편집기
- **펜 / 형광펜** — 자유 드로잉
- **텍스트** — 다크 테마 입력창, 마우스로 위치 이동 가능
- **사각형 / 타원** — 솔리드·점선, 채우기 토글
- **화살표 / 선** — 드래그 방향대로 생성
- **모자이크** — 범위 지정 블러
- **크롭** — 영역 잘라내기
- **지우개** — 원본 복원
- **실행취소 / 다시실행** — `Ctrl+Z` / `Ctrl+Y`
- **저장** — `Ctrl+S` (PNG)
- **클립보드 복사** — `Ctrl+C`

### 기타
- **OCR (텍스트 추출)** — Windows WinRT OCR 엔진 사용 (Windows 10+)
- **캡처 히스토리** — 세션 종료 후에도 기록 유지 (`%APPDATA%\CaptureApp\history`)
- **화면 녹화** — 사각형·단위영역·전체화면 녹화 (MP4)
- **작업표시줄 상주** — X 버튼 클릭 시 최소화, 우클릭으로 종료

## 실행 환경

- Windows 10 이상
- Python 3.10+

## 설치 및 실행

```bash
pip install pillow mss pywin32 numpy
python capture_tool.py
```

## 빌드 (단일 exe)

```bash
pip install pyinstaller
python -m PyInstaller CaptureApp.spec
# dist/캡처툴.exe 생성됨
```

## 배포 시 참고

- **Visual C++ 재배포 패키지** — 대부분의 PC에 이미 설치돼 있음 ([다운로드](https://aka.ms/vs/17/release/vc_redist.x64.exe))
- **OCR 기능** — Windows 10 이상에서만 동작
- **맑은 고딕 폰트** — 한국어 Windows 기본 탑재; 영문 Windows는 폰트가 다를 수 있음
