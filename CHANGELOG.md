# CHANGELOG - ExcelSlimmer

## 2025-11-15

### 초기 통합 및 파이프라인
- ExcelCleaner, ExcelImageOptimization, ExcelByteReduce 3개 도구를 통합하는 파이프라인 런처 `excel_suite_pipeline.py` 추가.
- 단일 UI에서 다음 단계를 선택적으로 수행하도록 구현:
  - 1) 이름 정의 정리 (ExcelCleaner 로직 호출)
  - 2) 이미지 최적화 (Image Slimmer 로직 호출)
  - 3) 정밀 슬리머 (Precision Plus 로직 호출)
- 파이프라인 완료 후 최종 결과 파일 1개만 남기고, 중간 산출물(.clean.xlsx, _slim.xlsx 등)은 자동 삭제하도록 처리.

### ExcelCleaner (이름 정의 정리)
- Desktop 아래 최상위 출력 폴더명을 `Excel이름관리자정리완료` → `ExcelSlimmed` 로 변경.
- 이전 구조(백업/정리본 서브폴더)를 제거하고, 타임스탬프 폴더(YYYY-MM-DD-HH-MM-SS) 안에 다음 두 파일만 생성하도록 변경:
  - `<원본이름>_backup.xlsx`
  - `<원본이름>_clean.xlsx`

### 정밀 슬리머 (Precision Plus) 안전 모드
- `excel_slimmer_precision_plus.py` 수정:
  - XML 정리 옵션이 **켜져 있을 때만** 다음 작업 수행:
    - calcChain.xml 제거
    - xl/printerSettings/*.bin 제거
    - docProps/thumbnail.jpeg 제거
    - docProps/custom.xml 제거
  - XML 정리 옵션이 **꺼져 있을 때는** 이미지 관련 최적화만 수행하고, 구조 관련 XML은 변경하지 않도록 조정.
- 숨은 XML 삭제(customXml) 옵션은 여전히 별도 고급 옵션으로 유지.
- 백업 파일 이름을 `.backup`/`.bak` 혼용에서 **`<원본이름>_backup.xlsx`** 한 가지 형식으로 통일.

### 파이프라인 UI/UX 개선
- 실행할 기능:
  - 정밀 슬리머 체크박스 아래에 경고 문구 추가:
    - `주의: 정밀 슬리머 사용 시 엑셀에서 복구 여부를 물어볼 수 있습니다.`
- 정밀 슬리머 옵션:
  - `숨은 XML 데이터 삭제 (customXml, 주의)` 아래에 경고 문구 추가:
    - `권장: 일반적인 경우 사용하지 마세요`
  - 정밀 슬리머 메인 체크박스가 꺼져 있을 때는 옵션 체크박스들을 비활성화.
- 체크박스/테마 정리:
  - Windows 느낌의 체크 스타일(vista 테마 시도) 적용.
  - 기본 상태에서 이름 정의 정리, 이미지 최적화만 ON, 정밀 슬리머 및 옵션은 OFF.
- 파이프라인 완료 후 동작:
  - 완료 메시지 박스 표시 후, 탐색기에서 최종 결과 파일 위치 열기.
  - 로그 텍스트는 유지하고, 대상 파일 경로/체크박스/진행률/상태 텍스트는 초기 상태로 리셋.

### 로그 및 임시 파일 처리
- 이미지 최적화 단계에서 생성되는 런타임 로그 파일(`*_image_slim.log`)은 파이프라인이 정상 완료되면 자동 삭제.
- 정밀/이미지 단계에서 생성된 중간 결과 파일들은 파이프라인이 성공적으로 끝난 경우에만 삭제하고, 오류 시에는 남겨두어 디버깅에 활용 가능하도록 유지.

### 빌드/실행 스크립트
- `install.bat`:
  - `.venv_suite` 가상환경 생성 및 pillow, lxml, pyinstaller 설치.
  - 배치 if/goto 구조 수정으로 Windows 10/11에서 안정적으로 동작하도록 정리.
- `run.bat`:
  - `.venv_suite` 활성화 후 `excel_suite_pipeline.py` 실행.
- `build.bat`:
  - PyInstaller를 사용해 단일 EXE 빌드하도록 구성.
  - EXE 이름을 `ExcelSlimmer.exe` 로 변경.
  - 아이콘 우선순위:
    - 현재 폴더의 `ExcelSlimmer.ico`가 존재하면 사용.
    - 없으면 `..\ExcelCleaner\icon.ico` 를 사용.

