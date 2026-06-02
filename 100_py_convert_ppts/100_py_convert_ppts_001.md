원하시는 대로 가상환경 설정부터 라이브러리 설치, PPT 일괄 변환 코드 실행까지 한 번에 처리하는 자동화 스크립트를 만들어 드리겠습니다.

Windows 기준으로 .bat 배치 파일 하나로 모든 과정을 자동화할 수 있습니다. (PowerShell보다 실행 정책 문제가 없어 편리합니다.)

---

📦 전체 자동화 배치 스크립트 (run_ppt_convert.bat)

아래 내용을 메모장에 복사하여 run_ppt_convert.bat (또는 원하는 이름.bat)로 저장하세요.

```batch
@echo off
chcp 65001 >nul
title PPT to Image Converter (AIP 지원)

echo ===================================================
echo   PPT 파일 → 이미지 일괄 변환 (AIP 보호 파일 지원)
echo ===================================================
echo.

REM ------------------------- 설정 -------------------------
REM 원본 PPT 폴더 경로 (여기를 직접 수정하거나, 실행 시 물어볼 수 있음)
set "SOURCE_DIR=C:\Users\%USERNAME%\Documents\ppt_원본"

REM 출력 루트 폴더 (비워두면 SOURCE_DIR + "_images" 자동 생성)
set "OUTPUT_DIR="

REM 이미지 포맷 (PNG 또는 JPG)
set "IMG_FORMAT=PNG"

REM 가상환경 이름
set "VENV_DIR=ppt_converter_venv"

REM --------------------------------------------------------

echo [1] 현재 폴더에 가상환경이 없으면 새로 생성합니다.
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo 가상환경 생성 중...
    python -m venv %VENV_DIR%
    if errorlevel 1 (
        echo 오류: Python이 설치되어 있지 않거나 PATH에 없습니다.
        pause
        exit /b 1
    )
    echo 완료.
) else (
    echo 이미 가상환경이 존재합니다.
)

echo.
echo [2] 가상환경 활성화 및 필요 라이브러리 설치
call "%VENV_DIR%\Scripts\activate.bat"
python -m pip install --upgrade pip >nul
pip install pywin32
if errorlevel 1 (
    echo 라이브러리 설치 실패.
    pause
    exit /b 1
)

echo.
echo [3] PPT 변환 Python 코드 실행
python -c ^
"import os, sys, subprocess, win32com.client; ^
from pathlib import Path; ^
 ^
def convert_all_ppts_in_folder(source_root, output_root=None, image_format='PNG'): ^
    source_root = Path(source_root).resolve(); ^
    if not source_root.is_dir(): ^
        print(f'오류: 폴더 없음 - {source_root}'); ^
        return; ^
    if output_root is None: ^
        output_root = source_root.parent / (source_root.name + '_images'); ^
    else: ^
        output_root = Path(output_root).resolve(); ^
    print(f'입력 폴더: {source_root}'); ^
    print(f'출력 루트: {output_root}'); ^
    ppt_files = list(source_root.rglob('*.ppt')) + list(source_root.rglob('*.pptx')); ^
    if not ppt_files: ^
        print('⚠️ PPT 파일 없음'); ^
        return; ^
    print(f'✅ 총 {len(ppt_files)}개 파일 발견'); ^
    powerpoint = win32com.client.Dispatch('PowerPoint.Application'); ^
    powerpoint.Visible = False; ^
    success = 0; ^
    for ppt_path in ppt_files: ^
        rel_path = ppt_path.parent.relative_to(source_root); ^
        file_stem = ppt_path.stem; ^
        out_dir = output_root / rel_path / file_stem; ^
        out_dir.mkdir(parents=True, exist_ok=True); ^
        pres = None; ^
        try: ^
            print(f'\n🔄 처리: {ppt_path.relative_to(source_root)}'); ^
            pres = powerpoint.Presentations.Open(str(ppt_path)); ^
            for i, slide in enumerate(pres.Slides, start=1): ^
                img_path = out_dir / f'{i}.{image_format.lower()}'; ^
                slide.Export(str(img_path), image_format); ^
            print(f'   ✅ 저장됨 ({len(pres.Slides)}개)'); ^
            success += 1; ^
        except Exception as e: ^
            print(f'   ❌ 실패: {e}'); ^
        finally: ^
            if pres: pres.Close(); ^
    powerpoint.Quit(); ^
    print(f'\n🎉 완료! 성공: {success} / 전체: {len(ppt_files)}'); ^
 ^
convert_all_ppts_in_folder(r'%SOURCE_DIR%', r'%OUTPUT_DIR%' if '%OUTPUT_DIR%'!='' else None, '%IMG_FORMAT%'); ^
echo. && echo 모든 작업이 끝났습니다. && pause"

REM 가상환경 비활성화 (선택)
call deactivate 2>nul

pause
```

---

📖 사용 방법 (아주 간단)

1. 배치 파일 저장
      위 코드를 메모장에 붙여넣고 run_ppt_convert.bat로 저장합니다. (인코딩은 ANSI 또는 UTF-8 with BOM 권장)
2. 원본 폴더 경로 수정 (선택)
      배치 파일 안의 set "SOURCE_DIR=..." 부분을 자신의 PPT 파일이 있는 최상위 폴더 경로로 바꿉니다.
      예: set "SOURCE_DIR=D:\업무자료\2025\프레젠테이션"
3. 배치 파일 실행
      run_ppt_convert.bat를 더블클릭하면:
   · 자동으로 가상환경이 생성됨 (최초 1회)
   · 가상환경 활성화
   · pywin32 설치
   · 모든 하위 폴더의 PPT/PPTX 파일을 찾아 이미지로 변환
   · 출력 폴더는 원본 폴더명에 _images 붙여서 생성됨

---

✅ 자동 생성되는 폴더 구조 예시

```
원본 폴더: D:\업무자료\ppt_원본
   ├── 영업팀\제안서.pptx
   └── 기획팀\계획안.ppt

출력 폴더: D:\업무자료\ppt_원본_images
   ├── 영업팀\제안서\
   │     ├── 1.png
   │     ├── 2.png
   │     └── 3.png
   └── 기획팀\계획안\
         ├── 1.png
         └── 2.png
```

---

⚠️ 주의사항

· PowerPoint 설치 필수 – Microsoft PowerPoint가 로컬 PC에 설치되어 있어야 합니다.
· AIP 권한 – 해당 PPT 파일을 직접 PowerPoint로 열 수 있는 권한이 있어야 변환 가능합니다.
· Python 환경 – 시스템에 Python 3.7 이상이 설치되어 있고 python 명령어가 PATH에 등록되어 있어야 합니다.
· 이미지 품질 – PNG는 고품질/투명배경, JPG는 파일 크기 작음 (배경 흰색).
· 보안 프로그램 – 일부 회사 보안 정책상 win32com이 차단될 수 있습니다. 그럴 경우 관리자에게 문의하세요.

---

🔧 만약 경로를 실행할 때마다 입력하고 싶다면?

배치 파일에서 set "SOURCE_DIR=..." 부분을 다음으로 교체하면 됩니다:

```batch
set /p "SOURCE_DIR=원본 PPT 폴더 경로를 입력하세요: "
```

이렇게 하면 실행할 때마다 폴더 경로를 직접 입력할 수 있습니다.

---

필요한 부분이 있으면 추가로 알려주세요. 특정 환경(예: macOS, 리눅스)에서 실행해야 한다면 그에 맞는 셸 스크립트도 제공 가능합니다.
