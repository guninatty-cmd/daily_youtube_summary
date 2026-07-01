@echo off
REM 이 파일을 daily_youtube_summary 저장소 최상위 폴더(scripts 폴더와 같은 위치)에 두고 더블클릭하세요.
REM 처음 한 번만: 이 폴더에 .env.example을 복사해 .env로 만들고 값을 채워두세요.

cd /d "%~dp0"

echo [1/2] 의존성 설치 확인 중...
pip install -r requirements.txt --quiet

echo [2/2] 유튜브 영상 + 자막 수집 시작...
python scripts\collect_videos.py

echo.
echo 완료되었습니다. 아무 키나 누르면 창이 닫힙니다.
pause >nul
