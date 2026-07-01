@echo off
REM Put this file in the daily_youtube_summary folder (same level as the scripts folder).
REM First time only: copy .env.example to .env and fill in the values.

cd /d "%~dp0"

echo [1/2] Installing dependencies...
pip install -r requirements.txt --quiet

echo [2/2] Collecting videos and transcripts...
python scripts\collect_videos.py

echo.
echo Done. Press any key to close this window.
pause >nul
