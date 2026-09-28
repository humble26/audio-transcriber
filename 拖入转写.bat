@echo off
rem ============================================================
rem  Audio transcription launcher (drag & drop target).
rem  Drop audio files or a folder onto the desktop shortcut
rem  "Audio Transcribe" to run a batch transcription.
rem
rem  NOTE: keep this file ASCII-only. cmd.exe decodes .bat files
rem  using the console code page, so non-ASCII bytes here can break
rem  parsing. All Chinese UI text is printed by transcribe_hq.py.
rem
rem  Optional: set ASR_MODEL=small  to use the faster model.
rem ============================================================
setlocal
set "PROJ=%~dp0"
set "FALLBACK=%APPDATA%\TRAE SOLO CN\ModularData\ai-agent\vm\tools\python\python.exe"
if not defined ASR_MODEL set "ASR_MODEL=large-v3"

set "PY="
if exist "%FALLBACK%" (
  "%FALLBACK%" -c "import faster_whisper" >nul 2>&1 && set "PY=%FALLBACK%"
)
if defined PY goto run
python -c "import faster_whisper" >nul 2>&1 && set "PY=python"
if defined PY goto run

echo.
echo   [ERROR] No Python environment with faster-whisper was found.
echo   Install it first:  pip install faster-whisper
echo.
pause
exit /b 1

:run
"%PY%" "%PROJ%transcribe_hq.py" -m %ASR_MODEL% %*
echo.
pause
exit /b 0