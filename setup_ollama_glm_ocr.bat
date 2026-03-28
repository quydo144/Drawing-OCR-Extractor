@echo off
setlocal EnableExtensions

set "MODEL=glm-ocr:latest"
set "OLLAMA_CMD="

echo =======================================
echo Kiem tra cai dat Ollama va model
echo Model yeu cau: %MODEL%
echo =======================================
echo.

call :resolve_ollama
if not defined OLLAMA_CMD (
  echo [INFO] Khong tim thay Ollama. Dang cai dat bang winget...
  winget install -e --id Ollama.Ollama --accept-package-agreements --accept-source-agreements
  if errorlevel 1 (
    echo [FAIL] Cai dat Ollama that bai.
    goto :exit_1
  )

  call :resolve_ollama
  if not defined OLLAMA_CMD (
    echo [FAIL] Da cai Ollama nhung chua tim thay lenh ollama.
    echo [GOI Y] Dong cua so CMD nay, mo terminal moi roi chay lai script.
    goto :exit_1
  )
)

echo [OK] Tim thay Ollama: %OLLAMA_CMD%

for /f "delims=" %%V in ('"%OLLAMA_CMD%" --version 2^>nul') do (
  echo [INFO] %%V
  goto :version_done
)
:version_done
echo.

set "TMP_FILE=%TEMP%\ollama_list_%RANDOM%_%RANDOM%.txt"
"%OLLAMA_CMD%" list > "%TMP_FILE%" 2>nul
if errorlevel 1 (
  echo [FAIL] Khong the lay danh sach model tu Ollama.
  if exist "%TMP_FILE%" del /q "%TMP_FILE%" >nul 2>nul
  goto :exit_3
)

findstr /I /B /C:"%MODEL% " "%TMP_FILE%" >nul
if errorlevel 1 (
  findstr /I /X /C:"%MODEL%" "%TMP_FILE%" >nul
  if errorlevel 1 (
    echo [INFO] Chua tim thay model %MODEL%. Dang cai dat model...
    "%OLLAMA_CMD%" pull %MODEL%
    if errorlevel 1 (
      echo [FAIL] Cai dat model %MODEL% that bai.
      if exist "%TMP_FILE%" del /q "%TMP_FILE%" >nul 2>nul
      goto :exit_2
    )

    "%OLLAMA_CMD%" list > "%TMP_FILE%" 2>nul
    findstr /I /B /C:"%MODEL% " "%TMP_FILE%" >nul
    if errorlevel 1 (
      findstr /I /X /C:"%MODEL%" "%TMP_FILE%" >nul
      if errorlevel 1 (
        echo [FAIL] Da pull nhung van khong tim thay model %MODEL%.
        if exist "%TMP_FILE%" del /q "%TMP_FILE%" >nul 2>nul
        goto :exit_2
      )
    )
  )
)

echo [OK] Da tim thay model %MODEL%.
echo [SUCCESS] Moi thu da san sang.

if exist "%TMP_FILE%" del /q "%TMP_FILE%" >nul 2>nul
goto :exit_0

:exit_0
echo.
echo Nhan phim bat ky de dong cua so...
pause >nul
exit /b 0

:exit_1
echo.
echo Nhan phim bat ky de dong cua so...
pause >nul
exit /b 1

:exit_2
echo.
echo Nhan phim bat ky de dong cua so...
pause >nul
exit /b 2

:exit_3
echo.
echo Nhan phim bat ky de dong cua so...
pause >nul
exit /b 3

:resolve_ollama
set "OLLAMA_CMD="
where ollama >nul 2>nul
if not errorlevel 1 (
  set "OLLAMA_CMD=ollama"
  goto :eof
)

if exist "%LOCALAPPDATA%\Programs\Ollama\ollama.exe" (
  set "OLLAMA_CMD=%LOCALAPPDATA%\Programs\Ollama\ollama.exe"
  goto :eof
)

if exist "%ProgramFiles%\Ollama\ollama.exe" (
  set "OLLAMA_CMD=%ProgramFiles%\Ollama\ollama.exe"
  goto :eof
)

goto :eof
