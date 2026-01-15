@echo off
setlocal

set "TARGET_DIR=C:\Program Files\GlobalHotkeys"
set "SRC_PROCESS=%TARGET_DIR%\ProcessList.txt"
set "SRC_SOUND=%TARGET_DIR%\SoundDevices.txt"

for %%C in (Debug Release) do (
  if not exist "GlobalHotkeys\bin\%%C" mkdir "GlobalHotkeys\bin\%%C" >nul 2>nul

  del /f /q "GlobalHotkeys\bin\%%C\ProcessList.txt" >nul 2>nul
  del /f /q "GlobalHotkeys\bin\%%C\SoundDevices.txt" >nul 2>nul

  echo Creating symlinks in GlobalHotkeys\bin\%%C...
  mklink "GlobalHotkeys\bin\%%C\ProcessList.txt" "%SRC_PROCESS%"
  mklink "GlobalHotkeys\bin\%%C\SoundDevices.txt" "%SRC_SOUND%"
  echo.
)

endlocal
