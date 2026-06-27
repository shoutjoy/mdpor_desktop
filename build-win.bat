@echo off
setlocal

cd /d "%~dp0"

set "APP_ICON=assets\icon.png"

echo.
echo [MDpro] Windows build started.
echo Project: %CD%
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Node.js is not installed or not available in PATH.
  echo Install Node.js LTS, then run this file again.
  pause
  exit /b 1
)

where npm >nul 2>nul
if errorlevel 1 (
  echo [ERROR] npm is not available in PATH.
  pause
  exit /b 1
)

if not exist "%APP_ICON%" (
  echo [ERROR] App icon not found: %CD%\%APP_ICON%
  pause
  exit /b 1
)

echo [MDpro] App icon: %CD%\%APP_ICON%

if not exist "node_modules" (
  echo [MDpro] node_modules not found. Installing dependencies...
  call npm ci
  if errorlevel 1 (
    echo [ERROR] npm ci failed.
    pause
    exit /b 1
  )
)

rem Force electron-builder to regenerate its ICO conversion when icon.png changes.
if exist "dist\.icon-ico" (
  echo [MDpro] Clearing generated icon cache...
  rmdir /s /q "dist\.icon-ico"
)

echo [MDpro] Building installer into dist with the MDpro icon...
call npm run build:win -- --config.win.icon="%APP_ICON%"
if errorlevel 1 (
  echo [ERROR] Build failed.
  pause
  exit /b 1
)

echo.
echo [MDpro] Build complete.
echo Output folder: %CD%\dist
echo.
dir /b "dist"
echo.
pause
endlocal
