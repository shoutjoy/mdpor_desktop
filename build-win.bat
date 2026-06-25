@echo off
setlocal

cd /d "%~dp0"

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

if not exist "node_modules" (
  echo [MDpro] node_modules not found. Installing dependencies...
  npm ci
  if errorlevel 1 (
    echo [ERROR] npm ci failed.
    pause
    exit /b 1
  )
)

echo [MDpro] Building installer into dist...
npm run build:win
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
