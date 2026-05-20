@echo off
setlocal
cd /d "%~dp0"

set "NODE_EXE=node"
where node >nul 2>nul
if errorlevel 1 if exist "%LOCALAPPDATA%\OpenAI\Codex\bin\node.exe" set "NODE_EXE=%LOCALAPPDATA%\OpenAI\Codex\bin\node.exe"

"%NODE_EXE%" --version >nul 2>nul
if errorlevel 1 (
  echo Node.js was not found in PATH.
  echo Please install Node.js or run this from an environment that includes node.
  pause
  exit /b 1
)

echo Starting offline preview at http://127.0.0.1:4174/
echo Keep this window open while previewing. Press Ctrl+C to stop.
start "" "http://127.0.0.1:4174/"
"%NODE_EXE%" scripts\preview-server.js 4174
