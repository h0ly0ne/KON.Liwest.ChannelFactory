@echo off
chcp 65001 >nul
setlocal

:: Prevent recursion
if "%PUBLISHING%"=="true" exit /b
if "%~1"=="StopLoop" exit /b

echo ==================================================
echo PROJECT PUBLISH
echo ==================================================

:: Mark that publishing has started
set PUBLISHING=true

:: Paths (works for any project)
set PROJ_FILE="KON.Liwest.ChannelFactory.csproj"
set PROJ_DIR=%~dp0
set OUTPUT_DIR=%PROJ_DIR%\bin\Debug\net10.0\win-x64\publish\
set MSBUILD_PATH="C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\amd64\msbuild.exe"

:: Create publish folder if it doesn’t exist
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

:: Clean previous publish
if exist "%OUTPUT_DIR%*" del /q "%OUTPUT_DIR%*" 2>nul

echo Output: %OUTPUT_DIR%
echo.

:: Publish
%MSBUILD_PATH% %PROJ_FILE% /t:Publish /p:Configuration=Debug /p:PublishProfile=FolderProfileDebug

:: Check result
if %errorlevel% neq 0 (
	echo ERROR! Code: %errorlevel%
	pause
	exit /b %errorlevel%
)