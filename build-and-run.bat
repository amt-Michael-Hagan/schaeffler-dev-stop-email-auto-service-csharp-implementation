@echo off
echo 🔨 Building C# Email Automation Service...
echo.

REM Check if MSBuild is available
where msbuild >nul 2>nul
if %errorlevel% neq 0 (
    echo ❌ MSBuild not found. Please run this from a Visual Studio Command Prompt.
    echo    Or install Visual Studio / Build Tools.
    pause
    exit /b 1
)

REM Build the project
echo Building project...
msbuild EmailAutomationLegacy.sln /p:Configuration=Release /p:Platform="Any CPU" /verbosity:minimal

if %errorlevel% neq 0 (
    echo.
    echo ❌ Build failed! Check the output above for errors.
    pause
    exit /b 1
)

echo.
echo ✅ Build successful!
echo.

REM Check if executable exists
if not exist "bin\Release\EmailAutomationLegacy.exe" (
    echo ❌ Executable not found at bin\Release\EmailAutomationLegacy.exe
    pause
    exit /b 1
)

echo 🏃‍♂️ Running Email Automation Service...
echo.
cd bin\Release
EmailAutomationLegacy.exe
cd ..\..

echo.
echo 🎉 Service execution completed!
pause