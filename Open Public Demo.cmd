@echo off
setlocal
where dotnet >nul 2>nul
if errorlevel 1 (
    echo Install the .NET 10 SDK to build and open this public source demo.
    pause
    exit /b 1
)
set "DEMO_BUILD_ROOT=%LOCALAPPDATA:\=/%/PropertyManagerDemo-Build"
echo Building the full Property Manager Demo application...
dotnet build "%~dp0PropertyManager.App\PropertyManager.App.csproj" --configuration Release -m:1 "-p:OutputPath=%DEMO_BUILD_ROOT%/bin/" "-p:BaseIntermediateOutputPath=%DEMO_BUILD_ROOT%/obj/" --verbosity quiet
if errorlevel 1 (
    echo The public demo could not be built. See the errors above.
    pause
    exit /b 1
)
start "" "%DEMO_BUILD_ROOT%/bin/PropertyManager.App.exe" %*
endlocal
