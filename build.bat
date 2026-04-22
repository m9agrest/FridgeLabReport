dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
if errorlevel 1 exit /b 1

"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "innoSetup.iss"
if errorlevel 1 exit /b 1

pause