dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "innoSetup.iss"

pause