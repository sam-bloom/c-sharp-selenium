for %%a in ("%~dp0..\") do set "PATH_ONE_LEVELS_UP=%%~fa"
set projectFolder=%PATH_ONE_LEVELS_UP%
set path=%path%;C:\Windows\Microsoft.NET\Framework\v4.0.30319
"C:\Program Files (x86)\MSBuild\14.0\Bin\MSBuild.exe" "%projectFolder%\TestAutomation.sln" /p:Configuration=Debug /t:Rebuild