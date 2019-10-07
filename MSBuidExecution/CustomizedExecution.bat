set startTime=%time%
for %%a in ("%~dp0..\") do set "PATH_ONE_LEVELS_UP=%%~fa"
set projectBinFolder=%PATH_ONE_LEVELS_UP%
echo Project location is: %projectBinFolder%
%projectBinFolder%packages\NUnit.ConsoleRunner.3.8.0\tools\nunit3-console.exe %projectBinFolder%TestAutomation\bin\Debug\TestAutomation.dll --where "cat=" --labels=All