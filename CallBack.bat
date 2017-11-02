echo off

set extName=".exe"
set fileName=%~n0
set sysdate=%DATE

%DATE 2017/09/01

cd %~dp0
start %fileName%%exName%

ping localhost -n 3 > nul:

DATE %sysdate%

rem pause
