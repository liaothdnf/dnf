@echo off
C:
taskkill /f /im Client.exe
taskkill /f /im aaa.exe
taskkill /f /im DNF.exe
taskkill /f /im DNFChina.exe
if exist z:\D\DNF\script (@echo ��ʼ) else (exit)
Z:
cd z:\D\DNF\script
for /r %%i in (*.*) do copy /y "%%i" "C:\Program Files\��������2014\QMScript"
rem ��ʵ�����������������������ã�/r����ʵ����Ѱ�������ļ���
cd ../update
set vScriptPath=C:\Program Files\��������2014
for /r %%i in (*.*) do copy /y "%%i" "%vScriptPath%"
for /r %%i in (*.exe) do (set scriptPath=%%i)
set scriptName=%scriptPath:~-13%
rem set vScriptPath=%vScriptPath%\%scriptName%
rem echo %vScriptPath%
c:
cd %vScriptPath%
start %scriptName%
exit



