@echo off


FOR /F "delims=" %%I IN ("svn.exe") DO (if exist %%~$PATH:I (CALL :copy %1 %2) else (echo û���ҵ�svn��������°�װsvn����ѡ��װ�����й��� & pause & exit 0))

:copy
cd /d %1
cd ..
cd data
svn up
copy *.lua %2 /y
goto:eof

@echo on