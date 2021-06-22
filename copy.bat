@echo off


FOR /F "delims=" %%I IN ("svn.exe") DO (if exist %%~$PATH:I (CALL :copy %1 %2) else (echo 没有找到svn命令，请重新安装svn，并选择安装命令行工具 & pause & exit 0))

:copy
cd /d %1
cd ..
cd data
svn up
copy *.lua %2 /y
goto:eof

@echo on