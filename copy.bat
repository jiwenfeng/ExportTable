@echo off
cd %1
cd ..
cd data
svn up
copy *.lua %2 /y
@echo on