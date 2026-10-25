@echo off
chcp 65001 >nul
echo ========================================================
echo  Запуск графического интерфейса iResolver
echo ========================================================
echo.
set JAVA_HOME=C:\jdk11
set PATH=%JAVA_HOME%\bin;%PATH%

cd /d "%~dp0iresolver_work\bin"
call run-ui.bat
