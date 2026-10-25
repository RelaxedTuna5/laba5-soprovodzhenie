@echo off
chcp 65001 >nul
echo ========================================================
echo  Генерация отчета iResolver (Apache Kafka)
echo ========================================================
echo.
set JAVA_HOME=C:\jdk11
set PATH=%JAVA_HOME%\bin;%PATH%

cd /d "%~dp0iresolver_work\bin"
call run.bat

echo.
echo ========================================================
echo  Готово! Открываем отчет в браузере...
echo ========================================================
start "" "%~dp0iresolver_work\output\report.html"
