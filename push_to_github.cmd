@echo off
chcp 65001 >nul
echo ========================================================
echo  GitHub'a Otomatik Yukleme Sihirbazi
echo ========================================================
echo.
set /p REPO_URL="GitHub Repo URL'nizi yapistirin (Orn: https://github.com/kullaniciadi/proje.git): "
if "%REPO_URL%"=="" goto end

git remote remove origin 2>nul
git remote add origin %REPO_URL%
git branch -M main
echo.
echo GitHub'a yukleniyor (Tarayici acilirsa 'Authorize' butonuna basmaniz yeterlidir)...
git push -u origin main

echo.
echo ========================================================
echo  Tebrikler! Projeniz GitHub'a basariyla yuklendi!
echo ========================================================
:end
pause
