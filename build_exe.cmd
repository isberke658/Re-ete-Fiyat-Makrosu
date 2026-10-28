@echo off
chcp 65001 >nul
echo ========================================================
echo Reçete Maliyeti Makrosu - Standalone EXE Derleyici
echo ========================================================
python -m pip install -r requirements.txt
pyinstaller --noconfirm Recete_Fiyat_Kopyalama_Makro.spec
echo.
echo Derleme tamamlandı! dist\Recete_Fiyat_Kopyalama_Makro.exe hazır.
pause
