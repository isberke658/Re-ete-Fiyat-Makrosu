@echo off
chcp 65001 >nul
echo ========================================================
echo Tiger & Logo ERP Otomasyon Araclari - Toplu EXE Derleyici
echo ========================================================
python -m pip install -r requirements.txt

echo.
echo [1/3] Recete Fiyat Kopyalama Makrosu Derleniyor...
pyinstaller --noconfirm Recete_Fiyat_Kopyalama_Makro.spec

echo.
echo [2/3] Logo Recete Makrosu Derleniyor...
pyinstaller --noconfirm --onefile --windowed --name "Logo_Recete_Makro" logo_recete_makro.py

echo.
echo [3/3] Netupp Recete Aktarim Derleniyor...
pyinstaller --noconfirm --onefile --windowed --name "Netupp_Recete_Aktarim" netupp_recete_aktarim.py

echo.
echo Tum Araclar Basariyla Derlendi! dist\ klasorunden ulasabilirsiniz.
pause
