@echo off
chcp 65001 >nul
echo ========================================================
echo    🔨 بناء نسخة ويندوز التنفيذية (.exe) BostaOrderManager
echo ========================================================

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [خطأ] بايثون غير مثبت. يرجى تثبيته أولاً.
    pause
    exit /b
)

if not exist ".venv" (
    python -m venv .venv
)

call .venv\Scripts\activate.bat
pip install -r requirements.txt
pip install pyinstaller

echo جاري بناء ملف exe...
pyinstaller --name "BostaOrderManager" --windowed --noconfirm --add-data "bosta_locations.json;." --add-data "Bosta_Template_Clean.xlsx;." --add-data "Bosta Upload Template AR V3.5.xlsx;." app.py

echo نسخ القوالب إلى مجلد dist\BostaOrderManager...
copy "Bosta_Template_Clean.xlsx" "dist\BostaOrderManager\"
copy "Bosta Upload Template AR V3.5.xlsx" "dist\BostaOrderManager\"
copy "bosta_locations.json" "dist\BostaOrderManager\"

echo ========================================================
echo  ✅ تم الانتهاء بنجاح! ستجد البرنامج في مجلد dist\BostaOrderManager
echo ========================================================
pause
