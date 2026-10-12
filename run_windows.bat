@echo off
chcp 65001 >nul
echo ========================================================
echo    🚀 تشغيل مدير أوردرات بوسطة الذكي (Windows)
echo ========================================================

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [خطأ] بايثون غير مثبت على هذا الجهاز. يرجى تثبيت Python 3 من python.org
    pause
    exit /b
)

if not exist ".venv" (
    echo [1/3] جاري إنشاء البيئة الافتراضية...
    python -m venv .venv
)

echo [2/3] جاري التأكد من تثبيت المكتبات المطلوبة...
call .venv\Scripts\activate.bat
pip install -r requirements.txt --quiet

echo [3/3] جاري تشغيل البرنامج...
python app.py
pause
