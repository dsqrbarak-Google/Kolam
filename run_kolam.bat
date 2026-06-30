rem @echo off
chcp 65001 > nul
echo ===================================================
echo "קולם" (Kolam) - הפעלת האפליקציה לגיל השלישי
echo ===================================================
echo.

:: Check if config.json exists. If not, create it
if not exist config.json (
    echo [מידע] קובץ ההגדרות config.json לא נמצא.
    echo [מידע] יוצר קובץ הגדרות חדש...
    echo {> config.json
    echo   "gemini_api_key": "YOUR_GEMINI_API_KEY",>> config.json
    echo   "document_path": "C:/Kolam/my_biography.docx",>> config.json
    echo   "backup_audio_dir": "C:/Kolam/backup_audio",>> config.json
    echo   "user_name": "אבא",>> config.json
    echo   "version": "0.1.14">> config.json
    echo }>> config.json
    echo [חשוב] אנא פתח את הקובץ config.json והזן את מפתח ה-API של Gemini שלך תחת השדה gemini_api_key.
    echo [חשוב] לאחר מכן, הפעל קובץ זה מחדש.
    pause
    exit /b
)

:: Check if venv exists
if not exist .venv (
    echo [מידע] יוצר סביבה וירטואלית עבור פייתון...
    py -m venv .venv
)

:: Install dependencies using uv if available, otherwise fallback to pip
where uv >nul 2>nul
if %errorlevel% equ 0 (
    echo [מידע] נמצא מנהל החבילות uv. מתקין חבילות במהירות...
    uv pip install --python .venv\Scripts\python.exe -r requirements.txt
) else (
    echo [מידע] מתקין חבילות באמצעות pip (זה עשוי לקחת מספר רגעים)...
    .venv\Scripts\python.exe -m pip install -r requirements.txt
)

echo [מידע] מפעיל את השרת המקומי...
start "" http://localhost:8000
.venv\Scripts\python.exe main.py
pause
