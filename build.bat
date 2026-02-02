@echo off
chcp 65001 >nul
cls
echo.
echo ╔══════════════════════════════════════════════════╗
echo ║   სამედიცინო დოკუმენტაცია - EXE Builder          ║
echo ╚══════════════════════════════════════════════════╝
echo.

REM Python-ის შემოწმება
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python არ არის დაყენებული!
    pause
    exit /b 1
)

echo [1/6] ძველი build ფაილების წაშლა...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec 2>nul

echo [2/6] ვირტუალური გარემოს შექმნა...
if not exist "venv" (
    python -m venv venv
)

echo [3/6] ვირტუალური გარემოს აქტივაცია...
call venv\Scripts\activate.bat

echo [4/6] დამოკიდებულებების დაყენება...
python -m pip install --upgrade pip
pip install flask==2.3.3
pip install python-docx==0.8.11
pip install werkzeug==2.3.7
pip install lxml
pip install pyinstaller

echo [5/6] EXE ფაილის აგება...
pyinstaller --noconfirm ^
    --onedir ^
    --console ^
    --name "MedicalDocs" ^
    --add-data "templates;templates" ^
    --add-data "static;static" ^
    --hidden-import=flask ^
    --hidden-import=flask.json ^
    --hidden-import=jinja2 ^
    --hidden-import=werkzeug ^
    --hidden-import=werkzeug.routing ^
    --hidden-import=docx ^
    --hidden-import=docx.document ^
    --hidden-import=docx.oxml ^
    --hidden-import=docx.oxml.ns ^
    --hidden-import=docx.shared ^
    --hidden-import=lxml ^
    --hidden-import=lxml._elementpath ^
    --hidden-import=lxml.etree ^
    --collect-all=docx ^
    --collect-all=flask ^
    app.py

echo [6/6] საქაღალდეების შექმნა...
if not exist "dist\MedicalDocs\documents" mkdir "dist\MedicalDocs\documents"
if not exist "dist\MedicalDocs\saved_templates" mkdir "dist\MedicalDocs\saved_templates"

REM გაშვების სკრიპტის შექმნა
echo @echo off > "dist\MedicalDocs\Start.bat"
echo cd /d "%%~dp0" >> "dist\MedicalDocs\Start.bat"
echo start "" "MedicalDocs.exe" >> "dist\MedicalDocs\Start.bat"

echo.
echo ╔══════════════════════════════════════════════════╗
echo ║   ✅ EXE წარმატებით შეიქმნა!                     ║
echo ║   📁 იხილეთ: dist\MedicalDocs\MedicalDocs.exe    ║
echo ╚══════════════════════════════════════════════════╝
echo.
pause