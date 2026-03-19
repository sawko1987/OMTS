@echo off
chcp 65001 >nul
echo Проверка установки Python...

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ОШИБКА: Python не найден в системе!
    echo Пожалуйста, установите Python с https://www.python.org/downloads/
    echo При установке обязательно отметьте 'Add Python to PATH'
    pause
    exit /b 1
)

python --version
echo.
echo Создание виртуального окружения...
python -m venv venv

if %errorlevel% neq 0 (
    echo ОШИБКА: Не удалось создать виртуальное окружение!
    pause
    exit /b 1
)

echo Виртуальное окружение создано успешно!
echo.
echo Обновление pip...
venv\Scripts\python.exe -m pip install --upgrade pip

echo.
echo Установка зависимостей из requirements.txt...
venv\Scripts\python.exe -m pip install -r requirements.txt

if %errorlevel% equ 0 (
    echo.
    echo ✓ Все зависимости успешно установлены!
    echo.
    echo Для активации виртуального окружения выполните:
    echo   venv\Scripts\activate.bat
) else (
    echo.
    echo ОШИБКА: Не удалось установить зависимости!
    pause
    exit /b 1
)

pause



