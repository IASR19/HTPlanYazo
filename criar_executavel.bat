@echo off
chcp 65001 >nul
echo ========================================
echo    PlanYazo - Criando Executável
echo ========================================
echo.

echo [1/4] Verificando dependências...
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Instalando PyInstaller...
    pip install pyinstaller
) else (
    echo PyInstaller ja esta instalado.
)

echo.
echo [2/4] Verificando arquivos necessarios...
if not exist "Front.py" (
    echo ERRO: Front.py nao encontrado!
    pause
    exit /b 1
)

if exist "ht.ico" (
    echo Icone personalizado encontrado: ht.ico
) else (
    echo AVISO: ht.ico nao encontrado - o executavel sera gerado sem icone personalizado
)

if not exist "Plan.xlsx" (
    echo AVISO: Plan.xlsx nao encontrado - sera criado um executavel sem este arquivo
)

if not exist "Yazo.xlsx" (
    echo AVISO: Yazo.xlsx nao encontrado - sera criado um executavel sem este arquivo
)

echo.
echo [3/4] Criando executavel...
python build_exe.py

echo.
echo [4/4] Verificando resultado...
if exist "PlanYazo_Executavel\PlanYazo.exe" (
    echo.
    echo ========================================
    echo    ✅ EXECUTAVEL CRIADO COM SUCESSO!
    echo ========================================
    echo.
    echo 📁 Localizacao: PlanYazo_Executavel\
    echo 📄 Arquivos incluidos:
    echo    - PlanYazo.exe
    if exist "Plan.xlsx" echo    - Plan.xlsx
    if exist "Yazo.xlsx" echo    - Yazo.xlsx
    echo    - README_Executavel.txt
    if exist "ht.ico" echo    - ht.ico (usado como icone)
    echo.
    echo 🚀 Para usar: De duplo clique em PlanYazo.exe
    echo.
) else (
    echo.
    echo ========================================
    echo    ❌ ERRO AO CRIAR EXECUTAVEL!
    echo ========================================
    echo.
    echo Verifique se:
    echo - Python esta instalado corretamente
    echo - PyInstaller foi instalado
    echo - Front.py existe na pasta atual
    echo.
)

pause 