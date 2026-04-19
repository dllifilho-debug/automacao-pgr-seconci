@echo off
:: Garante que roda sempre a partir da pasta PAI do refatoracao\
cd /d "%~dp0.."

echo ============================================
echo  Instalando refatoracao SST - Seconci GO
echo  Pasta de destino: %CD%
echo ============================================

:: Backup do app.py original
if exist app.py (
    copy app.py app.py.bkp >nul
    echo [OK] Backup criado: app.py.bkp
)

:: Criar pastas
for %%d in (config data utils modules) do (
    if not exist %%d (
        mkdir %%d
        echo [OK] Pasta criada: %%d
    )
)

:: Criar __init__.py em cada pasta
for %%d in (config data utils modules) do (
    if not exist %%d\__init__.py (
        type nul > %%d\__init__.py
    )
)

:: Copiar arquivos usando caminho absoluto baseado no local do bat
set "SRC=%~dp0"

copy /Y "%SRC%config\db.py"              config\db.py
echo [OK] config\db.py

copy /Y "%SRC%data\dicionario_cas.py"    data\dicionario_cas.py
echo [OK] data\dicionario_cas.py

copy /Y "%SRC%data\dicionario_campo.py"  data\dicionario_campo.py
echo [OK] data\dicionario_campo.py

copy /Y "%SRC%data\matriz_exames.py"     data\matriz_exames.py
echo [OK] data\matriz_exames.py

copy /Y "%SRC%utils\cas_utils.py"        utils\cas_utils.py
echo [OK] utils\cas_utils.py

copy /Y "%SRC%utils\exame_utils.py"      utils\exame_utils.py
echo [OK] utils\exame_utils.py

copy /Y "%SRC%utils\cargo_utils.py"      utils\cargo_utils.py
echo [OK] utils\cargo_utils.py

copy /Y "%SRC%utils\biologico_utils.py"  utils\biologico_utils.py
echo [OK] utils\biologico_utils.py

copy /Y "%SRC%utils\ia_client.py"        utils\ia_client.py
echo [OK] utils\ia_client.py

copy /Y "%SRC%modules\modulo_pcmso.py"   modules\modulo_pcmso.py
echo [OK] modules\modulo_pcmso.py

copy /Y "%SRC%app.py"                    app.py
echo [OK] app.py (novo)

echo.
echo ============================================
echo  INSTALACAO CONCLUIDA COM SUCESSO!
echo.
echo  Proximos passos:
echo    1. Verifique se modules\modulo_engenharia.py existe
echo    2. git add .
echo    3. git commit -m "refactor: separacao UI e logica"
echo    4. git push
echo ============================================
pause
